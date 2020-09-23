VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "smtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Public Function sendCreationNoticeTXT(acciRecID As Long) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "sendCreationNoticeTXT"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim qut As String
    qut = Chr$(34)
    
    Dim rsload As adodb.Recordset
    Dim HTML As Variant
    Dim br As Boolean
      
    HTML = HTML + "Customer Generation Notice"
    HTML = HTML + ""
    br = MySQL.OpenTable(directConn, rsload, , "select *, sysops.Username from accountinfo, sysops WHERE accountinfo.SysopID = sysops.Recid and accountinfo.RecID = " & acciRecID)
    
        HTML = HTML + "Account Name: "
        HTML = HTML + rsload!AccountName
      HTML = HTML + vbCrLf
      HTML = HTML + "DOB: "
        HTML = HTML + Format(rsload!DOB, "dd-mm-yyyy Hh:Nn:Ss")
      HTML = HTML + ""
        HTML = HTML + "Domain: "
        HTML = HTML + rsload!Realm
      HTML = HTML + vbCrLf
      HTML = HTML + "VISP: "
      Dim rsVISP As adodb.Recordset
        If rsload!VirtualID = 1000 Then
            HTML = HTML + "Number One Computer Services"
        Else
            br = MySQL.OpenTable(directConn, rsVISP, , "select * from virtualisp Where RecID = " & rsload!VirtualID)
            HTML = HTML + rsVISP!Description
        End If
      HTML = HTML + vbCrLf
      
        HTML = HTML + "Sysop: "
        HTML = HTML + rsload!Username
      HTML = HTML + vbCrLf
    HTML = HTML + vbCrLf
    HTML = HTML + "..::[ Addresses ]::.."
    HTML = HTML + vbCrLf
      HTML = HTML + vbCrLf
      br = MySQL.OpenTable(directConn, rsload, , "select * from acci_addresses Where AccI_RecID = " & acciRecID)
      
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + ""
              HTML = HTML + vbCrLf + rsload!ContactName + ", "
              HTML = HTML + IIf(IsNull(rsload!Street1), "", rsload!Street1) + ", "
              HTML = HTML + IIf(IsNull(rsload!Street2), "", rsload!Street2) + ", "
              HTML = HTML + IIf(IsNull(rsload!Suburb), "", rsload!Suburb) + ", "
              HTML = HTML + IIf(IsNull(rsload!PostCode), "", rsload!PostCode) + ", "
              HTML = HTML + IIf(IsNull(rsload!State), "", rsload!State) + ", "
              HTML = HTML + IIf(IsNull(rsload!Country), "", rsload!Country) + ""
            rsload.MoveNext
        Wend
      End If
    HTML = HTML + vbCrLf
    HTML = HTML + "..::[ Phone Numbers ]::.."
    HTML = HTML + vbCrLf
      
      br = MySQL.OpenTable(directConn, rsload, , "select * from acci_phonenumbers Where AccI_RecID = " & acciRecID)
      
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + vbCrLf
              HTML = HTML + IIf(IsNull(rsload!ContactName), "", rsload!ContactName) + ", "
              HTML = HTML + IIf(IsNull(rsload!PhoneNumber), "", rsload!PhoneNumber) + ", "
              HTML = HTML + IIf(IsNull(rsload!Extension), "", rsload!Extension) + ", "
              HTML = HTML + IIf(IsNull(rsload!ShortNote), "", rsload!ShortNote) + ", "
            rsload.MoveNext
        Wend
      End If
      
    HTML = HTML + "..::[ Email Addresses ]::.."
     HTML = HTML + vbCrLf
      
      br = MySQL.OpenTable(directConn, rsload, , "select * from acci_emailaddresses Where AccI_RecID = " & acciRecID)
      
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + vbCrLf
              HTML = HTML + IIf(IsNull(rsload!ContactName), "", rsload!ContactName) + ", "
              HTML = HTML + IIf(IsNull(rsload!EmailAddress), "", rsload!EmailAddress) + ", "
            rsload.MoveNext
        Wend
      End If
    HTML = HTML + vbCrLf
    HTML = HTML + "..::[ Services ]::.."
    HTML = HTML + vbCrLf
      bResult = MySQL.OpenTable(directConn, rsload, , "select acci_services.ServiceID, acci_services.RecID, acci_services.Checked, acci_services.ContactName, acci_services.Username, AES_DECRYPT(acci_services.Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as decPassword, acci_services.BaseURL, acci_services.DynamicField1, acci_services.DynamicField2, acci_services.DynamicField3, acci_services.DynamicField4, acci_services.DynamicField5, plantypes.Description, acci_services.ptRecID, plantypes.TemplateID from acci_services, plantypes Where plantypes.RecID = acci_services.ptRecID AND AccI_RecID = " & acciRecID)
      
      Dim rsload2 As adodb.Recordset
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + vbCrLf
              HTML = HTML + IIf(IsNull(rsload!ContactName), "", rsload!ContactName) + ", "
              HTML = HTML + IIf(IsNull(rsload!Description), "", rsload!Description) + ", "
              
              If MySQL.OpenTable(directConn, rsload2, , "select Description from plantemplates where RecID = " & rsload!TemplateID) = True Then
                HTML = HTML + IIf(IsNull(rsload2!Description), "", rsload2!Description) & ", "
              End If
              HTML = HTML + IIf(IsNull(rsload!Username), "", rsload!Username) + ", "
              HTML = HTML + IIf(IsNull(rsload!DecPassword), "", rsload!DecPassword) + ", "
              HTML = HTML + IIf(IsNull(rsload!DynamicField1), "", rsload!DynamicField1) + ", "
              HTML = HTML + IIf(IsNull(rsload!DynamicField2), "", rsload!DynamicField2) + ", "
              HTML = HTML + IIf(IsNull(rsload!DynamicField3), "", rsload!DynamicField3) + ", "
              HTML = HTML + IIf(IsNull(rsload!DynamicField4), "", rsload!DynamicField4) + ", "
              HTML = HTML + IIf(IsNull(rsload!DynamicField5), "", rsload!DynamicField5) + ", "
            rsload.MoveNext
        Wend
    End If
    HTML = HTML + vbCrLf

    sendCreationNoticeTXT = HTML
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function sendCreationNotice(acciRecID As Long) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "sendCreationNotice"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim qut As String
    qut = Chr$(34)
    
    Dim rsload As adodb.Recordset
    Dim HTML As Variant
    Dim br As Boolean
      
    HTML = HTML + "<html>"
    HTML = HTML + ""
    HTML = HTML + "<head>"
    HTML = HTML + "<meta http-equiv=" + qut + "Content-Language" + qut + "content=" + qut + "en-us" + qut + ">"
    HTML = HTML + "<meta http-equiv=" + qut + "Content-Type" + qut + " content=" + qut + "text/html; charset=windows-1252" + qut + ">"
    HTML = HTML + "<meta name=" + qut + "GENERATOR" + qut + " content=" + qut + "Microsoft FrontPage 4.0" + qut + ">"
    HTML = HTML + "<meta name=" + qut + "ProgId" + qut + " content=" + qut + "FrontPage.Editor.Document" + qut + ">"
    HTML = HTML + "<title>Customer Generation Notice</title>"
    HTML = HTML + "</head>"
    HTML = HTML + ""
    HTML = HTML + "<body bgcolor=" + qut + "#000000" + qut + " text=" + qut + "#FFFFFF" + qut + ">"
    HTML = HTML + ""
    HTML = HTML + "<table border=" + qut + "0" + qut + " width=" + qut + "100%" + qut + " cellspacing=" + qut + "0" + qut + " cellpadding=" + qut + "0" + qut + ">"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "19%" + qut + "><img border=" + qut + "0" + qut + " src=" + qut + "http://www.ep.net.au/The Nexus/images/tree.gif" + qut + " width=" + qut + "170" + qut + " height=" + qut + "123" + qut + "></td>"
        HTML = HTML + "<td width=" + qut + "39%" + qut + " valign=" + qut + "top" + qut + " align=" + qut + "right" + qut + "><img border=" + qut + "0" + qut + " src=" + qut + "http://www.ep.net.au/The Nexus/images/title.gif" + qut + " align=" + qut + "right" + qut + " width=" + qut + "226" + qut + " height=" + qut + "27" + qut + "><br>"
          HTML = HTML + "<br>"
          HTML = HTML + "<img border=" + qut + "0" + qut + " src=" + qut + "http://www.ep.net.au/The Nexus/images/ACN.gif" + qut + " width=" + qut + "134" + qut + " height=" + qut + "15" + qut + ">"
          HTML = HTML + "<p>&nbsp;</td>"
        HTML = HTML + "<td width=" + qut + "42%" + qut + ">"
          HTML = HTML + "<p align=" + qut + "center" + qut + "><b><font size=" + qut + "5" + qut + " face=" + qut + "Courier New" + qut + ">" + Format(sysnow, "dd-mm-yyyy Hh:Nn:Ss") + "</font></b></td>"
      HTML = HTML + "</tr>"
    HTML = HTML + "</table>"
    HTML = HTML + "<p align=" + qut + "center" + qut + "><b><u><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "6" + qut + "><sub>Customer"
    HTML = HTML + "Creation</sub></font></u></b></p>"
    HTML = HTML + "<hr>"
    HTML = HTML + "<p align=" + qut + "left" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>Account"
    
    br = MySQL.OpenTable(directConn, rsload, , "select *, sysops.Username from accountinfo, sysops WHERE accountinfo.SysopID = sysops.Recid and accountinfo.RecID = " & acciRecID)
    
        HTML = HTML + "<table border=" + qut + "0" + qut + " width=" + qut + "100%" + qut + " cellspacing=" + qut + "0" + qut + " cellpadding=" + qut + "0" + qut + ">"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "20%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>Account Name:&nbsp;</sub></b></font></td>"
        HTML = HTML + "<td width=" + qut + "80%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>" + rsload!AccountName + "</sub></b></font></td>"
      HTML = HTML + "</tr>"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "20%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>DOB:</sub></b></font></td>"
        HTML = HTML + "<td width=" + qut + "80%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>" + Format(rsload!DOB, "dd-mm-yyyy Hh:Nn:Ss") + "</sub></b></font></td>"
      HTML = HTML + "</tr>"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "20%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>Domain:&nbsp;</sub></b></font></td>"
        HTML = HTML + "<td width=" + qut + "80%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>" + rsload!Realm + "</sub></b></font></td>"
      HTML = HTML + "</tr>"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "20%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>VISP:</sub></b></font></td>"
            Dim rsVISP As adodb.Recordset
        If rsload!VirtualID = 1000 Then
            HTML = HTML + "<td width=" + qut + "80%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>" + "Number One Computer Services" + "</sub></b></font></td>"
        Else
            br = MySQL.OpenTable(directConn, rsVISP, , "select * from virtualisp Where RecID = " & rsload!VirtualID)
            HTML = HTML + "<td width=" + qut + "80%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>" + rsVISP!Description + "</sub></b></font></td>"
        End If
        
      HTML = HTML + "</tr>"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "20%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>Sysop:</sub></b></font></td>"
        HTML = HTML + "<td width=" + qut + "80%" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><sub>" + rsload!Username + "</sub></b></font></td>"
      HTML = HTML + "</tr>"
    HTML = HTML + "</table>"
    HTML = HTML + "<hr>"
    HTML = HTML + "<p align=" + qut + "left" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><u><sub>Addresses:</sub></u></b></font></p>"
    HTML = HTML + "<table border=" + qut + "1" + qut + " width=" + qut + "100%" + qut + " cellspacing=" + qut + "1" + qut + ">"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">"
          HTML = HTML + "<p align=" + qut + "center" + qut + ">Contact Name</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Street Addreess 1</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Street Address 2</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Suburb</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Postcode</td>"
        HTML = HTML + "<td width=" + qut + "15%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">State</td>"
        HTML = HTML + "<td width=" + qut + "15%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Country</td>"
      HTML = HTML + "</tr>"
      
      br = MySQL.OpenTable(directConn, rsload, , "select * from acci_addresses Where AccI_RecID = " & acciRecID)
      
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + "<tr>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!ContactName), "&nbsp;", rsload!ContactName) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!Street1), "&nbsp;", rsload!Street1) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!Street2), "&nbsp;", rsload!Street2) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!Suburb), "&nbsp;", rsload!Suburb) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!PostCode), "&nbsp;", rsload!PostCode) + "</td>"
              HTML = HTML + "<td width=" + qut + "15%" + qut + ">" + IIf(IsNull(rsload!State), "&nbsp;", rsload!State) + "</td>"
              HTML = HTML + "<td width=" + qut + "15%" + qut + ">" + IIf(IsNull(rsload!Country), "&nbsp;", rsload!Country) + "</td>"
            HTML = HTML + "</tr>"
            rsload.MoveNext
        Wend
      End If
    HTML = HTML + "</table>"
    HTML = HTML + "<p align=" + qut + "left" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><u><sub>Phone Numbers:</sub></u></b></font></p>"
    HTML = HTML + "<table border=" + qut + "1" + qut + " width=" + qut + "100%" + qut + " cellspacing=" + qut + "1" + qut + ">"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">"
          HTML = HTML + "<p align=" + qut + "center" + qut + ">Contact Name</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Phone Number</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Extension</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Position/Comment</td>"
      HTML = HTML + "</tr>"
      
      br = MySQL.OpenTable(directConn, rsload, , "select * from acci_phonenumbers Where AccI_RecID = " & acciRecID)
      
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + "<tr>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!ContactName), "&nbsp;", rsload!ContactName) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!PhoneNumber), "&nbsp;", rsload!PhoneNumber) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!Extension), "&nbsp;", rsload!Extension) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!ShortNote), "&nbsp;", rsload!ShortNote) + "</td>"
            HTML = HTML + "</tr>"
            rsload.MoveNext
        Wend
      End If
      HTML = HTML + "</table>"
    HTML = HTML + "<p align=" + qut + "left" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><u><sub>Email Addresses:</sub></u></b></font></p>"
    HTML = HTML + "<table border=" + qut + "1" + qut + " width=" + qut + "100%" + qut + " cellspacing=" + qut + "1" + qut + ">"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">"
          HTML = HTML + "<p align=" + qut + "center" + qut + ">Contact Name</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Email Address</td>"
     HTML = HTML + " </tr>"
      
      br = MySQL.OpenTable(directConn, rsload, , "select * from acci_emailaddresses Where AccI_RecID = " & acciRecID)
      
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + "<tr>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!ContactName), "&nbsp;", rsload!ContactName) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!EmailAddress), "&nbsp;", rsload!EmailAddress) + "</td>"
            HTML = HTML + "</tr>"
            rsload.MoveNext
        Wend
      End If
    HTML = HTML + "</table>"
    HTML = HTML + "<p align=" + qut + "left" + qut + "><font face=" + qut + "Trebuchet MS" + qut + " size=" + qut + "1" + qut + "><b><u><sub>Services:</sub></u></b></font></p>"
    HTML = HTML + "<table border=" + qut + "1" + qut + " width=" + qut + "100%" + qut + " cellspacing=" + qut + "1" + qut + ">"
      HTML = HTML + "<tr>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">"
          HTML = HTML + "<p align=" + qut + "center" + qut + ">Contact Name</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Description</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Template ID</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Username</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Password</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Field1</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Field2</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Field3</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Field4</td>"
        HTML = HTML + "<td width=" + qut + "14%" + qut + " bgcolor=" + qut + "#808000" + qut + " align=" + qut + "center" + qut + ">Field5</td>"
      HTML = HTML + "</tr>"
      bResult = MySQL.OpenTable(directConn, rsload, , "select acci_services.ServiceID, acci_services.RecID, acci_services.Checked, acci_services.ContactName, acci_services.Username, AES_DECRYPT(acci_services.Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as decPassword, acci_services.BaseURL, acci_services.DynamicField1, acci_services.DynamicField2, acci_services.DynamicField3, acci_services.DynamicField4, acci_services.DynamicField5, plantypes.Description, acci_services.ptRecID, plantypes.TemplateID from acci_services, plantypes Where plantypes.RecID = acci_services.ptRecID AND AccI_RecID = " & acciRecID)
      
      Dim rsload2 As adodb.Recordset
      If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            HTML = HTML + "<tr>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!ContactName), "&nbsp;", rsload!ContactName) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!Description), "&nbsp;", rsload!Description) + "</td>"
              
              
                HTML = HTML + "<td width=" + qut + "14%" + qut + ">" & IIf(IsNull(rsload!TemplateID), "&nbsp;", rsload!TemplateID) & "</td>"
              
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!Username), "&nbsp;", rsload!Username) + "</td>"
              HTML = HTML + "<td width=" + qut + "14%" + qut + ">" + IIf(IsNull(rsload!DecPassword), "&nbsp;", rsload!DecPassword) + "</td>"
              HTML = HTML + "<td width=" + qut + "15%" + qut + ">" + IIf(IsNull(rsload!DynamicField1), "&nbsp;", rsload!DynamicField1) + "</td>"
              HTML = HTML + "<td width=" + qut + "15%" + qut + ">" + IIf(IsNull(rsload!DynamicField2), "&nbsp;", rsload!DynamicField2) + "</td>"
              HTML = HTML + "<td width=" + qut + "15%" + qut + ">" + IIf(IsNull(rsload!DynamicField3), "&nbsp;", rsload!DynamicField3) + "</td>"
              HTML = HTML + "<td width=" + qut + "15%" + qut + ">" + IIf(IsNull(rsload!DynamicField4), "&nbsp;", rsload!DynamicField4) + "</td>"
              HTML = HTML + "<td width=" + qut + "15%" + qut + ">" + IIf(IsNull(rsload!DynamicField5), "&nbsp;", rsload!DynamicField5) + "</td>"
            
            HTML = HTML + "</tr>"
            rsload.MoveNext
        Wend
    End If
    HTML = HTML + "</table>"
    HTML = HTML + "</body>"
    HTML = HTML + "</html>"

    sendCreationNotice = HTML
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function


Public Function sendInvoiceHTML(oConn As adodb.Connection, acciRecID As Long, iTraxrID As Variant, Optional InvoiceFlag As Integer) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "sendInvoiceHTML"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim HTML As String
    Dim xHTML As Variant
    
    HTML = ""
    
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(oConn, rsload, , "select HTML from stationary where StationaryCode = 'INVOICE'")
    
    If rsload.RecordCount > 0 Then
    
        xHTML = rsload!HTML
        
    Else
        
        Exit Function
        
    End If

    Dim jVirtualID As Long
    
        Dim rsVISP As adodb.Recordset
        If MySQL.OpenTable(oConn, rsVISP, , "select virtualisp.VirtualID, virtualisp.Realm, accountinfo.SysopID, virtualisp.ACN, virtualisp.ABN, virtualisp.RecID, virtualisp.Description, virtualisp.LogoURL from virtualisp, accountinfo where virtualisp.RecID = accountinfo.VirtualID and accountinfo.RecID = " & acciRecID & " Limit 1") = True Then
              If rsVISP.RecordCount > 0 Then
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPRealm/", IIf(IsNull(rsVISP!Realm), "ep.net.au", rsVISP!Realm))
                  If IsNull(rsVISP!ACN) Then
                        If IsNull(rsVISP!ABN) Then
                            xHTML = MySQL.ReplaceString(xHTML, "/VISPRBN/", "No RBN on file")
                        Else
                            xHTML = MySQL.ReplaceString(xHTML, "/VISPRBN/", "ABN " & rsVISP!ABN)
                        End If
                    Else
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPRBN/", "ACN " & rsVISP!ACN)
                  End If
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPNAME/", IIf(IsNull(rsVISP!Description), "::", rsVISP!Description))
                  
                  jVirtualID = rsVISP!RecID
                  xHTML = MySQL.ReplaceString(xHTML, "/VirtualID/", "" & jVirtualID)
                    
                  If IsNull(rsVISP!LogoURL) Or rsVISP!LogoURL = "" Then
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", rsVISP!Description)
                  Else
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(IsNull(rsVISP!LogoURL), "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + "><BR>" & rsVISP!Description)
                  End If
                  
                  
                  Dim rsVispb As adodb.Recordset
                  Call MySQL.OpenTable(oConn, rsVispb, , "Select Description from virtualisp where RecID = '" & rsVISP!VirtualID & "'")
                  If rsVispb.State = adStateOpen Then
                    If rsVispb.RecordCount > 0 Then
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPSponsor/", rsVispb!Description)
                    Else
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPSponsor/", "Exitstencil Press Pty Ltd")
                    End If
                    rsVispb.Close
                  Else
                    xHTML = MySQL.ReplaceString(xHTML, "/VISPSponsor/", "Exitstencil Press Pty Ltd")
                  End If
              Else
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPNAME/", "Exitstencil Press Pty Ltd")
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPSponsor/", "Exitstencil Press Pty Ltd")
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPRBN/", "No RBN on file")
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPRealm/", "ep.net.au")
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(rsVISP.RecordCount = 0, "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + ">")
              End If
        End If
        
         If MySQL.OpenTable(oConn, rsload, , "select accountinfo.AccountName, MD5(AES_DECRYPT(accountinfo.gPassword,'" & odb.colSalts.ReturnSalt(PWSalt) & "')) as gPassword, sysops.Firstname, sysops.Surname, sysops.SecurityLevel, sysops.Email from accountinfo inner join sysops on accountinfo.SysopID = sysops.RecID where accountinfo.RecID = " & acciRecID) = True Then
          If rsload.RecordCount > 0 Then
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", IIf(IsNull(rsload!AccountName), "(null)", rsload!AccountName))
              xHTML = MySQL.ReplaceString(xHTML, "/USERMD5PASS/", IIf(IsNull(rsload!gPassword), "AHAHAHAHAHHA", rsload!gPassword))
              xHTML = MySQL.ReplaceString(xHTML, "/ConfSysop/", IIf(IsNull(rsload!eMail), "simon@projectalpha.com.au", rsload!eMail))
              xHTML = MySQL.ReplaceString(xHTML, "/SEC_SYSOP/", IIf(IsNull(rsload!SecurityLevel), "[0%]", "[" & rsload!SecurityLevel & "%]"))
              xHTML = MySQL.ReplaceString(xHTML, "/VISPSysop/", IIf(IsNull(rsload!Firstname), "", rsload!Firstname) + " " + IIf(IsNull(rsload!Surname), "", rsload!Surname))
            Else
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", ":: NO TITLE FOR ACCOUNT ESTABLISHED ::")
              xHTML = MySQL.ReplaceString(xHTML, "/USERMD5PASS/", "::PASS::")
              xHTML = MySQL.ReplaceString(xHTML, "/ConfSysop/", "simon@projectalpha.com.au")
              xHTML = MySQL.ReplaceString(xHTML, "/SEC_SYSOP/", "[0%]")
              xHTML = MySQL.ReplaceString(xHTML, "/VISPSysop/", "7 of 9 the borg")
          End If
        End If
        xHTML = MySQL.ReplaceString(xHTML, "/NOW/", Format(sysnow, "dd/mm/yyyy Hh:Nn:Ss"))
            
            
            
              If MySQL.OpenTable(oConn, rsload, , "select * from acci_addresses where Checked = -1 and AccI_RecID = " & acciRecID & " Limit 1") = True Then
                If rsload.RecordCount > 0 Then
                    HTML = HTML + vbCrLf + rsload!ContactName & "<br>" & rsload!Street1 & "<br>" & rsload!Street2 & "<br>" & rsload!Suburb & "<br>" & rsload!State & ", " & rsload!PostCode & ", " & rsload!Country & "<br><br>"
                
                
                End If
              End If
          xHTML = MySQL.ReplaceString(xHTML, "/Address/", HTML)
          HTML = ""
          
            HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
              HTML = HTML + "<tr>"
                HTML = HTML + "<td width=" + Chr$(34) + "30%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + ">Item Description</td>"
                HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + ">GST</td>"
                HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + ">Debit</td>"
                HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + ">Credited</td>"
              HTML = HTML + "</tr>"
              
              Dim cDebit As Currency
              Dim cPaid As Currency
              Dim cGST As Currency
              Dim cCredit As Currency
                Dim qSQL As String
                
                Select Case InvoiceFlag
                Case 0
                
              
                    If MySQL.OpenTable(oConn, rsload, , "select * from invoiceout where AccI_RecID = " & acciRecID & " and InvoiceFlagID = 0") = True Then
                      If rsload.RecordCount > 0 Then
                          While Not rsload.EOF And Err.Number = 0
                              HTML = HTML + "<tr>"
                                HTML = HTML + "<td width=" + Chr$(34) + "30%" + Chr$(34) + ">" & rsload!Description & "&nbsp; " & IIf(Format(rsload!StartCycle, "yyyy") <> "1899", "from: [" + Format(rsload!StartCycle, "dddd dd/mm/yyyy") + "] To: [" & Format(rsload!EndCycle, "dddd dd/mm/yyyy") + "]", "") & "</td>"
                                HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!GSTCharged, "Currency") & "</td>"
                                cGST = cGST + rsload!GSTCharged
                                HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!TotalDue, "Currency") & "</td>"
                                cDebit = cDebit + rsload!TotalDue
                                cPaid = cPaid + rsload!AmountPaid
                                HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!AmountRefunded + rsload!GSTRefunded, "Currency") & "</td>"
                                cDebit = cDebit - (rsload!AmountRefunded + rsload!GSTRefunded)
                                cCredit = cCredit + (rsload!AmountRefunded + rsload!GSTRefunded)
                              HTML = HTML + "</tr>"
                              MySQL.Execute oConn, "Update invoiceout Set InvoiceFlagID = 1, TraxrID = " & iTraxrID & "  where RecID = " & rsload!RecID
                              MySQL.Execute oConn, "Update invoicetraxr Set TotalDue=TotalDue+" & rsload!TotalDue & ",AmountPaid=AmountPaid+" & rsload!AmountPaid & " where RecID = " & iTraxrID
                              MySQL.Execute oConn, "Update invoicetraxr Set GSTDue = TotalDue * " & oTax(Login.TaxCode, Login.TaxCountry) & " where RecID = " & iTraxrID
                              rsload.MoveNext
                          Wend
                      End If
                    End If
                                    
                Case Else
                
                
                    If MySQL.OpenTable(oConn, rsload, , "select * from invoiceout where AccI_RecID = " & acciRecID & " and TraxrID = '" & iTraxrID & "'") = True Then
                          If rsload.RecordCount > 0 Then
                              While Not rsload.EOF And Err.Number = 0
                                  HTML = HTML + "<tr>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "30%" + Chr$(34) + ">" & rsload!Description & "&nbsp; " & IIf(Format(rsload!StartCycle, "yyyy") <> "1899", "from: [" + Format(rsload!StartCycle, "dddd dd/mm/yyyy") + "] To: [" & Format(rsload!EndCycle, "dddd dd/mm/yyyy") + "]", "") & "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!GSTCharged, "Currency") & "</td>"
                                    cGST = cGST + rsload!GSTCharged
                                    HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!TotalDue, "Currency") & "</td>"
                                    cDebit = cDebit + rsload!TotalDue
                                    cPaid = cPaid + rsload!AmountPaid
                                    HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!AmountRefunded + rsload!GSTRefunded, "Currency") & "</td>"
                                    cDebit = cDebit - (rsload!AmountRefunded + rsload!GSTRefunded)
                                    cCredit = cCredit + (rsload!AmountRefunded + rsload!GSTRefunded)
                                  HTML = HTML + "</tr>"
                                  'MySQL.Execute oConn, "Update invoiceout Set InvoiceFlagID = 1, TraxrID = " & iTraxrID & "  where RecID = " & rsLoad!RecID
                                  'MySQL.Execute oConn, "Update invoicetraxr Set TotalDue=TotalDue+" & rsLoad!TotalDue & ",AmountPaid=AmountPaid+" & rsLoad!AmountPaid & " where RecID = " & iTraxrID
                                  'MySQL.Execute oConn, "Update invoicetraxr Set GSTDue = TotalDue * " & oTax(login.TaxCode , login.TaxCountry ) & " where RecID = " & iTraxrID
                                  rsload.MoveNext
                              Wend
                          End If
                        End If
                
                End Select
                
            HTML = HTML + "<tr>"
              HTML = HTML + "<td width=" + Chr$(34) + "30%" + Chr$(34) + ">Totals</td>"
              HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(cGST, "Currency") & "</td>"
              HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(cDebit, "Currency") & "</td>"
              HTML = HTML + "<td width=" + Chr$(34) + "12%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(cCredit, "Currency") & "</td>"
            HTML = HTML + "</tr>"
              
            HTML = HTML + "</table>"
          
            xHTML = MySQL.ReplaceString(xHTML, "/InvoiceTable/", HTML)
            xHTML = MySQL.ReplaceString(xHTML, "/TotalCredited/", Format(cCredit, "Currency"))
            xHTML = MySQL.ReplaceString(xHTML, "/TotalDebit/", Format(cDebit, "Currency"))
            xHTML = MySQL.ReplaceString(xHTML, "/TotalPaid/", Format(cPaid, "Currency"))
            
            HTML = ""
            
            Dim rsINVTRX As adodb.Recordset
            
            Call MySQL.OpenTable(oConn, rsINVTRX, , "select * from invoicetraxr where RecID = '" & iTraxrID & "'")
            
            Call MySQL.OpenTable(oConn, rsload, , "SELECT invoicetraxr.InvoiceSerial, invoicetraxr.TotalDue, invoicetraxr.AmountPaid, invoicetraxr.PaymentDue, invoicetraxr.AmountCredited FROM invoicetraxr WHERE invoicetraxr.PaymentDue < '" + Format(rsINVTRX!PaymentDue, "yyyy-mm-dd ttttt") + "' and invoicetraxr.TotalDue > (invoicetraxr.AmountPaid + invoicetraxr.AmountCredited) and invoicetraxr.acci_RecID = '" & acciRecID & "'")
            If rsload.State = adStateOpen Then
                     HTML = HTML + "<table width=" + Chr(34) + "87%" + Chr(34) + " border=" + Chr(34) + "0" + Chr(34) + ">"
                    HTML = HTML + "<tr bgcolor=" + Chr(34) + "#000000" + Chr(34) + ">"
                    HTML = HTML + "<td width=" + Chr(34) + "25%" + Chr(34) + "><div align=" + Chr(34) + "center" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + "><strong>Invoice Number</strong></div></td>"
                    HTML = HTML + "<td width=" + Chr(34) + "12%" + Chr(34) + "><div align=" + Chr(34) + "center" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + "><strong>Total Due</strong></div></td>"
                    HTML = HTML + "<td width=" + Chr(34) + "15%" + Chr(34) + "><div align=" + Chr(34) + "center" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + "><strong>Amount Paid </strong></div></td>"
                    HTML = HTML + "<td width=" + Chr(34) + "22%" + Chr(34) + "><div align=" + Chr(34) + "center" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + "><strong>Payment Due </strong></div></td>"
                    HTML = HTML + "<td width=" + Chr(34) + "12%" + Chr(34) + "><div align=" + Chr(34) + "center" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + "><strong>Credited</strong></div></td>"
                    HTML = HTML + "<td width=" + Chr(34) + "14%" + Chr(34) + "><div align=" + Chr(34) + "center" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + "><strong>Status</strong></div></td>"
                    HTML = HTML + "</tr>"
                
                    Dim ttl_TotalDue As Currency
                    Dim ttl_AmountPaid As Currency
                    Dim ttl_Credited As Currency
                    
                While Not rsload.EOF
                
        
                    HTML = HTML + "<tr bgcolor=" + Chr(34) + "#FFCC99" + Chr(34) + " class=" + Chr(34) + "style40" + Chr(34) + ">"
                    HTML = HTML + "<td><span class=" + Chr(34) + "style60" + Chr(34) & rsload!InvoiceSerial & "</span></td>"
                    HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + "><span class=" + Chr(34) + "style60" + Chr(34) + "> " & Format(rsload!TotalDue, "Currency") & "</span></div></td>"
                    HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + "><span class=" + Chr(34) + "style60" + Chr(34) + "> " & Format(rsload!AmountPaid, "Currency") & "</span></div></td>"
                    HTML = HTML + "<td><div align=" + Chr(34) + "center" + Chr(34) + "><span class=" + Chr(34) + "style60" + Chr(34) + "" & Format(rsload!PaymentDue, "dd-mm-yyyy hh:nn") & "</span></div></td>"
                    HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + "><span class=" + Chr(34) + "style60" + Chr(34) + ">" & Format(rsload!AmountCredited, "Currency") & "</span></div></td>"
                    HTML = HTML + "<td><div align=" + Chr(34) + "center" + Chr(34) + " class=" + Chr(34) + "style60" + Chr(34) + ">"
                        If rsload!AmountPaid <= 0 Then
                            HTML = HTML + "Outstanding"
                        ElseIf rsload!AmountPaid < rsload!TotalDue Then
                            HTML = HTML + "Partially Paid " & Round((rsload!AmountPaid / rsload!TotalDue) * 100) & "%"
                        ElseIf rsload!AmountPaid >= rsload!TotalDue Then
                            HTML = HTML + "Finalised"
                        End If
                    ttl_TotalDue = ttl_TotalDue + IIf(IsNull(rsload!TotalDue), 0, rsload!TotalDue)
                    ttl_AmountPaid = ttl_AmountPaid + IIf(IsNull(rsload!AmountPaid), 0, rsload!AmountPaid)
                    ttl_Credited = ttl_Credited + IIf(IsNull(rsload!AmountCredited), 0, rsload!AmountCredited)
                    HTML = HTML + "</tr>"
                    rsload.MoveNext
                Wend
            
            
                HTML = HTML + "<tr bgcolor=" + Chr(34) + "#000000" + Chr(34) + ">"
                  HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + ">"
                    HTML = HTML + "<blockquote>"
                      HTML = HTML + "<p><strong>Totals: </strong></p>"
                    HTML = HTML + "</blockquote>"
                  HTML = HTML + "</div></td>"
                  HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + ">"
                    HTML = HTML + "<div align=" + Chr(34) + "right" + Chr(34) + "><strong>" & Format(ttl_TotalDue, "Currency") & "</strong></div>"
                  HTML = HTML + "</div></td>"
                  HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + ">"
                    HTML = HTML + "<div align=" + Chr(34) + "right" + Chr(34) + "><strong><strong></strong>" & Format(ttl_AmountPaid, "Currency") & "</strong></div>"
                  HTML = HTML + "</div></td>"
                  HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + "><span class=" + Chr(34) + "style44" + Chr(34) + "></span></div></td>"
                  HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + " class=" + Chr(34) + "style44" + Chr(34) + ">"
                    HTML = HTML + "<div align=" + Chr(34) + "right" + Chr(34) + "><strong>" & Format(ttl_Credited, "Currency") & "</strong></div>"
        
                  HTML = HTML + "</div></td>"
                  HTML = HTML + "<td><div align=" + Chr(34) + "right" + Chr(34) + "><span class=" + Chr(34) + "style44" + Chr(34) + "></span></div></td>"
                HTML = HTML + "</tr>"
              HTML = HTML + "</table>"
                
                xHTML = MySQL.ReplaceString(xHTML, "/InvSummaryTable/", HTML)
                xHTML = MySQL.ReplaceString(xHTML, "/TotalOutstanding/", Format(ttl_TotalDue - (ttl_AmountPaid + ttl_Credited), "Currency"))
                
            Else
                xHTML = MySQL.ReplaceString(xHTML, "/TotalOutstanding/", Format(0, "Currency"))
                xHTML = MySQL.ReplaceString(xHTML, "/InvSummaryTable/", "There is no previous invoices in the system")
            End If
                
                xHTML = MySQL.ReplaceString(xHTML, "/GRANDTOTAL/", Format((ttl_TotalDue - (ttl_AmountPaid + ttl_Credited)) + (cDebit - (cPaid + cCredit)), "Currency"))
            
            
            
                                    
                                    
                                     
              
              If MySQL.OpenTable(oConn, rsload, , "select BillingDate, PayInterval, PayIntervalType from accountinfo where RecID = " & acciRecID) = True Then
                If rsload.RecordCount > 0 Then
                    xHTML = MySQL.ReplaceString(xHTML, "/BillingDate/", Format(rsload!BillingDate, "dd-mm-yyyy Hh:Nn:Ss"))
                    xHTML = MySQL.ReplaceString(xHTML, "/DUEWHEN/", Format(DateAdd(rsload!PayIntervalType, rsload!PayInterval, sysnow), "dd-mm-yyyy Hh:Nn:Ss"))
                Else
                    
                    xHTML = MySQL.ReplaceString(xHTML, "/BillingDate/", Format(sysnow, "dd-mm-yyyy Hh:Nn:Ss"))
                    xHTML = MySQL.ReplaceString(xHTML, "/DUEWHEN/", Format(DateAdd("d", 14, sysnow), "dd-mm-yyyy Hh:Nn:Ss"))

                End If
              End If

            Dim sHTML As String
            sHTML = ""
            If MySQL.OpenTable(oConn, rsload, , "select * from visp_phonenumbers where ContactName like '%sale%' and visp_RecID = " & jVirtualID) = True Then
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        sHTML = sHTML + "Phone: " & IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!PhoneNumber), "", rsload!PhoneNumber) & IIf(IsNull(rsload!Extension), "<br>", "Ext." & rsload!Extension & "<br>") & "<br>"
                    End If
                End If
            End If

            If MySQL.OpenTable(oConn, rsload, , "select * from visp_addresses where ContactName like '%sale%' and visp_RecID = " & jVirtualID) = True Then
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        sHTML = sHTML + "Address: " & IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Street1), "", rsload!Street1) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Street2), "", rsload!Street2) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Suburb), "", rsload!Suburb) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Postcost), "", rsload!PostCode) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!State), "", rsload!State) & ", "
                    End If
                End If
            End If
            
            xHTML = MySQL.ReplaceString(xHTML, "/VISPSales/", sHTML)
            sHTML = ""
            
            If MySQL.OpenTable(oConn, rsload, , "select * from visp_phonenumbers where ContactName like '%account%' and visp_RecID = " & jVirtualID) = True Then
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        sHTML = sHTML + "Phone: " & IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!PhoneNumber), "", rsload!PhoneNumber) & IIf(IsNull(rsload!Extension), "<br>", "Ext." & rsload!Extension & "<br>") & "<br>"
                    End If
                End If
            End If

            If MySQL.OpenTable(oConn, rsload, , "select * from visp_addresses where ContactName like '%account%' and visp_RecID = " & jVirtualID) = True Then
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        sHTML = sHTML + "Address: " & IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Street1), "", rsload!Street1) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Street2), "", rsload!Street2) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Suburb), "", rsload!Suburb) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Postcost), "", rsload!PostCode) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!State), "", rsload!State) & ", "
                    End If
                End If
            End If
            
            xHTML = MySQL.ReplaceString(xHTML, "/VISPAccounts/", sHTML)
            sHTML = ""

            If MySQL.OpenTable(oConn, rsload, , "select * from visp_addresses where visp_RecID = " & jVirtualID & " limit 1,1") = True Then
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        sHTML = sHTML + IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Street1), "", rsload!Street1) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Street2), "", rsload!Street2) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Suburb), "", rsload!Suburb) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Postcost), "", rsload!PostCode) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!State), "", rsload!State) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Country), "", rsload!Country) & ". "
                    End If
                End If
            End If
            
            xHTML = MySQL.ReplaceString(xHTML, "/VISPPrimayAddy/", sHTML)
            sHTML = ""
            
            
            If MySQL.OpenTable(oConn, rsload, , "select * from visp_phonenumbers where ContactName like '%support%' and visp_RecID = " & jVirtualID) = True Then
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        sHTML = sHTML + "Phone: " & IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!PhoneNumber), "", rsload!PhoneNumber) & IIf(IsNull(rsload!Extension), "<br>", "Ext." & rsload!Extension & "<br>") & "<br>"
                    End If
                End If
            End If

            If MySQL.OpenTable(oConn, rsload, , "select * from visp_addresses where ContactName like '%support%' and visp_RecID = " & jVirtualID) = True Then
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        sHTML = sHTML + "Address: " & IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Street1), "", rsload!Street1) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Street2), "", rsload!Street2) & "<br>"
                        sHTML = sHTML + IIf(IsNull(rsload!Suburb), "", rsload!Suburb) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!Postcost), "", rsload!PostCode) & ", "
                        sHTML = sHTML + IIf(IsNull(rsload!State), "", rsload!State) & ", "
                    End If
                End If
            End If
            
            xHTML = MySQL.ReplaceString(xHTML, "/VISPSupport/", sHTML)
            sHTML = ""

    sendInvoiceHTML = xHTML

Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function sendReceiptHTML(ByRef oConn As adodb.Connection, ByVal acciRecID As Long, ByRef receiptNo As Double, bMode As Boolean) As String


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "sendReceiptHTML"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim HTML As String
    Dim xHTML As Variant
    
    HTML = ""
    
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(oConn, rsload, , "select HTML from stationary where StationaryCode = 'RECEIPT'")
    
    If rsload.RecordCount > 0 Then
    
        xHTML = rsload!HTML
        
    Else
        
        Exit Function
        
    End If
    
        Dim rsVISP As adodb.Recordset
        If MySQL.OpenTable(oConn, rsVISP, , "select LogoURL, virtualisp.Description from virtualisp, accountinfo where virtualisp.RecID = accountinfo.VirtualID and accountinfo.RecID = " & acciRecID & " Limit 1") = True Then
              If rsVISP.RecordCount > 0 Then
                  If IsNull(rsVISP!LogoURL) Or rsVISP!LogoURL = "" Then
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", rsVISP!Description)
                  Else
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(IsNull(rsVISP!LogoURL), "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + "><BR>" & rsVISP!Description)
                  End If
              Else
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(rsVISP.RecordCount = 0, "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + ">")
              End If
        End If
    
         If MySQL.OpenTable(oConn, rsload, , "select AccountName from accountinfo where RecID = " & acciRecID) = True Then
          If rsload.RecordCount > 0 Then
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", IIf(IsNull(rsload!AccountName), "(null)", rsload!AccountName))
            Else
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", "")
          End If
        End If
                
                        
                Dim rsCheck As adodb.Recordset
                
        If bMode = False Then
                
                If MySQL.OpenTable(oConn, rsCheck, , "select Max(ReceiptNo) from receipts") = True Then
                
                    If rsCheck.RecordCount > 0 Then
                        receiptNo = rsCheck("Max(ReceiptNo)") + 1
                    Else
                        receiptNo = 1
                    End If
                    
                    MySQL.Execute oConn, "Update receipts set ReceiptNo=" & receiptNo & " where ReceiptNo=0 and acci_RecId = " & acciRecID
                End If
                
        End If
                
                xHTML = MySQL.ReplaceString(xHTML, "/ReceiptNumber/", "" & receiptNo)
                
                
              If MySQL.OpenTable(oConn, rsload, , "select * from acci_addresses where Checked = -1 and AccI_RecID = " & acciRecID & " Limit 1") = True Then
                If rsload.RecordCount > 0 Then
                    HTML = HTML + vbCrLf + IIf(IsNull(rsload!ContactName), "", rsload!ContactName & "<br>") & IIf(IsNull(rsload!Street1), "", rsload!Street1 & "<br>") & IIf(IsNull(rsload!Street2), "", rsload!Street2 & "<br>") & IIf(IsNull(rsload!Suburb), "", rsload!Suburb & "<br>") & IIf(IsNull(rsload!State), "", rsload!State & ", ") & IIf(IsNull(rsload!PostCode), "", rsload!PostCode & ", ") & ", " & IIf(IsNull(rsload!Country), "", rsload!Country)
                Else
                    HTML = HTML + "[No Active Address Found in Database]"
                End If
              End If

            xHTML = MySQL.ReplaceString(xHTML, "/Address/", "" & HTML)
            HTML = ""
            
                
                If MySQL.OpenTable(oConn, rsload, , "select * from receipts where TraxrID <> 0 and ReceiptNo = " & receiptNo) Then
                    If rsload.RecordCount > 0 Then
                        HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "77" + Chr$(34) + ">"
                          HTML = HTML + "<tr>"
                            HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Invoice"
                              HTML = HTML + "Number</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Amount"
                              HTML = HTML + "</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Date"
                              HTML = HTML + "</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Money"
                              HTML = HTML + "Order/Cheque No/Payment Type.</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Status</b></font></td>"
                          HTML = HTML + "</tr>"
                                                  
                          While Not rsload.EOF And Err.Number = 0
                                HTML = HTML + "<tr>"
                                
                                  HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Hex(rsload!TraxrID) & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Format(rsload!Paid, "Currency") & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & IIf(IsNull(rsload!SerialNumber), IIf(IsNull(rsload!paymenttype), " ", rsload!paymenttype), rsload!SerialNumber) & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & IIf(IsNull(rsload!Description), rsload!paymenttype, rsload!Description) & "</font>" & "</td>"
                                HTML = HTML + "</tr>"
                                rsload.MoveNext
                          Wend
                        HTML = HTML + "</table>"
                   End If
                End If
                
            xHTML = MySQL.ReplaceString(xHTML, "/InvoicePayment/", "" & HTML)
            HTML = ""
                
                
                If MySQL.OpenTable(oConn, rsload, , "select * from receipts where invoiceinID <> 0 and ReceiptNo = " & receiptNo) Then
                    If rsload.RecordCount > 0 Then
                        
                        HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "77" + Chr$(34) + ">"
                        HTML = HTML + "<tr>"
                            HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Item "
                              HTML = HTML + "Number</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Amount "
                              HTML = HTML + "</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Date "
                              HTML = HTML + "</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Money "
                              HTML = HTML + "Order/Cheque No/Payment Type.</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Status</b></font></td>"
                          HTML = HTML + "</tr>"

                          While Not rsload.EOF And Err.Number = 0
                                HTML = HTML + "<tr>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & rsload!RecID & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Format(IIf(rsload!Paid = 0, rsload!Refunded, rsload!Paid), "Currency") & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & IIf(IsNull(rsload!SerialNumber), IIf(IsNull(rsload!paymenttype), " ", rsload!paymenttype), rsload!SerialNumber) & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & IIf(IsNull(rsload!Description), rsload!paymenttype, rsload!Description) & "</font>" & "</td>"
                                HTML = HTML + "</tr>"
                                rsload.MoveNext
                          Wend
                        HTML = HTML + "</table>"
                   End If
                End If
                
                
            xHTML = MySQL.ReplaceString(xHTML, "/DepositAccount/", "" & HTML)
            HTML = ""
            
                If MySQL.OpenTable(oConn, rsload, , "select receipts.*, invoiceout.Description as invDesc from receipts, invoiceout where receipts.invoiceoutID <> 0 and invoiceout.RecId = receipts.invoiceoutID and receipts.ReceiptNo = " & receiptNo) Then
                    If rsload.RecordCount > 0 Then
                
                        HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "77" + Chr$(34) + ">"
                          HTML = HTML + "<tr>"
                            HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Item "
                              HTML = HTML + "Number</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Amount "
                              HTML = HTML + "</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Date "
                              HTML = HTML + "</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Money "
                              HTML = HTML + "Order/Cheque No./Payment Type</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Item "
                              HTML = HTML + "Description</b></font></td>"
                          HTML = HTML + "</tr>"
                        
                          While Not rsload.EOF And Err.Number = 0
                                HTML = HTML + "<tr>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & rsload!RecID & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Format(rsload!Paid, "Currency") & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & IIf(IsNull(rsload!SerialNumber), IIf(IsNull(rsload!paymenttype), " ", rsload!paymenttype), rsload!SerialNumber) & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & rsload!invDesc & "</font>" & "</td>"
                                HTML = HTML + "</tr>"
                                rsload.MoveNext
                          Wend
                        HTML = HTML + "</table>"
                   End If
                End If
                
            xHTML = MySQL.ReplaceString(xHTML, "/ItemPayments/", "" & HTML)
            HTML = ""
                
                
                
                If MySQL.OpenTable(oConn, rsload, , "select receipts.* from receipts where receipts.RefundID <> 0 and receipts.ReceiptNo = " & receiptNo) Then
                    If rsload.RecordCount > 0 Then
                
                        HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "77" + Chr$(34) + ">"
                          HTML = HTML + "<tr>"
                            HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Item "
                              HTML = HTML + "Number</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Amount "
                              HTML = HTML + "Refunded</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Date "
                              HTML = HTML + "Refunded</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Money "
                              HTML = HTML + "Order/Cheque No./Payment Type</b></font></td>"
                            HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + " height=" + Chr$(34) + "21" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Description</b></font></td>"
                          HTML = HTML + "</tr>"
                        
                          While Not rsload.EOF And Err.Number = 0
                                HTML = HTML + "<tr>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "15%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & rsload!RecID & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & rsload!Refunded & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "24%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "22%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & IIf(IsNull(rsload!SerialNumber), IIf(IsNull(rsload!paymenttype), " ", rsload!paymenttype), rsload!SerialNumber) & "</font>" & "</td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "26%" + Chr$(34) + " height=" + Chr$(34) + "19" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & "<font size=" + Chr$(34) + "1" + Chr$(34) + ">" & IIf(IsNull(rsload!Description), rsload!paymenttype, rsload!Description) & "</font>" & "</td>"
                                HTML = HTML + "</tr>"
                                rsload.MoveNext
                          Wend
                        HTML = HTML + "</table>"
                   End If
                End If
                
            xHTML = MySQL.ReplaceString(xHTML, "/Credits/", "" & HTML)
            HTML = ""
                
              If MySQL.OpenTable(oConn, rsload, , "select BillingDate, PayInterval, PayIntervalType from accountinfo where RecID = " & acciRecID) = True Then
                If rsload.RecordCount > 0 Then
                    xHTML = MySQL.ReplaceString(xHTML, "/BillingDate/", Format(rsload!BillingDate, "dd-mm-yyyy Hh:Nn:Ss"))
                    xHTML = MySQL.ReplaceString(xHTML, "/DUEWHEN/", Format(DateAdd(rsload!PayIntervalType, rsload!PayInterval, sysnow), "dd-mm-yyyy Hh:Nn:Ss"))
                Else
                    
                    xHTML = MySQL.ReplaceString(xHTML, "/BillingDate/", Format(sysnow, "dd-mm-yyyy Hh:Nn:Ss"))
                    xHTML = MySQL.ReplaceString(xHTML, "/DUEWHEN/", Format(DateAdd("d", 14, sysnow), "dd-mm-yyyy Hh:Nn:Ss"))

                End If
              End If

    sendReceiptHTML = xHTML
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function sendStatementHTML(ByRef oConn As adodb.Connection, ByRef acciRecID As Long, ByRef dStatementNo As Double, Optional bMode As Boolean) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "sendStatementHTML"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim HTML As String
    Dim xHTML As Variant
    
    HTML = ""
    
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(oConn, rsload, , "select HTML from stationary where StationaryCode = 'STATEMENT'")
    
    If rsload.RecordCount > 0 Then
    
        xHTML = rsload!HTML
        
    Else
        
        Exit Function
        
    End If
    
    
        Dim rsVISP As adodb.Recordset
        If MySQL.OpenTable(oConn, rsVISP, , "select LogoURL, virtualisp.Description from virtualisp, accountinfo where virtualisp.RecID = accountinfo.VirtualID and accountinfo.RecID = " & acciRecID & " Limit 1") = True Then
              If rsVISP.RecordCount > 0 Then
                  If IsNull(rsVISP!LogoURL) Or rsVISP!LogoURL = "" Then
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", rsVISP!Description)
                  Else
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(IsNull(rsVISP!LogoURL), "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + "><BR>" & rsVISP!Description)
                  End If
              Else
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(rsVISP.RecordCount = 0, "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + ">")
              End If
        End If
        
    
         If MySQL.OpenTable(oConn, rsload, , "select AccountName, BillingDate from accountinfo where RecID = " & acciRecID) = True Then
          If rsload.RecordCount > 0 Then
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", IIf(IsNull(rsload!AccountName), "(null)", rsload!AccountName))
              xHTML = MySQL.ReplaceString(xHTML, "/BillingDate/", Format(rsload!BillingDate, "dd-mm-yyyy Hh:Nn:Ss"))
            Else
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", "")
              xHTML = MySQL.ReplaceString(xHTML, "/BillingDate/", Format(sysnow, "dd-mm-yyyy Hh:Nn:Ss"))
          End If
        End If
        
        xHTML = MySQL.ReplaceString(xHTML, "/NOW/", Format(sysnow, "dd/mm/yyyy Hh:Nn:Ss"))
            
            
            
        If MySQL.OpenTable(oConn, rsload, , "select * from acci_addresses where Checked = -1 and AccI_RecID = " & acciRecID & " Limit 1") = True Then
            If rsload.RecordCount > 0 Then
                HTML = HTML + vbCrLf + rsload!ContactName & "<br>" & rsload!Street1 & "<br>" & rsload!Street2 & "<br>" & rsload!Suburb & "<br>" & rsload!State & ", " & rsload!PostCode & ", " & rsload!Country & "<br><br>"
            End If
        End If
        xHTML = MySQL.ReplaceString(xHTML, "/Address/", HTML)
        HTML = ""
              
              
        Dim rsCheck As adodb.Recordset
        Dim StateID As Double
        
        If bMode = False Then
        
            Call MySQL.OpenTable(oConn, rsCheck, , "select max(CycleEnd) as CycleEnd from statementtraxr where acci_RecID = '" & acciRecID & "'")
            
            If rsCheck.State = adStateOpen Then
                If rsCheck.RecordCount > 0 Then
                    If IsNull(rsCheck!CycleEnd) Then
                        Call MySQL.OpenTable(oConn, rsCheck, , "select CreationDate as CycleEnd from accountinfo where RecID = '" & acciRecID & "'")
                    Else
                    
                    End If
                Else
                
                    Call MySQL.OpenTable(oConn, rsCheck, , "select CreationDate as CycleEnd from accountinfo where RecID = '" & acciRecID & "'")
                
                End If
            End If
        Else
        
            Call MySQL.OpenTable(oConn, rsCheck, , "select CycleEnd from statementtraxr where RecID = '" & dStatementNo & "'")
            
            If rsCheck.State = adStateOpen Then
                If rsCheck.RecordCount > 0 Then
                    If IsNull(rsCheck!CycleEnd) Then
                        Call MySQL.OpenTable(oConn, rsCheck, , "select CreationDate as CycleEnd from accountinfo where RecID = '" & acciRecID & "'")
                    Else
                    
                    End If
                Else
                
                    Call MySQL.OpenTable(oConn, rsCheck, , "select CreationDate as CycleEnd from accountinfo where RecID = '" & acciRecID & "'")
                
                End If
            End If
            
        End If
        On Error Resume Next
        
        If bMode = False Then
            Do
                Err.Clear
                StateID = MySQL.GetTMPRecID("statementtraxr", directConn)
                directConn.Execute "Insert Into statementtraxr (RecID, acci_RecID, CycleStart, CycleEnd) VALUES ('" & StateID & "','" & acciRecID & "','" & Format(DateAdd("s", 1, rsCheck!CycleEnd), "yyyy-mm-dd ttttt") & "','" & Format(sysnow, "yyyy-mm-dd ttttt") & "')"
                gSleep
                gSleep
            Loop Until Err.Number = 0
        
            dStatementNo = StateID
        End If
        
        
        
        xHTML = MySQL.ReplaceString(xHTML, "/ReceiptNumber/", "" & dStatementNo)
        
        Dim ttlInvoiced As Currency
        
        If bMode = False Then
        
                                
                    If MySQL.OpenTable(oConn, rsload, , "select * from invoicetraxr Where acci_RecID = " & acciRecID & " and Created >= '" & Format(DateAdd("m", -1, NextCycle), "yyyy-mm-dd Hh:Nn:Ss") & "' and Created <= '" & Format(NextCycle, "yyyy-mm-dd Hh:Nn:Ss") & "'") Then
                        If rsload.RecordCount > 0 Then
                            
                            HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
                              HTML = HTML + "<tr>"
                                HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Invoice"
                                  HTML = HTML + "Number</font></td>"
                                HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Date</font></td>"
                                HTML = HTML + "<td width=" + Chr$(34) + "34%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Amount</font></td>"
                              HTML = HTML + "</tr>"
                              HTML = HTML + "<tr>"
                            
                            Do
                                HTML = HTML + "<tr>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" + rsload!InvoiceSerial + "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" + Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") + "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "34%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" + Format(rsload!TotalDue, "Currency") + "</td>"
                                    ttlInvoiced = ttlInvoiced + rsload!TotalDue
                                    If bMode = False Then MySQL.Execute oConn, "Update invoicetraxr Set StatementID = " & dStatementNo & " Where RecID = " & rsload!RecID
                                    rsload.MoveNext
                                HTML = HTML + "</tr>"
                            Loop Until rsload.EOF Or Err.Number <> 0
                        
                            
                            HTML = HTML + "</table>"
                        End If
                    End If

                xHTML = MySQL.ReplaceString(xHTML, "/Invoices/", HTML)
                HTML = ""
            Else
            
                xHTML = MySQL.ReplaceString(xHTML, "/StatementNumber/", "" & dStatementNo)
        
                
                    If MySQL.OpenTable(oConn, rsload, , "select * from invoicetraxr Where StatementID = '" & dStatementNo & "'") Then
                        If rsload.RecordCount > 0 Then
                            
                            HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
                              HTML = HTML + "<tr>"
                                HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Invoice"
                                  HTML = HTML + "Number</font></td>"
                                HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Date</font></td>"
                                HTML = HTML + "<td width=" + Chr$(34) + "34%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Amount</font></td>"
                              HTML = HTML + "</tr>"
                              HTML = HTML + "<tr>"
                            
                            Do
                                HTML = HTML + "<tr>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" + rsload!InvoiceSerial + "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" + Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") + "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "34%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" + Format(rsload!TotalDue, "Currency") + "</td>"
                                    ttlInvoiced = ttlInvoiced + rsload!TotalDue
                                    If bMode = False Then MySQL.Execute oConn, "Update invoicetraxr Set StatementID = " & dStatementNo & " Where RecID = " & rsload!RecID
                                    rsload.MoveNext
                                HTML = HTML + "</tr>"
                            Loop Until rsload.EOF Or Err.Number <> 0
                        
                            
                            HTML = HTML + "</table>"
                        End If
                    End If

                xHTML = MySQL.ReplaceString(xHTML, "/Invoices/", HTML)
                HTML = ""
            
            
            
            End If
            
            
                Dim ttlPaid As Currency
                Dim ttlRefund As Currency
                    
                    
            If bMode = False Then
            
                    If MySQL.OpenTable(oConn, rsload, , "select * from receipts Where acci_RecID = " & acciRecID & " and Created >= '" & Format(DateAdd("m", -1, NextCycle), "yyyy-mm-dd Hh:Nn:Ss") & "' and Created <= '" & Format(NextCycle, "yyyy-mm-dd Hh:Nn:Ss") & "'") Then
                        If rsload.RecordCount > 0 Then
              
              
                            HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
                              HTML = HTML + "<tr>"
            
                                  HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Reciept"
                                  HTML = HTML + " Number</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Date</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "17%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Paid</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Refunded</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "77%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Transaction"
                                    HTML = HTML + " Text</font></td>"
                                HTML = HTML + "</tr>"
                                
                                
                                While Not rsload.EOF And Err.Number = 0
                                    HTML = HTML + "<tr>"
                                        HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & rsload!receiptNo & "</td>"
                                        HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") & "</td>"
                                        HTML = HTML + "<td width=" + Chr$(34) + "17%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!Paid, "Currency") & "</td>"
                                        ttlPaid = ttlPaid + rsload!Paid
                                        HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!Refunded, "Currency") & "</td>"
                                        ttlRefund = ttlRefund + rsload!Refunded
                                        HTML = HTML + "<td width=" + Chr$(34) + "77%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & rsload!paymenttype & "</td>"
                                        MySQL.Execute oConn, "Update receipts Set StatementID = " & dStatementNo & " Where RecID = " & rsload!RecID
                                        
                                    rsload.MoveNext
                                    HTML = HTML + "</tr>"
                                Wend
                                
                              
                            HTML = HTML + "</table>"
                            
                        End If
                    End If
                            
                xHTML = MySQL.ReplaceString(xHTML, "/Payments/", HTML)
                HTML = ""
            Else
            
                If MySQL.OpenTable(oConn, rsload, , "select * from receipts Where StatementID = '" & dStatementNo & "'") Then
                        If rsload.RecordCount > 0 Then
              
              
                            HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
                              HTML = HTML + "<tr>"
            
                                  HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Reciept"
                                  HTML = HTML + " Number</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Date</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "17%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Paid</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Refunded</font></td>"
                                  HTML = HTML + "<td width=" + Chr$(34) + "77%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Transaction"
                                    HTML = HTML + " Text</font></td>"
                                HTML = HTML + "</tr>"
                                
                                
                                While Not rsload.EOF And Err.Number = 0
                                    HTML = HTML + "<tr>"
                                        HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & rsload!receiptNo & "</td>"
                                        HTML = HTML + "<td width=" + Chr$(34) + "46%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!Created, "dd-mm-yyyy Hh:Nn:Ss") & "</td>"
                                        HTML = HTML + "<td width=" + Chr$(34) + "17%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!Paid, "Currency") & "</td>"
                                        ttlPaid = ttlPaid + rsload!Paid
                                        HTML = HTML + "<td width=" + Chr$(34) + "13%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & Format(rsload!Refunded, "Currency") & "</td>"
                                        ttlRefund = ttlRefund + rsload!Refunded
                                        HTML = HTML + "<td width=" + Chr$(34) + "77%" + Chr$(34) + " align=" + Chr$(34) + "right" + Chr$(34) + ">" & rsload!paymenttype & "</td>"
                                        'MySQL.Execute oConn, "Update receipts Set StatementID = " & dStatementNo & " Where RecID = " & rsLoad!RecID
                                        
                                    rsload.MoveNext
                                    HTML = HTML + "</tr>"
                                Wend
                                
                              
                            HTML = HTML + "</table>"
                            
                        End If
                    End If
                            
                xHTML = MySQL.ReplaceString(xHTML, "/Payments/", HTML)
                HTML = ""
            End If
            
                HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
                  HTML = HTML + "<tr>"
                    HTML = HTML + "<td width=" + Chr$(34) + "33%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Total "
                      HTML = HTML + "Invoiced</font></td>"
                    HTML = HTML + "<td width=" + Chr$(34) + "33%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Total "
                      HTML = HTML + "</font></td>"
                    HTML = HTML + "<td width=" + Chr$(34) + "34%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008000" + Chr$(34) + "><font color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Amount"
                      HTML = HTML + " Outstanding</font></td>"
                      If bMode = False Then MySQL.Execute oConn, "Update statementtraxr Set amtInvoiced = " & ttlInvoiced & ", amtreceipts = " & ttlPaid & ", amtRefunded = " & ttlRefund & " Where RecID = " & dStatementNo
                  HTML = HTML + "</tr>"
                  HTML = HTML + "<tr>"
                    HTML = HTML + "<td width=" + Chr$(34) + "33%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & Format(ttlInvoiced, "Currency") & "</td>"
                    HTML = HTML + "<td width=" + Chr$(34) + "33%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & Format(ttlPaid, "Currency") & "</td>"
                    HTML = HTML + "<td width=" + Chr$(34) + "34%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & Format(ttlInvoiced - ttlPaid - ttlRefund, "Currency") & "</td>"
                  HTML = HTML + "</tr>"
                HTML = HTML + "</table>"
              
                xHTML = MySQL.ReplaceString(xHTML, "/Totals/", HTML)
                HTML = ""
                            
    If ttlInvoiced > 0 Or ttlPaid > 0 Or ttlRefunded > 0 Then sendStatementHTML = xHTML
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function SaveHTML(HTML As Variant, Filename As String, dDate As Date) As String


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveHTML"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim stmp As String
    
    Filename = MySQL.ReplaceString(Filename, "#", "")
    Filename = MySQL.ReplaceString(Filename, "/", "")
    Filename = MySQL.ReplaceString(Filename, "\", "")
    Filename = MySQL.ReplaceString(Filename, "?", "")
    Filename = MySQL.ReplaceString(Filename, ":", "")
    
    If Len(Dir(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(dDate, "yyyy-mmmm-dd"), vbDirectory)) > 0 Then
        stmp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(dDate, "yyyy-mmmm-dd") + "\" + Filename
    Else
        MkDir IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(sysnow, "yyyy-mmmm-dd")
        stmp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(sysnow, "yyyy-mmmm-dd") + "\" + Filename
    End If
        
    If Dir(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(dDate, "yyyy-mmmm-dd") + "\" + Filename, vbNormal) <> "" Then
        Randomize Now / Len(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(dDate, "yyyy-mmmm-dd") + "\" + Filename)
               
        Do
            gSleep
            If Dir(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(dDate, "yyyy-mmmm-dd") + "\" + Format(dDate, "+hh-nn-ss") + "_" + Filename, vbNormal) <> "" Then
            
            Else
                stmp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + Format(dDate, "yyyy-mmmm-dd") + "\" + Format(dDate, "+hh-nn-ss") + "_" + Filename
                Exit Do
            End If
        Loop Until Err.Number <> 0
    Else

    End If
    
    Dim iFileNum As Long
    
    iFileNum = FreeFile + 1
    'MsgBox stmp
    stmp = MySQL.ReplaceString(stmp, "*", "")
    
'    MsgBox stmp
    Open stmp For Output As #iFileNum
    
        Print #iFileNum, HTML
        
    Close #iFileNum
    
    SaveHTML = stmp
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function SendQuotaHTML(oConn As adodb.Connection, acciRecID As Long, QuotaMSGID As Long) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SendQuotaHTML"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim HTML As String
    Dim xHTML As Variant
    
    HTML = ""
    
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(oConn, rsload, , "select HTML from stationary where StationaryCode = 'QUOTA'")
    
    If rsload.RecordCount > 0 Then
    
        xHTML = rsload!HTML
        
    Else
        
        Exit Function
        
    End If

        Dim rsVISP As adodb.Recordset
        If MySQL.OpenTable(oConn, rsVISP, , "select LogoURL, virtualisp.Description from virtualisp, accountinfo where virtualisp.RecID = accountinfo.VirtualID and accountinfo.RecID = " & acciRecID & " Limit 1") = True Then
              If rsVISP.RecordCount > 0 Then
                  If IsNull(rsVISP!LogoURL) Or rsVISP!LogoURL = "" Then
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", rsVISP!Description)
                  Else
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(IsNull(rsVISP!LogoURL), "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + "><BR>" & rsVISP!Description)
                  End If
              Else
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(rsVISP.RecordCount = 0, "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + ">")
              End If
        End If
        
         If MySQL.OpenTable(oConn, rsload, , "select AccountName from accountinfo where RecID = " & acciRecID) = True Then
          If rsload.RecordCount > 0 Then
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", IIf(IsNull(rsload!AccountName), "(null)", rsload!AccountName))
            Else
              xHTML = MySQL.ReplaceString(xHTML, "/AccountName/", "")
          End If
        End If
        xHTML = MySQL.ReplaceString(xHTML, "/NOW/", Format(sysnow, "dd/mm/yyyy Hh:Nn:Ss"))
            

            HTML = HTML + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
              HTML = HTML + "<tr>"
                HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + "><font size=" + Chr$(34) + "2" + Chr$(34) + " face=" + Chr$(34) + "Verdana, Arial, Helvetica, sans-serif" + Chr$(34) + " color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Time</b></font></td>"
                HTML = HTML + "<td width=" + Chr$(34) + "40%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + "><font size=" + Chr$(34) + "2" + Chr$(34) + " face=" + Chr$(34) + "Verdana, Arial, Helvetica, sans-serif" + Chr$(34) + " color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Description</b></font></td>"
                HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + "><font size=" + Chr$(34) + "2" + Chr$(34) + " face=" + Chr$(34) + "Verdana, Arial, Helvetica, sans-serif" + Chr$(34) + " color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Unit's"
                  HTML = HTML + " Left</b></font></td>"
                HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + " bgcolor=" + Chr$(34) + "#008080" + Chr$(34) + "><font size=" + Chr$(34) + "2" + Chr$(34) + " face=" + Chr$(34) + "Verdana, Arial, Helvetica, sans-serif" + Chr$(34) + " color=" + Chr$(34) + "#FFFFFF" + Chr$(34) + "><b>Unit's "
                  HTML = HTML + "Used</b></font></td>"
              HTML = HTML + "</tr>"
              
              Dim rsQuota As adodb.Recordset
              
              If MySQL.OpenTable(oConn, rsQuota, , "select * from acci_quotareceipt where QuotaMSGID = '" & QuotaMSGID & "'") = True Then
                    If rsQuota.State = adStateOpen Then
                    
                        If rsQuota.RecordCount > 0 Then
                            While Not rsQuota.EOF And Err.Number = 0
                                HTML = HTML + "<tr>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" + IIf(IsNull(rsQuota!Created), "&nbsp;", Format(rsQuota!Created, "Hh:Nn:Ss")) + "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "40%" + Chr$(34) + ">" + IIf(IsNull(rsQuota!Description), "&nbsp;", rsQuota!Description) + "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & IIf(IsNull(rsQuota!UnitsLeft), "&nbsp;", rsQuota!UnitsLeft) & IIf(IsNull(rsQuota!UnitType), "&nbsp;", rsQuota!UnitType) & "</td>"
                                    HTML = HTML + "<td width=" + Chr$(34) + "20%" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">" & IIf(IsNull(rsQuota!UnitsUsed), "&nbsp;", rsQuota!UnitsUsed) & IIf(IsNull(rsQuota!UnitType), "&nbsp;", rsQuota!UnitType) & "</td>"
                                HTML = HTML + "</tr>"
                                rsQuota.Update
                                rsQuota!QuotaMSGSent = True
                                rsQuota.MoveNext
                            Wend
                        End If
                    End If
              End If
              
              HTML = HTML + "</table>"

            xHTML = MySQL.ReplaceString(xHTML, "/Quota/", HTML)

SendQuotaHTML = xHTML

Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function GetPOHTML(POID As Double, acciRecID As Long, ShippingID As Long, oConn As adodb.Connection) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "GetPOHTML"
    Const ContainerName = "smtp"
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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


    Dim xHTML As Variant
    Dim HTML As String
    Dim sHTML As String
    
    
    HTML = ""
    
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(oConn, rsload, , "select HTML from stationary where StationaryCode = 'PURCHASEORDER'")
    
    If rsload.RecordCount > 0 Then
    
        xHTML = rsload!HTML
        
    Else
        
        Exit Function
        
    End If


    
    
        Dim rsVISP As adodb.Recordset
        If MySQL.OpenTable(oConn, rsVISP, , "select LogoURL, virtualisp.Description from virtualisp, accountinfo where virtualisp.RecID = accountinfo.VirtualID and accountinfo.RecID = " & acciRecID & " Limit 1") = True Then
              If rsVISP.RecordCount > 0 Then
                  If IsNull(rsVISP!LogoURL) Or rsVISP!LogoURL = "" Then
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", rsVISP!Description)
                  Else
                        xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(IsNull(rsVISP!LogoURL), "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + "><BR>" & rsVISP!Description)
                  End If
              Else
                  xHTML = MySQL.ReplaceString(xHTML, "/VISPLogo/", "<img src=" + Chr$(34) + IIf(rsVISP.RecordCount = 0, "http://www.ep.net.au/projectalpha/stationary/ident.jpg", rsVISP!LogoURL) + Chr$(34) + " width=" + Chr$(34) + "100" + Chr$(34) + " height=" + Chr$(34) + "100" + Chr$(34) + ">")
              End If
        End If
        
        If MySQL.OpenTable(oConn, rsload, , "select * from acci_addresses where RecID = " & ShippingID & " Limit 1") = True Then
          If rsload.RecordCount > 0 Then
              HTML = HTML + vbCrLf + vbCrLf + IIf(IsNull(rsload!ContactName), "", rsload!ContactName) & "<br>" & IIf(IsNull(rsload!Street1), "", rsload!Street1) & "<br>" & IIf(IsNull(rsload!Street2), "", rsload!Street2) & "<br>" & IIf(IsNull(rsload!Suburb), "", rsload!Suburb) & "<br>" & IIf(IsNull(rsload!State), "", rsload!State) & ", " & IIf(IsNull(rsload!PostCode), "", rsload!PostCode) & ", " & IIf(IsNull(rsload!Country), "", rsload!Country) & "<br><br>"
          End If
        End If
        
        
          xHTML = MySQL.ReplaceString(xHTML, "/ShippingID/", "<ShippingAddress>" & HTML & "</ShippingAddress>")
          xHTML = MySQL.ReplaceString(xHTML, "/NOW/", "<SysTime>" & Format(sysnow, "ddd dd-mm-yyyy ttttt") & "</SysTime>")
          xHTML = MySQL.ReplaceString(xHTML, "/POID/", "<PONumber>" & MySQL.ReplaceString("" & POID, "-", "J-") & "</PONumber>")
          
          HTML = ""
    
    
    Dim rsCnt As adodb.Recordset
    
    Dim SQL As String
    
    
        SQL = "select distinct acci_services.DomainID, acci_services.RecID, (acci_services.PeriodFee * " & oTax(Login.TaxCode, Login.TaxCountry) & ") + acci_services.PeriodFee as SalePrice, (acci_services.JoiningFee * " & oTax(Login.TaxCode, Login.TaxCountry) & ") + acci_services.JoiningFee as JoiningInc, acci_services.DefaultShippingID, acci_services.RadiusID, acci_services.ContractID, vendors.RecID as VendorID, accountinfo.RecID as AccI_RecID, servicetypes.ServiceKey, plantemplates.VendorPartID, " + _
                "plantemplates.SubPartID, plantypes.CostPrice, plantemplates.Description, vendors.vName, plantemplates.PeriodFee, plantemplates.MBQuota, acci_services.Username, AES_DECRYPT(Password, '" & odb.colSalts.ReturnSalt("md5Password") & "') as Password " + _
                "from plantemplates left join plantypes on plantemplates.RecID = plantypes.TemplateID " + _
                "left join vendors on vendors.RecID = plantypes.VendorID " + _
                "left join acci_services on acci_services.ptRecID = plantypes.RecID " + _
                "left join accountinfo on acci_services.AccI_RecID = accountinfo.RecID " + _
                "left join servicetypes on acci_services.ServiceID = servicetypes.RecID " + _
                "where acci_services.POID = " & Val(POID) & " order by servicetypes.ServiceKey"

        
    Call MySQL.OpenTable(oConn, rsload, , SQL)
    

    
 HTML = "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "95%" + Chr$(34) + "><tr>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + "><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + "><b>Your"
    HTML = HTML + vbCrLf + " Part Code</b></font></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + "><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + "><b>Sub"
      HTML = HTML + vbCrLf + " Part Code</b></font></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "30%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + " align=" + Chr$(34) + "center" + Chr$(34) + ">"
      HTML = HTML + vbCrLf + "<p align=" + Chr$(34) + "center" + Chr$(34) + "><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + "><b>Description</b></font></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + ">"
      HTML = HTML + vbCrLf + "<p align=" + Chr$(34) + "center" + Chr$(34) + "><b><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + ">Cost Price<br>"
      HTML = HTML + vbCrLf + "(Ex Tax)</font></b></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + ">"
      HTML = HTML + vbCrLf + "<p align=" + Chr$(34) + "center" + Chr$(34) + "><b><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + ">Username</font></b></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + ">"
      HTML = HTML + vbCrLf + "<p align=" + Chr$(34) + "center" + Chr$(34) + "><b><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + ">Password</font></b></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + ">"
      HTML = HTML + vbCrLf + "<p align=" + Chr$(34) + "center" + Chr$(34) + "><b><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + ">MB's</font></b></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + ">"
      HTML = HTML + vbCrLf + "<p align=" + Chr$(34) + "center" + Chr$(34) + "><b><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + ">Sale Price (Inc Tax)</font></b></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + " bgcolor=" + Chr$(34) + "#C0C0C0" + Chr$(34) + " bordercolor=" + Chr$(34) + "#808000" + Chr$(34) + ">"
      HTML = HTML + vbCrLf + "<p align=" + Chr$(34) + "center" + Chr$(34) + "><b><font face=" + Chr$(34) + "Arial" + Chr$(34) + " color=" + Chr$(34) + "#000000" + Chr$(34) + ">Joining Fee (Inc Tax)</font></b></td>"
  
  HTML = HTML + vbCrLf + "</tr>"
  
  Dim iTotal As Single
  Dim HTML2 As String
  Dim rsDSL As adodb.Recordset
  Dim rsDomain As adodb.Recordset
  Dim rsXML As adodb.Recordset
  Dim Xtra As String
  Dim iServices As Long
  Dim rsSysops As adodb.Recordset
  
  If rsload.RecordCount > 0 Then
    
    While Not rsload.EOF And Err.Number = 0
        
        iServices = iServices + 1
        Call MySQL.OpenTable(directConn, rsCnt, , "select contracttemplates.Description from contractsruntime inner join contracttemplates on contractsruntime.ContractID = contracttemplates.RecID where contractsruntime.RecID = '" & rsload!ContractID & "'")
        If rsCnt.State = adStateOpen Then
            If rsCnt.RecordCount > 0 Then
                Xtra = "<br>" & rsCnt!Description
            End If
        End If
        
        HTML = HTML + vbCrLf + "<Service" & iServices & ">"
        HTML = HTML + vbCrLf + "<tr>"
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + ">" & IIf(IsNull(rsload!VendorPartID), "&nbsp;", rsload!VendorPartID) & "</td>"
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + ">" & IIf(IsNull(rsload!SubPartID), "&nbsp;", rsload!SubPartID) & "</td>"
        
        HTML2 = ""
        Select Case UCase(rsload!ServiceKey)
        '**** NOTE DOMAIN EXTENTIONS WAS ADDED TO THIS ROUTINE WHERE THE DATABASE IS EXTRACTED..
        '***  REMINDER The Domain Reference table is in XML
        Case "DOMAIN"
            
            
            If MySQL.OpenTable(directConn, rsXML, , "select * from domaindocs where DomainID = '" & rsload!DomainID & "' order by DocType, RecID") = True Then
                
                
                If rsXML.RecordCount > 0 Then
                    If MySQL.OpenTable(directConn, rsDomain, , "select * from domainlist where RecID = '" & rsload!DomainID & "'") = True Then
                    
                        HTML2 = HTML2 + vbCrLf + "<table width=" + Chr$(34) + "100%" + Chr$(34) + "  border=" + Chr$(34) + "0" + Chr$(34) + " cellspacing=" + Chr$(34) + "1" + Chr$(34) + " cellpadding=" + Chr$(34) + "0" + Chr$(34) + ">"
                          HTML2 = HTML2 + vbCrLf + "<tr bordercolor=" + Chr$(34) + "#0033CC" + Chr$(34) + " bgcolor=" + Chr$(34) + "#ECE9D8" + Chr$(34) + ">"
                            HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "52%" + Chr$(34) + ">Domian: <domainname>" & IIf(IsNull(rsDomain!Domain), "", rsDomain!Domain) & "</domainname> </td>"
                            HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "48%" + Chr$(34) + ">Admins Email Address: <adminemail>" & IIf(IsNull(rsDomain!AdminEmail), "", rsDomain!AdminEmail) & "</adminemail></td>"
                          HTML2 = HTML2 + vbCrLf + "</tr>"
                          HTML2 = HTML2 + vbCrLf + "<tr bordercolor=" + Chr$(34) + "#0033CC" + Chr$(34) + " bgcolor=" + Chr$(34) + "#ECE9D8" + Chr$(34) + ">"
                            HTML2 = HTML2 + vbCrLf + vbTab + "<td>Tech Name: <techname>" & IIf(IsNull(rsDomain!TechName), "", rsDomain!TechName) & "</techname></td>"
                            Call MySQL.OpenTable(directConn, rsSysops, , "select Username from sysops where RecID = '" & rsDomain!SysopID & "'")
                            HTML2 = HTML2 + vbCrLf + vbTab + "<td>Served by Sysops: <sysopusername>" & IIf(IsNull(rsSysops!Username), "Unknown", rsSysops!Username) & "</sysopusername> </td>"
                          HTML2 = HTML2 + vbCrLf + "</tr>"
                          HTML2 = HTML2 + vbCrLf + "<tr bordercolor=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">"
                            HTML2 = HTML2 + vbCrLf + vbTab + "<td>&nbsp;</td>"
                            HTML2 = HTML2 + vbCrLf + vbTab + "<td>&nbsp;</td>"
                          HTML2 = HTML2 + vbCrLf + "</tr>"
                    
                    End If
                    
                    While Not rsXML.EOF
                        Select Case rsXML!DocType
                        Case "ADDRESS"
                        
                            HTML2 = HTML2 + vbCrLf + "<tr>"
                              HTML2 = HTML2 + vbCrLf + vbTab + "<td colspan=" + Chr$(34) + "2" + Chr$(34) + "><table width=" + Chr$(34) + "100%" + Chr$(34) + "  border=" + Chr$(34) + "0" + Chr$(34) + " cellspacing=" + Chr$(34) + "1" + Chr$(34) + " cellpadding=" + Chr$(34) + "0" + Chr$(34) + ">"
                                HTML2 = HTML2 + vbCrLf + "<tr>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td><strong><XMLDOC_" & rsXML!RecID & ">:: <itemtext>" & IIf(IsNull(rsXML!ItemText), "Address", rsXML!ItemText) & "</itemtext> :: <docdescription>" & IIf(IsNull(rsXML!Description), "Address", rsXML!Description) & "</docdescription> ::</strong></td>"
                                HTML2 = HTML2 + vbCrLf + "</tr>"
                                HTML2 = HTML2 + vbCrLf + "<tr>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td><div align=" + Chr$(34) + "center" + Chr$(34) + "><Address><contact>" + XMLTag(1, rsXML!DocText, "Contact") + "</contact><br />"
                                      HTML2 = HTML2 + vbCrLf + "<Street1>" + XMLTag(1, rsXML!DocText, "Street1") + "</Street1><br />"
                                      HTML2 = HTML2 + vbCrLf + "<Street2>" + XMLTag(1, rsXML!DocText, "Street2") + "</Street2>"
                                      HTML2 = HTML2 + vbCrLf + "<br />"
                                      HTML2 = HTML2 + vbCrLf + "<Suburb>" + XMLTag(1, rsXML!DocText, "Suburb") + "</Suburb> <State>" + XMLTag(1, rsXML!DocText, "State") + "</State> <Country>" + XMLTag(1, rsXML!DocText, "Country") + "</Country></address></div></td>"
                                HTML2 = HTML2 + vbCrLf + "</tr>"
                                HTML2 = HTML2 + vbCrLf + "<tr>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td><strong></XMLDOC_" & rsXML!RecID & "></strong></td>"
                                HTML2 = HTML2 + vbCrLf + "</tr>"
                              HTML2 = HTML2 + vbCrLf + "</table></td>"
                            HTML2 = HTML2 + vbCrLf + "</tr>"
                      
                      Case "PHONE"
                      
                            HTML2 = HTML2 + vbCrLf + "<tr>"
                              HTML2 = HTML2 + vbCrLf + vbTab + "<td colspan=" + Chr$(34) + "2" + Chr$(34) + "><table width=" + Chr$(34) + "100%" + Chr$(34) + "  border=" + Chr$(34) + "0" + Chr$(34) + " cellspacing=" + Chr$(34) + "1" + Chr$(34) + " cellpadding=" + Chr$(34) + "0" + Chr$(34) + ">"
                                HTML2 = HTML2 + vbCrLf + "<tr>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td colspan=" + Chr$(34) + "4" + Chr$(34) + "><div align=" + Chr$(34) + "left" + Chr$(34) + "><strong> :: Phone List for This Domain :: <docdescription>Doc Description</docdescription> ::<phonelist> </strong></div></td>"
                                  HTML2 = HTML2 + vbCrLf + "</tr>"
                                HTML2 = HTML2 + vbCrLf + "<tr bordercolor=" + Chr$(34) + "#000000" + Chr$(34) + ">"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "32%" + Chr$(34) + "><div align=" + Chr$(34) + "center" + Chr$(34) + "><strong>Contact Name</strong></div></td>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "23%" + Chr$(34) + "><div align=" + Chr$(34) + "center" + Chr$(34) + "><strong>Phone Number </strong></div></td>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "13%" + Chr$(34) + "><div align=" + Chr$(34) + "center" + Chr$(34) + "><strong>Extension</strong></div></td>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "32%" + Chr$(34) + "><div align=" + Chr$(34) + "center" + Chr$(34) + "><strong>Short Note </strong></div></td>"
                                HTML2 = HTML2 + vbCrLf + "</tr>"
                                While rsXML!DocType = "PHONE"
                                    
                                    HTML2 = HTML2 + vbCrLf + "</XMLDOC_" & rsXML!RecID & "><tr bordercolor=" + Chr$(34) + "#000000" + Chr$(34) + ">"
                                      HTML2 = HTML2 + vbCrLf + vbTab + "<td><div align=" + Chr$(34) + "center" + Chr$(34) + "><contact>" + XMLTag(1, rsXML!DocText, "Contact") + "</contact></div></td>"
                                      HTML2 = HTML2 + vbCrLf + vbTab + "<td><div align=" + Chr$(34) + "center" + Chr$(34) + "><phoneno>" + XMLTag(1, rsXML!DocText, "PhoneNumber") + "</phoneno></div></td>"
                                      HTML2 = HTML2 + vbCrLf + vbTab + "<td><div align=" + Chr$(34) + "center" + Chr$(34) + "><extn>" + XMLTag(1, rsXML!DocText, "Extension") + "</extn></div></td>"
                                      HTML2 = HTML2 + vbCrLf + vbTab + "<td><div align=" + Chr$(34) + "center" + Chr$(34) + "><note>" + XMLTag(1, rsXML!DocText, "ShortNote") + "</note></div></td>"
                                    HTML2 = HTML2 + vbCrLf + "</tr><XMLDOC_" & rsXML!RecID & "></strong>"
                                    rsXML.MoveNext
                                    
                                    
                               Wend
                               rsXML.MovePrevious
                               HTML2 = HTML2 + vbCrLf + "</table></td>"
                            HTML2 = HTML2 + vbCrLf + "</tr>"
                            HTML2 = HTML2 + vbCrLf + "<tr>"
                              HTML2 = HTML2 + vbCrLf + vbTab + "<td colspan=" + Chr$(34) + "2" + Chr$(34) + "></phonelist><strong></strong></td>"
                            HTML2 = HTML2 + vbCrLf + "</tr>"
                            
                      Case "XML"
                      
                            HTML2 = HTML2 + vbCrLf + "<tr>"
                              HTML2 = HTML2 + vbCrLf + vbTab + "<td colspan=" + Chr$(34) + "2" + Chr$(34) + "><table width=" + Chr$(34) + "100%" + Chr$(34) + "  border=" + Chr$(34) + "0" + Chr$(34) + " cellspacing=" + Chr$(34) + "1" + Chr$(34) + " cellpadding=" + Chr$(34) + "0" + Chr$(34) + ">"
                                HTML2 = HTML2 + vbCrLf + "<tr>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td><strong><XMLDOC_" & rsXML!RecID & ">:: XML Doc :: <docdescription>" & IIf(IsNull(rsXML!ItemText), "XML Document", rsXML!Description) & "</docdescription> :: </strong></td>"
                                HTML2 = HTML2 + vbCrLf + "</tr>"
                                HTML2 = HTML2 + vbCrLf + "<tr>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td><XML>" & IIf(IsNull(rsXML!DocText), "", rsXML!DocText) & "<XML></td>"
                                HTML2 = HTML2 + vbCrLf + "</tr>"
                                HTML2 = HTML2 + vbCrLf + "<tr>"
                                  HTML2 = HTML2 + vbCrLf + vbTab + "<td><strong></XMLDOC_" & rsXML!RecID & "></strong></td>"
                                HTML2 = HTML2 + vbCrLf + "</tr>"
                              HTML2 = HTML2 + vbCrLf + "</table></td>"
                            HTML2 = HTML2 + vbCrLf + "</tr>"
                    End Select
                    rsXML.MoveNext
                    
                Wend
                
                HTML2 = HTML2 + vbCrLf + "</tr>"
                HTML2 = HTML2 + vbCrLf + "</table>"
                'HTML2 = HTML2 + vbCrLf +"</tr>"
                'HTML2 = HTML2 + vbCrLf +"</table>"
                End If
            End If
            
        Case "ADSL", "SHDSL"
            If MySQL.OpenTable(directConn, rsDSL, , "select * from acci_dslconnections where RadiusID = " & rsload!RadiusID) = True Then
                If rsDSL.RecordCount > 0 Then
                    HTML2 = HTML2 + vbCrLf + "<DSLConn" & rsDSL!RecID & ">"
                    HTML2 = HTML2 + vbCrLf + "<table border=" + Chr$(34) + "1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + ">"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "100%" + Chr$(34) + " colspan=" + Chr$(34) + "2" + Chr$(34) + "><font face=" + Chr$(34) + "Arial" + Chr$(34) + " size=" + Chr$(34) + "2" + Chr$(34) + ">" & IIf(IsNull(rsload!ServiceKey), "", "[<ServiceKey>" & rsload!ServiceKey & "</ServiceKey>] - ") & IIf(IsNull(rsload!Description), "", "[<ServiceDesc>" & rsload!Description & "</ServiceDesc>] - ") & Xtra & "</font></td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Account Name:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + ">" & IIf(IsNull(rsDSL!AccountName), "&nbsp;", "<AccountName>" & rsDSL!AccountName & "</AccountName>") & "</td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">DSL Line Number:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + ">(" & IIf(IsNull(rsDSL!AreaCode), "&nbsp;", "<AreaCode>" & rsDSL!AreaCode & "</AreaCode>") & ") " & IIf(IsNull(rsDSL!PhoneNumber), "", "<Landline>" & rsDSL!PhoneNumber & "</Landline>") & "</td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Email Address to notify:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + ">" & IIf(IsNull(rsDSL!eMail), "&nbsp;", "<Notify>" & rsDSL!eMail & "</Notify>") & "</td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Street Line 1:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + ">" & IIf(IsNull(rsDSL!UnitNo), "&nbsp;", "<UnitNo>" & rsDSL!UnitNo & "</UnitNo>") & "</td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Street Line 2:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + ">" & IIf(IsNull(rsDSL!StreetNo), "&nbsp;", "<StreetNo>" & rsDSL!StreetNo) & "</StreetNo> <StreetName>" & IIf(IsNull(rsDSL!StreetName), "", rsDSL!StreetName) & "</StreetName> <StreetType>" & IIf(IsNull(rsDSL!StreetType), "", rsDSL!StreetType) & "</StreetType></td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Suburb:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + "><Suburb>" & IIf(IsNull(rsDSL!Suburb), "&nbsp;", rsDSL!Suburb) & "</Suburb></td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Postcode:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + "><Postcode>" & IIf(IsNull(rsDSL!PostCode), "&nbsp;", rsDSL!PostCode) & "</Postcode></td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Country:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + "><Country>" & IIf(IsNull(rsDSL!Country), "&nbsp;", rsDSL!Country) & "</Country></td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">State:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + "><State>" & IIf(IsNull(rsDSL!State), "&nbsp;", rsDSL!State) & "</State></td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                      HTML2 = HTML2 + vbCrLf + "<tr>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "22%" + Chr$(34) + ">Churn:</font></td>"
                        HTML2 = HTML2 + vbCrLf + vbTab + "<td width=" + Chr$(34) + "78%" + Chr$(34) + "><Churn>" & IIf(IsNull(Val(rsDSL!Churn)), "No", IIf(Val(rsDSL!Churn) = 0, "No", "Yes")) & "</Churn></td>"
                      HTML2 = HTML2 + vbCrLf + "</tr>"
                    HTML2 = HTML2 + vbCrLf + "</table>"
                    HTML2 = HTML2 + vbCrLf + "</DSLConn" & rsDSL!RecID & ">"
                    
                    
                End If
                
            End If
        Case Else
            HTML2 = IIf(IsNull(rsload!ServiceKey), "", "[<ServiceKey>" & rsload!ServiceKey & "</ServiceKey>] - <ServiceDesc>") & IIf(IsNull(rsload!Description), "", "[" & rsload!Description & "</ServiceDesc>] - ") & Xtra
        End Select
    
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "30%" + Chr$(34) + ">" & HTML2 & "</td>"

        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "><PeriodFee>" & IIf(IsNull(rsload!PeriodFee), "&nbsp;", Format(rsload!PeriodFee, "$ ###,###,###,###,###.####0")) & "</PeriodFee></td>"
        iTotal = iTotal + IIf(IsNull(rsload!CostPrice), 0, rsload!CostPrice)
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "><Username>" & IIf(IsNull(rsload!Username), "(n/a)", rsload!Username) & "</Username></td>"
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "><Password>" & IIf(IsNull(rsload!Password), "(n/a)", rsload!Password) & "</Password></td>"
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "><MBQuota>" & IIf(IsNull(rsload!MBQuota), "0", rsload!MBQuota) & "</MBQuota></td>"
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "><SalePrice>" & IIf(IsNull(rsload!SalePrice), "0", Format(rsload!SalePrice, "$ ###,###,###,###,###.####0")) & "</SalePrice></td>"
        HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "><JoiningInc>" & IIf(IsNull(rsload!JoiningInc), "0", Format(rsload!JoiningInc, "$ ###,###,###,###,###.####0")) & "</JoiningInc></td>"
        HTML = HTML + vbCrLf + "</tr>"
        HTML = HTML + vbCrLf + "</Service" & iServices & ">"
        rsload.MoveNext
    Wend
  
    xHTML = MySQL.ReplaceString(xHTML, "/TotalServices/", "<TotalServices>" & iServices & "</TotalServices>")
    xHTML = MySQL.ReplaceString(xHTML, "<title>", "<title><NumServices>" & iServices & "</NumServices> Services On Purchase Order - ")
  Else
  
    HTML = HTML + vbCrLf + "<tr>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "30%" + Chr$(34) + "></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "></td>"
    HTML = HTML + vbCrLf + vbTab + "<td width=" + Chr$(34) + "10%" + Chr$(34) + "></td>"
    HTML = HTML + vbCrLf + "</tr>"
    
    xHTML = MySQL.ReplaceString(xHTML, "/TotalServices/", "<TotalServices>" & 0 & "</TotalServices>")
  
  End If
  
  HTML = HTML + vbCrLf + "</table>"
  
  xHTML = MySQL.ReplaceString(xHTML, "/Items/", HTML)
  
'  rsload.MoveLast
  
    xHTML = MySQL.ReplaceString(xHTML, "/POValue/", "<POValue>" & iTotal & "</POValue>")
    xHTML = MySQL.ReplaceString(xHTML, "/POGST/", "<POGST>" & iTotal * oTax(Login.TaxCode, Login.TaxCountry) & "</POGST>")
    iTotal = iTotal * oTax(Login.TaxCode, Login.TaxCountry) + iTotal
    xHTML = MySQL.ReplaceString(xHTML, "/Total/", "<POTotal>" & Format(iTotal, "Currency") & "</POTotal>")
  
  GetPOHTML = xHTML
  
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
