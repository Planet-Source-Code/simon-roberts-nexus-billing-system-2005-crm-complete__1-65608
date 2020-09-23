VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmGroup 
   Caption         =   "Group Payment"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12735
   ControlBox      =   0   'False
   Icon            =   "frmGroup_new.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Group Payment Calculator"
      Height          =   8565
      Left            =   8490
      TabIndex        =   3
      Top             =   60
      Width           =   4155
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Close"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   150
         TabIndex        =   15
         Top             =   7860
         Width           =   3765
      End
      Begin VB.CommandButton cmdmakePayment 
         Caption         =   "Make Group Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   14
         Top             =   7260
         Width           =   3765
      End
      Begin VB.TextBox txtCalc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Text            =   "$ 0.00"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtCalc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   180
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Text            =   "$ 0.00"
         Top             =   3330
         Width           =   3795
      End
      Begin VB.Frame Frame3 
         Caption         =   "Payment Calculator"
         Height          =   3375
         Left            =   180
         TabIndex        =   5
         Top             =   3750
         Width           =   3825
         Begin VB.CommandButton cmdEqual 
            Cancel          =   -1  'True
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   180
            TabIndex        =   13
            Top             =   2490
            Width           =   735
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   180
            TabIndex        =   12
            Top             =   1710
            Width           =   735
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1800
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   300
            Width           =   1905
         End
         Begin MSComctlLib.ListView lvPaymentType 
            Height          =   2565
            Left            =   990
            TabIndex        =   16
            Tag             =   "0"
            Top             =   690
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   4524
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   4763
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Sub"
               Object.Width           =   1411
            EndProperty
         End
         Begin VB.Label Label2 
            Caption         =   "Payment Type:"
            Height          =   525
            Index           =   2
            Left            =   150
            TabIndex        =   11
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Amount Debit:"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   9
            Top             =   390
            Width           =   1410
         End
      End
      Begin VB.TextBox txtCalc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Index           =   0
         Left            =   210
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   660
         Width           =   3795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount Debit:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   330
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Active Unpaid Invoices"
      Height          =   7785
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   8355
      Begin VB.Frame Frame4 
         Caption         =   "Statistics"
         Height          =   645
         Left            =   180
         TabIndex        =   17
         Top             =   330
         Width           =   7965
         Begin VB.Label lblAllocated 
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1770
            TabIndex        =   19
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Amount unallocated: "
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   270
            Width           =   1500
         End
      End
      Begin MSComctlLib.ListView lvInvoices 
         Height          =   6525
         Left            =   180
         TabIndex        =   1
         Top             =   1080
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   11509
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilTreeview"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "VISPDesc"
            Text            =   "ViSP"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "InvoiceSerial"
            Text            =   "Invoice Number"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Object.Tag             =   "CreatedWhen"
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Payment Due"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "AccountName"
            Text            =   "Account Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "TotalDue"
            Text            =   "Debit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "AmountPaid"
            Text            =   "Paid"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Credit"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilTreeview 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":1ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":27AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":3084
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":395E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":4238
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":4B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":53EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":5CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":65A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":6E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":7754
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":802E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":8908
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":91E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":9ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":A396
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":AC70
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":B54A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":BE24
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":C6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":CFD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":D8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":E18C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":EA66
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":F340
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":FC1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":104F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":10DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":116A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":11F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":1285C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":13136
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":13A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":142EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":14BC4
            Key             =   "book"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":14EDE
            Key             =   "news"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":151F8
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":15512
            Key             =   "world"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":1582C
            Key             =   "Finalised"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":15C7E
            Key             =   "Unfinalised"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":160D0
            Key             =   "Overdue"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup_new.frx":16522
            Key             =   "Partially"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   990
      X2              =   8430
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Payment"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   990
      TabIndex        =   2
      Top             =   120
      Width           =   2280
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   60
      Picture         =   "frmGroup_new.frx":16C74
      Stretch         =   -1  'True
      Top             =   30
      Width           =   780
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim csPay As New clsPay

Public Function LoadFlags()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadFlags"
    Const ContainerName = "frmGroup"
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


    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    Dim rsload As adodb.Recordset
    Dim itmX As ListItem
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from paymenttype")
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvPaymentType.ListItems.Add(, IIf(rsload!CreditCard = True, "c", "a") & rsload!RecID, rsload!Description)
            itmX.SubItems(1) = IIf(rsload!sub = 0, "", Format(rsload!sub, "Currency"))
            itmX.Tag = rsload!sub
            rsload.MoveNext
        Wend
    End If


    Dim ix As Integer
    
    If MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct virtualisp.Description as VISPDescription ,accountinfo.AccountName ,invoicetraxr.* from accountinfo inner join invoicetraxr on invoicetraxr.acci_RecID = accountinfo.RecID, virtualisp where invoicetraxr.Finalised = '0'", "accountinfo", True) + " group by RecID Order By accountinfo.AccountName DESC") = True Then
        If Not rsload.EOF And Not rsload.BOF Then
            If rsload.RecordCount > 0 Then
                lvInvoices.ListItems.Clear
                While Not rsload.EOF And Err.Number = 0
                    
                        Set itmX = lvInvoices.ListItems.Add(, "i" & rsload!RecID, rsload!VISPDescription)
                        itmX.SubItems(1) = IIf(IsNull(rsload!InvoiceSerial), "", rsload!InvoiceSerial)
                        itmX.SubItems(2) = IIf(IsNull(rsload!Created), "", Format(rsload!Created, "yyyy-mm-dd dddd"))
                        itmX.SubItems(3) = IIf(IsNull(rsload!PaymentDue), "", Format(rsload!PaymentDue, "yyyy-mm-dd dddd"))
                        itmX.SubItems(4) = IIf(IsNull(rsload!AccountName), "", rsload!AccountName)
                        itmX.SubItems(5) = IIf(IsNull(rsload!TotalDue), "$ 0.00", Format(rsload!TotalDue, "Currency"))
                        itmX.SubItems(6) = IIf(IsNull(rsload!AmountPaid), "$ 0.00", Format(rsload!AmountPaid, "Currency"))
                        itmX.SubItems(7) = IIf(IsNull(rsload!AmountCredited), "$ 0.00", Format(rsload!AmountCredited, "Currency"))
                        
                        itmX.SmallIcon = IIf(Val(rsload!Finalised) = 0, "Unfinalised", "Finalised")
                        If rsload!AmountPaid <> 0 And rsload!AmountPaid < rsload!TotalDue Then
                            itmX.SmallIcon = "Partially"
                        End If
                        If rsload!AmountPaid < rsload!TotalDue And DateDiff("s", rsload!PaymentDue, sysnow) > 0 Then itmX.SmallIcon = "Overdue"
    
                        Call csPay.clsGroup.colInvoices.Add("i" & rsload!RecID, CLng(rsload!VirtualID), CLng(rsload!acci_RecID), CLng(rsload!InvoiceID), "", IIf(IsNull(rsload!TotalDue), 0, rsload!TotalDue), IIf(IsNull(rsload!AmountCredited), 0, rsload!AmountCredited), IIf(IsNull(rsload!AmountPaid), 0, rsload!AmountPaid), False, "i" & rsload!RecID)
                    
                    rsload.MoveNext
                Wend
            End If
        End If
    End If

    lblAllocated.Caption = Format(csPay.clsGroup.colInvoices.TotalDebit(False) - csPay.clsGroup.colInvoices.TotalCredit(False) - csPay.clsGroup.colInvoices.TotalPaid(False), "Currency")
    
    If Err.Number = 0 Then Exit Function
    

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

Private Sub cmdAdd_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_Click"
    Const ContainerName = "frmGroup"
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


    Dim ix As Long
    Dim iPayType As Long
    Dim bGetSerial As Boolean
    Dim bGetCC As Boolean
    
    If Val(txtAmt.Text) = 0 Then
        MsgBox "You must enter an amount to add to the transactions list"
        Exit Sub
    End If
    
    If CCur(txtAmt.Text) + csPay.clsPayments.colPayments.TotalAmount > CCur(txtCalc(2).Text) Then
        MsgBox "You cannot go higher than the total amount selected for Debit Payment", vbCritical
        Exit Sub
    End If
    
    For ix = 1 To lvPaymentType.ListItems.Count
        If lvPaymentType.ListItems(ix).Checked = True Then
            If lvPaymentType.ListItems(ix).Text = "Cheque" Or lvPaymentType.ListItems(ix).Text = "Money Order" Then bGetSerial = True
            If lvPaymentType.ListItems(ix).Text = "Credit Card" Then bGetCC = True
            iPayType = ix
            Exit For
        End If
    Next
    
    If iPayType = 0 Then
        MsgBox "You must first select a type of payment."
        Exit Sub
    End If
    
    Dim fSN As New frmSN
    Dim fCC As New frmCC
    
    If bGetSerial = True Then
        fSN.Show 1
    End If
    
    If bGetCC = True Then
        fCC.Show 1
        fSN.Caption = "Manual Transaction Receipt Number"
        fSN.Show 1
    End If
    
    Call csPay.clsPayments.colPayments.Add("p" & csPay.clsPayments.colPayments.Count + 1, Val(Mid(lvPaymentType.ListItems(iPayType).Key, 2)), Val(txtAmt.Text), fSN.Num, fCC.CCID)
    
    txtCalc(0).Text = txtCalc(0).Text + vbCrLf
    txtCalc(0).SelStart = Len(txtCalc(0))
    txtCalc(0).Text = txtCalc(0).Text + Format(Val(txtAmt.Text), "Currency") + " + "
    txtCalc(1).Text = Format(csPay.clsPayments.colPayments.TotalAmount, "Currency")
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub cmdEqual_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdEqual_Click"
    Const ContainerName = "frmGroup"
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


    csPay.clsPayments.colPayments.Clear
    txtCalc(0).Text = ""
    txtCalc(1).Text = Format(csPay.clsPayments.colPayments.TotalAmount, "Currency")
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub cmdExit_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdExit_Click"
    Const ContainerName = "frmGroup"
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

    
    Unload Me
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub cmdmakePayment_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdmakePayment_Click"
    Const ContainerName = "frmGroup"
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


    If CCur(txtCalc(1)) = 0 Then Exit Sub
    
    If Val(txtCalc(1)) < Val(txtCalc(2)) Then
        Select Case MsgBox("The amount entered is less than the amount debited, do you wish me to proceed with the payment, I will do transactions from the furtherest payment date possible?", vbQuestion + vbYesNo, "Amounts do not match!")
        Case vbNo
            Exit Sub
        End Select
    End If
    
    Dim rsTraxr As adodb.Recordset
    
    Dim ix As Long
    Dim SQL As String
    Dim rsSave As adodb.Recordset
    Dim rsInv As adodb.Recordset
    Dim cDiff As Currency
    
    Dim iPayCount As Long
    
    SQL = ""
    For ix = 1 To csPay.clsGroup.colInvoices.Count
        If csPay.clsGroup.colInvoices(ix).Checked = True Then
            SQL = SQL + " invoiceout.TraxrID = " & Mid(csPay.clsGroup.colInvoices(ix).Key, 2)
            SQL = SQL + " OR"
        End If
    Next ix
    
    SQL = Left(SQL, Len(SQL) - 3)
    If MySQL.OpenTable(directConn, rsInv, , "select * from invoiceout where " & SQL + " ORDER BY PaymentDue") = True Then
        If rsInv.EOF And rsInv.BOF Then
        
        Else
            If rsInv.RecordCount > 0 Then
                Do While Not rsInv.EOF And Err.Number = 0
                    If iPayCount + 1 <= csPay.clsPayments.colPayments.Count Then
                        iPayCount = iPayCount + 1
                    Else
                        GoTo exitdo
                    End If
                    
                    Do
                        cDiff = 0
                        If rsInv!AmountPaid <= rsInv!TotalDue Then
                            If csPay.clsPayments.colPayments(iPayCount).Amount >= rsInv!TotalDue - rsInv!AmountPaid Then
                                cDiff = rsInv!TotalDue - rsInv!AmountPaid
                                rsInv!AmountPaid = rsInv!AmountPaid + cDiff
                                csPay.clsPayments.colPayments(iPayCount).Amount = csPay.clsPayments.colPayments(iPayCount).Amount - cDiff
                            Else
                                rsInv!AmountPaid = rsInv!AmountPaid + csPay.clsPayments.colPayments(iPayCount).Amount
                                cDiff = csPay.clsPayments.colPayments(iPayCount).Amount
                                csPay.clsPayments.colPayments(iPayCount).Amount = 0
                            End If
                            
                            bResult = MySQL.OpenTable(directConn, rsSave, , "select * from invout_payment Limit 1")
                            
                            rsSave.AddNew
                            rsSave!InvOut_RecID = rsInv!RecID
                            rsSave!Amount = cDiff
                            
                            rsSave!GST = cDiff * oTax(Login.TaxCode, Login.TaxCountry)
                            
                            rsSave!TotalPaid = cDiff
                            rsSave!acci_RecID = rsInv!acci_RecID
                            rsSave!WhenPaid = sysnow
                            rsSave.Update
                            rsInv!PaidWhen = sysnow
                            
                            rsInv.Update
                            
                            For ix = 1 To lvPaymentType.ListItems.Count
                                If lvPaymentType.ListItems(ix).Checked = True Then
                                    If Left(lvPaymentType.ListItems(ix).Key, 1) = "c" Then
                                        MySQL.AddReceiptItem directConn, rsInv!acci_RecID, rsInv!RecID, , , , cDiff, , lvPaymentType.ListItems(ix).Text, csPay.clsPayments.colPayments(iPayCount).SerialNumber, Val(Mid(csPay.clsGroup.colInvoices(iPayCount).Key, 2))
                                    Else
                                        MySQL.AddReceiptItem directConn, rsInv!acci_RecID, rsInv!RecID, , , , cDiff, , lvPaymentType.ListItems(ix).Text, csPay.clsPayments.colPayments(iPayCount).SerialNumber, Val(Mid(csPay.clsGroup.colInvoices(iPayCount).Key, 2))
                                    End If
                                End If
                            Next
                            
                            If rsInv!TraxrID <> 0 Then
                                bResult = MySQL.OpenTable(directConn, rsTraxr, , "select * from invoicetraxr Where RecID = " & rsInv!TraxrID)
                                If rsTraxr.RecordCount > 0 Then
                                    MySQL.Execute directConn, "UPDATE invoicetraxr Set AmountPaid=AmountPaid+" & IIf(rsTraxr!AmountPaid = 0 And rsInv!AmountPaid <> 0, rsInv!AmountPaid, cDiff) & " where RecID = " & rsInv!TraxrID
                                    bResult = MySQL.OpenTable(directConn, rsTraxr, , "select * from invoicetraxr Where RecID = " & rsInv!TraxrID)
                                    If rsTraxr!AmountPaid >= rsTraxr!TotalDue Then
                                        MySQL.Execute directConn, "UPDATE invoicetraxr Set Finalised = -1 where RecID = " & rsInv!TraxrID
                                    End If
                                End If
                            End If
                        Else
                            If rsInv!TraxrID <> 0 Then
                                bResult = MySQL.OpenTable(directConn, rsTraxr, , "select * from invoicetraxr Where RecID = " & rsInv!TraxrID)
                                If rsTraxr.RecordCount > 0 Then
                                    rsTraxr!AmountPaid = rsTraxr!AmountPaid + cDiff
                                    If rsTraxr!AmountPaid >= rsTraxr!TotalDue Then rsTraxr!Finalised = True
                                    rsTraxr.Update
                                End If
                            End If

                        End If
                        If csPay.clsPayments.colPayments(iPayCount).Amount = 0 Then Exit Do
                        If rsInv!AmountPaid >= rsInv!TotalDue Then rsInv.MoveNext
                    Loop Until rsInv.EOF Or Err.Number <> 0
                Loop
exitdo:
            End If
        End If
    End If
    
    Unload Me
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
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
    Const ContainerName = "frmGroup"
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


    Set csPay = New clsPay
    
    Call GUI.LoadColWidths(lvInvoices, Me)
    
    Call LoadFlags
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmGroup"
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


    Call GUI.SaveColWidths(lvInvoices, Me)
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub lvInvoices_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvInvoices_ColumnClick"
    Const ContainerName = "frmGroup"
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


    Call GUI.ColumnSort(ColumnHeader, lvInvoices)
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub lvInvoices_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvInvoices_ItemCheck"
    Const ContainerName = "frmGroup"
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


    csPay.clsGroup.colInvoices(Item.Key).Checked = Item.Checked
    
    txtCalc(2).Text = Format(csPay.clsGroup.colInvoices.TotalDebit(True) - csPay.clsGroup.colInvoices.TotalCredit(True) - csPay.clsGroup.colInvoices.TotalPaid(True), "Currency")
    lblAllocated.Caption = Format(csPay.clsGroup.colInvoices.TotalDebit(False) - csPay.clsGroup.colInvoices.TotalCredit(False) - csPay.clsGroup.colInvoices.TotalPaid(False), "Currency")
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub lvpaymenttype_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvpaymenttype_ItemCheck"
    Const ContainerName = "frmGroup"
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


    Dim ix As Long
    For ix = 1 To lvPaymentType.ListItems.Count
        If Item.Key = lvPaymentType.ListItems(ix).Key Then
        
        Else
            lvPaymentType.ListItems(ix).Checked = False
        End If
    Next
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtAmt_KeyPress"
    Const ContainerName = "frmGroup"
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


    Select Case KeyAscii
    Case 8, 48 To 57
    Case Asc(".")
        If InStr(txtAmt.Text, ".") > 0 Then KeyAscii = 0
    Case Else
        KeyAscii = 0
    End Select
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub
