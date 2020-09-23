VERSION 5.00
Begin VB.Form frmEXPpay 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Expendure Payment"
   ClientHeight    =   6945
   ClientLeft      =   2655
   ClientTop       =   4335
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "frmEXPpay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEXPpay.frx":0ECA
   ScaleHeight     =   6945
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame txtPayleft 
      BackColor       =   &H00000000&
      Caption         =   "Amount Left to Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3300
      TabIndex        =   16
      Top             =   5310
      Width           =   4095
      Begin VB.TextBox txtAmountLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0968F&
         Height          =   525
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   31
         TabIndex        =   17
         Tag             =   "AmountPaid"
         Text            =   "0.00"
         Top             =   240
         Width           =   3915
      End
   End
   Begin VB.Frame frmFields 
      BackColor       =   &H00000000&
      Caption         =   "Amount Paid Today"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   7
      Left            =   7680
      TabIndex        =   14
      Top             =   5310
      Width           =   3975
      Begin VB.TextBox frmField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0968F&
         Height          =   525
         Index           =   3
         Left            =   90
         MaxLength       =   31
         TabIndex        =   15
         Tag             =   "AmountPaid"
         Text            =   "0.00"
         Top             =   240
         Width           =   3795
      End
   End
   Begin VB.Frame frmFields 
      BackColor       =   &H00000000&
      Caption         =   "Payment Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Index           =   3
      Left            =   3270
      TabIndex        =   6
      Top             =   3540
      Width           =   8385
      Begin VB.Frame frmFields 
         BackColor       =   &H00000000&
         Caption         =   "Name of Bodyment Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Index           =   6
         Left            =   5070
         TabIndex        =   12
         Top             =   840
         Width           =   3165
         Begin VB.TextBox txtpaytxt 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   150
            MaxLength       =   86
            TabIndex        =   13
            Tag             =   "N"
            Top             =   240
            Width           =   2925
         End
      End
      Begin VB.Frame frmFields 
         BackColor       =   &H00000000&
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Index           =   5
         Left            =   2640
         TabIndex        =   10
         Top             =   840
         Width           =   2325
         Begin VB.TextBox txtpaytxt 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   150
            MaxLength       =   42
            TabIndex        =   11
            Tag             =   "A"
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame frmFields 
         BackColor       =   &H00000000&
         Caption         =   "BSB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Index           =   4
         Left            =   210
         TabIndex        =   8
         Top             =   840
         Width           =   2325
         Begin VB.TextBox txtpaytxt 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   150
            MaxLength       =   42
            TabIndex        =   9
            Tag             =   "B"
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.ComboBox cmbMethod 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   330
         Width           =   8025
      End
   End
   Begin VB.Frame frmFields 
      BackColor       =   &H00000000&
      Caption         =   "Serial Number of Payment   (ie. Bank Cheque No)  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   2
      Left            =   3270
      TabIndex        =   4
      Top             =   2790
      Width           =   8385
      Begin VB.TextBox frmField 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   90
         MaxLength       =   64
         TabIndex        =   5
         Tag             =   "SerialNo"
         Top             =   240
         Width           =   8205
      End
   End
   Begin VB.Frame frmFields 
      BackColor       =   &H00000000&
      Caption         =   "Paid To Name or Details of Correesponding Reciepiant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Index           =   1
      Left            =   3270
      TabIndex        =   2
      Top             =   1560
      Width           =   8385
      Begin VB.TextBox frmField 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   1
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "PaidTo"
         Top             =   240
         Width           =   8205
      End
   End
   Begin VB.Frame frmFields 
      BackColor       =   &H00000000&
      Caption         =   "Transaction Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Index           =   0
      Left            =   3270
      TabIndex        =   0
      Top             =   120
      Width           =   8385
      Begin VB.TextBox frmField 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1035
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   1
         Tag             =   "TransactionText"
         Top             =   240
         Width           =   8205
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0066B0DD&
      BorderWidth     =   3
      Height          =   345
      Index           =   1
      Left            =   3330
      Shape           =   4  'Rounded Rectangle
      Top             =   6420
      Width           =   2205
   End
   Begin VB.Label lblBttn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0043D143&
      Height          =   255
      Index           =   1
      Left            =   3510
      TabIndex        =   19
      Top             =   6450
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0066B0DD&
      BorderWidth     =   3
      Height          =   345
      Index           =   0
      Left            =   9390
      Shape           =   4  'Rounded Rectangle
      Top             =   6420
      Width           =   2205
   End
   Begin VB.Label lblBttn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make &Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0043D143&
      Height          =   255
      Index           =   0
      Left            =   9570
      TabIndex        =   18
      Top             =   6450
      Width           =   1875
   End
End
Attribute VB_Name = "frmEXPpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AmountLeft As Currency
Public AmountPaid As Currency
Public InBookID As Double
Public Cancel As Boolean
Public RecID As Double

Private Sub cmbMethod_Change()

    cmbMethod.Tag = Trim(Mid(cmbMethod.Text, InStr(cmbMethod.Text, "[") + 1, InStr(cmbMethod.Text, "]") - InStr(cmbMethod.Text, "[") - 2))
    
    
    If InStr(cmbMethod.Tag, "B") > 0 Then
        txtpaytxt(0).Locked = False
        txtpaytxt(0).Text = ""
    Else
        txtpaytxt(0).Locked = True
        txtpaytxt(0).Text = ""
    End If
    
    If InStr(cmbMethod.Tag, "A") > 0 Then
        txtpaytxt(1).Locked = False
        txtpaytxt(1).Text = ""
    Else
        txtpaytxt(1).Locked = True
        txtpaytxt(1).Text = ""
    End If
    
    If InStr(cmbMethod.Tag, "N") > 0 Then
        txtpaytxt(2).Locked = False
        txtpaytxt(2).Text = ""
    Else
        txtpaytxt(2).Locked = True
        txtpaytxt(2).Text = ""
    End If
    
End Sub

Private Sub cmbMethod_Click()

    Call cmbMethod_Change

End Sub

Private Sub Form_Load()

    Dim rsLoad As ADODB.Recordset
    
    Call MySQL.OpenTable(ADOConn, rsLoad, , "select * from paymenttype where CreditCard = 0")
    If rsLoad.State = adStateOpen Then
        If rsLoad.RecordCount > 0 Then
            Do
                cmbMethod.AddItem rsLoad!Description & " [ " & IIf(Val(rsLoad!HasBSB) <> 0, "B ", "") & IIf(Val(rsLoad!HasSerial) <> 0, "S ", "") & IIf(Val(rsLoad!HasAcc) <> 0, "A ", "") & IIf(Val(rsLoad!HasName) <> 0, "N ", "") & "]"
                cmbMethod.ItemData(cmbMethod.ListCount - 1) = rsLoad!RecID
                rsLoad.MoveNext
            Loop Until rsLoad.EOF Or Err.Number <> 0
        End If
        rsLoad.Close
    End If
    
    txtAmountLeft = Format(Me.AmountLeft, "Currency")
End Sub

Private Sub frmField_Change(Index As Integer)

    If frmField(Index).Tag = "AmountPaid" Then
    
            If Val(frmField(Index).Text) > Me.AmountLeft Then frmField(Index).Text = "" & Me.AmountLeft
        txtAmountLeft.Text = Format(Me.AmountLeft - Val(frmField(Index).Text), "Currency")

    End If
End Sub

Private Sub frmField_KeyPress(Index As Integer, KeyAscii As Integer)

    If frmField(Index).Tag = "AmountPaid" Then
    
        Select Case KeyAscii
        Case 8, Asc("0") To Asc("9")
        Case Asc(".")
            If InStr(frmField(Index).Text, ".") > 0 Then KeyAscii = 0
        Case Else
            keyacii = 0
        End Select
    
            If Val(frmField(Index).Text) > Me.AmountLeft Then frmField(Index).Text = "" & Me.AmountLeft
        txtAmountLeft.Text = Format(Me.AmountLeft - Val(frmField(Index).Text), "Currency")

    End If

    

End Sub

Private Sub lblBttn_Click(Index As Integer)

    If Index <> 0 Then
        Me.Cancel = True
        Unload Me
        Exit Sub
    End If
    
    
    If cmbMethod.ListIndex = -1 Then
        MsgBox "You must select a payment method and complete the fields."
        cmbMethod.SetFocus
        Exit Sub
    End If
    
    If InStr(cmbMethod.Tag, "S") > 0 Then
        If Trim(frmField(2).Text) = "" Then
            MsgBox "You must enter the serial number of the payment for this method."
            frmField(2).SetFocus
            Exit Sub
        End If
    End If
    
    If InStr(cmbMethod.Tag, "N") > 0 Then
        If Trim(txtpaytxt(2).Text) = "" Then
            MsgBox "You must enter the name of the account that the payment was made into."
            txtpaytxt(2).SetFocus
            Exit Sub
        End If
    End If
    
    If InStr(cmbMethod.Tag, "A") > 0 Then
        If Trim(txtpaytxt(2).Text) = "" Then
            MsgBox "You must enter the Account Number of the account that the payment was made into."
            txtpaytxt(2).SetFocus
            Exit Sub
        End If
    End If
        
    If InStr(cmbMethod.Tag, "B") > 0 Then
        If Trim(txtpaytxt(0).Text) = "" Then
            MsgBox "You must enter the BSB or department code that is required for this payment to be correctly addressed."
            txtpaytxt(0).SetFocus
            Exit Sub
        End If
    End If
        
    On Error Resume Next
        
        Do
            Err.Clear
            Me.RecID = MySQL.GetTMPRecID("exp_flags", ADOConn)
            MySQL.Execute ADOConn, "insert into exp_flags (RecID, InBookID, TransactionText, PaidTo, SerialNo, PaymentMethod, AmountPaid)" + _
                                    "Values ('" & Me.RecID & "','" & Me.InBookID & "','" & MySQL.ESC(frmField(0).Text) & "','" & MySQL.ESC(frmField(1).Text) & "','" & MySQL.ESC(frmField(2).Text) & "','" & cmbMethod.ItemData(cmbMethod.ListIndex) & "','" & Val(frmField(3).Text) & "')", False
        Loop Until Err.Number = 0
    
       Me.AmountPaid = Val(frmField(3).Text)
       
    Unload Me
    
End Sub

Private Sub txtpaytxt_Change(Index As Integer)

    Dim st As String
    
    st = st + "Pay By: " & Left(cmbMethod.Text, InStr(cmbMethod.Text, "[") - 1)
    
    If txtpaytxt(2).Locked = False Then
        st = st + vbCrLf + "Pay To: " & txtpaytxt(2).Text
    End If
    
    If txtpaytxt(0).Locked = False Then
        st = st + vbCrLf + "BSB/Code: " & txtpaytxt(0).Text
    End If

    If txtpaytxt(1).Locked = False Then
        st = st + IIf(InStr(st, "BSB") > 0, "  -  ", vbCrLf) + "Account: " & txtpaytxt(1).Text
    End If
    
    If frmField(0) <> st Then frmField(0).Text = st
    
End Sub
