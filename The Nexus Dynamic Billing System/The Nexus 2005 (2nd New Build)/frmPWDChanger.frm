VERSION 5.00
Begin VB.Form frmPWDChanger 
   BackColor       =   &H00B4E9F8&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4890
   ClientLeft      =   1575
   ClientTop       =   1755
   ClientWidth     =   6645
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPWDChanger.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00B4E9F8&
      Caption         =   "&Verify && Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4260
      Width           =   2265
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00B4E9F8&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4260
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B4E9F8&
      Caption         =   "Comfirm New Password"
      Height          =   825
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   3330
      Width           =   5475
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BackColor       =   &H00B4E9F8&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   150
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   300
         Width           =   5145
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B4E9F8&
      Caption         =   "New Password"
      Height          =   825
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   2400
      Width           =   5475
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BackColor       =   &H00B4E9F8&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   150
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   300
         Width           =   5145
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B4E9F8&
      Caption         =   "Verify Your Current Password"
      Height          =   825
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   1470
      Width           =   5475
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BackColor       =   &H00B4E9F8&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   150
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   300
         Width           =   5145
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H006FAFEA&
      BorderWidth     =   3
      X1              =   630
      X2              =   6750
      Y1              =   4530
      Y2              =   4530
   End
   Begin VB.Line Line1 
      BorderColor     =   &H006FAFEA&
      BorderWidth     =   3
      X1              =   630
      X2              =   630
      Y1              =   1200
      Y2              =   4530
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H006FAFEA&
      BorderWidth     =   3
      Height          =   1125
      Left            =   90
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPWDChanger.frx":030A
      Height          =   855
      Left            =   1290
      TabIndex        =   1
      Top             =   540
      Width           =   5115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00B4E9F8&
      Caption         =   "Password Verify and Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1290
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "frmPWDChanger.frx":0391
      Stretch         =   -1  'True
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmPWDChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PassWRD As String
Public Capt As String

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
    Case 0
        Unload Me
    Case 1
        Dim bPassed As Boolean
        
        If PassWRD = txtField(0) Then
            If txtField(1) = txtField(2) Then
                PassWRD = txtField(1)
                Unload Me
                bPassed = True
            End If
        End If
        
        If bPassed = False Then
            
            If PassWRD = txtField(0) Then
                If txtField(1) <> txtField(2) Then
                    MsgBox "The new password you have entered does not match with the comfirm new password. Please revise.", vbCritical, "New Password did not verify"
                End If
            Else
                MsgBox "You have entered in an incorrect password.", vbCritical, "Old Password did not match."
            End If
        
        End If
        
    End Select
End Sub

Private Sub Form_Load()

    If PassWRD = "" Then
        txtField(0).Enabled = False
        txtField(0).Locked = True
    End If
    
    Me.Caption = Capt
        
End Sub
