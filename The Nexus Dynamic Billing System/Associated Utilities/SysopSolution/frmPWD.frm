VERSION 5.00
Begin VB.Form frmPWD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password"
   ClientHeight    =   2565
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Comfirm New Password"
      Height          =   735
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   1710
      Width           =   4155
      Begin VB.TextBox txtPWD 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   240
         Width           =   3945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Password"
      Height          =   735
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   900
      Width           =   4155
      Begin VB.TextBox txtPWD 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   240
         Width           =   3945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Password"
      Height          =   735
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4155
      Begin VB.TextBox txtPWD 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   3945
      End
   End
End
Attribute VB_Name = "frmPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public inPassword As String
Public outPassword As String

Private Sub txtPWD_GotFocus(Index As Integer)

    txtPWD(Index).SelStart = 0
    txtPWD(Index).SelLength = Len(txtPWD(Index).Text)
    
End Sub

Private Sub txtPWD_KeyPress(Index As Integer, KeyAscii As Integer)


   
    Select Case KeyAscii
    Case 13
        If Len(Trim(txtPWD(1).Text)) = 0 Or Len(Trim(txtPWD(2).Text)) = 0 Then
            MsgBox "You must complete the fields required to set this password."
            Exit Sub
        End If
        
        Select Case inPassword
        Case ""
        
                If txtPWD(1).Text = txtPWD(2).Text Then
                
                    outPassword = txtPWD(1).Text
                    Unload Me
                    
                Else
                    
                    MsgBox "Your new password and the confirmation password do not match!", vbCritical, "Password Mismatch"
                        
                End If
                
        Case Else
        
            If Login.bMaster = True Then
            
                If txtPWD(1).Text = txtPWD(2).Text Then
                
                    outPassword = txtPWD(1).Text
                    Unload Me
                    
                Else
                    
                    MsgBox "Your new password and the confirmation password do not match!", vbCritical, "Password Mismatch"
                        
                End If
                
            ElseIf txtPWD(0).Text = inPassword Then
                
                If txtPWD(1).Text = txtPWD(2).Text Then
                
                    outPassword = txtPWD(1).Text
                    Unload Me
                    
                Else
                    
                    MsgBox "Your new password and the confirmation password do not match!", vbCritical, "Password Mismatch"
                        
                End If
            Else
            
                MsgBox "Your old password does not match you current exisiting password!", vbCritical, "Password Mismatch"
            
            End If
        End Select
        
    End Select
    
End Sub
