VERSION 5.00
Begin VB.Form frmSMTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMTP Server Settings - Local Configuration"
   ClientHeight    =   4275
   ClientLeft      =   2160
   ClientTop       =   6810
   ClientWidth     =   8355
   Icon            =   "frmSMTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Email Domain Name (ie. ep.net.au)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1950
      Width           =   8145
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   7875
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel && Close"
      Height          =   375
      Index           =   1
      Left            =   6060
      TabIndex        =   12
      Top             =   3810
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save and Close"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3810
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Password for username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   4260
      TabIndex        =   11
      Top             =   2850
      Width           =   3975
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   150
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   330
         Width           =   3675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Username (not normally required)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   150
      TabIndex        =   10
      Top             =   2850
      Width           =   4035
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   3765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reply Email Address (ie. support@ep.net.au)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   8145
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   330
         Width           =   7875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6630
      TabIndex        =   8
      Top             =   90
      Width           =   1635
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   330
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SMTP Server on this network (ie. 202.172.123.25)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6435
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   6105
      End
   End
End
Attribute VB_Name = "frmSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    Select Case Index
    Case 0
        reg.smtpServer = txtField(0).Text
        reg.smtpPort = txtField(1).Text
        reg.ReplyAddress = txtField(2).Text
        reg.smtpUsername = txtField(3).Text
        reg.smtpPassword = txtField(4).Text
        reg.smtpDomain = txtField(5).Text
        reg.smtpSetYet = True
        
        SaveRegistry
    
    End Select
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    txtField(0).Text = reg.smtpServer
    txtField(1).Text = reg.smtpPort
    txtField(2).Text = reg.ReplyAddress
    txtField(3).Text = reg.smtpUsername
    txtField(4).Text = reg.smtpPassword
    txtField(5).Text = reg.smtpDomain
    
End Sub

Private Sub txtField_GotFocus(Index As Integer)

    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
    
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub
