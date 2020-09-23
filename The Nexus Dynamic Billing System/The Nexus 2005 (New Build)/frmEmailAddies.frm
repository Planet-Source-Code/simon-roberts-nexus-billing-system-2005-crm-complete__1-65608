VERSION 5.00
Begin VB.Form frmEmailAddies 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter in the email addresses in the fields below."
   ClientHeight    =   4230
   ClientLeft      =   3255
   ClientTop       =   7050
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel Sending"
      Height          =   345
      Index           =   1
      Left            =   5250
      TabIndex        =   7
      Top             =   3810
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send Email"
      Height          =   345
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   3810
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Blind Carbon Copy [BCC]: (ie someone@ep.net.au)"
      Height          =   1155
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   2580
      Width           =   7155
      Begin VB.TextBox Text1 
         Height          =   795
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carbon Copy [CC]: (ie someone@ep.net.au)"
      Height          =   1155
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   1350
      Width           =   7155
      Begin VB.TextBox Text1 
         Height          =   795
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "To: (ie someone@ep.net.au)"
      Height          =   1155
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7155
      Begin VB.TextBox Text1 
         Height          =   795
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frmEmailAddies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ButPush As Integer
Public sTo As String
Public sCC As String
Public sBC As String

Private Sub Command1_Click(Index As Integer)

    ButPush = Index
    
    Select Case Index
    Case 0
    
        If Trim(Text1(0)) <> "" Or Trim(Text1(1)) <> "" Or Trim(Text1(2)) <> "" Then
            Me.sBC = Trim(Text1(2).Text)
            Me.sCC = Trim(Text1(1).Text)
            Me.sTo = Trim(Text1(0).Text)
        
            Unload Me
        End If
        
    Case 1
    
        Unload Me
        
    End Select
    
End Sub

Private Sub Form_Load()

    ButPush = 1
    
    Text1(2).Text = Me.sBC
    Text1(1).Text = Me.sCC
    Text1(0).Text = Me.sTo
    
End Sub
