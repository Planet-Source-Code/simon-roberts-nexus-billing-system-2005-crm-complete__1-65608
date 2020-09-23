VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDateTime 
   BackColor       =   &H00BA3F3F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select a Date and Time"
   ClientHeight    =   3645
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   3210
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   3210
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Top             =   2610
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63832066
      CurrentDate     =   38021
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   63832066
      CurrentDate     =   38021
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dDate As Date
Public Cancel As Boolean

Private Sub cmdOk_Click()

    dDate = CDate(Format(MonthView1.Value, "yyyy-mm-dd ") & Format(DTPicker1.Value, "ttttt"))
    Unload Me
    
End Sub

Private Sub Command1_Click()

    Me.Cancel = True
    Unload Me
    
End Sub

Private Sub Form_Load()

    MonthView1.Value = dDate
    DTPicker1.Value = dDate
    
    
End Sub
