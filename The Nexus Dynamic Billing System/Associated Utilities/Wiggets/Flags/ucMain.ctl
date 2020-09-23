VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucMain 
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1920
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucMain.ctx":0000
   Begin VB.CommandButton cmdForm 
      Height          =   480
      Left            =   4200
      Picture         =   "ucMain.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   480
   End
   Begin VB.PictureBox picFlag 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   90
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      ToolTipText     =   "Click here to load the Flag X-Editor"
      Top             =   90
      Width           =   480
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1215
      Left            =   60
      TabIndex        =   2
      Top             =   630
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   2143
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList il16x16 
      Left            =   780
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList il32x32 
      Left            =   1380
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lblFlagDesc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   630
      TabIndex        =   3
      Top             =   60
      Width           =   3480
   End
End
Attribute VB_Name = "ucMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public oCommand As New oFlagsExt


