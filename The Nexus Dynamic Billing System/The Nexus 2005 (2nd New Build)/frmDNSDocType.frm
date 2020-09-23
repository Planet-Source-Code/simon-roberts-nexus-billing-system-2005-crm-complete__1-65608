VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDNSDocType 
   BackColor       =   &H00BA3F3F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create DNS Document"
   ClientHeight    =   6705
   ClientLeft      =   4725
   ClientTop       =   4500
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084E8E8&
      Height          =   1965
      Left            =   150
      TabIndex        =   8
      Top             =   4110
      Width           =   3945
      Begin VB.PictureBox picIcon 
         BackColor       =   &H00BA3F3F&
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   120
         ScaleHeight     =   1605
         ScaleWidth      =   3435
         TabIndex        =   9
         Top             =   240
         Width           =   3435
         Begin VB.Image imgIcon 
            BorderStyle     =   1  'Fixed Single
            Height          =   540
            Index           =   0
            Left            =   60
            Picture         =   "frmDNSDocType.frx":0000
            Top             =   60
            Width           =   540
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00BA3F3F&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   90
         ScaleHeight     =   1695
         ScaleWidth      =   3765
         TabIndex        =   10
         Top             =   210
         Width           =   3765
         Begin VB.VScrollBar vslIcon 
            Height          =   1665
            LargeChange     =   5
            Left            =   3510
            Max             =   100
            TabIndex        =   11
            Top             =   30
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Document Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084E8E8&
      Height          =   1485
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   2520
      Width           =   3945
      Begin VB.TextBox txtDesc 
         Height          =   915
         Left            =   150
         MaxLength       =   255
         TabIndex        =   7
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Document"
      Height          =   375
      Left            =   390
      TabIndex        =   5
      Top             =   6210
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Select Document Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084E8E8&
      Height          =   2325
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3945
      Begin VB.OptionButton optDoc 
         BackColor       =   &H00BA3F3F&
         Caption         =   "Create XML Document"
         ForeColor       =   &H0084E8E8&
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Tag             =   "XML"
         Top             =   1800
         Value           =   -1  'True
         Width           =   3555
      End
      Begin VB.OptionButton optDoc 
         BackColor       =   &H00BA3F3F&
         Caption         =   "Create Rich Text Document"
         Enabled         =   0   'False
         ForeColor       =   &H0084E8E8&
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Tag             =   "RTF"
         Top             =   1350
         Width           =   3585
      End
      Begin VB.OptionButton optDoc 
         BackColor       =   &H00BA3F3F&
         Caption         =   "Create Phone Number Record"
         ForeColor       =   &H0084E8E8&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Tag             =   "PHONE"
         Top             =   900
         Width           =   3555
      End
      Begin VB.OptionButton optDoc 
         BackColor       =   &H00BA3F3F&
         Caption         =   "Create Address Document"
         ForeColor       =   &H0084E8E8&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Tag             =   "ADDRESS"
         Top             =   450
         Width           =   3615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":2B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":2F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":33C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":381A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":3B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":3F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":43D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":482A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":4C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":50CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":5520
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNSDocType.frx":5972
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDNSDocType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sType As String
Public sDesc As String
Public bIcon As Byte


Private Sub Command1_Click()


    Dim Xnt As Integer
    
    For Xnt = optDoc.LBound To optDoc.UBound
        If optDoc(Xnt).Value = True Then
            sType = optDoc(Xnt).Tag
            Exit For
        End If
    Next
    
    For Xnt = imgIcon.LBound To imgIcon.UBound
        If imgIcon(Xnt).BorderStyle = 1 Then
            bIcon = Val(Xnt)
            Exit For
        End If
    Next
    
    sDesc = txtDesc.Text
    
    Unload Me
    
    
End Sub

Private Sub Form_Load()

    Dim I As Integer
    Static Level As Integer
    
    For I = 1 To ImageList1.ListImages.count
        If I = 1 Then
            imgIcon(I - 1).Picture = ImageList1.ListImages(I).ExtractIcon
            imgIcon(I - 1).Move 120, 0
        Else
            Load imgIcon(I - 1)
            Set imgIcon(I - 1).Container = picIcon
            imgIcon(I - 1).Picture = ImageList1.ListImages(I).ExtractIcon
            imgIcon(I - 1).Move imgIcon(0).Left + imgIcon(I - 2).Left + imgIcon(0).Width, Level * (imgIcon(0).Width + 120)
            If imgIcon(I - 1).Left > picIcon.ScaleWidth - imgIcon(I - 1).Width Then
                Level = Level + 1
                imgIcon(I - 1).Move imgIcon(0).Left, Level * (imgIcon(0).Width + 120)
            End If
            imgIcon(I - 1).Visible = True
            imgIcon(I - 1).BorderStyle = 0
        End If
    
    Next I
    
    picIcon.Height = imgIcon(ImageList1.ListImages.count - 1).Top + imgIcon(ImageList1.ListImages.count - 1).Height
    Set picIcon.Container = Picture1
    picIcon.Move 0, 0
    Picture1.Move 90, 240, Frame3.Width - 180, Frame3.Height - 330
End Sub

Private Sub imgIcon_Click(Index As Integer)

    Dim ix As Integer
    
    For ix = imgIcon.LBound To imgIcon.UBound
        imgIcon(ix).BorderStyle = 0
    Next
    
    imgIcon(Index).BorderStyle = 1
    
End Sub

Private Sub vslIcon_Change()


    picIcon.Top = -((vslIcon.Value / 100) * (picIcon.Height - Picture1.Height))
    
    
    
End Sub
