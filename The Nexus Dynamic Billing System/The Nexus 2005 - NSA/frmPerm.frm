VERSION 5.00
Begin VB.Form frmPerm 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Permissioning for Sysop Account"
   ClientHeight    =   7335
   ClientLeft      =   870
   ClientTop       =   3945
   ClientWidth     =   9315
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmPerm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPerm.frx":0ECA
   ScaleHeight     =   7335
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpt 
      Caption         =   "&Close and Cancel"
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   46
      Top             =   6300
      Width           =   2025
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "&Save"
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   45
      Top             =   6780
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Maintenance"
      ForeColor       =   &H00FFFFFF&
      Height          =   1425
      Index           =   14
      Left            =   7200
      TabIndex        =   42
      Tag             =   "bMaster"
      Top             =   120
      Width           =   2025
      Begin VB.OptionButton optPerm 
         BackColor       =   &H00BA3F3F&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   29
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   810
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H00BA3F3F&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   28
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   43
         Tag             =   "bMaintain"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0066B0DD&
      Caption         =   "This is a Master Server mode account"
      Height          =   885
      Index           =   13
      Left            =   3720
      TabIndex        =   39
      Tag             =   "bMaster"
      Top             =   6270
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0066B0DD&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   27
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "Master"
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0066B0DD&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   26
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0066B0DD&
      Caption         =   "Can View and Create Templates"
      Height          =   885
      Index           =   12
      Left            =   3720
      TabIndex        =   36
      Tag             =   "bMaster"
      Top             =   5250
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0066B0DD&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   25
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0066B0DD&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   24
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "bTemplates"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Can View and Create ViSP's"
      Height          =   885
      Index           =   11
      Left            =   3720
      TabIndex        =   33
      Top             =   4230
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   23
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "bVISP"
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   22
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Can View Vendor Information"
      Height          =   885
      Index           =   10
      Left            =   3720
      TabIndex        =   30
      Top             =   3210
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   21
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   20
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "bVendors"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Can View and Explore Customer Base"
      Height          =   885
      Index           =   9
      Left            =   3720
      TabIndex        =   27
      Top             =   2190
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   19
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "bHoldings"
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   18
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Extra Reporing on ViSP Level"
      Height          =   885
      Index           =   8
      Left            =   3720
      TabIndex        =   24
      Top             =   1170
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   17
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "bVISPFiscal"
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   16
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Can enter Account Settings Window"
      Height          =   885
      Index           =   7
      Left            =   3720
      TabIndex        =   21
      Top             =   150
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   15
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   14
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "bAccSettings"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Able to Change Ownership"
      Height          =   885
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   6270
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   13
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "bOwnership"
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   12
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Able to Create Customers"
      Height          =   885
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   5250
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   11
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   10
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "bAddCust"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Access To Refund Systems"
      Height          =   885
      Index           =   4
      Left            =   210
      TabIndex        =   12
      Top             =   4230
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   9
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "bRefund"
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   8
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Access To Commission Systems"
      Height          =   885
      Index           =   3
      Left            =   210
      TabIndex        =   9
      Top             =   3210
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   7
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "bComm"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Access To Expendure Systems"
      Height          =   885
      Index           =   2
      Left            =   210
      TabIndex        =   6
      Top             =   2190
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "bExpenditure"
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Access To Invoice Systems"
      Height          =   885
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   1170
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "bInvoice"
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0094BD7B&
      Caption         =   "Access To Recievables"
      Height          =   885
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   3375
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optPerm 
         BackColor       =   &H0094BD7B&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "bRecievables"
         Top             =   300
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecID As Double

Private Sub cmdOpt_Click(Index As Integer)

    Select Case Index
    Case 0
    
        Dim bx As Byte
        
        For bx = optPerm.LBound To optPerm.UBound
        If optPerm(bx).Tag <> "" Then
            Select Case optPerm(bx).Value
            Case 1
                Call oMySQL.Execute(oConn, "update sysops set " & optPerm(bx).Tag & " = '-1' where RecID = '" & Me.RecID & "'")
            Case 0
                Call oMySQL.Execute(oConn, "update sysops set " & optPerm(bx).Tag & " = '0' where RecID = '" & Me.RecID & "'")
            End Select
        End If
        Next

    
    Case 1
    
        
        
        
    End Select
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim rsLoad As ADODB.Recordset
    
    Dim bx As Byte
    
    For bx = Frame1.LBound To Frame1.UBound
        If Frame1(bx).Tag = "bMaster" Then
            Frame1(bx).Enabled = Login.bMaster
        End If
    Next
    
    
    Call oMySQL.OpenTable(oConn, rsLoad, , "select * from sysops where RecID = '" & Me.RecID & "'")
    
    If rsLoad.State = adStateOpen Then
        If rsLoad.RecordCount > 0 Then
            For bx = optPerm.LBound To optPerm.UBound
                If optPerm(bx).Tag <> "" Then
                    If Val(rsLoad(optPerm(bx).Tag)) <> 0 Then
                        optPerm(bx).Value = True
                    Else
                        optPerm(bx).Value = False
                    End If
                End If
            Next
        
        End If
    End If
        
    
End Sub
