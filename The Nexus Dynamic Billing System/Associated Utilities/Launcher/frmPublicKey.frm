VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPublicKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generating Key"
   ClientHeight    =   720
   ClientLeft      =   2025
   ClientTop       =   3015
   ClientWidth     =   6585
   ControlBox      =   0   'False
   Icon            =   "frmPublicKey.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3930
      Top             =   180
   End
   Begin VB.ListBox lstB 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   3240
      TabIndex        =   2
      Top             =   750
      Width           =   3165
   End
   Begin VB.ListBox lstA 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   150
      TabIndex        =   1
      Top             =   750
      Width           =   2835
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmPublicKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KeyName As String
Public PrimaryKey As String
Dim bPop As Boolean

Private Sub Form_Paint()

    
    
    If bPop = False Then
    
        bPop = True
        
        Dim bChar As Byte
        
        For bChar = 48 To 57
            lstA.AddItem IIf(Len("" & bChar) < 3, String(3 - Len("" & bChar), "0") & bChar, "" & bChar) & " = '" & Chr$(bChar) & "'"
            lstA.ItemData(lstA.ListCount - 1) = bChar
        Next
        
        For bChar = 65 To 90
            lstA.AddItem IIf(Len("" & bChar) < 3, String(3 - Len("" & bChar), "0") & bChar, "" & bChar) & " = '" & Chr$(bChar) & "'"
            lstA.ItemData(lstA.ListCount - 1) = bChar
        Next
        
        For bChar = 97 To 122
            lstA.AddItem IIf(Len("" & bChar) < 3, String(3 - Len("" & bChar), "0") & bChar, "" & bChar) & " = '" & Chr$(bChar) & "'"
            lstA.ItemData(lstA.ListCount - 1) = bChar
        Next
                
        For bChar = 97 To 122
            lstA.AddItem IIf(Len("" & bChar) < 3, String(3 - Len("" & bChar), "0") & bChar, "" & bChar) & " = '" & Chr$(bChar) & "'"
            lstA.ItemData(lstA.ListCount - 1) = bChar
        Next

        For bChar = 161 To 172
            lstB.AddItem IIf(Len("" & bChar) < 3, String(3 - Len("" & bChar), "0") & bChar, "" & bChar) & " = '" & Chr$(bChar) & "'"
            lstB.ItemData(lstB.ListCount - 1) = bChar
        Next
        
        For bChar = 174 To 253
            lstB.AddItem IIf(Len("" & bChar) < 3, String(3 - Len("" & bChar), "0") & bChar, "" & bChar) & " = '" & Chr$(bChar) & "'"
            lstB.ItemData(lstB.ListCount - 1) = bChar
        Next
        
        Timer1.Enabled = True
        
    End If
    
End Sub

Private Sub Timer1_Timer()

    Static ComboNo As Integer
    Static NumPasses As Integer
    
    If ComboNo = 0 Then
        
        pb.Max = lstA.ListCount
            
    End If
    
    ComboNo = ComboNo + 1
    
    Randomize Now / ComboNo
    
    If lstA.ListCount = 0 Then
        
        Timer1.Enabled = False
        
        Me.Caption = "Testing Key - ASCII Conversion"
        Dim hs As String
        Dim bChar As Integer
        For bChar = 48 To 57
            hs = hs & Chr(bChar)
        Next
        
        For bChar = 65 To 90
            hs = hs & Chr(bChar)
        Next
        
        For bChar = 97 To 122
            hs = hs & Chr(bChar)
        Next
                
        For bChar = 97 To 122
            hs = hs & Chr(bChar)
        Next
        
        
'123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz

'Îö¿Á4íªÄËîæ®£·Ò¬¦ÖûÑ©´ØÙìÞÍ¹éôñ¯è¢ÌðãÇÆë¸½ÅÈ«Ð¶×²àÜ¥á÷øòåÀ¡óçÉãÇÆë¸½ÅÈ«Ð¶×²àÜ¥á÷øòåÀ¡óçÉ

        If Crypt(Crypt(hs, True, KeyName, Me.PrimaryKey), False, KeyName, Me.PrimaryKey) = hs Then
            Me.Caption = "Testing Key - ASCII Test Pass"
            Unload Me
            Exit Sub
        Else
            NumPasses = NumPasses + 1
            Me.Caption = "Generating Key - Pass/Attempt: " & (NumPasses + 1)
            bPop = False
            Call Form_Paint
            pb.Value = 0
            ComboNo = 0
        End If
        
    End If
    
    Dim iPosa As Integer
    Dim iPosb As Integer
    
    iPosa = Round((lstA.ListCount - 1) * Rnd)
    iPosb = Round((lstB.ListCount - 1) * Rnd)
    
    SaveSetting Me.PrimaryKey, KeyName, "A-" & ComboNo - 1, lstA.ItemData(iPosa)
    SaveSetting Me.PrimaryKey, KeyName, "B-" & ComboNo - 1, lstB.ItemData(iPosb)
    
    lstA.RemoveItem iPosa
    lstB.RemoveItem iPosb
    
    If pb.Value + 1 < pb.Max Then pb.Value = pb.Value + 1
    
End Sub
