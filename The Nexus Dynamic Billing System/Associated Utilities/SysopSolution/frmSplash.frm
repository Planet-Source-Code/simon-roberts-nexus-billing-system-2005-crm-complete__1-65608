VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project Alpha - Sysops"
   ClientHeight    =   4920
   ClientLeft      =   4275
   ClientTop       =   5205
   ClientWidth     =   7515
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0442
   ScaleHeight     =   4920
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   -60
      TabIndex        =   0
      Top             =   4650
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public bFinished As Boolean


Private Sub Form_Load()

    pb.Max = 2
    
        
    Show
    
    
    Me.Caption = "Project Alpha - Connecting to oMySQL Server"
    
    If InStr(Command, "/testbench") > 0 Then
        Call oMySQL.Connection("projectalpha", "localhost", "pa2004", "p0st41", oConn)
    Else
    
        Call oMySQL.Connection(, , , , oConn)
    End If
    
    gSleep
    pb.Value = 1
    DirName = GetWindowsDir()
    
    
    frmLogin.Show 1
    pb.Value = 2
    frmMain.Show
    Unload Me
    
    
End Sub

Function GetWindowsDir() As String
    Dim Temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    Temp = String$(MAX_LENGTH, 0)
    Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, Ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        GetWindowsDir = Temp & "\"
    Else
        GetWindowsDir = Temp
    End If
End Function
