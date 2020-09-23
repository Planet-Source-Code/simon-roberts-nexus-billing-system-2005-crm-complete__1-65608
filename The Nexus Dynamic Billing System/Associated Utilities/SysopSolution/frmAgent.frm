VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmAgent 
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin AgentObjectsCtl.Agent Agent 
      Index           =   0
      Left            =   480
      Top             =   90
   End
End
Attribute VB_Name = "frmAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sChar() As String
Public oChar As IAgentCtlCharacter
Dim IDAgent

Public Function LoadChar(ASCFile As String, X As Single, Y As Single)

'    On Error Resume Next
    
       
    'ReDim Preserve sChar(Agent.UBound)
    'ReDim Preserve oChar(oChar.UBound + 1)
    'sChar(Agent.UBound) = ASCFile
    
    IDAgent = Agent(Agent.UBound).Characters.Load("CharacterID", ASCFile)
    
    Set oChar = Agent(Agent.UBound).Characters("CharacterID")
    oChar.MoveTo X, Y
    'oChar.Show
    
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    
    'oChar.StopAll
    'oChar.Hide


End Sub

