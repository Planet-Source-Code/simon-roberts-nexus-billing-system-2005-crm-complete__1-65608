Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Global MySQL As New clsMySQL
Global oErr As New clsErrors

Sub cDebug(txt As String)

    Debug.Print txt
    
End Sub

Public Sub ShellLauncher()

    Dim Path As String, Extension As String
    
    Path = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    
            SaveSetting "projectalpha", "db", "ConnectionString", ""
    
    If Dir(Path + "dblauncher.exe", vbNormal) <> "" Then
            
            
            ShellExecute frmAgent.hwnd, vbNullString, Path + "dblauncher.exe " + Extension, vbNullString, "C:\", SW_SHOWNORMAL
            
    End If
    
            
End Sub
