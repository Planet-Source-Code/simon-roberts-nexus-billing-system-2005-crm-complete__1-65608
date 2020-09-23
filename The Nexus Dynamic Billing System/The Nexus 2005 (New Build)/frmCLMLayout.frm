VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCLMLayout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Column Layouts"
   ClientHeight    =   10875
   ClientLeft      =   2250
   ClientTop       =   2460
   ClientWidth     =   12915
   Icon            =   "frmCLMLayout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   12915
   Begin MSComctlLib.ListView lvColumn 
      Height          =   7275
      Left            =   120
      TabIndex        =   1
      Top             =   3510
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   12832
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "!Description!"
         Text            =   "Description/Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "!ServiceKey!"
         Text            =   "Service Key"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "!FieldName!"
         Text            =   "Fieldname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "!sFormat!"
         Text            =   "Formating On nResult"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "!Width!"
         Text            =   "Width"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "!SECLevel!"
         Text            =   "Security Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "!formcode!"
         Text            =   "formcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "!cOrder!"
         Text            =   "Order"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Column Code for Listview"
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   12675
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   10260
         TabIndex        =   20
         Top             =   2700
         Width           =   2300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   10260
         TabIndex        =   19
         Top             =   1680
         Width           =   2300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   10260
         TabIndex        =   18
         Top             =   2190
         Width           =   2300
      End
      Begin VB.Frame Frame2 
         Caption         =   "Security Level"
         Height          =   675
         Index           =   7
         Left            =   6060
         TabIndex        =   15
         Top             =   2430
         Width           =   1455
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   16
            Tag             =   "SECLevel"
            Top             =   210
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Width"
         Height          =   675
         Index           =   6
         Left            =   6030
         TabIndex        =   13
         Top             =   1710
         Width           =   2565
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Tag             =   "width"
            Top             =   210
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Order"
         Height          =   675
         Index           =   5
         Left            =   7590
         TabIndex        =   11
         Top             =   2430
         Width           =   1005
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   12
            Tag             =   "cOrder"
            Top             =   210
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Format (i.e. Currency or nResult Select Statement)"
         Height          =   1365
         Index           =   4
         Left            =   6030
         TabIndex        =   9
         Top             =   270
         Width           =   6525
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   1005
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Tag             =   "sFormat"
            Top             =   210
            Width           =   6285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Primary Field for query in incoming SQL Statement for this column"
         Height          =   675
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   2430
         Width           =   5745
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   7
            Left            =   90
            TabIndex        =   17
            Tag             =   "FieldName"
            Top             =   240
            Width           =   5505
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Formcode"
         Height          =   675
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   1710
         Width           =   5745
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Tag             =   "formcode"
            Top             =   210
            Width           =   5505
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Service Key"
         Height          =   675
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   990
         Width           =   5745
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Tag             =   "ServiceKey"
            Top             =   210
            Width           =   5505
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Column Title/Description"
         Height          =   675
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Width           =   5745
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Tag             =   "Description"
            Top             =   210
            Width           =   5505
         End
      End
   End
End
Attribute VB_Name = "frmclmLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

    If odb.colDBObjects.Count > 0 Then
    
        Combo2.Clear
        
        Dim i As Long
        
        For i = 1 To odb.colDBObjects.Count
            If Combo1.List(Combo1.ListIndex) = odb.colDBObjects(i).Tablename Then
                Combo2.AddItem odb.colDBObjects(i).FieldName
            End If
        Next i
        
        Combo2.ListIndex = 0
    End If

End Sub

Private Sub Combo1_Click()

    Call Combo1_Change
    
End Sub

Private Sub Command1_Click(Index As Integer)

    Dim SQL As String
    Dim sql2 As String
    Dim bx As Byte

    Select Case Index
    Case 1 ' Add
        
        SQL = ""
        For bx = txtField.LBound To txtField.UBound
            sql2 = sql2 + "`" + txtField(bx).Tag + "`,"
            SQL = SQL + "'" + MySQL.ESC(txtField(bx).Text) + "',"
        Next bx
    
        SQL = "iNSERT INTO columnlayout (" & Left(sql2, Len(sql2) - 1) & ") VALUES (" + Left(SQL, Len(SQL) - 1) & ")"
        
    Case 0 ' Save
    
        SQL = ""
        
        For bx = txtField.LBound To txtField.UBound
            SQL = SQL + "`" + txtField(bx).Tag + "` = '" + MySQL.ESC(txtField(bx).Text) + "',"
        Next bx
        
        SQL = "update columnlayout set " + Left(SQL, Len(SQL) - 1) & " where RecID = '" & Mid(lvColumn.SelectedItem.Key, 2) & "'"
        
    Case 2 ' Delete
        
        SQL = "delete from columnlayout where RecID = '" & Mid(lvColumn.SelectedItem.Key, 2) & "'"
        
    End Select
        
        
    Call MySQL.Execute(ADOConn, SQL, False)
    
    
    If Index = 2 Then
        lvColumn.ListItems.Remove lvColumn.SelectedItem.Index
        Exit Sub
        
    End If
        
    
    If Index = 0 Then
    
        SQL = "select * from  columnlayout where RecID = '" & Mid(lvColumn.SelectedItem.Key, 2) & "'"
        
    Else
        SQL = ""
            
        For bx = txtField.LBound To txtField.UBound
            SQL = SQL + "`" + txtField(bx).Tag + "` = '" + MySQL.ESC(txtField(bx).Text) + "' and "
        Next bx
        
        SQL = "select * from  columnlayout where " + Left(SQL, Len(SQL) - 5) & ""
    End If
    
    Dim rsload As ADODB.Recordset
    
    Call MySQL.OpenTable(ADOConn, rsload, , SQL)
    
    Call MySQL.fillLV(ADOConn, rsload, lvColumn, True, IIf(Index = 0, lvColumn.SelectedItem, Nothing))
    
    
End Sub

Private Sub Form_Load()

    If bBigFont = True Then
        
        lvColumn.Font.Size = 16
        
    End If
   

    Call GUI.LoadColWidths(lvColumn, Me)
    
    Dim rsload As ADODB.Recordset
    
    Call MySQL.OpenTable(ADOConn, rsload, , "select * from columnlayout order by formcode, cOrder")
    
    Call MySQL.fillLV(ADOConn, rsload, lvColumn, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call GUI.SaveColWidths(lvColumn, Me)
    
End Sub

Private Sub lvColumn_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim bx As Byte
    Dim by As Byte
    
    For bx = 1 To lvColumn.ColumnHeaders.Count
        For by = txtField.LBound To txtField.UBound
            If InStr(LCase(lvColumn.ColumnHeaders(bx).Tag), LCase(txtField(by).Tag)) > 0 Then
                Select Case bx
                Case 1
                    txtField(by) = Item.Text
                Case Else
                    txtField(by) = Item.SubItems(bx - 1)
                End Select
                Exit For
            End If
        Next by
    Next bx
    
    
End Sub

