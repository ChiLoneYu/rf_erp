VERSION 5.00
Begin VB.Form FrmFiledEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据表字段长度修改"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "FrmFiledEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3570
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Field_New_Len 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1485
      TabIndex        =   2
      Top             =   1380
      Width           =   1905
   End
   Begin VB.TextBox Field_Old_Len 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1485
      TabIndex        =   1
      Top             =   900
      Width           =   1905
   End
   Begin VB.TextBox Field_No 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1485
      TabIndex        =   0
      Top             =   420
      Width           =   1905
   End
   Begin VB.CommandButton cmd_quit 
      Caption         =   "退出"
      Height          =   315
      Left            =   2010
      TabIndex        =   4
      Top             =   2130
      Width           =   930
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "确定"
      Height          =   315
      Left            =   660
      TabIndex        =   3
      Top             =   2130
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "字段编号新长度:"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1395
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "字段编号原长度:"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   930
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "字段编号:"
      Height          =   180
      Left            =   660
      TabIndex        =   5
      Top             =   450
      Width           =   810
   End
End
Attribute VB_Name = "FrmFiledEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()
Dim w_rs As New ADODB.Recordset
If check_ok Then
    w_rs.Open "select * from sysobjects where xtype='u' ", G_Con
    Do Until w_rs.EOF
        If checkcolumnintable(w_rs!Name, Trim(Field_No.Text)) Then
            G_Con.Execute "ALTER TABLE " & w_rs!Name & " ALTER COLUMN " & Trim(Field_No.Text) & " nvarchar(" & Val(Field_New_Len.Text) & ") "
        End If
        w_rs.MoveNext
    Loop
    MsgBox "数据修改完毕", vbInformation, "提示信息"
End If
End Sub

Private Sub cmd_quit_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "TDBGrid1" Then
    
    If ActiveControl.Name = "note" Then
        If ActiveControl.MultiLine = False Then
            SendKeys "{TAB}"
        End If
    Else
        SendKeys "{TAB}"
    End If
    Exit Sub
End If

End Sub

Private Function check_ok() As Boolean
 If Val(Field_Old_Len.Text) >= Val(Field_New_Len.Text) Then
    MsgBox "新长度不能小于或等于原长度", vbInformation, "提示信息"
    Field_New_Len.SetFocus
    Field_New_Len.SelStart = 0
    Field_New_Len.SelLength = Len(Field_New_Len.Text)
    check_ok = False
    Exit Function
 End If
 check_ok = True
End Function

Private Function checkcolumnintable(ByVal stablename As String, ByVal sfieldname As String) As Boolean
    Dim RS As New ADODB.Recordset
    
    Set RS = G_Con.Execute("sp_columns @table_name='" & stablename & "',@column_name='" & sfieldname & "' ")
    checkcolumnintable = Not (RS.EOF And RS.BOF)
    RS.Close
End Function

