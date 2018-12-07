VERSION 5.00
Begin VB.Form FrmPassEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户密码修改"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "FrmPassEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3570
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox C_New_Pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1380
      Width           =   1380
   End
   Begin VB.TextBox New_Pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   900
      Width           =   1380
   End
   Begin VB.TextBox Old_Pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   420
      Width           =   1380
   End
   Begin VB.CommandButton cmd_quit 
      Caption         =   "退出"
      Height          =   315
      Left            =   1755
      TabIndex        =   4
      Top             =   2130
      Width           =   930
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "确定"
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   2130
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "确定密码:"
      Height          =   180
      Left            =   420
      TabIndex        =   7
      Top             =   1395
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "新密码:"
      Height          =   180
      Left            =   600
      TabIndex        =   6
      Top             =   930
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "旧密码:"
      Height          =   180
      Left            =   600
      TabIndex        =   5
      Top             =   450
      Width           =   630
   End
End
Attribute VB_Name = "FrmPassEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()
If check_ok Then
   G_Con.Execute "update mmst901 set password='" & Trim(C_New_Pass.Text) & "' where user_id='" & G_User_ID & "'"
   MsgBox "密码修改成功", vbInformation, "提示信息"
   Unload Me
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
 Dim w_rs As New ADODB.Recordset

   w_rs.Open "select password from mmst901 where password='" & Trim(Old_Pass.Text) & "' and user_id='" & G_User_ID & "'", G_Con
   If w_rs.EOF = True Then
      MsgBox "请检查你输入的密码,旧密码不对", vbInformation, "提示信息"
      Old_Pass.SetFocus
      Old_Pass.SelStart = 0
      Old_Pass.SelLength = Len(Old_Pass.Text)
      check_ok = False
      Set w_rs = Nothing
      Exit Function
   End If
   Set w_rs = Nothing

 
 If Trim(New_Pass.Text) <> Trim(C_New_Pass.Text) Then
    MsgBox "请确定输入的密码是否相同", vbInformation, "提示信息"
    C_New_Pass.SetFocus
    C_New_Pass.SelStart = 0
    C_New_Pass.SelLength = Len(New_Pass.Text)
    check_ok = False
    Exit Function
 End If
 
 check_ok = True
 
End Function
