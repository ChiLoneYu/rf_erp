VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form mmss907 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户程式授权"
   ClientHeight    =   6180
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   907
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   8655
   Begin VB.Frame Frame2 
      Caption         =   "程式权限设置"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   4290
      TabIndex        =   3
      Top             =   660
      Width           =   2235
      Begin VB.CheckBox chk_delete1 
         Caption         =   "  删  除(G)(BOM)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   27
         Top             =   4550
         Width           =   1770
      End
      Begin VB.CheckBox Chk_pick 
         Caption         =   "  特  采(K)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   26
         Top             =   4165
         Width           =   1530
      End
      Begin VB.CheckBox chk_price 
         Caption         =   "  单  价(F)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   25
         Top             =   3780
         Width           =   1530
      End
      Begin VB.CheckBox Chk_Save 
         Caption         =   "  存  档(S)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   20
         Top             =   3395
         Width           =   1530
      End
      Begin VB.CheckBox chk_see 
         Caption         =   "  可  见(O)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   12
         Top             =   4935
         Width           =   1530
      End
      Begin VB.CheckBox Chk_print 
         Caption         =   "  列  印(P)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   11
         Top             =   3010
         Width           =   1530
      End
      Begin VB.CheckBox Chk_preview 
         Caption         =   "  预  览(V)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   10
         Top             =   2625
         Width           =   1530
      End
      Begin VB.CheckBox Chk_query 
         Caption         =   "  查  询(I)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   9
         Top             =   2240
         Width           =   1530
      End
      Begin VB.CheckBox Chk_reset 
         Caption         =   "  重  置(R)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   1855
         Width           =   1530
      End
      Begin VB.CheckBox Chk_check 
         Caption         =   "  审  核(C)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   7
         Top             =   1470
         Width           =   1530
      End
      Begin VB.CheckBox Chk_delete 
         Caption         =   "  删  除(D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   6
         Top             =   1085
         Width           =   1530
      End
      Begin VB.CheckBox Chk_edit 
         Caption         =   "  修  改(U)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   5
         Top             =   700
         Width           =   1530
      End
      Begin VB.CheckBox Chk_add 
         Caption         =   "  新  增(A)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Top             =   315
         Width           =   1530
      End
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8565
      Begin VB.TextBox user_name 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         TabIndex        =   28
         Top             =   210
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   180
         Left            =   3570
         TabIndex        =   24
         Top             =   285
         Visible         =   0   'False
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   990
         TabIndex        =   0
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "用户代号:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   6540
      TabIndex        =   21
      Top             =   690
      Width           =   2025
      Begin VB.CommandButton Command1 
         Caption         =   " 确 认(&Y)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   390
         TabIndex        =   13
         Top             =   690
         Width           =   1350
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取 消(&N)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   390
         TabIndex        =   14
         Top             =   1332
         Width           =   1350
      End
      Begin VB.CommandButton Command3 
         Caption         =   "全部授权(&A)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   390
         TabIndex        =   15
         Top             =   1974
         Width           =   1350
      End
      Begin VB.CommandButton Command4 
         Caption         =   "退 出(Q)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   390
         TabIndex        =   19
         Top             =   4545
         Width           =   1350
      End
      Begin VB.CommandButton Command5 
         Caption         =   "单项授权(&A)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   390
         TabIndex        =   16
         Top             =   2616
         Width           =   1350
      End
      Begin VB.CommandButton Command6 
         Caption         =   "全部移除(&D)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   390
         TabIndex        =   17
         Top             =   3258
         Width           =   1350
      End
      Begin VB.CommandButton Command7 
         Caption         =   "单项移除(&A)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   390
         TabIndex        =   18
         Top             =   3900
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   30
      TabIndex        =   1
      Top             =   690
      Width           =   4200
      Begin MSComctlLib.TreeView tvwDB 
         Height          =   5205
         Left            =   45
         TabIndex        =   2
         Top             =   180
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   9181
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
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
   End
End
Attribute VB_Name = "mmss907"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim St_C03 As New ADODB.Recordset
Dim St_C05 As New ADODB.Recordset
Dim st_temp As New ADODB.Recordset
Dim W_Sav_Rec As Variant
Dim W_Text As String
Dim Mytree_Nodx As Integer
Dim Mynode As Node

Private Type TmpData
     system_id As Integer
     menu_id As String
     menu_name As String
     menu_ename As String
     prog_id As String
     prog_type As String
     Rights As String
     user_id As String
     new_rights As String
End Type

Dim Mytmpdata() As TmpData

'Private Sub chk_delete1_Click()
'If chk_delete1.Value = 0 Then
'       Call del_qx("G", chk_delete1)
' ElseIf chk_delete1.Value = 1 Then
'       Call add_qx("G", chk_delete1)
' End If
'End Sub

Private Sub chk_delete1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chk_delete1.Value = 0 Then
       Call del_qx("G", chk_delete1)
 ElseIf chk_delete1.Value = 1 Then
       Call add_qx("G", chk_delete1)
 End If
End Sub

'Private Sub Chk_pick_Click()
'If Chk_pick.Value = 0 Then
'       Call del_qx("K", Chk_pick)
' ElseIf Chk_pick.Value = 1 Then
'       Call add_qx("K", Chk_pick)
' End If
'End Sub

Private Sub Chk_pick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_pick.Value = 0 Then
       Call del_qx("K", Chk_pick)
 ElseIf Chk_pick.Value = 1 Then
       Call add_qx("K", Chk_pick)
 End If
End Sub

'Private Sub chk_price_Click()
''If chk_price.Value = 0 Then
''       Call del_qx("F", chk_price)
'' ElseIf chk_price.Value = 1 Then
''       Call add_qx("F", chk_price)
'' End If
'End Sub

Private Sub chk_price_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chk_price.Value = 0 Then
       Call del_qx("F", chk_price)
 ElseIf chk_see.Value = 1 Then
       Call add_qx("F", chk_price)
 End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'修改表单各标识
If Shift = 1 And KeyCode = vbKeyM Then
    If G_Userid = "A001" Then
        With mmss_mofify
            Set .Parent_form = Me
            .Show vbModal
        End With
    End If
End If
End Sub

Private Sub Form_Load()
Dim st_901 As New ADODB.Recordset

With mmss907
    .Left = Int(sys_main.Width - mmss907.Width) / 2
    .Top = Int(sys_main.Height - mmss907.Height - 1500) / 2
End With

With st_901
    .ActiveConnection = G_Con
    .Open "select * from mmst901 order by user_id"
    Do While .EOF = False
        Combo1.AddItem .Fields("user_id").Value
        .MoveNext
    Loop
    .Close
End With
Set st_901 = Nothing

'Call change_form(mmss907)

Combo1.ListIndex = 0
Set Mynode = tvwDB.HitTest(0, 0)
Call tvwdb_NodeClick(Mynode)

End Sub
Private Sub Chk_add_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Debug.Print Chk_add.Value
If Chk_add.Value = 0 Then
        Call del_qx("A")
 ElseIf Chk_add.Value = 1 Then
         Call add_qx("A")
 End If

End Sub

Private Sub Chk_check_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_check.Value = 0 Then
        Call del_qx("C", Chk_check)
 ElseIf Chk_check.Value = 1 Then
        Call add_qx("C", Chk_check)
 End If

End Sub

Private Sub Chk_delete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_delete.Value = 0 Then
         Call del_qx("D", Chk_delete)
 ElseIf Chk_delete.Value = 1 Then
        Call add_qx("D", Chk_delete)
 End If

End Sub

Private Sub Chk_edit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_edit.Value = 0 Then
         Call del_qx("U", Chk_edit)
 ElseIf Chk_edit.Value = 1 Then
         Call add_qx("U", Chk_edit)
 End If
End Sub

Private Sub Chk_preview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_preview.Value = 0 Then
        Call del_qx("V", Chk_preview)
 ElseIf Chk_preview.Value = 1 Then
        Call add_qx("V", Chk_preview)
 End If

End Sub

Private Sub Chk_print_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_print.Value = 0 Then
        Call del_qx("P", Chk_print)
 ElseIf Chk_print.Value = 1 Then
        Call add_qx("P", Chk_print)
 End If

End Sub

Private Sub Chk_query_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_query.Value = 0 Then
        Call del_qx("I", Chk_query)
 ElseIf Chk_query.Value = 1 Then
        Call add_qx("I", Chk_query)
 End If
End Sub

Private Sub Chk_reset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_reset.Value = 0 Then
        Call del_qx("R", Chk_reset)
 ElseIf Chk_reset.Value = 1 Then
       Call add_qx("R", Chk_reset)
 End If

End Sub

Private Sub chk_see_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If chk_see.Value = 0 Then
Call check_node(Mynode)
'End If
End Sub

Private Sub chk_save_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Chk_Save.Value = 0 Then
       Call del_qx("S", Chk_Save)
 ElseIf chk_see.Value = 1 Then
       Call add_qx("S", Chk_Save)
 End If
End Sub

Private Sub check_node(p_Node As MSComctlLib.Node)
On Error Resume Next
Dim X_Node As MSComctlLib.Node

If p_Node.Children > 0 Then       '当有子记录时
    Set X_Node = p_Node.Child
    Mytree_Nodx = Val(p_Node.Tag)
    If chk_see.Value = 0 Then
        Call del_qx("O", chk_see)
    ElseIf chk_see.Value = 1 Then
        Call add_qx("O", chk_see)
    End If
    For i = 1 To p_Node.Children
        Call check_node(X_Node)
        Set X_Node = X_Node.Next
    Next i
Else  '没有子记录时
    Mytree_Nodx = Val(p_Node.Tag)
    If chk_see.Value = 0 Then
        Call del_qx("O", chk_see)
    ElseIf chk_see.Value = 1 Then
        Call add_qx("O", chk_see)
    End If
End If
End Sub


Private Sub Combo1_Click()
Dim i As Integer
Dim W_RS As New ADODB.Recordset
i = 0
If W_Text <> Combo1.Text Then
    MyStr = "SELECT A.*,B.rights, B.[User_id] " & _
            "From (SELECT  system_id,menu_id,menu_name,prog_id, prog_type, menu_ename " & _
                  "From mmstc02 " & _
                  "WHERE list_visible=1 and menu_id NOT IN ('menu_add', 'menu_modify','menu_edit', 'menu_delete', 'menu_v','menu_v1', 'menu_v2','menu_loadpic','menu_picture','menu_unloadpic')) A " & _
                  "LEFT JOIN  " & _
                  "(SELECT system_id,menu_id,rights,[user_id] " & _
                   "FROM mmstc03 " & _
                   "WHERE [user_id]='" & Trim(Combo1.Text) & "') B " & _
                   "ON A.system_id = B.system_id AND " & _
                   "A.menu_id = B.menu_id  " & _
             "ORDER BY A.system_id,A.menu_id "
    Set st_temp = Nothing
    With st_temp
        .ActiveConnection = G_Con
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open MyStr
    End With
    If st_temp.AbsolutePosition <> -1 Then
        ReDim Mytmpdata(st_temp.RecordCount)
        st_temp.MoveFirst
    End If
    Do Until st_temp.EOF
       
       Mytmpdata(i).system_id = st_temp!system_id
       Mytmpdata(i).menu_id = st_temp!menu_id
       Mytmpdata(i).menu_name = st_temp!menu_name
       Mytmpdata(i).menu_ename = NullSetValue(st_temp!menu_ename, "")
       If st_temp!prog_id <> " " Then
           Mytmpdata(i).prog_id = st_temp!prog_id
       End If
       Mytmpdata(i).prog_type = IIf(IsNull(st_temp!prog_type), "", st_temp!prog_type)
       If NullSetValue(st_temp!Rights, "") <> "" Then
            Mytmpdata(i).Rights = st_temp!Rights
            Mytmpdata(i).new_rights = st_temp!Rights
       End If
       If Not IsNull(st_temp!user_id) Then
           Mytmpdata(i).user_id = st_temp!user_id
       Else
           Mytmpdata(i).user_id = ""
       End If
       st_temp.MoveNext
       i = i + 1
    Loop
    
    Call treeshow
End If
W_Text = Combo1.Text

Set W_RS = Nothing
W_RS.Open "select user_name from mmst901 where user_id='" & Combo1.Text & "' ", G_Con
If W_RS.EOF = False Then
   user_name = W_RS!user_name
End If
Set W_RS = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 13
End Sub

Private Sub Command1_Click()

    Dim i As Integer
    Bar1.Visible = True
    Do Until i = UBound(Mytmpdata)
       If Trim(Mytmpdata(i).new_rights) <> Trim(Mytmpdata(i).Rights) Then
            Mytmpdata(i).Rights = Mytmpdata(i).new_rights
            If CheckRight(Mytmpdata(i).user_id, Mytmpdata(i).menu_id, Mytmpdata(i).system_id) = True Then
                MyStr = "update mmstc03 set rights='" & Trim(Mytmpdata(i).Rights) & "' " & _
                        "where User_id='" & Trim(Mytmpdata(i).user_id) & "' and " & _
                        "menu_id='" & Trim(Mytmpdata(i).menu_id) & "' and " & _
                        "system_id=" & Mytmpdata(i).system_id
            Else
                MyStr = "insert mmstc03(rights,user_id,menu_id,system_id) values('" & Trim(Mytmpdata(i).Rights) & "','" & _
                        Trim(Mytmpdata(i).user_id) & "','" & _
                        Trim(Mytmpdata(i).menu_id) & "'," & _
                        Mytmpdata(i).system_id & ")"
            End If
            G_Con.Execute MyStr
            Bar1.Value = i / UBound(Mytmpdata) * 100
       End If
     i = i + 1
    Loop
    If Mynode.Key <> "nMENU" Then Call tvwdb_NodeClick(Mynode)
    tvwDB.SetFocus
    Bar1.Visible = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = True
    Command6.Enabled = True

End Sub

Private Sub Command2_Click()
Dim i As Integer
Do Until i = UBound(Mytmpdata)
 Mytmpdata(i).new_rights = Mytmpdata(i).Rights
 i = i + 1
Loop
Call tvwdb_NodeClick(Mynode)
tvwDB.SetFocus
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command6.Enabled = True
MsgBox "取消完成", 48, "提示"
End Sub

Private Sub Command3_Click()
On Error GoTo myerr:
Dim i As Integer
Dim M As Long
M = UBound(Mytmpdata)
Do Until i = M
    Mytmpdata(i).new_rights = Mytmpdata(i).prog_type
    Mytmpdata(i).user_id = Trim(Combo1.Text)
    i = i + 1
Loop
Call tvwdb_NodeClick(Mynode)
tvwDB.SetFocus
Command3.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Exit Sub
myerr:

Command3.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
End Sub


Private Sub Command4_Click()
Unload Me
End Sub
Private Sub Command5_Click()
Dim i As Integer
Dim Node_Len As Integer
Dim Node_Key As String
If UCase(Left(CStr(Mynode.Key), 6)) = "MENU_P" Or UCase(Left(CStr(Mynode.Key), 6)) = "MENU_J" Or UCase(Left(CStr(Mynode.Key), 6)) = "MENU_K" Or UCase(Left(CStr(Mynode.Key), 6)) = "MENU_C" Then
   Node_Key = Left(Trim(Mynode.Key), Len(CStr(Mynode.Key)) - 2)
Else
   Node_Key = Left(Trim(Mynode.Key), Len(CStr(Mynode.Key)) - 1)
End If

Node_Len = Len(Node_Key)
Do Until i = UBound(Mytmpdata)
  If UCase(Left(Trim(Mytmpdata(i).menu_id), Node_Len)) = UCase(Node_Key) Then
      Mytmpdata(i).new_rights = Mytmpdata(i).prog_type
      Mytmpdata(i).user_id = Trim(Combo1.Text)
  End If
  
 i = i + 1
Loop
Call tvwdb_NodeClick(Mynode)
tvwDB.SetFocus
Command5.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command6_Click()
On Error GoTo myerr:
Dim i As Integer
Dim M As Long
M = UBound(Mytmpdata)
Do Until i = M
  Mytmpdata(i).new_rights = ""
  Mytmpdata(i).user_id = Trim(Combo1.Text)
  i = i + 1
Loop
Call tvwdb_NodeClick(Mynode)
tvwDB.SetFocus
Command6.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Exit Sub
myerr:

Command6.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command7_Click()
Dim i As Integer
Dim Node_Len As Integer
Dim Node_Key As String
If UCase(Left(CStr(Mynode.Key), 6)) = "MENU_P" Or UCase(Left(CStr(Mynode.Key), 6)) = "MENU_J" Or UCase(Left(CStr(Mynode.Key), 6)) = "MENU_K" Then
   Node_Key = Left(Trim(Mynode.Key), Len(CStr(Mynode.Key)) - 2)
Else
   Node_Key = Left(Trim(Mynode.Key), Len(CStr(Mynode.Key)) - 1)
End If

Node_Len = Len(Node_Key)
Do Until i = UBound(Mytmpdata)
  If UCase(Left(Trim(Mytmpdata(i).menu_id), Node_Len)) = UCase(Node_Key) Then
      Mytmpdata(i).new_rights = ""
      Mytmpdata(i).user_id = Trim(Combo1.Text)
  End If
  
 i = i + 1
Loop
Call tvwdb_NodeClick(Mynode)
tvwDB.SetFocus
Command7.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
End Sub

Sub add_qx(w_qx As String, Optional chk As CheckBox)
Mytmpdata(Mytree_Nodx).new_rights = Mytmpdata(Mytree_Nodx).new_rights & w_qx
If Mytmpdata(Mytree_Nodx).user_id = "" Then
    Mytmpdata(Mytree_Nodx).user_id = Trim(Combo1.Text)
End If
Command1.Enabled = True
Command2.Enabled = True
End Sub
Sub del_qx(w_qx As String, Optional chk As CheckBox)
Mytmpdata(Val(Mytree_Nodx)).new_rights = del_str_qx(Mytmpdata(Val(Mytree_Nodx)).new_rights, w_qx)
If Mytmpdata(Mytree_Nodx).user_id = "" Then
    Mytmpdata(Mytree_Nodx).user_id = Trim(Combo1.Text)
End If
Command1.Enabled = True
Command2.Enabled = True
End Sub
Function del_str_qx(str_qx As String, w_qx1 As String) As String
Dim W_I As Integer, str As String

W_I = InStr(1, str_qx, w_qx1, vbTextCompare)
If W_I > 0 Then
    del_str_qx = Mid(str_qx, 1, W_I - 1) & Mid(str_qx, W_I + 1)
Else
    del_str_qx = str_qx
End If
End Function

Private Sub treeshow()
Dim M As Integer
Dim Nodx As Node
'On Error Resume Next
M = 0
tvwDB.Nodes.Clear

Set Nodx = tvwDB.Nodes.Add(, , "mMENU", "系统授权栏")

'Do Until m = UBound(mytmpdata)
'    If Len(mytmpdata(m).system_id) > 2 Then
'        If Right(mytmpdata(m).system_id, 2) = "00" Then
'            w_nodkey = Trim(mytmpdata(m).menu_id)
'            Set nodx = tvwDB.Nodes.Add("mMENU", tvwChild, w_nodkey & mytmpdata(m).system_id, IIf(g_Language = "C", Trim(mytmpdata(m).menu_name), Trim(mytmpdata(m).menu_ename)))
'            nodx.Tag = m
'        End If
'    End If
'    m = m + 1
'Loop
'm = 0

Do Until M = UBound(Mytmpdata)
   w_nodkey = LCase(Trim(Mytmpdata(M).menu_id))
'   On Error Resume Next
  If Trim(Mytmpdata(M).prog_type) = "MO" And Len(w_nodkey) = 6 Then
      Set Nodx = tvwDB.Nodes.Add("mMENU", tvwChild, w_nodkey & Mytmpdata(M).system_id, Trim(Mytmpdata(M).menu_name))
      Nodx.Tag = M
  Else
      If Len(w_nodkey) < 9 Then
         '为何不把menu_a1之类的menu_id改为menu_a_1呢?多堋有规律的编码就这样被破坏了
          Set Nodx = tvwDB.Nodes.Add(Left(w_nodkey, 6) & Mytmpdata(M).system_id, tvwChild, w_nodkey & Mytmpdata(M).system_id, Trim(Mytmpdata(M).menu_name))
      Else
          Set Nodx = tvwDB.Nodes.Add(PartString(w_nodkey, "_", True, False) & Mytmpdata(M).system_id, tvwChild, w_nodkey & Mytmpdata(M).system_id, Trim(Mytmpdata(M).menu_name))
      End If
      Nodx.Tag = M
  End If
  M = M + 1
Loop
Set Nodx = Nothing
tvwDB.Nodes(1).Expanded = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Command1.Enabled = True Then
    Dim Msg As String
    '当有数据改动时.询问是否要退出系统
    Msg = "当前纪录尚未存储,您要退出吗?"
    If MsgBox(Msg, vbQuestion + vbYesNo, "提示") = vbNo Then
      Cancel = 1
    Else
      Cancel = 0
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
W_Text = ""
Set St_C03 = Nothing
Set St_C05 = Nothing
Set st_temp = Nothing
End Sub


Private Sub tvwdb_NodeClick(ByVal Node As MSComctlLib.Node)
'Debug.Print Node.Key
Dim W_Keyval As Integer

W_Keyval = Val(Node.Tag)

Command5.Enabled = True
Command7.Enabled = True

Set Mynode = Node

If InStr(1, Mytmpdata(W_Keyval).prog_type, "A") <> 0 Then
    Chk_add.Enabled = True
Else
    Chk_add.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "U") <> 0 Then
    Chk_edit.Enabled = True
Else
    Chk_edit.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "D") <> 0 Then
    Chk_delete.Enabled = True
Else
    Chk_delete.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "C") <> 0 Then
    Chk_check.Enabled = True
Else
    Chk_check.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "R") <> 0 Then
    Chk_reset.Enabled = True
Else
    Chk_reset.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "I") <> 0 Then
    Chk_query.Enabled = True
Else
    Chk_query.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "V") <> 0 Then
    Chk_preview.Enabled = True
Else
    Chk_preview.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "P") <> 0 Then
    Chk_print.Enabled = True
Else
   Chk_print.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "O") <> 0 Then
    chk_see.Enabled = True
Else
   chk_see.Enabled = False
End If

If InStr(1, Mytmpdata(W_Keyval).prog_type, "S") <> 0 Then
    Chk_Save.Enabled = True
Else
   Chk_Save.Enabled = False
End If

'判断是否涉及单价授权
If InStr(1, Mytmpdata(W_Keyval).prog_type, "F") <> 0 Then
    chk_price.Enabled = True
Else
    chk_price.Enabled = False
End If

'判断是否涉及特采授权
If InStr(1, Mytmpdata(W_Keyval).prog_type, "K") <> 0 Then
    Chk_pick.Enabled = True
Else
    Chk_pick.Enabled = False
End If

'用於bom删除
If InStr(1, Mytmpdata(W_Keyval).prog_type, "G") <> 0 Then
    chk_delete1.Enabled = True
Else
    chk_delete1.Enabled = False
End If


'******************************
If InStr(1, Mytmpdata(W_Keyval).new_rights, "A") <> 0 Then
    Chk_add.Value = 1
Else
    Chk_add.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "U") <> 0 Then
    Chk_edit.Value = 1
Else
    Chk_edit.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "D") <> 0 Then
    Chk_delete.Value = 1
Else
    Chk_delete.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "C") <> 0 Then
    Chk_check.Value = 1
Else
    Chk_check.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "R") <> 0 Then
    Chk_reset.Value = 1
Else
    Chk_reset.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "I") <> 0 Then
    Chk_query.Value = 1
Else
    Chk_query.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "V") <> 0 Then
    Chk_preview.Value = 1
Else
    Chk_preview.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "P") <> 0 Then
    Chk_print.Value = 1
Else
    Chk_print.Value = 0
End If
If InStr(1, Mytmpdata(W_Keyval).new_rights, "O") <> 0 Then
    chk_see.Value = 1
Else
    chk_see.Value = 0
End If

If InStr(1, Mytmpdata(W_Keyval).new_rights, "S") <> 0 Then
    Chk_Save.Value = 1
Else
    Chk_Save.Value = 0
End If
'用於设定是否查看或录入单价
If InStr(1, Mytmpdata(W_Keyval).new_rights, "F") <> 0 Then
    chk_price.Value = 1
Else
    chk_price.Value = 0
End If

'用於设定是否查看或录入特采
If InStr(1, Mytmpdata(W_Keyval).new_rights, "K") <> 0 Then
    Chk_pick.Value = 1
Else
    Chk_pick.Value = 0
End If

'用於bom
If InStr(1, Mytmpdata(W_Keyval).new_rights, "G") <> 0 Then
    chk_delete1.Value = 1
Else
    chk_delete1.Value = 0
End If

Mytree_Nodx = W_Keyval
'warning.Panels(1).Text = "提示:现在接受授权的项目是'" & Node & "'"
End Sub

Sub init_form()
'设定表单
Call change_form(mmss907)
'改变Treeview
Call treeshow
End Sub
Private Function CheckRight(user_id As String, menu_id As String, system_id As Integer) As Boolean
Dim W_RS As New ADODB.Recordset
W_RS.Open "select user_id from mmstc03 where user_id='" & user_id & "' and menu_id='" & menu_id & "' and system_id='" & system_id & "'", G_Con
CheckRight = (W_RS.EOF = False)
W_RS.Close
Set W_RS = Nothing
End Function
