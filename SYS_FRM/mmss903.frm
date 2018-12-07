VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mmss903 
   BorderStyle     =   1  '单线固定
   Caption         =   "系统解锁档(903)"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   HelpContextID   =   903
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8385
   Begin VB.Frame Frame1 
      Height          =   3225
      Left            =   3780
      TabIndex        =   6
      Top             =   1260
      Width           =   4575
      Begin VB.CommandButton Cmd_quit 
         Caption         =   "退出(&Q)"
         Height          =   350
         Left            =   2310
         TabIndex        =   10
         Top             =   1260
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_unlock 
         Caption         =   "解锁(&U)"
         Height          =   350
         Left            =   750
         TabIndex        =   9
         Top             =   1260
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_cancel 
         Caption         =   "取消(&N)"
         Height          =   350
         Left            =   2310
         TabIndex        =   8
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_ok 
         Caption         =   "确认(&Y)"
         Height          =   350
         Left            =   750
         TabIndex        =   7
         Top             =   840
         Width           =   1000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "解锁档描述"
      Height          =   1245
      Left            =   3780
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      Begin VB.Label Label1 
         Caption         =   "名称 :"
         Height          =   225
         Left            =   300
         TabIndex        =   5
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label2 
         Height          =   495
         Left            =   900
         TabIndex        =   4
         Top             =   450
         Width           =   3525
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "程式档案列表"
      Height          =   4485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Frame Frame3 
         Height          =   4185
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   3615
         Begin ComctlLib.TreeView tvwDB 
            Height          =   4065
            Left            =   0
            TabIndex        =   2
            Top             =   90
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   7170
            _Version        =   327682
            LabelEdit       =   1
            Style           =   6
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新细明体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "mmss903"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称: 解锁档(mmst903)(mmstlock)
'*编写日期:
'*制作人员:
'*修改日期:
'*修改人员:
'***********************************************
'定义欲打开的数据库及数据表名称
Dim st_prg As Recordset
Dim st_dbf As Recordset
Dim Tmp_Rb As Recordset

'定义树属性
Dim Nodx As ComctlLib.Node
Dim mitem As ListItem

'定义按钮变量
Dim c_unlock As Boolean

Dim c_off_unlock As Boolean

'定义当前树所指向的档案名称
Dim dbf_no As String

Private Sub treeshow()
Dim W_Rs As New ADODB.Recordset
Dim w_str As String
Dim i As Integer

tvwDB.Nodes.Clear

w_str = " select distinct system_id,menu_type from mmstlock order by system_id "
Set W_Rs = open_RecordSet(w_str)

i = 0
Set Nodx = tvwDB.Nodes.Add(, , "mMENU", "系统档案")
Do Until W_Rs.EOF
    i = i + 1
    w_str = "A00" + CStr(W_Rs!system_id)
    Set Nodx = tvwDB.Nodes.Add("mMENU", tvwChild, w_str, CStr(i) + Space(2) + W_Rs!menu_type)
    Call treeshow1(w_str)
    W_Rs.MoveNext
Loop
W_Rs.Close
Set W_Rs = Nothing
tvwDB.Nodes(1).Expanded = True
End Sub

Private Sub treeshow1(p_key As String)
Dim w_rs1 As New ADODB.Recordset
Dim w_str As String
Dim i As Integer

w_str = " select menu_no,menu_name,table_name from mmstlock where system_id=" & Right(p_key, 1) & " order by menu_no "
Set W_Rs = open_RS(w_str)

i = 0
Do Until W_Rs.EOF
    i = i + 1
    w_str = "child" + Format(i, "00")
    Set Nodx = tvwDB.Nodes.Add(p_key, tvwChild, w_str + W_Rs!Table_Name, CStr(i) + Space(2) + W_Rs!menu_name)
    W_Rs.MoveNext
Loop


Set w_rs1 = Nothing
tvwDB.Nodes(1).Expanded = True
End Sub





Private Sub Form_Load()
'将窗口置中
Call CenterWindow(Me)

'设置树
Call treeshow

'将按钮变量赋初值
c_unlock = False

'将确认取消按钮设成尽能
Cmd_unlock.Enabled = False
cmd_ok.Enabled = False
cmd_cancel.Enabled = False

End Sub
Private Sub readshow()

If c_unlock = True Then
    Cmd_unlock.Enabled = False
    cmd_ok.Enabled = True
    cmd_cancel.Enabled = True
Else
    Cmd_unlock.Enabled = True
    cmd_ok.Enabled = False
    cmd_cancel.Enabled = False
End If

End Sub

Private Sub cmd_cancel_Click()
Call vcontrol("N")
End Sub

Private Sub cmd_ok_Click()
Call vcontrol("Y")
End Sub

Private Sub Cmd_unlock_Click()
Call vcontrol("U")
End Sub

Sub cmd_quit_Click()
Call vcontrol("Q")
End Sub

Private Sub vcontrol(p_choice As String)
Select Case p_choice
    Case "Y"            '确定
        If check_ok() = True Then
            Call upd_data
            Frame3.Enabled = True
        Else
           Call cmd_cancel_Click
        End If
        
    Case "N"            '取消
        c_unlock = False

        Call readshow
        Frame3.Enabled = True
        
    Case "U"
    
        c_unlock = True    '解锁
        
        Call readshow
        Frame3.Enabled = False
        
    Case "Q"
    
    Unload Me
End Select

End Sub

Private Function check_ok()
If Label2.Caption = "系统档案" Or Label2.Caption = "" Then
  MsgBox "请选择系统档案中一个档案", 64, "信息"
  check_ok = False
  Exit Function
End If
check_ok = True
End Function

Private Sub upd_data()

G_Con.Errors.Clear
On Error GoTo ERRDO:
If dbf_no <> "" Then
    G_Con.Execute " update " & Right(dbf_no, Len(dbf_no) - 7) & " set lock='No' "
    MsgBox "操作成功!", 64, "信息"
End If

ERRDO:
    If err.Number <> 0 Then
        MsgBox "解锁失败", 64, "提示信息"
    End If

c_unlock = False
Call readshow
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If c_unlock Then
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

Private Sub tvwdb_NodeClick(ByVal Node As ComctlLib.Node)
Label2.Caption = Node.Text

dbf_no = Node.Key

If UCase(Left(dbf_no, 3)) = "A00" Then
    Cmd_unlock.Enabled = False
Else
    Cmd_unlock.Enabled = True
End If

End Sub

