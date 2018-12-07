VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mmss906 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "系統解鎖檔(903)"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   HelpContextID   =   903
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   11850
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "解鎖檔描述"
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   5430
      TabIndex        =   3
      Top             =   0
      Width           =   6345
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "名稱 :"
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   450
         Width           =   5055
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "程式檔案列表"
      ForeColor       =   &H80000008&
      Height          =   6705
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5385
      Begin ComctlLib.TreeView tvwDB 
         Height          =   6375
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   11245
         _Version        =   327682
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   5430
      TabIndex        =   8
      Top             =   5070
      Width           =   6405
      Begin VB.CommandButton Cmd_quit 
         Caption         =   "退出(&Q)"
         Height          =   345
         Left            =   4920
         TabIndex        =   7
         Top             =   690
         Width           =   1125
      End
      Begin VB.CommandButton Cmd_unlock 
         Caption         =   "解鎖(&U)"
         Height          =   345
         Left            =   390
         TabIndex        =   5
         Top             =   690
         Width           =   1125
      End
      Begin VB.CommandButton Cmd_cancel 
         Caption         =   "取消(&N)"
         Height          =   345
         Left            =   3410
         TabIndex        =   9
         Top             =   690
         Width           =   1125
      End
      Begin VB.CommandButton Cmd_ok 
         Caption         =   "确認(&Y)"
         Height          =   345
         Left            =   1920
         TabIndex        =   6
         Top             =   690
         Width           =   1125
      End
   End
End
Attribute VB_Name = "mmss906"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名稱: 解鎖檔(mmss903)
'*編寫日期: 2002年7月29日
'*制作人員: 危俊剛
'*修改日期:
'*修改人員:
'***********************************************
'定義欲打開的數據庫及數據表名稱

Dim st_prg As New ADODB.Recordset
Dim st_dbf As New ADODB.Recordset
Dim tmp_rb As New ADODB.Recordset

'定義樹屬性
Dim w_nodx  As ComctlLib.Node
Dim mitem As ListItem

'定義按鈕變量
Dim c_unlock As Boolean

Dim c_off_unlock As Boolean

'定義當前樹所指向的檔案名稱
Dim dbf_no As String

Private Sub treeshow()
Dim w_frm_name As String
Dim w_menu As String
'On Error Resume Next
tvwDB.Nodes.Clear

Set w_nodx = tvwDB.Nodes.Add(, , "A", "訂單出貨系統")

Dim w_tmp As New ADODB.Recordset
w_tmp.Open "select distinct frm_name,list_no from mmstprg order by list_no", G_Con
Do While w_tmp.EOF = False
    Set w_nodx = tvwDB.Nodes.Add("A", 4, "A" & w_tmp!list_no, w_tmp!Frm_Name)
    w_tmp.MoveNext
Loop
    
'If st_prg.EOF = False Then
'    st_prg.MoveFirst
'End If
'
'Do While st_prg.EOF = False
'    Set w_nodx = tvwDB.Nodes.Add("A" & st_prg!list_no, tvwChild, "A" & CStr(st_prg!frm_no), st_prg!Frm_Name)
'    st_prg.MoveNext
'Loop

tvwDB.Nodes(1).Expanded = True
End Sub

Private Sub Form_Activate()
g_active = True
End Sub

Private Sub Form_Load()
'將窗口置中
Call CenterWindow(Me)


'打開數據庫及相應數據表

 st_prg.Open "select * from mmstprg order by list_no ", G_Con, adOpenKeyset, adLockOptimistic
 st_dbf.Open "select * from mmstdbf order by list_no ", G_Con, adOpenKeyset, adLockOptimistic

'設置樹
Call treeshow

'將按鈕變量賦初值
c_unlock = False

'將確認取消按鈕設成盡能
cmd_ok.Enabled = False
Cmd_cancel.Enabled = False

End Sub
Private Sub readshow()

If c_unlock = True Then
    Cmd_unlock.Enabled = False
    cmd_ok.Enabled = True
    Cmd_cancel.Enabled = True
Else
    Cmd_unlock.Enabled = True
    cmd_ok.Enabled = False
    Cmd_cancel.Enabled = False
End If

End Sub

Private Sub cmd_cancel_Click()
Call vcontrol("N")
End Sub

Private Sub Cmd_OK_Click()
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
            Frame4.Enabled = True
        Else
           Call cmd_cancel_Click
        End If
        
    Case "N"            '取消
        c_unlock = False

        Call readshow
        Frame4.Enabled = True
        
    Case "U"
    
        c_unlock = True    '解鎖
        
        Call readshow
        Frame4.Enabled = False
        
    Case "Q"
    
    Unload Me
End Select

End Sub

Private Function check_ok()
If Label2.Caption = "訂單出貨系統" Or Label2.Caption = "" Then
  MsgBox "請選擇系統檔案中一個檔案", 64, "信息"
  check_ok = False
  Exit Function
End If
check_ok = True
End Function

Private Sub upd_data()

Dim temp As New ADODB.Recordset
On Error Resume Next
temp.Open "select * from mmstdbf where frm_no='" & dbf_no & "'", G_Con, adOpenKeyset, adLockOptimistic
 Do Until temp.EOF
'    If temp!Filter <> "" Then
'        G_Con.Execute "update " & temp!dbf_name & " set lock = 'No' WHERE " & temp!Filter
'    Else
        G_Con.Execute "update " & temp!dbf_name & " set lock = 'No' "
'    End If
    
    temp.MoveNext
 Loop
 
temp.Close
MsgBox "操作成功!", 64, "信息"


c_unlock = False
Call readshow
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If c_unlock Then
    Dim Msg As String
    '當有數据改動時.詢問是否要退出系統
    Msg = "當前紀錄尚未存儲,您要退出嗎?"
    If MsgBox(Msg, vbQuestion + vbYesNo, "提示") = vbNo Then
      Cancel = 1
    Else
      Cancel = 0
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
g_active = False
Set st_prg = Nothing
Set st_dbf = Nothing
Set tmp_rb = Nothing



End Sub

Private Sub tvwDB_NodeClick(ByVal Node As ComctlLib.Node)
Dim W_Rs As New ADODB.Recordset

Label2.Caption = Node.Text
With W_Rs
    .CursorLocation = adUseClient
    .Open "select frm_no from mmstprg where frm_name='" & Trim(Node.Text) & "'", G_Con, adOpenDynamic
End With
If W_Rs.EOF = False Then
    dbf_no = W_Rs!frm_no
End If
W_Rs.Close
End Sub


