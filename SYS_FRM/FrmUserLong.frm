VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmUserLong 
   BorderStyle     =   1  '单线固定
   Caption         =   "看谁在线上"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7290
   Begin VB.CheckBox Check1 
      Caption         =   "全删除"
      Height          =   255
      Left            =   1470
      TabIndex        =   3
      Top             =   4170
      Width           =   915
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "删除(&D)"
      Height          =   315
      Left            =   4290
      TabIndex        =   2
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2370
      Top             =   4110
   End
   Begin VB.CommandButton cmd_quit 
      Caption         =   "退出(&Q)"
      Height          =   315
      Left            =   5820
      TabIndex        =   1
      Top             =   4140
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6800
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新细明体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "用户姓名"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "模块名称"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "登入日期"
         Object.Width           =   6068
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUserLong.frx":0000
            Key             =   "main"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmUserLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Del_Click()
If Check1.Value = 1 Then
  G_Con.Execute "delete from userlong "
Else
  G_Con.Execute "delete from userlong " & _
              "where user_name='" & ListView1.SelectedItem.Text & "' " & _
              "and code_name='" & ListView1.SelectedItem.ListSubItems.Item(1).Text & "'" & _
              "and upd_date='" & ListView1.SelectedItem.ListSubItems.Item(2).Text & "'"
End If
End Sub

Private Sub cmd_quit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim w_rs As New ADODB.Recordset

Call CenterWindow(FrmUserLong, sys_main)
w_rs.Open "select * from userlong order by user_name,code_name ", G_Con, adOpenKeyset, adLockOptimistic
i = 1
While w_rs.EOF <> True
   ListView1.ListItems.Add , , w_rs!user_name
   ListView1.ListItems(i).ListSubItems.Add , , w_rs!code_name
   ListView1.ListItems(i).ListSubItems.Add , , w_rs!upd_date
 
  w_rs.MoveNext
  i = i + 1
Wend
End Sub

Private Sub Timer1_Timer()
Dim w_rs As New ADODB.Recordset
ListView1.ListItems.Clear
Set w_rs = Nothing
w_rs.Open "select * from userlong order by user_name,code_name ", G_Con, adOpenKeyset, adLockOptimistic
i = 1
While w_rs.EOF <> True
   ListView1.ListItems.Add , , w_rs!user_name
   ListView1.ListItems(i).ListSubItems.Add , , w_rs!code_name
   ListView1.ListItems(i).ListSubItems.Add , , w_rs!upd_date
 
  w_rs.MoveNext
  i = i + 1
Wend
End Sub
