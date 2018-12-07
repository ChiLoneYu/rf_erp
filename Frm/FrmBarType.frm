VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmBarType 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "仓别资料选择(FrmBarType)"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_ok 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      Picture         =   "FrmBarType.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmd_quit 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      Picture         =   "FrmBarType.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
   Begin ComctlLib.TreeView tvwdb 
      Height          =   5085
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   8969
      _Version        =   327682
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   -30
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmBarType.frx":2B44
            Key             =   "book"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmBarType.frx":2E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmBarType.frx":2F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmBarType.frx":3052
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmBarType.frx":314C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmBarType.frx":3246
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmBarType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_Node As ComctlLib.Node
Dim W_key As String
Dim W_Bar_No As String
Dim W_Bar_name As String
Dim W_Bar_ID As String

Public Property Get Bar_No() As String
Bar_No = W_Bar_No
End Property

Public Property Get Bar_Name() As String
Bar_Name = W_Bar_name
End Property

Public Property Get Bar_ID() As String
Bar_ID = W_Bar_ID
End Property

Private Sub Form_Load()
Me.KeyPreview = True
Call Set_Color_Frm(Me)
Call brow_tree
Set W_Node = tvwdb.Nodes(1)
Call tvwDB_NodeClick(W_Node)

'展开节点
For i = 1 To tvwdb.Nodes.Count
    If i <= 3 Then
        tvwdb.Nodes(i).Expanded = True
    End If
Next i
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Me.ActiveControl.Name = "tvwdb" Then
    Call tvwdb_DblClick
End If
If KeyCode = vbKeyEscape Then
    Unload Me
End If

If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               '确认
       Call Cmd_OK_Click
    Case vbKeyEscape           '取消
       Call cmd_quit_Click
    End Select
End If
End Sub
'加载 TreeView 数据
Private Sub brow_tree()
Dim st_903_1 As New ADODB.Recordset
Dim w_type As ComctlLib.Node

With st_903_1
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open "select * from mmst903 where len(Bar_No) = 4 and bar_name<>'总务仓'order by Bar_No"
End With

tvwdb.Nodes.Clear
Set w_type = tvwdb.Nodes.Add(, , "A", "仓别资料", 1)
Do While st_903_1.EOF = False
    If Trim(st_903_1!type_type) = "0" Then '小类
        Set w_type = tvwdb.Nodes.Add("A", tvwChild, st_903_1!Bar_No, st_903_1!Bar_Name, 3)
        w_type.Tag = "X"
    Else
        Set w_type = tvwdb.Nodes.Add("A", tvwChild, st_903_1!Bar_No, st_903_1!Bar_Name, 1)
        w_type.Tag = "E"
'        Call browtree(st_903_1!Bar_No)
    End If
    st_903_1.MoveNext
Loop
tvwdb.Sorted = True
End Sub

Private Sub browtree(P_Key As String)
Dim w_aa As New ADODB.Recordset
Dim w_type1 As ComctlLib.Node
With w_aa
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open "select * from mmst903 where left(Bar_No,len('" & P_Key & "')) = '" & P_Key & "' AND Bar_No<> '" & P_Key & "' and len(Bar_No)=(len('" & P_Key & "')+3)"
End With

Do Until w_aa.EOF
    If Trim(w_aa!type_type) = "0" Then
        Set w_type1 = tvwdb.Nodes.Add(Trim(P_Key), tvwChild, w_aa!Bar_No, w_aa!Bar_Name, 3)
        w_type1.Tag = "X"
        
    Else
        Set w_type1 = tvwdb.Nodes.Add(Trim(P_Key), tvwChild, w_aa!Bar_No, w_aa!Bar_Name, 1)
        w_type1.Tag = "E"
        'Call browtree(w_aa!Bar_No)
    End If
    w_aa.MoveNext
Loop
w_aa.Close
End Sub



Private Sub tvwDB_Collapse(ByVal Node As ComctlLib.Node)
    Node.Image = 1
End Sub

Private Sub tvwdb_DblClick()
If tvwdb.Nodes.Count > 1 Then
    If W_key <> "A" Then
        W_Bar_No = W_Node.Key
        W_Bar_name = W_Node.Text
        W_Bar_ID = W_Node.Key
        '加载下阶
        If W_Node.Tag = "E" Then
            Call browtree(W_key)
            W_Node.Expanded = True
            W_Node.Tag = "X"
            Exit Sub
        End If
'        Unload Me
    End If
Else
    MsgBox "请先输入仓别资料!", vbInformation, g_CON_IniFile
    Unload Me
End If
End Sub

Private Sub tvwDB_Expand(ByVal Node As ComctlLib.Node)
Node.Image = 2
Node.Sorted = True
End Sub

Private Sub tvwDB_NodeClick(ByVal Node As ComctlLib.Node)
    Set W_Node = Node
    W_key = Node.Key
    If W_key <> "A" Then
        W_Bar_No = W_Node.Key
        W_Bar_name = W_Node.Text
        W_Bar_ID = W_Node.Key
    End If
End Sub

Private Sub cmd_quit_Click()
Unload Me
End Sub

Private Sub Cmd_OK_Click()
    If W_key <> "A" Then
        Unload Me
    Else
        tvwdb.SetFocus
    End If
End Sub
