VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form mmss90a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "单据提示授权(90a)"
   ClientHeight    =   6180
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   907
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   6855
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6765
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   180
         Left            =   3960
         TabIndex        =   8
         Top             =   285
         Visible         =   0   'False
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label user_name 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2415
         TabIndex        =   10
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "用户代号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   4740
      TabIndex        =   1
      Top             =   690
      Width           =   2025
      Begin VB.CommandButton Command1 
         Caption         =   " 确 认(&Y)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   300
         TabIndex        =   4
         Top             =   480
         Width           =   1500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取 消(&N)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   300
         TabIndex        =   3
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         Caption         =   "退 出(Q)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   300
         TabIndex        =   2
         Top             =   4605
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   30
      TabIndex        =   0
      Top             =   690
      Width           =   4545
      Begin MSComctlLib.TreeView tvwDB 
         Height          =   5205
         Left            =   45
         TabIndex        =   9
         Top             =   180
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   9181
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
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
   End
End
Attribute VB_Name = "mmss90a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ST_Tmp As New ADODB.Recordset
Dim W_Text As String
Dim User_List As Integer

Private Type TmpData
     D_List As String
     Inv_Type As String
     Rights As Boolean
End Type

Dim W_TmpData() As TmpData


Private Sub Form_Load()
Dim ST_901 As New ADODB.Recordset

With mmss90a
    .Left = Int(sys_main.Width - mmss90a.Width) / 2
    .Top = Int(sys_main.Height - mmss90a.Height - 1500) / 2
End With

With ST_901
    .ActiveConnection = G_Con
    .Open "select * from mmst901 order by user_id"
    Do While .EOF = False
        Combo1.AddItem .Fields("user_id").Value
        .MoveNext
    Loop
    .Close
End With
Set ST_901 = Nothing

If Combo1.ListCount > 0 Then
    Combo1.ListIndex = 0
End If


End Sub

Private Sub Combo1_Click()
Dim i As Integer
Dim M As Integer
Dim w_901 As New ADODB.Recordset

w_901.ActiveConnection = G_Con
w_901.Open "select user_name,list_no as d_list from mmst901 where user_id = '" & Combo1.Text & "'"

If w_901.EOF Then
    user_name.Caption = ""
Else
    user_name.Caption = w_901!user_name
    User_List = w_901!D_List
End If

w_901.Close


i = 0
If W_Text <> Combo1.Text Then

    MyStr = "SELECT D_List,Inv_Type,table_name " & _
            " From mmst905 " & _
             " ORDER BY Inv_Type "
             
    Set ST_Tmp = Nothing
    With ST_Tmp
        .ActiveConnection = G_Con
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open MyStr
    End With
    
    If Not ST_Tmp.EOF Then
       If ST_Tmp.AbsolutePosition <> -1 Then
           ReDim W_TmpData(ST_Tmp.RecordCount)
           ST_Tmp.MoveFirst
       End If
       
       Do Until ST_Tmp.EOF
          W_TmpData(i).D_List = CStr(ST_Tmp!D_List)
          W_TmpData(i).Inv_Type = ST_Tmp!Inv_Type
          ST_Tmp.MoveNext
          i = i + 1
       Loop
       
       Call treeshow
    End If
End If
W_Text = Combo1.Text
Call Ini_form
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 13
End Sub

Private Sub Command1_Click()
Dim Tmp_RB As New ADODB.Recordset
Dim i As Integer
i = 2
For i = 2 To tvwDB.Nodes.Count
    If tvwDB.Nodes(i).Checked Then
       If Not W_TmpData(i - 2).Rights Then
           G_Con.Execute ("Insert into mmst906(User_list,Inv_list) values('" & User_List & "','" & W_TmpData(i - 2).D_List & "')")
       End If
    Else
       If W_TmpData(i - 2).Rights Then
           G_Con.Execute ("delete from  mmst906 where user_list='" & User_List & "' and inv_list='" & W_TmpData(i - 2).D_List & "'")
       End If
    End If
Next

Call treeshow

Command2.Enabled = False
Command1.Enabled = False

MsgBox "授权已经完成", 48, "提示"
End Sub

Private Sub Command2_Click()
tvwDB.SetFocus
MsgBox "取消完成", 48, "提示"
Command1.Enabled = False
Command2.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub treeshow()
Dim M As Integer
M = 0
tvwDB.Nodes.Clear
Set Nodx = tvwDB.Nodes.Add(, , "mMENU", "系统提示审核授权栏")
    Do Until M = UBound(W_TmpData())
    
           Set Nodx = tvwDB.Nodes.Add("mMENU", tvwChild, "M" & Trim(W_TmpData(M).D_List), Trim(W_TmpData(M).Inv_Type))
                      
           If CheckRight("'" & User_List & "'", Trim(W_TmpData(M).D_List)) Then
             tvwDB.Nodes(M + 2).Checked = True
             W_TmpData(M).Rights = True
           Else
             W_TmpData(M).Rights = False
           End If
           
           Nodx.Tag = M
      M = M + 1
    Loop
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
Set ST_Tmp = Nothing
End Sub

Private Sub tvwDB_NodeCheck(ByVal node As MSComctlLib.node)
Dim i As Integer
i = 1
If node.Index = 1 Then
    Do While i <= tvwDB.Nodes.Count
       tvwDB.Nodes(i).Checked = tvwDB.Nodes(1).Checked
       i = i + 1
    Loop
End If
Command1.Enabled = True
Command2.Enabled = True
End Sub

Private Function CheckRight(User_List As String, Inv_list As String) As Boolean
Dim W_rs As New ADODB.Recordset

W_rs.Open "select user_list from mmst906 where user_list=" & User_List & " And inv_list='" & Inv_list & "' ", G_Con
CheckRight = (W_rs.EOF = False)
W_rs.Close
Set W_rs = Nothing
End Function

Private Function Ini_form()
Dim i As Integer
i = 1

For i = 1 To tvwDB.Nodes.Count
  If W_TmpData(i - 1).Rights Then
     tvwDB.Nodes(i + 1).Checked = True
  End If
Next

End Function


