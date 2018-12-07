VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmUserList 
   BackColor       =   &H80000009&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "当前系统用户列表"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   7020
   Icon            =   "FrmUserList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Material List"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   8220
      Top             =   3330
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   1440
      Picture         =   "FrmUserList.frx":3C52
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Cancel"
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   180
      Picture         =   "FrmUserList.frx":51F4
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   90
      Width           =   1155
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   7005
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   6975
      _cx             =   12303
      _cy             =   12356
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   16711680
      BackColorSel    =   49152
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16711680
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   13
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   1
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7005
      Y1              =   510
      Y2              =   525
   End
End
Attribute VB_Name = "FrmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'用户资料查询
'*************************************************************************************************
'表单级的局部变量
Public W_Select_Data As String
Private W_Cancel_Status As Boolean
Private W_User_Count As Long
Public W_User_name As String
'只读属性返回值
Public Property Get user_count() As Long
    user_count = W_User_Count
End Property
'是否按了Cancel
Public Property Get cancel_status() As Boolean
    cancel_status = False
End Property

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        Dim i As Long
        Dim j As Long
        i = 1
        j = 0
        '置初值
        For i = 0 To 100
            W_User_List(i) = ""
        Next i
        '加入用户选择的值
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 0)) = -1 Then
                    W_User_List(j) = Trim(Grid1.TextMatrix(i, 2))
                j = j + 1
            End If
        Next i
        W_User_Count = j
        W_Cancel_Status = False
    Else
        W_User_Count = 0
        W_Cancel_Status = True
    End If
    '关闭窗口
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyReturn
       'SendKeys "{tab}"
    Case vbKeyF5               '确认
        Call Command1_Click(0)
    Case vbKeyEscape           '取消
        Call Command1_Click(1)
    End Select
End If
End Sub

Private Sub Form_Load()
Dim i As Long
'load图片
Set Me.Picture = G_MDIForm.Picture
'Call Set_Color_Frm(Me)
W_Cancel_Status = False
W_User_Count = 0
Me.KeyPreview = True

   
    
Call Select_date

    For i = 1 To Grid1.Rows - 1
        If InStr(1, W_User_name, Trim(Grid1.TextMatrix(i, 2))) > 0 Then
                Grid1.TextMatrix(i, 0) = -1
        End If
    Next i
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call ResizeListWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
W_Select_Data = ""
Set Grid1.DataSource = Nothing
End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub Grid1_DblClick()
'    Call Command1_Click(0)
End Sub

Private Sub Select_date()
On Error Resume Next
    Dim W_Rs As New ADODB.Recordset
    Dim W_Str As String
'
'    If mess002.CallForm.Name = "mmss304" Then
'        W_Str = "select cast(0 as bit) as 选取,mmst901.user_id as 用户编号,user_name as 用户名称  from mmstc03 " & _
'                                    "inner join mmstc02 on mmstc02.menu_id=mmstc03.menu_id " & _
'                                    " INNER JOIN  mmst901 ON MMSTC03.USER_ID=MMST901.user_id  " & _
'                                "where prog_id='304' and charindex('C',rights)>0 And MMST901.user_id<>'" & G_User_ID & "' order by mmst901.user_id "
'    ElseIf mess002.CallForm.Name = "mmss303" Then
'        W_Str = "select cast(0 as bit) as 选取,mmst901.user_id as 用户编号,user_name as 用户名称  from mmstc03 " & _
'                                    "inner join mmstc02 on mmstc02.menu_id=mmstc03.menu_id " & _
'                                    " INNER JOIN  mmst901 ON MMSTC03.USER_ID=MMST901.user_id  " & _
'                                "where prog_id='303' and charindex('C',rights)>0 And MMST901.user_id<>'" & G_User_ID & "' order by mmst901.user_id "
'    ElseIf mess002.CallForm.Name = "mmss343" Then
'        W_Str = "select cast(0 as bit) as 选取,mmst901.user_id as 用户编号,user_name as 用户名称  from mmstc03 " & _
'                                    "inner join mmstc02 on mmstc02.menu_id=mmstc03.menu_id " & _
'                                    " INNER JOIN  mmst901 ON MMSTC03.USER_ID=MMST901.user_id  " & _
'                                "where prog_id='343' and charindex('C',rights)>0 And MMST901.user_id<>'" & G_User_ID & "' order by mmst901.user_id "
'    ElseIf mess002.CallForm.Name = "mmss344" Then
'        W_Str = "select cast(0 as bit) as 选取,mmst901.user_id as 用户编号,user_name as 用户名称  from mmstc03 " & _
'                                    "inner join mmstc02 on mmstc02.menu_id=mmstc03.menu_id " & _
'                                    " INNER JOIN  mmst901 ON MMSTC03.USER_ID=MMST901.user_id  " & _
'                                "where prog_id='344' and charindex('C',rights)>0 And MMST901.user_id<>'" & G_User_ID & "' order by mmst901.user_id "
'    Else
    
        W_Str = " select cast(0 as bit) as 选取,user_id as 用户编号,user_name as 用户名称 " & _
                " from mmst901 " & _
                "  " & _
                " order by user_id "
'    End If
    Set W_Rs = open_RS(W_Str)
    Set Grid1.DataSource = W_Rs
    Grid1.Editable = flexEDKbdMouse
    Grid1.ExtendLastCol = True
End Sub
