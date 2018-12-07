VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "AresButtonPro.ocx"
Begin VB.Form mmss902 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "部门资料维护档(902)"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15165
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8325
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   2475
      TabIndex        =   14
      Top             =   810
      Width           =   2505
      Begin ARESBUTTONLib.AresButton cmd_ok 
         Height          =   360
         Left            =   690
         TabIndex        =   4
         Top             =   420
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
         ToolTipString   =   "确认存盘"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_cancel 
         Height          =   420
         Left            =   690
         TabIndex        =   5
         Top             =   1260
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
         ToolTipString   =   "放弃存盘"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_add 
         Height          =   420
         Left            =   690
         TabIndex        =   6
         Top             =   2115
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         ToolTipString   =   "增加一笔记录"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   127107100
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_edit 
         Height          =   420
         Left            =   690
         TabIndex        =   7
         Top             =   2955
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         ToolTipString   =   "修改该记录"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_delete 
         Height          =   420
         Left            =   690
         TabIndex        =   8
         Top             =   3810
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         ToolTipString   =   "删除该记录"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_quit 
         Height          =   360
         Left            =   690
         TabIndex        =   9
         Top             =   4650
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
         ToolTipString   =   "退出该程式"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton AresButton1 
         Height          =   825
         Left            =   780
         TabIndex        =   15
         Top             =   6420
         Width           =   1050
         _Version        =   327687
         BackGroundColor =   16777215
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\erp_proj.gif"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\erp_proj.gif"
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   90897964
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2490
         Y1              =   5370
         Y2              =   5370
      End
      Begin VB.Label Help_txt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   300
         TabIndex        =   16
         Top             =   5700
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   2490
      Picture         =   "mmss902.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   12645
      TabIndex        =   10
      Top             =   810
      Width           =   12675
      Begin VB.TextBox Dpt_Right 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2220
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1327
         Width           =   5700
      End
      Begin VB.TextBox dpt_id 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2220
         MaxLength       =   6
         TabIndex        =   0
         Top             =   510
         Width           =   1005
      End
      Begin VB.TextBox dpt_name 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6015
         MaxLength       =   8
         TabIndex        =   1
         Top             =   510
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "部门名称:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5100
         TabIndex        =   13
         Top             =   570
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "部门代号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   570
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "部门职能:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   1380
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   210
      Top             =   9330
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
      Bindings        =   "mmss902.frx":437B2
      Height          =   6195
      Left            =   2490
      TabIndex        =   3
      Top             =   2940
      Width           =   12675
      _cx             =   22357
      _cy             =   10927
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483634
      ForeColorFixed  =   -2147483630
      BackColorSel    =   65280
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"mmss902.frx":437C7
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   5
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "部门资料维护"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   465
      Left            =   7320
      TabIndex        =   17
      Top             =   180
      Width           =   2985
   End
End
Attribute VB_Name = "mmss902"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称: 部门资料档(MMSS902)(标准程序)
'*编写日期: 2002年07月29日
'*制作人员: 于
'*修改日期:
'*修改人员:
'***********************************************
'定义欲打开的数据库及数据表名称
Dim st_902 As New ADODB.Recordset

'定义按钮变量
Dim c_add As Boolean
Dim c_edit As Boolean
Dim c_delete As Boolean

'存放TDBGRID1 的旧字符
Dim W_Old_Str As String
Dim w_dpt_id As String

Dim c_off_add As Boolean
Dim c_off_edit As Boolean
Dim c_off_delete As Boolean

'纪录当前行列
Dim W_col As Double
Dim W_Row As Double

'定义窗体打开变量
Dim Gridc_902(127) As Grid_Data '存放 Grid 属性值
Dim Row_Height As Double        'Grid 高度变量

Private Sub dpt_id_LostFocus()
'定位处理
If Not (c_add Or c_edit) Then
    w_curr_row = TDBGrid1.Row
    w_find_row = TDBGrid1.FindRow(dpt_id.Text, 0, 1, False)
    If w_find_row > 0 Then
        TDBGrid1.TopRow = w_find_row
        TDBGrid1.Row = w_find_row
        TDBGrid1.Col = 1
    Else
        TDBGrid1.Row = w_curr_row
        Call readshow
    End If
    
End If
End Sub

Public Sub Form_Activate()
'当窗口激活时,刷新TDBGrid
Call GetVSGridSetting("mmss902", "TDBGrid1", Gridc_902, g_CON_IniFile9)
Row_Height = Gridc_902(0).Grid_RowHeight
Call readactive
'刷新表格
Call RefreshGrid
TDBGrid1.Col = 1
If TDBGrid1.Rows > 1 Then
    TDBGrid1.Row = 1
End If

End Sub

Private Sub Form_Load()
'装载图片
Call load_picture
'将窗口置中
Call CenterWindow(mmss902, sys_main)

'将按钮变量赋初值
c_add = False
c_edit = False
c_delete = False
c_off_add = False
c_off_edit = False
c_off_delete = False

'MDI子窗口按钮权限设订
c_off_add = lopcheck("A", "902", G_User_ID)
c_off_edit = lopcheck("U", "902", G_User_ID)
c_off_delete = lopcheck("D", "902", G_User_ID)

If c_off_add = True Then
    cmd_add.Enabled = False
End If

If c_off_edit = True Then
    cmd_edit.Enabled = False
End If

If c_off_delete = True Then
    cmd_delete.Enabled = False
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Form_KeyDown可照搬
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

If Shift = 0 Then
    Select Case KeyCode
        '按 ESC 时取消操作
        Case vbKeyEscape
            If (c_add Or c_edit Or c_delete) Then
                Call vcontrol("N")
                KeyCode = 0
            End If
        '当按 F2 时,新　记录
        Case vbKeyF2
            If cmd_add.Enabled Then
                Call vcontrol("A")
                KeyCode = 0
            End If
        '当按 F3 时 , 修改记录
        Case vbKeyF3
            If Me.ActiveControl.Name = "dpt_id" Then
                Call dpt_id_LostFocus
            End If

            If cmd_edit.Enabled Then
                Call vcontrol("U")
                KeyCode = 0
            End If
        '当按 F4 时 , 删除记录
        Case vbKeyF4
            If Me.ActiveControl.Name = "dpt_id" Then
                Call dpt_id_LostFocus
            End If

            If cmd_delete.Enabled Then
                Call vcontrol("D")
                KeyCode = 0
            End If
        '当按 F5 时,保存记录
        Case vbKeyF5
            If cmd_ok.Enabled Then
                Call vcontrol("Y")
                KeyCode = 0
            End If
        '当按 F6 时,退出系统
        Case vbKeyF6
            If cmd_quit.Enabled Then
                Call vcontrol("Q")
                KeyCode = 0
            End If
    End Select
End If
End Sub

Sub readshow()
'对控件置值
If c_add = True Or Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then
    dpt_id.Text = ""
    Dpt_Name.Text = ""
    Dpt_Right.Text = ""
Else
    dpt_id.Text = st_902!dpt_id
    Dpt_Name.Text = NullSetValue(st_902!Dpt_Name, "")
    Dpt_Right.Text = NullSetValue(st_902!Dpt_Right, "")
End If

'设定按键的 Enabled 属性
If c_add Or c_edit Or c_delete Then
    cmd_add.Enabled = False
    cmd_edit.Enabled = False
    cmd_delete.Enabled = False
    cmd_ok.Enabled = True
    cmd_cancel.Enabled = True
Else
    cmd_add.Enabled = True
    cmd_edit.Enabled = True
    cmd_delete.Enabled = True
    cmd_ok.Enabled = False
    cmd_cancel.Enabled = False
    
    '当数据表无记录时
    If st_902.EOF Then
        cmd_edit.Enabled = False
        cmd_delete.Enabled = False
    Else
        cmd_edit.Enabled = True
        cmd_delete.Enabled = True
    End If
End If

'通过权限设定按键的 Enabled
If c_off_add = True Then
    cmd_add.Enabled = False
End If
    
If c_off_edit = True Then
    cmd_edit.Enabled = False
End If
    
If c_off_delete = True Then
    cmd_delete.Enabled = False
End If

If Not (c_add Or c_edit) Then
    Dpt_Name.Locked = True
    Dpt_Right.Locked = True
Else
    Dpt_Name.Locked = False
    Dpt_Right.Locked = False
End If

If c_edit Then
    dpt_id.Locked = True
Else
    dpt_id.Locked = False
End If
End Sub
'刷新表格
Private Sub RefreshGrid()
Call readactive
Call readshow
End Sub

Private Sub readactive()
Set st_902 = Nothing
With st_902
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
    .Open "select dpt_id  ," & _
                 "dpt_name," & _
                 "dpt_right," & _
                 "upd_name ," & _
                 "upd_date " & _
            "FROM mmst902  ORDER BY dpt_id"

End With

'设置tdbgrid1的数据来源
Set Adodc1.Recordset = st_902

Call SetVSGridSetting(TDBGrid1, Gridc_902)

'刷新全部 ROW 的高度 包括 HEADER
For i = 1 To TDBGrid1.Rows
    TDBGrid1.Row = i - 1
    TDBGrid1.RowHeight(i - 1) = Row_Height
    
    If i < TDBGrid1.Rows Then
        TDBGrid1.TextMatrix(i, 0) = i
    End If
Next i
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

End Sub

'命令按键事件
Private Sub Cmd_Add_MouseClick()
Call vcontrol("A")
End Sub

Private Sub Cmd_Edit_MouseClick()
Call vcontrol("U")
End Sub

Private Sub Cmd_Delete_MouseClick()
Call vcontrol("D")
End Sub

Private Sub Cmd_OK_MouseClick()
Call vcontrol("Y")
End Sub

Private Sub cmd_cancel_MouseClick()
Call vcontrol("N")
End Sub

Private Sub cmd_quit_MouseClick()
Call vcontrol("Q")
End Sub

Private Sub vcontrol(p_choice As String)
Select Case p_choice
    Case "Y"            '确定
        If check_ok() = True Then
            Call upd_data
            TDBGrid1.Enabled = True
        End If
        
    Case "N"            '取消
        '解锁处理
        If c_edit Or c_delete Then
            Call UnLockRecord("mmst902", "dpt_id='" & dpt_id.Text & "'")
        End If
        c_add = False
        c_edit = False
        c_delete = False
        
        TDBGrid1.Enabled = True
        Call readshow
        
    Case "A"             '增加
        c_add = True
        Call readshow
        dpt_id.SetFocus
        TDBGrid1.Enabled = False
        
    Case "U"             '修改
        '加锁
        If LockRecord("mmst902", "dpt_id='" & dpt_id.Text & "'") Then
            W_Row = TDBGrid1.Row
            W_col = TDBGrid1.Col
            
            c_edit = True
            TDBGrid1.Enabled = False
            Call readshow
            Dpt_Name.SetFocus
        End If
        
    Case "D"             '删除
        '加锁
        If LockRecord("mmst902", "dpt_id='" & dpt_id.Text & "'") = True Then
            If MsgBox(g_CON_CDelete, vbYesNo + vbDefaultButton2 + vbInformation, g_CON_CTitle) = vbNo Then
                Call UnLockRecord("mmst902", "dpt_id='" & dpt_id.Text & "'")
                Exit Sub
            End If
            
            '判断是否可以删除
            c_delete = True
            If check_ok = False Then
                Call UnLockRecord("mmst902", "dpt_id='" & dpt_id.Text & "'")
                c_delete = False
                Exit Sub
            End If
            
            '删除记录
            G_Con.Execute "DELETE FROM mmst902 WHERE dpt_id='" & Trim(dpt_id.Text) & "'"
            c_delete = False
            '刷新数据
            
            Call RefreshGrid
            
            '删除后移动到第一笔记录
            TDBGrid1.Col = 1
            If TDBGrid1.Rows > 1 Then
                TDBGrid1.TopRow = 1
                TDBGrid1.Row = 1
            End If
            
        End If
    Case "Q"            '退出
        Unload Me
End Select
End Sub

'当修改或删除或新增时进行一致性判断
Private Function check_ok() As Boolean
Dim w_tmp As New ADODB.Recordset

If c_delete Then
    w_tmp.Open "select * from mmst901 where dpt_id = '" & dpt_id.Text & "'", G_Con
    If w_tmp.EOF = False Then
        MsgBox "该部门在用户资料中使用!", 64, "提示信息"
        Call vcontrol("N")
        check_ok = False
        w_tmp.Close
        Exit Function
    End If
    check_ok = True
    w_tmp.Close
End If
    
'新增时判断
If c_add = True Then
    If Trim(dpt_id.Text) = "" Then
        MsgBox "请输入部门标识", 64, "提示信息"
        dpt_id.SetFocus
        check_ok = False
        Exit Function
    Else
        '判断代号是否重复
        w_tmp.CursorLocation = adUseClient
        w_tmp.Open "select dpt_id from mmst902 where dpt_id = '" & Trim(dpt_id.Text) & "'", G_Con, adOpenForwardOnly
        If w_tmp.EOF = False Then
            MsgBox "部门标识重复!", 64, "提示信息"
            dpt_id.SetFocus
            check_ok = False
            Set w_tmp = Nothing
            Exit Function
        End If
        Set w_tmp = Nothing
    End If
End If

'新增和修改时判断
If Dpt_Name.Text = "" Then
    MsgBox "请输入部门名称", 64, "提示信息"
    Dpt_Name.SetFocus
    check_ok = False
    Exit Function
Else
    '判断部门名称是否重复
    w_tmp.CursorLocation = adUseClient
    w_tmp.Open "select dpt_name from mmst902 where dpt_name= '" & Trim(Dpt_Name.Text) & "' and dpt_id <> '" & dpt_id.Text & "'", G_Con, adOpenForwardOnly
    If w_tmp.EOF = False Then
        MsgBox "部门名称重复!", 64, "提示信息"
        Dpt_Name.SetFocus
        Set w_tmp = Nothing
        check_ok = False
        Exit Function
    End If
    Set w_tmp = Nothing
End If

check_ok = True
End Function

'对数据库进行更新
Private Sub upd_data()
Dim st_902_1 As New ADODB.Recordset
Dim W_Find As String

W_Find = dpt_id.Text

With st_902_1
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
    .Open "select * from mmst902 where dpt_id='" & dpt_id.Text & "'"
End With

'新增一笔记录到数据库
If c_add = True Then
    With st_902_1
        .AddNew
        !dpt_id = UCase(Trim(dpt_id.Text))
        !Dpt_Name = Dpt_Name.Text
        !Dpt_Right = Dpt_Right.Text
        
        !upd_name = Trim(G_User_Name)
        !upd_date = Get_SQLDATE
        !lock = "No"
        .Update
    End With
    Set st_902_1 = Nothing
    c_add = False
End If

'修改记录
If c_edit = True Then
    With st_902_1
        !Dpt_Name = Dpt_Name.Text
        !Dpt_Right = Dpt_Right.Text
        
        !upd_name = Trim(G_dpt_name)
        !upd_date = Get_SQLDATE
        !lock = "No"
        .Update
    End With
    Set st_902_1 = Nothing
    c_edit = False
End If

'刷新数据表
Call RefreshGrid

TDBGrid1.Row = TDBGrid1.FindRow(W_Find, 0, 1, False)
TDBGrid1.Col = W_col
TDBGrid1.TopRow = TDBGrid1.FindRow(W_Find, 0, 1, False)


End Sub

'表单的 QueryUnload 和 Unload 事件
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If c_add Or c_edit Or c_delete Then
    '当有数据改动时.询问是否要退出系统
    If MsgBox(g_CON_CQuit, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
        Cancel = 1
    Else
        '当有修改或删除时未解锁时,解除锁定
        If c_edit Or c_delete Then
            Call UnLockRecord("mmst902", "dpt_id='" & dpt_id.Text & "'")
        End If
        Cancel = 0
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

'退出时，存储 TDBGrid 属性
Call SaveGridSetting("mmss902", "TDBGrid1", Gridc_902, g_CON_IniFile9)

Set TDBGrid1.DataSource = Nothing
Set st_902 = Nothing
Set mmss902 = Nothing
End Sub

'各控件的相关事件
Private Sub dpt_id_KeyPress(KeyAscii As Integer)
'控制输入为字母或数字
If c_add Then
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
           (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or _
           (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = vbKeyBack Or _
            KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TDBGrid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow Then
    If NewRow >= 0 Then
        TDBGrid1.TextMatrix(OldRow, 0) = W_Old_Str
        W_Old_Str = TDBGrid1.TextMatrix(NewRow, 0)
        TDBGrid1.TextMatrix(NewRow, 0) = "★"
        TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
                
    End If
    '当点击TDBGRID1 cell 时,移动 ADODC1.Recordset 指针
    If Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Move TDBGrid1.Row - 1
        TDBGrid1.FocusRect = flexFocusRaised
    End If
    Call readshow
End If
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
End Sub

Private Sub TDBGrid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'移动COl改变宽度
If Col > 0 Then
    If Col > Gridc_902(0).Grid_Columns Then
        Cancel = 1
    Else
        If UCase(Mid(Gridc_902(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_902(Col - 1).Grid_Visible = "" Then
            Cancel = 1
        Else
            Gridc_902(Col - 1).Grid_Width = TDBGrid1.ColWidth(Col)
        End If
    End If
End If

'移动ROW改变高度
If Row >= 0 Then
    w_cur_row = TDBGrid1.Row
    Row_Height = TDBGrid1.RowHeight(Row)
    Gridc_902(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.Row = i - 1
        TDBGrid1.RowHeight(i - 1) = Row_Height
    Next i
    TDBGrid1.Row = w_cur_row
End If

End Sub

Private Sub TDBGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'鼠标点在HEADER上
If X > 0 And Y < Row_Height Then
   
    '存储 TDBGrid 属性
    Call SaveVSGridSetting("mmss902", "TDBGrid1", Gridc_902, g_CON_IniFile9)
    
    '调用 TDBGrid 属性设定
    With mmss_set
        Set .Parent_form = mmss902
        .Get_FormName = "mmss902"
        .Get_GridName = "TDBGrid1"
        .Gridc_File = g_CON_IniFile9
        .Show vbModal
    End With
End If
End Sub

Private Sub TDBGrid1_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'不许更改第0行COl的宽度
If Col = 0 Then
    Cancel = True
End If
End Sub


'**********************************************************************
'装载按钮图片
'**********************************************************************
Private Sub load_picture()
cmd_ok.PictureURL = App.Path + "\Picture\Norm\Ok_norm.bmp"
cmd_ok.PictureDisableURL = App.Path + "\Picture\Dis\Ok_dis.bmp"
cmd_ok.PictureOverURL = App.Path + "\Picture\Over\Ok_Over.bmp"

cmd_cancel.PictureURL = App.Path + "\Picture\Norm\cancel_norm.bmp"
cmd_cancel.PictureDisableURL = App.Path + "\Picture\Dis\cancel_dis.bmp"
cmd_cancel.PictureOverURL = App.Path + "\Picture\Over\cancel_Over.bmp"

cmd_add.PictureURL = App.Path + "\Picture\Norm\add_norm.bmp"
cmd_add.PictureDisableURL = App.Path + "\Picture\Dis\add_dis.bmp"
cmd_add.PictureOverURL = App.Path + "\Picture\Over\add_Over.bmp"

cmd_edit.PictureURL = App.Path + "\Picture\Norm\edit_norm.bmp"
cmd_edit.PictureDisableURL = App.Path + "\Picture\Dis\edit_dis.bmp"
cmd_edit.PictureOverURL = App.Path + "\Picture\Over\edit_Over.bmp"

cmd_delete.PictureURL = App.Path + "\Picture\Norm\delete_norm.bmp"
cmd_delete.PictureDisableURL = App.Path + "\Picture\Dis\delete_dis.bmp"
cmd_delete.PictureOverURL = App.Path + "\Picture\Over\delete_Over.bmp"

cmd_quit.PictureURL = App.Path + "\Picture\Norm\Quit_norm.bmp"
cmd_quit.PictureDisableURL = App.Path + "\Picture\Dis\Quit_dis.bmp"
cmd_quit.PictureOverURL = App.Path + "\Picture\Over\Quit_Over.bmp"

AresButton1.PictureURL = App.Path + "\Picture\file.gif"
AresButton1.GifAnimationPlay
End Sub

'**********************************************************************
'更改提示符
'**********************************************************************

Private Sub cmd_add_MouseEnter()
Help_txt.Caption = cmd_add.ToolTipString
Help_txt.Refresh

End Sub

Private Sub Cmd_Add_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_edit_MouseEnter()
Help_txt.Caption = cmd_edit.ToolTipString
Help_txt.Refresh
End Sub

Private Sub cmd_edit_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_delete_MouseEnter()
Help_txt.Caption = cmd_delete.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_delete_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_ok_MouseEnter()
Help_txt.Caption = cmd_ok.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_ok_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_cancel_MouseEnter()
Help_txt.Caption = cmd_cancel.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_cancel_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_quit_MouseEnter()
Help_txt.Caption = cmd_quit.ToolTipString
Help_txt.Refresh

End Sub

Private Sub Cmd_quit_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

'**********************************************************************
'按 ENTER调用 Mouse_click事件
'**********************************************************************
Private Sub Cmd_OK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_OK_MouseClick
End If
End Sub


Private Sub Cmd_cancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_cancel_MouseClick
End If
End Sub


Private Sub Cmd_add_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Add_MouseClick
End If
End Sub

Private Sub Cmd_edit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Edit_MouseClick
End If
End Sub

Private Sub Cmd_delete_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Delete_MouseClick
End If
End Sub

Private Sub Cmd_quit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_quit_MouseClick
End If
End Sub


