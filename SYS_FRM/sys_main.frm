VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "AresButtonPro.ocx"
Begin VB.MDIForm sys_main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000014&
   Caption         =   "系统管理"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7725
   Icon            =   "sys_main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Align           =   1  'Align Top
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   465
      Visible         =   0   'False
      Width           =   7725
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   7725
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin ARESBUTTONLib.AresButton AresButton1 
         Height          =   330
         Left            =   450
         TabIndex        =   1
         Top             =   120
         Width           =   345
         _Version        =   327687
         MoveOnDown      =   -1  'True
         ToolTipBackColor=   12648447
         ToolTipTextColor=   0
         ToolTipGradientColor=   12648447
         PictureURL      =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\关于.bmp"
         PictureOverURL  =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\关于1.bmp"
         PictureDownURL  =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\关于2.bmp"
         PictureBaseURL  =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\关于.bmp"
         ToolTipString   =   "关于系统"
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
         PictureRES      =   "sys_main.frx":0442
         PictureOverRES  =   "sys_main.frx":0AC4
         PictureDownRES  =   "sys_main.frx":1146
         HoldingFlag     =   7
         PrevPointer     =   220434756
         _ExtentX        =   609
         _ExtentY        =   582
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton Cmd_Quit 
         Height          =   330
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   345
         _Version        =   327687
         MoveOnDown      =   -1  'True
         ToolTipBackColor=   12648447
         ToolTipTextColor=   0
         ToolTipGradientColor=   12648447
         ToolTipBorderColor=   4210752
         PictureURL      =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\退出.bmp"
         PictureOverURL  =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\退出1.bmp"
         PictureDownURL  =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\退出2.bmp"
         PictureBaseURL  =   "Y:\c_sys\billy\xsh_erp\Picture\FRM_PICTURE\退出.bmp"
         ToolTipString   =   "退出基本资料系统"
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
         PictureRES      =   "sys_main.frx":17C8
         PictureOverRES  =   "sys_main.frx":1E4A
         PictureDownRES  =   "sys_main.frx":24CC
         HoldingFlag     =   7
         PrevPointer     =   220434756
         _ExtentX        =   609
         _ExtentY        =   582
         _StockProps     =   80
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000006&
         X1              =   0
         X2              =   12000
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   12000
         Y1              =   75
         Y2              =   75
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   1830
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSComctlLib.ImageList B_Imagelist 
      Left            =   60
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":2B4E
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":3092
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":35D6
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":36EA
            Key             =   "save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":3C2E
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":4172
            Key             =   "check"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":428A
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":47CE
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":4D12
            Key             =   "print"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":5256
            Key             =   "quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":536E
            Key             =   "help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":5482
            Key             =   "find"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys_main.frx":5596
            Key             =   "ok"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5895
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7761
            MinWidth        =   7761
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "12:22"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   810
      Visible         =   0   'False
      Width           =   7725
      _ExtentX        =   13626
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
   Begin VB.Menu menu_i1 
      Caption         =   "系统(&S)"
      Begin VB.Menu menu_i1_1 
         Caption         =   "关于ERP系统(&A)"
      End
      Begin VB.Menu menu_i1_2 
         Caption         =   "退出系统(&Q)"
      End
   End
   Begin VB.Menu menu_k1 
      Caption         =   "用户管理(&O)"
      Begin VB.Menu menu_k1_1 
         Caption         =   "部门资料设定(&A)"
      End
      Begin VB.Menu menu_k1_2 
         Caption         =   "用户资料设定(&B)"
      End
   End
   Begin VB.Menu menu_k2 
      Caption         =   "系统管理"
      Begin VB.Menu menu_k2_1 
         Caption         =   "用户权限设定"
      End
      Begin VB.Menu menu_k2_a 
         Caption         =   "用户提示授权"
      End
      Begin VB.Menu menu_k2_b 
         Caption         =   "单据审核消息反馈授权"
      End
      Begin VB.Menu menu_k2_2 
         Caption         =   "系统数据解锁"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_i3_3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_k2_3 
         Caption         =   "公司名称维护"
      End
      Begin VB.Menu menu_k2_4 
         Caption         =   "用户密码修改"
      End
      Begin VB.Menu menu_k2_c 
         Caption         =   "物料料号冻结"
      End
      Begin VB.Menu menu_k2_d 
         Caption         =   "物料料号替代"
      End
      Begin VB.Menu menu_k2_5 
         Caption         =   "物料料号修改"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_k2_7 
         Caption         =   "客户编号修改"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_k2_8 
         Caption         =   "厂商编号修改"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_k2_9 
         Caption         =   "帐务重复删除"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_k2_6 
         Caption         =   "系统补丁"
      End
   End
   Begin VB.Menu menu_i4 
      Caption         =   "数据维护(&M)"
      Visible         =   0   'False
      Begin VB.Menu menu_i4_1 
         Caption         =   "系统数据备份(&B)"
      End
      Begin VB.Menu menu_i4_2 
         Caption         =   "系统数据恢复(&R)"
      End
   End
   Begin VB.Menu menu_v 
      Caption         =   "检视(&V)"
      Enabled         =   0   'False
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu menu_v1 
         Caption         =   "工具栏(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu menu_v2 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "sys_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
If App.PrevInstance = True Then
    End
End If


'取得 Window 目录
G_Windir = GetWinDir()
G_Path = App.Path
'取得电脑名称
G_Pc_Name = Get_ComputerName

G_system_id = 12
Call Get_mod_setting

'**********************************************************************************
'判断用户 连接并取得用户连接
If Not Check_Login_Status_conn Then
    End
End If

'**********************************************************************************
'取得连接属性
If Not Get_SQL_conn Then
    End
End If


'**********************************************************************************
'取得公司信息
Call Get_comp_info
'取得MDI窗体
Set G_MDIForm = GetMdiForm
'设置MDI窗体的StatusBar
Call Set_MDI_StatusBar
'设置MDI窗体的特殊字符
'Call Set_MDI_Comb_Singn



'**********************************************************************************
'获得用户详细资料,包含权限
'G_User_Data.User_Id = G_User_ID
'G_User_Data = Get_User_Data(G_User_Data.User_Id)


 '**********************************************************************************
'刷新MDI权限
Call Set_init_form

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)

Call Set_Mdi_unload

Call ActiveMainEXE
End Sub



Private Sub AresButton1_MouseClick()
aboutmms.Show 1
End Sub


'各菜单事件
Private Sub menu_add_Click()
Call sys_main.ActiveForm.menu_add_Click
End Sub

Private Sub menu_Delete_Click()
Call sys_main.ActiveForm.menu_Delete_Click
End Sub

Private Sub menu_edit_Click()
Call sys_main.ActiveForm.menu_edit_Click
End Sub

Private Sub cmd_quit_MouseClick()
Call menu_i1_2_Click
End Sub

Private Sub menu_i1_1_Click()
aboutmms.Show 1
End Sub

Private Sub menu_i1_2_Click()
Dim w_title As String
Dim w_info As String
If g_Language = "C" Then
    w_title = "你真要退出吗?"
    w_info = "提示信息"
Else
    w_title = "Do you really quit?"
    w_info = "Information"
End If
If MsgBox(w_title, vbYesNo + vbQuestion, w_info) = vbNo Then
    Exit Sub
End If

Unload Me

End Sub
Private Sub menu_k1_1_Click()
mmss902.ZOrder 0
End Sub
Private Sub menu_k1_2_Click()
mmss901.ZOrder 0
End Sub
Private Sub menu_k2_1_Click()
mmss907.ZOrder 0

End Sub
Private Sub menu_k2_2_Click()
mmss903.ZOrder 0
End Sub
Private Sub menu_k2_3_Click()
mmss810.Show
End Sub

Private Sub menu_k2_4_Click()
FrmPassEdit.Show vbModal
End Sub

Private Sub menu_k2_5_Click()
mmss904.ZOrder 0
End Sub

Private Sub menu_k2_6_Click()

FrmPack.Show 1
End Sub

Private Sub menu_k2_7_Click()
mmss905.ZOrder 0
End Sub

Private Sub menu_k2_8_Click()
mmss908.ZOrder 0
End Sub

Private Sub menu_k2_9_Click()
mmss909.ZOrder 0
End Sub

Private Sub menu_k2_a_Click()
mmss90a.ZOrder 0
End Sub

Private Sub menu_k2_b_Click()
mmss90b.ZOrder 0
End Sub

Private Sub menu_k2_c_Click()
mmss90c.ZOrder 0

End Sub

Private Sub menu_k2_d_Click()
mmss904.ZOrder 0
End Sub

Private Sub menu_v2_Click()
If menu_v2.Checked Then
    StatusBar1.Visible = False
    menu_v2.Checked = False
Else
    StatusBar1.Visible = True
    menu_v2.Checked = True
End If
End Sub

Private Sub menu_v1_Click()
If menu_v1.Checked Then
    CoolBar.Visible = False
    menu_v1.Checked = False
Else
    CoolBar.Visible = True
    menu_v1.Checked = True
End If
End Sub

