VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "AresButtonPro.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.MDIForm Erp_Deliv 
   BackColor       =   &H80000014&
   Caption         =   "出货管理"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8880
   Icon            =   "Erp_Deliv.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "Erp_Deliv.frx":030A
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
      Top             =   795
      Visible         =   0   'False
      Width           =   8880
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
      ScaleWidth      =   8880
      TabIndex        =   0
      Top             =   0
      Width           =   8880
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
         PictureRES      =   "Erp_Deliv.frx":43ABC
         PictureOverRES  =   "Erp_Deliv.frx":4413E
         PictureDownRES  =   "Erp_Deliv.frx":447C0
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
         PictureRES      =   "Erp_Deliv.frx":44E42
         PictureOverRES  =   "Erp_Deliv.frx":454C4
         PictureDownRES  =   "Erp_Deliv.frx":45B46
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
      Begin MSForms.ComboBox Comb_Singn 
         Height          =   315
         Left            =   10020
         TabIndex        =   4
         Top             =   135
         Width           =   1845
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3254;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   3
         FontName        =   "新细明体"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "特殊符号:"
         Height          =   195
         Left            =   9180
         TabIndex        =   3
         Top             =   210
         Width           =   825
      End
   End
   Begin Crystal.CrystalReport Rpt1 
      Left            =   0
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin MSComctlLib.ImageList B_Imagelist 
      Left            =   510
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
            Picture         =   "Erp_Deliv.frx":461C8
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":4670C
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":46C50
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":46D64
            Key             =   "save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":472A8
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":477EC
            Key             =   "check"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":47904
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":47E48
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":4838C
            Key             =   "print"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":488D0
            Key             =   "quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":489E8
            Key             =   "help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":48AFC
            Key             =   "find"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Erp_Deliv.frx":48C10
            Key             =   "ok"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6465
      Width           =   8880
      _ExtentX        =   15663
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
            TextSave        =   "14:25"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   465
      Visible         =   0   'False
      Width           =   8880
      _ExtentX        =   15663
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
   Begin VB.Menu menu_s1 
      Caption         =   "系统(&S)"
      Begin VB.Menu menu_s1_1 
         Caption         =   "退出系统(&Q)"
      End
      Begin VB.Menu sp_qqqq 
         Caption         =   "-"
      End
      Begin VB.Menu menu_s1_2 
         Caption         =   "关于ERP系统(&A)"
      End
   End
   Begin VB.Menu menu_i1 
      Caption         =   "成品管理(&A)"
      Visible         =   0   'False
      Begin VB.Menu menu_i1_1 
         Caption         =   "成品入库单(&A)"
      End
      Begin VB.Menu menu_i1_2 
         Caption         =   "成品出库单(&B)"
      End
      Begin VB.Menu menu_sp3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_i1_3 
         Caption         =   "成品入库一级审核(&C)"
      End
      Begin VB.Menu menu_i1_4 
         Caption         =   "成品出库一级审核(&D)"
      End
      Begin VB.Menu GGG 
         Caption         =   "-"
      End
      Begin VB.Menu menu_i1_5 
         Caption         =   "成品入库二级审核(&E)"
      End
      Begin VB.Menu menu_i1_6 
         Caption         =   "成品出库二级审核(&F)"
      End
   End
   Begin VB.Menu menu_i2 
      Caption         =   "出货管理(&B)"
      Begin VB.Menu menu_i2_1 
         Caption         =   "成品出货单(&A)"
      End
      Begin VB.Menu menu_i2_2 
         Caption         =   "成品退货单(&B)"
      End
      Begin VB.Menu menu_sp236 
         Caption         =   "-"
      End
      Begin VB.Menu menu_i2_3 
         Caption         =   "成品出货一级审核(&C)"
      End
      Begin VB.Menu menu_i2_4 
         Caption         =   "成品退货一级审核(&D)"
      End
      Begin VB.Menu ddddd 
         Caption         =   "-"
      End
      Begin VB.Menu menu_i2_5 
         Caption         =   "成品出货二级审核(&E)"
      End
      Begin VB.Menu menu_i2_6 
         Caption         =   "成品退货二级审核(&F)"
      End
      Begin VB.Menu dddddd 
         Caption         =   "-"
      End
      Begin VB.Menu menu_i2_7 
         Caption         =   "成品出退货查询(&G)"
      End
      Begin VB.Menu menu_i2_8 
         Caption         =   "成品出退货(不含单¤)查询(&G)"
      End
   End
   Begin VB.Menu menu_i3 
      Caption         =   "单据查询(&C)"
      Visible         =   0   'False
      Begin VB.Menu menu_i3_1 
         Caption         =   "成品入库查询(&B)"
      End
      Begin VB.Menu menu_i3_2 
         Caption         =   "成品出货查询(&B)"
      End
   End
   Begin VB.Menu menu_i4 
      Caption         =   "发票制作(&D)"
      Visible         =   0   'False
      Begin VB.Menu menu_i4_1 
         Caption         =   "发票单制作(&A)"
      End
      Begin VB.Menu menu_i4_2 
         Caption         =   "发票单查询(&B)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_Url 
      Caption         =   "登陆我们网站(&W)"
      Visible         =   0   'False
   End
   Begin VB.Menu Menu_Email 
      Caption         =   "寄信给我们(&Email)"
      Visible         =   0   'False
   End
   Begin VB.Menu menu_v 
      Caption         =   "检视(&V)"
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
   Begin VB.Menu menu_modify 
      Caption         =   "编辑"
      Visible         =   0   'False
      Begin VB.Menu menu_add 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu menu_edit 
         Caption         =   "修改(&U)"
      End
      Begin VB.Menu menu_delete 
         Caption         =   "删除(&D)"
      End
   End
   Begin VB.Menu menu_picture 
      Caption         =   "图片"
      Visible         =   0   'False
      Begin VB.Menu menu_loadpic 
         Caption         =   "加载图片"
      End
      Begin VB.Menu menu_sp20 
         Caption         =   "-"
      End
      Begin VB.Menu menu_unloadpic 
         Caption         =   "移除图片"
      End
   End
End
Attribute VB_Name = "Erp_Deliv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AresButton1_MouseClick()
aboutmms.Show 1
End Sub

Private Sub MDIForm_Load()
If App.PrevInstance = True Then
    End
End If

Dim W_000 As New ADODB.Recordset
Dim W_SQL As DAO.Recordset
Dim W_RS As New ADODB.Recordset

Dim W_Lre As Long

Dim W_Server As String   '用户名
Dim W_Data As String     '数据库名
Dim W_Uid As String      '用户ID
Dim W_Pass As String     '用户密码
Dim W_Conn As String     '连接字符串

'取得 Window 目录
G_Windir = GetWinDir()
G_Path = App.Path
'打开临时打印库
Set G_PrintDb = OpenDatabase(G_Windir & g_CON_PrintPath)
Set mms_run = G_PrintDb.OpenRecordset("SELECT * FROM mms_run")
'Set G_TitleDb = OpenDatabase(G_Path & g_CON_TitlePath)
'Set mms_run = G_TitleDb.OpenRecordset("select * from mms_run")
'取得登入用户 ID 和 用户名称
If mms_run.EOF Then
    MsgBox "请执行主程序!" & Chr(10) & Chr(13) & "Please run the main program first!", 64, "提示信息(Information)"
    End
Else
    G_User_ID = mms_run!G_User_ID
    G_User_Name = mms_run!G_UserName
End If

'取得连接属性
Set W_SQL = G_PrintDb.OpenRecordset("SELECT * FROM sql_data")

'Set W_Sql = G_TitleDb.OpenRecordset("select * from sql_data")
If W_SQL.EOF = False Then
    W_Server = NullSetValue(W_SQL!server_name, "chance")
    W_Data = NullSetValue(W_SQL!data_name, "xc_mms")
    W_Uid = NullSetValue(W_SQL!user_id, "sa")
    W_Pass = NullSetValue(W_SQL!user_pass, "")
    W_Conn = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & W_Uid & ";Initial Catalog=" & W_Data & ";Data Source=" & W_Server & ";Pwd=" & W_Pass
    G_Con.CursorLocation = adUseClient
    G_Con.Open W_Conn
    
    G_Con.Execute "delete from userlong where user_name='" & G_User_Name & "' and code_name='出货模块' "

    Set W_RS = Nothing
    W_RS.Open "select * from userlong ", G_Con, adOpenKeyset, adLockOptimistic
    W_RS.AddNew
    W_RS!user_name = G_User_Name
    W_RS!code_name = "出货模块"
    W_RS!upd_date = Date
    W_RS.Update
    W_RS.Close
    
End If

'取得公司相关属性值
W_000.Open "SELECT * FROM mmst000", G_Con, adOpenForwardOnly

G_Loc_Com = W_000!loc_name
G_Loc_Tel = W_000!cmp_tel
G_Loc_Fax = W_000!cmp_fax

w_cmp_cname = W_000!cmp_cname

'Form_sale_flow.Show
Set W_000 = Nothing

StatusBar1.Panels(1).Picture = Me.Icon
StatusBar1.Panels(1).Text = "公司名称:" & w_cmp_cname
StatusBar1.Panels(2).Text = "开发商: " & "同盛软件有限公司"
StatusBar1.Panels(3).Text = "当前用户:(" & G_User_ID & ")" & G_User_Name
StatusBar1.Panels(4).Text = "当前日期:" & DateToChar(Date) & "   " & WeekName(Weekday(Date, vbMonday))

A_Sign(0).decript = "货币数字符号"
Set A_Sign(0).form_Name = mmssigns

A_Sign(1).decript = "希腊拉丁符号"
Set A_Sign(1).form_Name = mmssigns1

A_Sign(2).decript = "其他常用符号"
Set A_Sign(2).form_Name = mmssigns2

Comb_Singn.AddItem A_Sign(2).decript
Comb_Singn.AddItem A_Sign(1).decript
Comb_Singn.AddItem A_Sign(0).decript

Set G_MDIForm = GetMdiForm


Call init_form

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next

If G_Con.State = adStateOpen Then
    G_Con.Close
End If
Set G_Con = Nothing
Call ActiveMainEXE
End Sub

Private Sub Cmd_quit_MouseClick()
If MsgBox(g_CON_CQuit, vbYesNo + vbQuestion, g_CON_CTitle) = vbNo Then
    Exit Sub
End If
Unload Me
End Sub


'各菜单事件
Private Sub menu_a1_1_Click()
If MsgBox(g_CON_CQuit, vbYesNo + vbQuestion, g_CON_CTitle) = vbNo Then
    Exit Sub
End If
Unload Me
End Sub

Private Sub menu_add_Click()
On Error Resume Next
Call Erp_Deliv.ActiveForm.menu_add_Click
End Sub

Private Sub menu_delete_Click()
On Error Resume Next
Call Erp_Deliv.ActiveForm.menu_delete_Click
End Sub

Private Sub menu_edit_Click()
On Error Resume Next
Call Erp_Deliv.ActiveForm.menu_edit_Click
End Sub

Private Sub Menu_Email_Click()
   On Error Resume Next
    Dim Success As Long
    Dim URL As String
    URL = "mailto:"
    URL = URL & "szlifeng@pub.dgnet.gd.cn"
    Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", 0)

End Sub

Private Sub menu_i1_2_Click()
mmss602.ZOrder 0
End Sub

Private Sub menu_i1_3_Click()
    mmss641.ZOrder 0
End Sub

Private Sub menu_i1_4_Click()
    mmss642.ZOrder 0
End Sub

Private Sub menu_i1_5_Click()
mmss681.ZOrder 0
End Sub

Private Sub menu_i1_6_Click()
mmss682.ZOrder 0
End Sub

Private Sub menu_i2_1_Click()
mmss603.ZOrder 0
End Sub

Private Sub menu_i2_2_Click()
    mmss604.ZOrder 0
End Sub

Private Sub menu_i2_3_Click()
    mmss643.ZOrder 0
End Sub

Private Sub menu_i2_4_Click()
    mmss644.ZOrder 0
End Sub

Private Sub menu_i2_5_Click()
mmss683.ZOrder 0
End Sub

Private Sub menu_i2_6_Click()
mmss684.ZOrder 0
End Sub

Private Sub menu_i2_7_Click()
mmss653.ZOrder 0
End Sub

Private Sub menu_i2_8_Click()
    mmss654.ZOrder 0
End Sub

Private Sub menu_i3_1_Click()
    mmss651.ZOrder 0
End Sub

Private Sub menu_i3_2_Click()
    mmss653.ZOrder 0
End Sub

Private Sub menu_i4_1_Click()
    mmss606.ZOrder 0
End Sub

Private Sub menu_i4_2_Click()
    mmss656.ZOrder 0
End Sub

Private Sub menu_loadpic_click()
On Error Resume Next
Call Erp_Deliv.ActiveForm.menu_loadpic_click
End Sub

Private Sub menu_s1_1_Click()
If MsgBox(g_CON_CQuit, vbYesNo + vbQuestion, g_CON_CTitle) = vbNo Then
    Exit Sub
End If
Unload Me

End Sub

Private Sub menu_s1_2_Click()
aboutmms.Show 1
End Sub

Private Sub menu_unloadpic_click()
On Error Resume Next
Call Erp_Deliv.ActiveForm.menu_unloadpic_click
End Sub

Private Sub Menu_Url_Click()
    On Error Resume Next
    Dim Success As Long
    Dim URL As String
    URL = "http://www.newchancesoft.com"
    Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", 0)


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
Private Sub menu_i1_1_Click()
    mmss601.ZOrder 0
End Sub

Public Sub init_form()
a = Me.Count
Dim wbook As Variant
Dim wbook_b As Boolean
Dim w_tmp As New ADODB.Recordset
w_tmp.Open "select * from mmstc03 where user_id = '" & G_User_ID & "'  ", G_Con, adOpenKeyset

wbook_b = False
If w_tmp.EOF = False Then
   wbook_b = True
   wbook = w_tmp.Bookmark
End If
        
Do While w_tmp.EOF = False
    If InStr(1, w_tmp!rights, "O") = 0 Then
        For i = 0 To a - 1
            If UCase(Trim(Me(i).Name)) = UCase(Trim(w_tmp!menu_id)) Then
                Me(i).Enabled = False
              
            End If
        Next
    Else
        For i = 0 To a - 1
            If UCase(Trim(Me(i).Name)) = UCase(Trim(w_tmp!menu_id)) Then
                Me(i).Enabled = True
               
            End If
        Next
    End If
    w_tmp.MoveNext
Loop

If wbook_b Then
   w_tmp.Bookmark = wbook
End If

If w_tmp.EOF = True Then
   For i = 0 To a - 1
      If UCase(Left(Trim(Me(i).Name), 6)) = UCase("menu_i") Then
         Me(i).Enabled = False
      End If
   Next
End If



End Sub

Private Sub Comb_singn_Click()
Select Case Comb_Singn.Text
    Case A_Sign(0).decript
        A_Sign(0).form_Name.Show
    Case A_Sign(1).decript
        A_Sign(1).form_Name.Show
    Case A_Sign(2).decript
        A_Sign(2).form_Name.Show
End Select

End Sub


