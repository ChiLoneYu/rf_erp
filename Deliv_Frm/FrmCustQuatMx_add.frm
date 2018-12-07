VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmCustQuatMx_add 
   BackColor       =   &H00FFFFFF&
   Caption         =   "批量报价明细资料"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   15495
   LinkTopic       =   "Form1"
   Picture         =   "FrmCustQuatMx_add.frx":0000
   ScaleHeight     =   6435
   ScaleWidth      =   15495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "新增空白记录"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      Left            =   3600
      TabIndex        =   11
      Top             =   270
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12600
      TabIndex        =   10
      Top             =   270
      Width           =   1020
   End
   Begin VB.TextBox Cnt_amt 
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
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   1200
      MaxLength       =   21
      TabIndex        =   3
      ToolTipText     =   "不能超过21个字符"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Cust_Order_No 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4890
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9615
      Picture         =   "FrmCustQuatMx_add.frx":0DA3
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Cancel"
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton CmdOK 
      Height          =   375
      Left            =   7845
      Picture         =   "FrmCustQuatMx_add.frx":2345
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&OK"
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton cmd_brow_order 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      Left            =   3000
      TabIndex        =   1
      Top             =   270
      Width           =   300
   End
   Begin VB.TextBox mtr_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1200
      MaxLength       =   21
      TabIndex        =   0
      ToolTipText     =   "不能超过21个字符"
      Top             =   270
      Width           =   2115
   End
   Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
      Height          =   5625
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   15345
      _cx             =   27067
      _cy             =   9922
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
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   65280
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   0
      TreeColor       =   255
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
      Rows            =   20
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmCustQuatMx_add.frx":38E7
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      OwnerDraw       =   0
      Editable        =   2
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   240
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Image pic_cmd 
         Height          =   330
         Left            =   480
         Picture         =   "FrmCustQuatMx_add.frx":39CF
         Top             =   -1320
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "pcs/cnt:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   6630
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "订单单号:"
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
      Left            =   3975
      TabIndex        =   7
      Top             =   330
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "产品型号:"
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
      Left            =   270
      TabIndex        =   6
      Top             =   315
      Width           =   885
   End
End
Attribute VB_Name = "FrmCustQuatMx_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称: 产品出退货明细 (mmss502)
'*编写日期:
'*制作人员:
'*修改日期:
'*修改人员:
'***********************************************

Dim W_UpdateMode As Byte '0:add 出货,1: add 退货,3:edit 出货,4:edit 退货

Dim W_CallForm As Form

Dim W_Bar_No As String
Public W_LIST_NO As Double

Public w_mtr_name As String
Public W_Cust_No As String
Public W_Cust_mtr_name As String
Dim Gridc_frmquatmx_1(127) As Grid_Data

Dim Row_Height As Double

Public Property Get UpdateMode() As Byte
    UpdateMode = W_UpdateMode
End Property

Public Property Let UpdateMode(b As Byte)
    W_UpdateMode = b
    Select Case b
        Case 0
            Me.Caption = "新增产品出货明细"
        Case 1
            Me.Caption = "修改产品出货明细"
            Mtr_Name.Locked = True
    End Select
End Property

Public Property Get CallForm() As Form
    Set CallForm = W_CallForm
End Property

Public Property Set CallForm(f As Form)
    Set W_CallForm = f
End Property

Private Sub cmd_brow_order_Click()
    If Me.UpdateMode = 0 Then
        With FrmSectList_bat
             .W_edit_able = True
             .Quer_status = True
'             .W_Field2 = "mtr_name"
'             .W_Field3 = "mtr_dim"
             .W_Select_Data = " select '' as 选择,mtr_name as 产品型号,mtr_no as 物料代号,mtr_dim as 产品规格 ," & _
                                     " 'pcs' as 单位" & _
                              " from mmsp611 " & _
                              " where mtr_type like 'A008%'  " & _
                                    "  and mtr_name like '" & Trim(Mtr_Name.Text) & "%' " & _
                              " and mtr_no not in (select mtr_no from mmst034_tmp where pc_name='" & G_Pc_Name & "')  " & _
                              " and mtr_no not in (select mtr_no from mmst034 where quat_no='" & Me.CallForm.quat_no.Text & "')  "
'             .Grid1.Editable = flexEDNone
             .W_Forma_name = "frmquatmx_add"
             .Show vbModal
             If .cancel_status = False And .List1 <> "" Then
                Mtr_Name.Text = .List1
'                Cust_mtr_name.Text = .List2
'                cust_mtr_no.Text = .List3
'                old_mtr_no.Text = .List4
'                Call mtr_name_LostFocus
                Call RefreshGrid
'                mtr_amt.SetFocus
             End If
        End With
    End If
End Sub
Private Sub RefreshGrid()

Dim W_Rd As New ADODB.Recordset
Dim StrSQL As String

StrSQL = "sELECT  mtr_no,mtr_name,mtr_dim,item_no,color_name,mtr_prs,g_weight,n_weight," & _
            " Cont_Amt , Box_Dim,  OUT_STATUS,  SIZE, Cuft, Cart_Dim, Fabric, Packing, FT_40, HQ_40, HQ_45, FT_20,NOTE ,D_LIST " & _
            " FROM MMST034_TMP where pc_name='" & G_Pc_Name & "'" & _
        "  "

Set W_Rd = Open_Rs(StrSQL)

Set Adodc1.Recordset = W_Rd
Set TDBGrid1.DataSource = Adodc1

'Call readactive
Call SetVSGridSetting(TDBGrid1, Gridc_frmquatmx_1)



For i = 1 To TDBGrid1.Rows
    TDBGrid1.RowHeight(i - 1) = Row_Height
    If i < TDBGrid1.Rows Then
        TDBGrid1.TextMatrix(i, 0) = i
    End If
Next i
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColWidth(0) = 700
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
TDBGrid1.MergeCells = flexMergeFree

'TDBGrid1.ColComboList(8) = "..."
'TDBGrid1.ColComboList(8) = "...."
'TDBGrid1.ColComboList(9) = "...."
End Sub

Private Sub Command1_Click()
Dim W_List As Double

If MsgBox(g_CON_CDelete, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
               
               Exit Sub
End If

W_List = TDBGrid1.TextMatrix(TDBGrid1.RowSel, TDBGrid1.Cols - 1)

G_Con.Execute "delete from mmst034_tmp where d_List =" & Val(W_List)

Call RefreshGrid
End Sub

Private Sub Command2_Click()
    G_Con.Execute "Insert into mmst034_tmp(pc_name,mtr_no,MTR_NAME,MTR_DIM,item_no,color_name,mtr_prs,g_weight,n_weight," & _
                    " Cont_Amt , Box_Dim, pic_url, OUT_STATUS, RINV_NO, SIZE, Cuft, Cart_Dim, Fabric, Packing, FT_40, HQ_40, HQ_45, FT_20,NOTE) " & _
                                        " select  top 1 '" & G_Pc_Name & "','','','','','',0,0,0,0,'','','',inv_No,'','','','','','','','','','' from mmst663 order by list_no desc"
    Call RefreshGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If LCase(TypeName(ActiveControl)) = "textbox" Then
        If ActiveControl.MultiLine = True Then
            Exit Sub
        End If
    End If
    
    If LCase(TypeName(ActiveControl)) = "combobox" And Not TypeOf ActiveControl Is ComboBox Then
        Exit Sub
    End If
    SendKeys "{TAB}"
    KeyCode = 0
End If

If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               '确认
         If CmdOK.Enabled = True Then
             Call CmdOk_Click
             KeyCode = 0
         End If
    Case vbKeyEscape           '取消
         If CmdCancel.Enabled = True Then
             Call CmdCancel_Click
             KeyCode = 0
         End If
    End Select
End If
End Sub

Private Sub Form_Load()


Me.KeyPreview = True
Set Me.Picture = Erp_Sale.Picture
'加载币别数据

G_Con.Execute "Delete from mmst034_tmp where pc_name='" & G_Pc_Name & "' "

Call GetVSGridSetting("frmquatmx", "TDBGrid1", Gridc_frmquatmx_1, g_CON_IniFile)
Row_Height = Gridc_frmquatmx_1(0).Grid_RowHeight

Call RefreshGrid

End Sub

Private Sub Form_Resize()
TDBGrid1.Width = Me.Width - 500
End Sub

Private Sub TDBGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim T_LIST As Double

T_LIST = Val(TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1))
'mtr_no
If Col = 1 Then
    
    G_Con.Execute "Update mmst034_tmp set mtr_no='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'mtr_name
If Col = 2 Then
    
    G_Con.Execute "Update mmst034_tmp set mtr_name='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'mtr_dim
If Col = 3 Then
    
    G_Con.Execute "Update mmst034_tmp set mtr_dim='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'item_no
If Col = 4 Then
    
    G_Con.Execute "Update mmst034_tmp set item_no='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'color_name,mtr_prs,g_weight,n_weight," & _
            " Cont_Amt , Box_Dim, pic_url, OUT_STATUS, RINV_NO, SIZE, Cuft, Cart_Dim, Fabric, Packing, FT_40, HQ_40, HQ_45, FT_20,NOTE  " & _

If Col = 5 Then
    
    G_Con.Execute "Update mmst034_tmp set color_name='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'mtr_prs
If Col = 6 Then
    
    G_Con.Execute "Update mmst034_tmp set mtr_prs='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'g_weight
If Col = 7 Then
    
    G_Con.Execute "Update mmst034_tmp set g_weight='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'n_weight
If Col = 8 Then
    
    G_Con.Execute "Update mmst034_tmp set n_weight='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'Cont_Amt
If Col = 9 Then
    
    G_Con.Execute "Update mmst034_tmp set cont_amt='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If



'Box_Dim
If Col = 10 Then
    
    G_Con.Execute "Update mmst034_tmp set box_dim='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'OUT_STATUS"FOB YT|FOB SK|FOB XIAMEN|FOB HK"
If Col = 11 Then
    Dim TMP_S As String
    Dim TMP_Z As String
    
    TMP_S = TDBGrid1.TextMatrix(Row, Col)
    If TMP_S = "FOB YT" Then
        TMP_Z = 1
    ElseIf TMP_S = "FOB SK" Then
        TMP_Z = 2
    ElseIf TMP_S = "FOB XIAMEN" Then
        TMP_Z = 4
    ElseIf TMP_S = "FOB HK" Then
        TMP_Z = 3
    Else
    End If
    G_Con.Execute "Update mmst034_tmp set OUT_STATUS= '" & TMP_Z & "' where D_LIST= " & T_LIST
    
End If
'SIZECuft, Cart_Dim, Fabric, Packing, FT_40, HQ_40, HQ_45, FT_20,NOTE
If Col = 12 Then
    
    G_Con.Execute "Update mmst034_tmp set SIZE='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'Cuft
If Col = 13 Then
    
    G_Con.Execute "Update mmst034_tmp set Cuft='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'Cart_Dim
If Col = 14 Then
    
    G_Con.Execute "Update mmst034_tmp set Cart_Dim='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'Fabric
If Col = 15 Then
    
    G_Con.Execute "Update mmst034_tmp set Fabric='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'Packing
If Col = 16 Then
    
    G_Con.Execute "Update mmst034_tmp set Packing=" & Trim(TDBGrid1.TextMatrix(Row, Col)) & " where D_LIST= " & T_LIST
    
End If

'FT_40
If Col = 17 Then
    
    G_Con.Execute "Update mmst034_tmp set FT_40='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'HQ_40
If Col = 18 Then
    
    G_Con.Execute "Update mmst034_tmp set HQ_40='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If


'HQ_45
If Col = 19 Then
    
    G_Con.Execute "Update mmst034_tmp set HQ_45='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If



'FT_20
If Col = 20 Then
    
    G_Con.Execute "Update mmst034_tmp set FT_20='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'备注
If Col = 21 Then
    
    G_Con.Execute "Update mmst034_tmp set note='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

End Sub

Private Sub TDBGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Col < 3 Then
'    Cancel = True
'End If

'生产线
'Dim W_Status_List1 As String
'Dim W_List1 As New ADODB.Recordset
'Dim Tmp_Str As String
'
'Tmp_Str = "  select distinct money_no  from mmst621 where isnull(money_no,'')<>''order by money_no"
'Set W_List1 = open_RS(Tmp_Str)
'If Not W_List1.EOF Then
'    W_Status_List1 = TDBGrid1.BuildComboList(W_List1, "money_no") & "|"
'    TDBGrid1.ColComboList(12) = W_Status_List1
'End If
'不良类型
If Col = 11 Then
    TDBGrid1.ColComboList(Col) = "FOB YT|FOB SK|FOB XIAMEN|FOB HK"
End If

End Sub
Private Sub CmdOk_Click()
    If check_ok Then
    
    G_Con.Execute "exec  ts_Upd_mmst034_bat '" & G_Pc_Name & "','" & Me.CallForm.quat_no.Text & "','" & G_User_Name & "' "
    
    
        Set w_rs = Nothing
        Call Me.CallForm.RefreshGrid

        Unload Me

    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Public Sub ClearFields()
   
End Sub

Private Function check_ok() As Boolean
   Dim W_Status As String
   Dim w_rs As New ADODB.Recordset
   W_Status = CheckStatus("QUAT_NO", Me.CallForm.quat_no.Text, "mmst033", "status")
   If W_Status <> "0" Then
        If W_Status = "9" Then
            MsgBox "单据" & Me.CallForm.quat_no.Text & "已被其它用户删除.", vbInformation, g_CON_CTitle
        Else
            MsgBox "单据" & Me.CallForm.quat_no.Text & "已被审核,不能新增或修改明细.", vbExclamation, g_CON_CTitle
        End If
        check_ok = False
        Unload Me
        Exit Function
    End If
    
    If Me.UpdateMode = 0 Then
        Set w_rs = Nothing
        '
        w_rs.Open "SELECT mtr_name ,mtr_no FROM mmst034 WHERE QUAT_NO='" & Trim(Me.CallForm.quat_no.Text) & "' and mtr_no in (select mtr_no from mmst034_tmp where pc_name='" & Trim(G_Pc_Name) & "')  ", G_Con, , , adCmdText
        If w_rs.EOF = False Then
                MsgBox w_rs!Mtr_Name & "  " & w_rs!Mtr_No & "资料重复,该报价单内已有此物料的报价单!", vbExclamation, g_CON_CTitle
'                Mtr_Name.SetFocus
                Set w_rs = Nothing
                Exit Function
        End If
        w_rs.Close
        Set w_rs = Nothing
        
        Set w_rs = Open_Rs("Select *from mmst034_tmp where pc_name='" & G_Pc_Name & "' and isnull(item_no,'')=''")
        
        If Not w_rs.EOF Then
            MsgBox w_rs!Mtr_Name & "  " & w_rs!Mtr_No & "请输入PO#号!", 64, "提示信息"
'            PO_NO.SetFocus
            check_ok = False
            Exit Function

        End If
        
        
         Set w_rs = Open_Rs("Select *from mmst034_tmp where pc_name='" & G_Pc_Name & "' and mtr_prs<=0")
        
        If Not w_rs.EOF Then
            MsgBox w_rs!Mtr_Name & "  " & w_rs!Mtr_No & "请录入本次报价单价!", 64, "提示信息"
'            PO_NO.SetFocus
            check_ok = False
            Exit Function

        End If
        
        
    End If
        
'Set w_rs = open_RS("Select *from mmst034_tmp where pc_name='" & G_Pc_Name & "' and isnull(money_no,'')=''")
'
'
'If Not w_rs.EOF Then
'    MsgBox w_rs!Mtr_Name & "请选择结算币别!", 64, "提示信息"
''    Lab_Unit.SetFocus
'    check_ok = False
'    Exit Function
'End If
'
'Set w_rs = open_RS("Select *from mmst034_tmp where pc_name='" & G_Pc_Name & "' and (isnull(net_weight,0)=0 or isnull(grass_weight,0)=0)")
'
'
'If Not w_rs.EOF Then
'    MsgBox w_rs!Mtr_Name & "净重或毛重不能为0!", 64, "提示信息"
''    Lab_Unit.SetFocus
'    check_ok = False
'    Exit Function
'End If
    
    check_ok = True
End Function
