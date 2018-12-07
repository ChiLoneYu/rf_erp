VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frm10OMX_add 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出货产品明细资料"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   15495
   LinkTopic       =   "Form1"
   Picture         =   "Frm10OMx_add.frx":0000
   ScaleHeight     =   6435
   ScaleWidth      =   15495
   StartUpPosition =   1  'CenterOwner
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
      Width           =   1665
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9615
      Picture         =   "Frm10OMx_add.frx":0DA3
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Cancel"
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton CmdOK 
      Height          =   375
      Left            =   7845
      Picture         =   "Frm10OMx_add.frx":2345
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
   Begin VB.TextBox order_no 
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
      FormatString    =   $"Frm10OMx_add.frx":38E7
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
         Picture         =   "Frm10OMx_add.frx":39CF
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
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "制单单号:"
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
      Width           =   825
   End
End
Attribute VB_Name = "Frm10OMX_add"
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

Public W_order_no As String
Public W_Cust_No As String
Public W_Cust_Order_No As String
Dim Gridc_10O_1(127) As Grid_Data

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
            order_no.Locked = True
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
             .W_Field2 = "prod_name"
             .W_Field3 = "prod_dim"
             .W_Select_Data = " select '' as 选择,order_no as 制单单号,cust_order_no as 订单单号,cust_mtr_no as 客户型号,mtr_no as 本厂型号," & _
                                     " mtr_amt as 订购数量,isnull(out_amt,0) as 已出数量,isnull(mtr_amt,0)-isnull(out_amt,0) as 未交数量,'pcs' as 单位" & _
                              " from mmsp011 " & _
                              " where status='2' and isnull(out_amt,0)<mtr_amt " & _
                                    " and order_type=1 and cust_no='" & W_Cust_No & "' and order_no like '" & Trim(order_no.Text) & "%' " & _
                              " and order_no not in (select order_no from mmst10a_tmp where pc_name='" & G_Pc_Name & "')  " & _
                              " and order_no not in (select order_no from mmst10a where deliv_no='" & Me.CallForm.Deliv_no.Text & "')  "
'             .Grid1.Editable = flexEDNone
             .Show vbModal
             If .cancel_status = False And .List1 <> "" Then
                order_no.Text = .List1
                cust_order_No.Text = .List2
'                cust_mtr_no.Text = .List3
'                old_mtr_no.Text = .List4
'                Call Order_No_LostFocus
                Call RefreshGrid
'                mtr_amt.SetFocus
             End If
        End With
    End If
End Sub
Private Sub RefreshGrid()

Dim W_Rd As New ADODB.Recordset
Dim StrSQL As String

StrSQL = "sELECT b.order_no,cust_order_no,cust_mtr_no,b.prod_name,prod_dim,color_script," & _
                   " b.mtr_amt as  order_amt,a.mtr_amt as out_amt,Pmtr_scrpt,po_no,order_name,b.mtr_prs,money_no,net_weight,grass_weight,mtr_meas,cnt_amt,order_name,note,D_lIST FROM mmst10a_tmp " & _
         " a inner join mmsp011 b on a.order_no=b.order_no  where pc_name='" & G_Pc_Name & "'" & _
        "  "

Set W_Rd = Open_Rs(StrSQL)

Set Adodc1.Recordset = W_Rd
Set TDBGrid1.DataSource = Adodc1

'Call readactive
Call SetVSGridSetting(TDBGrid1, Gridc_10O_1)



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

G_Con.Execute "delete from mmst10a_tmp where d_List =" & Val(W_List)

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

G_Con.Execute "Delete from mmst10a_tmp where pc_name='" & G_Pc_Name & "' "

Call GetVSGridSetting("frm10o", "TDBGrid1", Gridc_10O_1, g_CON_IniFile4)
Row_Height = Gridc_10O_1(0).Grid_RowHeight

Call RefreshGrid

End Sub

Private Sub Form_Resize()
TDBGrid1.Width = Me.Width - 500
End Sub

Private Sub TDBGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim T_LIST As Double

T_LIST = Val(TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1))




'出货数量
If Col = 8 Then
    
    G_Con.Execute "Update mmst10a_tmp set mtr_amt=" & Val(TDBGrid1.TextMatrix(Row, Col)) & " where D_LIST= " & T_LIST
    
End If
'用量1Pmtr_scrpt产品描述
If Col = 9 Then
    
    G_Con.Execute "Update mmst10a_tmp set Pmtr_scrpt='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'po_no#
If Col = 10 Then
    
    G_Con.Execute "Update mmst10a_tmp set po_no='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'order_name
If Col = 11 Then
    
    G_Con.Execute "Update mmst10a_tmp set order_name='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'order_name
If Col = 12 Then
    
    G_Con.Execute "Update mmst10a_tmp set mtr_prs='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'money_no
If Col = 13 Then
    
    G_Con.Execute "Update mmst10a_tmp set money_no='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'net_weight
If Col = 14 Then
    
    G_Con.Execute "Update mmst10a_tmp set net_weight='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

'grass_weight
If Col = 15 Then
    
    G_Con.Execute "Update mmst10a_tmp set grass_weight='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If
'mtr_meas
If Col = 16 Then
    
    G_Con.Execute "Update mmst10a_tmp set mtr_meas=" & Val(TDBGrid1.TextMatrix(Row, Col)) & " where D_LIST= " & T_LIST
    
End If
'cnt_amt
If Col = 17 Then
    
    G_Con.Execute "Update mmst10a_tmp set cnt_amt=" & Val(TDBGrid1.TextMatrix(Row, Col)) & " where D_LIST= " & T_LIST
    
End If



'备注
If Col = 18 Then
    
    G_Con.Execute "Update mmst10a_tmp set note='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' where D_LIST= " & T_LIST
    
End If

End Sub

Private Sub TDBGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col < 8 Then
    Cancel = True
End If

'生产线
Dim W_Status_List1 As String
Dim W_List1 As New ADODB.Recordset
Dim Tmp_Str As String

Tmp_Str = "  select distinct money_no  from mmst621 where isnull(money_no,'')<>''order by money_no"
Set W_List1 = Open_Rs(Tmp_Str)
If Not W_List1.EOF Then
    W_Status_List1 = TDBGrid1.BuildComboList(W_List1, "money_no") & "|"
    TDBGrid1.ColComboList(13) = W_Status_List1
End If


End Sub
Private Sub CmdOk_Click()
    If check_ok Then
    
    G_Con.Execute "exec  ts_Upd_mmst10a_bat '" & G_Pc_Name & "','" & Me.CallForm.Deliv_no.Text & "','" & G_User_Name & "' "
    
    
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
   W_Status = CheckStatus("Deliv_No", Me.CallForm.Deliv_no.Text, "mmst109", "status")
   If W_Status <> "0" Then
        If W_Status = "9" Then
            MsgBox "单据" & Me.CallForm.Deliv_no.Text & "已被其它用户删除.", vbInformation, g_CON_CTitle
        Else
            MsgBox "单据" & Me.CallForm.Deliv_no.Text & "已被审核,不能新增或修改明细.", vbExclamation, g_CON_CTitle
        End If
        check_ok = False
        Unload Me
        Exit Function
    End If
    
    If Me.UpdateMode = 0 Then
        Set w_rs = Nothing
        '
        w_rs.Open "SELECT deliv_no  FROM mmst10a WHERE deliv_no='" & Trim(Me.CallForm.Deliv_no.Text) & "' and order_no='" & Trim(order_no.Text) & "'  ", G_Con, , , adCmdText
        If w_rs.EOF = False Then
                MsgBox "资料重复,该出货单内已有此编号的制单!", vbExclamation, g_CON_CTitle
                order_no.SetFocus
                Set w_rs = Nothing
                Exit Function
        End If
        w_rs.Close
        Set w_rs = Nothing
        
        Set w_rs = Open_Rs("Select *from mmst10a_tmp where pc_name='" & G_Pc_Name & "' and isnull(po_No,'')=''")
        
        If Not w_rs.EOF Then
            MsgBox w_rs!order_no & "请输入PO#号!", 64, "提示信息"
'            PO_NO.SetFocus
            check_ok = False
            Exit Function

        End If
        
        
         Set w_rs = Open_Rs("Select *from mmst10a_tmp where pc_name='" & G_Pc_Name & "' and mtr_amt<=0")
        
        If Not w_rs.EOF Then
            MsgBox w_rs!order_no & "请录入本次出货数量!", 64, "提示信息"
'            PO_NO.SetFocus
            check_ok = False
            Exit Function

        End If
        
        
    End If
        
Set w_rs = Open_Rs("Select *from mmst10a_tmp where pc_name='" & G_Pc_Name & "' and isnull(money_no,'')=''")

    
If Not w_rs.EOF Then
    MsgBox w_rs!order_no & "请选择结算币别!", 64, "提示信息"
'    Lab_Unit.SetFocus
    check_ok = False
    Exit Function
End If

Set w_rs = Open_Rs("Select *from mmst10a_tmp where pc_name='" & G_Pc_Name & "' and (isnull(net_weight,0)=0 or isnull(grass_weight,0)=0)")

    
If Not w_rs.EOF Then
    MsgBox w_rs!order_no & "净重或毛重不能为0!", 64, "提示信息"
'    Lab_Unit.SetFocus
    check_ok = False
    Exit Function
End If
    
    check_ok = True
End Function
