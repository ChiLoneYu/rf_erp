VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm603Mx 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "成品出库单(Frm603Mx)"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox outbar_amt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1200
   End
   Begin VB.TextBox old_mtr_no 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "不能超过21个字符"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmd_brow_bar 
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
      Height          =   315
      Left            =   6525
      TabIndex        =   12
      Top             =   4950
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox Bar_No 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5025
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Mtr_Amt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1350
      TabIndex        =   10
      Top             =   2760
      Width           =   1200
   End
   Begin VB.TextBox Inbar_Amt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2265
      Width           =   1200
   End
   Begin VB.TextBox Max_Amt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2265
      Width           =   1200
   End
   Begin VB.TextBox Cust_No 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   1350
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "不能超过21个字符"
      Top             =   750
      Width           =   1920
   End
   Begin VB.TextBox Cust_Name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox Cust_Mtr_NO 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1245
      Width           =   1815
   End
   Begin VB.TextBox Cust_Order_No 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1245
      Width           =   1935
   End
   Begin VB.TextBox mtr_dim 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   345
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1755
      Width           =   5250
   End
   Begin VB.TextBox note 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3270
      Width           =   5250
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
      Height          =   315
      Left            =   2955
      TabIndex        =   1
      Top             =   270
      Width           =   300
   End
   Begin VB.CommandButton CmdOK 
      Height          =   405
      Left            =   2160
      Picture         =   "Frm603Mx.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4260
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancel 
      Height          =   405
      Left            =   3840
      Picture         =   "Frm603Mx.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4260
      Width           =   1140
   End
   Begin VB.TextBox Order_No 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1350
      MaxLength       =   21
      TabIndex        =   0
      ToolTipText     =   "不能超过21个字符"
      Top             =   240
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已出库数:"
      Height          =   195
      Left            =   3810
      TabIndex        =   29
      Tag             =   "Qty:"
      Top             =   2820
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "产品型号:"
      Height          =   195
      Left            =   3810
      TabIndex        =   27
      Top             =   330
      Width           =   825
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "本次出库:"
      Height          =   195
      Left            =   420
      TabIndex        =   26
      Tag             =   "Qty:"
      Top             =   2820
      Width           =   825
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已入库数:"
      Height          =   195
      Left            =   3810
      TabIndex        =   25
      Tag             =   "Qty:"
      Top             =   2320
      Width           =   825
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "最大出库数:"
      Height          =   195
      Left            =   225
      TabIndex        =   24
      Tag             =   "Qty:"
      Top             =   2320
      Width           =   1020
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "客户名称:"
      Height          =   195
      Left            =   3810
      TabIndex        =   23
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "客户编号:"
      Height          =   195
      Left            =   420
      TabIndex        =   22
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户料号:"
      Height          =   195
      Left            =   3810
      TabIndex        =   21
      Tag             =   "Product Code:"
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "客户订单:"
      Height          =   195
      Left            =   420
      TabIndex        =   20
      Top             =   1335
      Width           =   825
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备   注:"
      Height          =   195
      Left            =   495
      TabIndex        =   19
      Tag             =   "Remark:"
      Top             =   3360
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "产品规格:"
      Height          =   195
      Left            =   420
      TabIndex        =   18
      Tag             =   "Standard:"
      Top             =   1815
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "制单单号:"
      Height          =   195
      Left            =   420
      TabIndex        =   17
      Tag             =   "Order No.:"
      Top             =   360
      Width           =   825
   End
   Begin MSForms.Label Label6 
      Height          =   195
      Left            =   4050
      TabIndex        =   16
      Top             =   4980
      Visible         =   0   'False
      Width           =   765
      BackColor       =   -2147483639
      VariousPropertyBits=   276824083
      Caption         =   "仓       别:"
      Size            =   "1349;344"
      FontName        =   "新细明体"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "Frm603Mx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*适用於产品入仓单新增/修改入仓产品明细
'*
'*
'*********************************************************
Dim W_UpdateMode As Byte '0:add,1:edit
Dim W_CallForm As Form
Dim W_Bar_No As String
Dim W_Inv_No As String

Public G_List_No As Double
Public Property Get UpdateMode() As Byte
UpdateMode = W_UpdateMode
End Property

Public Property Let UpdateMode(b As Byte)
    W_UpdateMode = b
    Select Case b
        Case 0
            Me.Caption = "新增产品入库明细"
        Case 1
            Me.Caption = "修改产品出库明细"
            order_no.Locked = True
            cmd_brow_order.Enabled = False
    End Select
End Property

Public Property Get CallForm() As Form
    Set CallForm = W_CallForm
End Property

Public Property Set CallForm(f As Form)
    Set W_CallForm = f
End Property

Public Property Let Inv_No(f As String)
    W_Inv_No = f
End Property

Private Sub cmd_brow_bar_Click()
With FrmBarType
    .Show vbModal
'    Bar_Name.Text = .Bar_Name
    Bar_No.Text = .Bar_Name
End With
End Sub




Private Sub cmd_brow_order_Click()
If W_UpdateMode = 1 Then
    Exit Sub
End If

With FrmSectList
         .W_edit_able = False
         .Quer_status = False
    .W_Select_Data = " select  a.order_no as 制单编号,mtr_Amt as 订单数量," & _
                    "       cust_no as 客户编号,cust_name as 客户简称," & _
                    "       cust_order_no as 客户订单,cust_mtr_no as 客户料号," & _
                    "       prod_name as 产品类型,prod_dim as 产品规格 " & _
                    "  from order_inbar_amt  a inner join  mmsp011 b on a.order_no=b.order_no      " & _
                    "       and a.inbar_Amt>0 " & _
                    "  order by a.order_no "
    .Grid1.Editable = flexEDNone
    .Show vbModal
    If .cancel_status = False And .List1 <> "" Then
        order_no.Text = .List1
        Call Order_No_LostFocus
        Mtr_Amt.SetFocus
    End If
End With
End Sub

Private Sub CmdCaNcel_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Set Me.Picture = GetMdiForm.Picture

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 34
    End If
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
             Call CmdOK_Click
             KeyCode = 0
         End If
    Case vbKeyEscape           '取消
         If CmdCancel.Enabled = True Then
             Call CmdCaNcel_Click
             KeyCode = 0
         End If
    End Select
End If
End Sub

Private Sub CmdOK_Click()

If check_ok = False Then
    Exit Sub
End If

If W_UpdateMode = 0 Then
    G_Con.Execute " insert mmst534 (inv_no,order_no,mtr_amt,bar_no,note,upd_name,upd_date) " & _
                  " values ('" & W_Inv_No & "','" & Trim(order_no.Text) & "'," & Val(Mtr_Amt.Text) & "," & _
                  "             '" & W_Bar_No & "','" & Trim(Note.Text) & "','" & G_User_Name & "','" & Get_SQLDATE & "')  "
End If


If W_UpdateMode = 1 Then
    G_Con.Execute " update mmst534 set mtr_amt=" & Val(Mtr_Amt.Text) & ",bar_no='" & W_Bar_No & "'," & _
                  "         note='" & Trim(Note.Text) & "',upd_name='" & G_User_Name & "',upd_date='" & Get_SQLDATE & "' " & _
                  " where inv_no='" & W_Inv_No & "' and order_no='" & Trim(order_no.Text) & "'  "
End If


Call Me.CallForm.RefreshGrid

    If W_UpdateMode = 0 Then
        Call ClearFields
        order_no.SetFocus
    Else
        Unload Me
    End If

End Sub

Private Function check_ok() As Boolean
Dim W_RS As New ADODB.Recordset
Dim W_Str  As String

check_ok = False

If W_UpdateMode = 0 Then
    If Trim(order_no.Text) = "" Then
        MsgBox "请输入制单号!", 64, "提示信息"
        Exit Function
    End If
    Set W_RS = open_RS(" select order_no from mmst011 where order_no='" & Trim(order_no.Text) & "' and status=2 and order_type<>8 ")
    If W_RS.EOF Then
        MsgBox " 输入的制单号并不存在或者没有审核!", 64, "提示信息"
        Set W_RS = Nothing
        Exit Function
    End If
    
    Set W_RS = open_RS(" select order_no from mmst534 where inv_no='" & W_Inv_No & "' and order_no='" & Trim(order_no.Text) & "' ")
    If Not W_RS.EOF Then
        Set W_RS = Nothing
        MsgBox "输入资料重复!", 64, "提示信息"
        Exit Function
    End If
End If

'检查是否已经有入库
Set W_RS = open_RS(" select order_no from order_inbar_amt where order_no='" & Trim(order_no.Text) & "' and inbar_amt>0  ")
If W_RS.EOF Then
    MsgBox "该订单并没有入库,无法出库! ", 64, "提示信息"
    Set W_RS = Nothing
    Exit Function
End If

Set W_RS = open_RS(" select a.inbar_amt,isnull(b.outbar_amt,0) as outbar_amt " & _
                 " from order_inbar_amt a left join order_outbar_amt b on a.order_no=b.order_no " & _
                 " where a.order_no='" & Trim(order_no.Text) & "'   ")
If Not W_RS.EOF Then
    If Val(Mtr_Amt.Text) > W_RS!Inbar_Amt - W_RS!outbar_amt Then
        MsgBox "库存不够,欠数:" & Val(Mtr_Amt.Text) + W_RS!outbar_amt - W_RS!Inbar_Amt & "  ", 64, "提示信息"
        Exit Function
    End If
End If

Set W_RS = Nothing
check_ok = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set Frm603Mx = Nothing
End Sub

Public Sub ClearFields()

    order_no.Text = ""
    Mtr_Amt.Text = ""

    cust_no.Text = ""
    cust_name.Text = ""
    
    Cust_Order_No.Text = ""
    Cust_Mtr_No.Text = ""
    
    
    Mtr_Dim.Text = ""
    Bar_No.Text = ""
    Max_Amt.Text = ""
    Inbar_Amt.Text = ""

    Note.Text = ""
    
End Sub






Public Sub Order_No_LostFocus()

Dim tmp_rs As New ADODB.Recordset
Set tmp_rs = open_RS(" select  order_no ,mtr_no,mtr_Amt  ," & _
                    "       cust_no  ,cust_name  ," & _
                    "       cust_order_no  ,cust_mtr_no  ," & _
                    "       prod_name  ,prod_dim   " & _
                    "  from mmsp011 where status=2 and order_type<>8       " & _
                    "        " & _
                    "       and order_no = '" & Trim(order_no.Text) & "' " & _
                    "  order by order_no ")
                    
If Not tmp_rs.EOF Then
    old_mtr_no.Text = NullVal(tmp_rs!old_mtr_no, "")
    cust_no.Text = tmp_rs!cust_no
    cust_name.Text = tmp_rs!cust_name
    Cust_Order_No = NullVal(tmp_rs!Cust_Order_No, "")
    Cust_Mtr_No.Text = NullVal(tmp_rs!Cust_Mtr_No, "")
    Mtr_Dim.Text = NullVal(tmp_rs!prod_dim, "")
    Max_Amt.Text = tmp_rs!Mtr_Amt
Else
    old_mtr_no.Text = ""
    cust_no.Text = ""
    cust_name.Text = ""
    Cust_Order_No.Text = ""
    Cust_Mtr_No.Text = ""
    Mtr_Dim.Text = ""
    Max_Amt.Text = ""
End If
Set tmp_rs = Nothing

Set tmp_rs = open_RS(" select  *  from  order_inbar_amt where order_no='" & Trim(order_no.Text) & "'  ")
If Not tmp_rs.EOF Then
    Inbar_Amt.Text = NullVal(tmp_rs!Inbar_Amt, 0)
Else
    Inbar_Amt.Text = 0
End If

Set tmp_rs = open_RS(" select  *  from  order_outbar_amt where order_no='" & Trim(order_no.Text) & "'  ")
If Not tmp_rs.EOF Then
    outbar_amt.Text = NullVal(tmp_rs!outbar_amt, 0)
Else
    outbar_amt.Text = 0
End If


Set tmp_rs = Nothing
End Sub

