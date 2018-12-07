VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm401Mx 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   3  '雙線固定對話方塊
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm401Mx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton CmdOK 
      Height          =   345
      Left            =   1313
      Picture         =   "Frm401Mx.frx":000C
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   4260
      Width           =   1110
   End
   Begin VB.CommandButton CmdCancel 
      Height          =   345
      Left            =   3098
      Picture         =   "Frm401Mx.frx":15AE
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   4260
      Width           =   1110
   End
   Begin VB.CommandButton Cmd_Qc_Brow 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   24
      Top             =   2505
      Width           =   285
   End
   Begin VB.TextBox Lot_No 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3630
      MaxLength       =   21
      TabIndex        =   21
      Top             =   2925
      Width           =   1560
   End
   Begin VB.TextBox Unit_Name 
      Appearance      =   0  '平面
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2055
      Width           =   630
   End
   Begin VB.TextBox Mtr_Amt1 
      Appearance      =   0  '平面
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3630
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2055
      Width           =   915
   End
   Begin VB.TextBox Supl_Unit 
      Appearance      =   0  '平面
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2295
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2055
      Width           =   570
   End
   Begin VB.CommandButton Cmd_Pcs_Brow 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2970
      TabIndex        =   17
      Top             =   330
      Width           =   285
   End
   Begin VB.TextBox Note 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   3945
   End
   Begin VB.TextBox Mtr_Amt 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   9
      Top             =   2055
      Width           =   1035
   End
   Begin VB.TextBox Mtr_Dim 
      Appearance      =   0  '平面
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1620
      Width           =   4005
   End
   Begin VB.TextBox Mtr_Name 
      Appearance      =   0  '平面
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1185
      Width           =   4005
   End
   Begin VB.CommandButton Cmd_Mtr_Brow 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2970
      TabIndex        =   18
      Top             =   765
      Width           =   285
   End
   Begin VB.TextBox Mtr_No 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      MaxLength       =   21
      TabIndex        =   3
      Top             =   750
      Width           =   2040
   End
   Begin VB.TextBox Pcs_No 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      MaxLength       =   12
      TabIndex        =   1
      Top             =   315
      Width           =   2040
   End
   Begin VB.TextBox Qc_No 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3630
      MaxLength       =   21
      TabIndex        =   19
      Top             =   2490
      Width           =   1560
   End
   Begin MSForms.ComboBox Spe_Let 
      Height          =   300
      Left            =   1230
      TabIndex        =   28
      Top             =   2940
      Width           =   1665
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2937;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Bar_No 
      Height          =   300
      Left            =   1230
      TabIndex        =   27
      Top             =   2520
      Width           =   1665
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2937;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lab_Spe_Let 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "特        採:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   23
      Top             =   2970
      Width           =   765
   End
   Begin VB.Label Lab_Lot_no 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "LOT NO:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2910
      TabIndex        =   22
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "QC單號:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2910
      TabIndex        =   20
      Top             =   2565
      Width           =   645
   End
   Begin VB.Label Lab_note 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備        注:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   15
      Top             =   3405
      Width           =   765
   End
   Begin VB.Label lblBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "倉        別:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   14
      Top             =   2535
      Width           =   765
   End
   Begin VB.Label lblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "數        量:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   8
      Top             =   2115
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規        格:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   6
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "品        名:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   4
      Top             =   1245
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "材料代號:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   2
      Top             =   825
      Width           =   765
   End
   Begin VB.Label lblPo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "採購單號:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   315
      TabIndex        =   0
      Top             =   390
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "= 本廠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3000
      TabIndex        =   11
      Top             =   2100
      Width           =   495
   End
End
Attribute VB_Name = "Frm401Mx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名稱: 採購入庫/退貨明細(Frm401Mx)
'*編寫日期: 2002/07/04
'*制作人員: 張  峰
'*修改日期:
'*修改人員:
'***********************************************
'定義記錄集與命令對象
Dim W_UpdateMode As Byte '0: 新增明細 , 1:修改明細

'定義呼叫窗體
Dim W_CallForm As Form

'定義工作變量
Dim W_Unit_ID As String
Dim W_Bar_No As String

Public W_Trans_Val As Single
Public G_Qc_Type As String

Dim W_supl_No As String

'定義屬性
Public Property Get UpdateMode() As Byte
UpdateMode = W_UpdateMode
End Property

Public Property Let UpdateMode(b As Byte)
W_UpdateMode = b
If b = 0 Then
    Me.Caption = "新增材料(" & Me.Name & ")"
Else
    Me.Caption = "修改材料(" & Me.Name & ")"
    Pcs_No.Locked = True
    mtr_no.Locked = True
    mtr_dim.Locked = True
    
    Pcs_No.TabStop = False
    mtr_no.TabStop = False
    mtr_dim.TabStop = False
    
    Cmd_Pcs_Brow.Enabled = False
    Cmd_Mtr_Brow.Enabled = False
End If
End Property

Public Property Get CallForm() As Form
Set CallForm = W_CallForm
End Property

Public Property Set CallForm(f As Form)
Set W_CallForm = f
On Error Resume Next
W_supl_No = f.Supl_No.Text
End Property

Private Sub Cmd_Mtr_Brow_Click()
With FrmPoMtrList
    .G_Pcs_No = Pcs_No.Text
    .G_Mtr_Filter = " a.mtr_no like '" & Trim(mtr_no.Text) & "%'  "
   .Show vbModal
    If .mtr_no <> "" Then
        mtr_no.Text = .mtr_no
        mtr_name.Text = .mtr_name
        mtr_dim.Text = .mtr_dim
        Supl_Unit.Text = .Supl_Unit
        W_Trans_Val = .Trans_Val
        Unit_Name.Text = .Unit_Name
        Call mtr_no_LostFocus
        mtr_amt.SetFocus
    End If
End With

End Sub

Private Sub Cmd_Qc_Brow_Click()
With FrmIQCList
    .g_Qc_Filter = "Qc_No Like '" & qc_no.Text & "%'"
    .G_Supl_No = W_supl_No
    .G_Qc_Type = G_Qc_Type
    
    .G_Mtr_No = mtr_no.Text
    .G_Mtr_Dim = mtr_dim.Text
    .Show vbModal
    If Not .Cancel_Click Then
        qc_no.Text = .qc_no
        
    End If
End With
End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
Me.Picture = Erp_Bar.Picture

Call AddRsToList(Me.bar_no, "select bar_name from mmst903 order by bar_name")

spe_let.Clear
spe_let.AddItem "是"
spe_let.AddItem "否"
spe_let.ListIndex = 1
End Sub

Private Sub Form_KeyDowN(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And ActiveControl.Name <> "note" Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cmd_Pcs_brow_Click()
With FrmPoList
    .G_Po_Filter = "Pcs_No Like '" & Pcs_No.Text & "%'"
    .G_Supl_No = W_supl_No
    .Show vbModal
    If Not .Cancel_Click Then
        Pcs_No.Text = .Pcs_No
        Call Pcs_No_LostFocus
        mtr_no.SetFocus
    End If
End With
End Sub

Private Sub CmdOk_Click()
If check_ok Then
    Dim w_rs As New ADODB.Recordset
    With w_rs
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        Set .ActiveConnection = G_Con
        
        If Me.UpdateMode = 0 Then
            .Open "select * from mmst302 where mtr_no= '' "
            .AddNew
            !inv_no = Me.CallForm.inv_no.Text
            !Pcs_No = Trim(Pcs_No.Text)
            !mtr_no = Trim(mtr_no.Text)
            !mtr_dim = Trim(mtr_dim.Text)
        Else
            .Open "select * from mmst302 where Inv_no='" & Me.CallForm.inv_no.Text & "' and mtr_no='" & mtr_no.Text & "' and Pcs_No='" & Pcs_No.Text & "' and Mtr_Dim ='" & mtr_dim.Text & "'"
        End If
        !mtr_amt = Val(mtr_amt.Text)
        W_Unit_ID = FromNameGetID("mmst602", "Unit_Id", "Unit_name", Supl_Unit.Text)
        !Unit_Id = W_Unit_ID
        !Trans_Val = W_Trans_Val
        !bar_no = W_Bar_No
        !qc_no = qc_no.Text
        !spe_let = spe_let.Text
        !Lot_No = Lot_No.Text
        !note = note.Text
        !upd_name = G_User_Name
        !upd_date = Date
        .Update
        .Close
    End With
    '更改質檢單使用狀態
    G_Con.Execute "UPDATE mmsta11 SET Use_State = '1' WHERE Qc_No = '" & qc_no.Text & "'"
    
    Set w_rs = Nothing
    Call Me.CallForm.RefreshGrid
    If Me.UpdateMode = 0 Then
        Call ClearFields
        mtr_no.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Function check_ok() As Boolean
Dim w_rs As New ADODB.Recordset

If Me.UpdateMode = 0 Then
    If Pcs_No.Text = "" Then
        MsgBox "請輸入採購單號.", vbExclamation, g_CON_CTitle
        Pcs_No.SetFocus
        check_ok = False
        Exit Function
    Else
        Set w_rs = Nothing
        w_rs.CursorLocation = adUseClient
        w_rs.Open "SELECT Pcs_No " & _
                    " FROM mmst205 " & _
                    " WHERE mmst205.Pcs_No = '" & Pcs_No.Text & "' " & _
                        "AND mmst205.Status = '2' " & _
                        "AND Supl_No = '" & W_supl_No & "'", G_Con, adOpenDynamic
                        
        If w_rs.EOF Then
            MsgBox "無此採購單號.", vbExclamation, g_CON_CTitle
            Pcs_No.SetFocus
            check_ok = False
            Exit Function
        End If
        w_rs.Close
    End If
    
    If Trim(mtr_no) = "" Then
        MsgBox "必須輸入材料編號.", vbExclamation, g_CON_CTitle
        mtr_no.SetFocus
        check_ok = False
        Exit Function
    Else
        Set w_rs = Nothing
        w_rs.CursorLocation = adUseClient
        w_rs.Open "SELECT Mtr_No " & _
                      " FROM mmst206 " & _
                      " WHERE Mtr_no='" & Trim(mtr_no.Text) & "'" & _
                        "AND Pcs_No='" & Trim(Pcs_No.Text) & "'", G_Con, , , adCmdText
                        
        If w_rs.EOF Then
             w_rs.Close
             MsgBox "採購單" & Pcs_No.Text & " 未訂購此材料.", vbExclamation, g_CON_CTitle
             mtr_no.SetFocus
             check_ok = False
             Exit Function
        End If
        w_rs.Close
    End If
    
    If Trim(mtr_dim) = "" Then
        MsgBox "必須輸入規格.", vbExclamation, g_CON_CTitle
        mtr_dim.SetFocus
        check_ok = False
        Exit Function
    Else
        Set w_rs = Nothing
        w_rs.CursorLocation = adUseClient
        w_rs.Open "SELECT Mtr_No,Mtr_Dim " & _
                      "FROM mmst206 " & _
                      "WHERE Mtr_No = '" & Trim(mtr_no.Text) & "' " & _
                        "AND Mtr_Dim = '" & Trim(mtr_dim.Text) & "' " & _
                        "AND Pcs_No = '" & Trim(Pcs_No.Text) & "'", G_Con, , , adCmdText
                        
        If w_rs.EOF Then
             w_rs.Close
             MsgBox "採購單" & Pcs_No.Text & "未訂購此規格材料.", vbExclamation, g_CON_CTitle
             mtr_dim.SetFocus
             check_ok = False
             Exit Function
        End If
        w_rs.Close
    End If
    
    Set w_rs = Nothing
    w_rs.CursorLocation = adUseClient
    w_rs.Open "SELECT * " & _
                "FROM mmst302 " & _
                "WHERE Inv_No = '" & Me.CallForm.inv_no.Text & "' " & _
                    "AND Mtr_no = '" & Trim(mtr_no.Text) & "' " & _
                    "AND Pcs_No = '" & Pcs_No.Text & "' " & _
                    "AND Mtr_Dim = '" & mtr_dim.Text & "' ", G_Con, , , adCmdText
                    
    If w_rs.EOF = False Then
        MsgBox "輸入資料重複(採購單號+材料代號+規格).", vbExclamation, g_CON_CTitle
        mtr_no.SetFocus
        check_ok = False
        Exit Function
    End If
    w_rs.Close
    
    Set w_rs = Nothing
End If

If Val(mtr_amt.Text) <= 0 Then
    MsgBox "請輸入正確的數量.", vbExclamation, g_CON_CTitle
    mtr_amt.SetFocus
    check_ok = False
    Exit Function
End If
     
If bar_no.Text = "" Then
    MsgBox "請選擇倉別", vbExclamation, g_CON_CTitle
    bar_no.SetFocus
    check_ok = False
    Exit Function
Else
    W_Bar_No = FromNameGetID("mmst903", "bar_no", "bar_name", bar_no.Text)
    If W_Bar_No = "" Then
        MsgBox "無此倉別資料.", vbExclamation, g_CON_CTitle
        bar_no.Clear
        Call AddRsToList(bar_no, "select bar_name from mmst903 order by bar_name")
        Exit Function
    End If
End If
    
If qc_no.Text = "" Then
    MsgBox "請輸入質檢單號", vbExclamation, g_CON_CTitle
    qc_no.SetFocus
    check_ok = False
    Exit Function
Else
    Set w_rs = Nothing
    w_rs.CursorLocation = adUseClient
    w_rs.Open "SELECT Qc_No " & _
                "FROM mmsta11 " & _
                "WHERE Qc_No = '" & qc_no.Text & "' " & _
                    "AND Qc_Type =  '" & G_Qc_Type & "' " & _
                    "AND Supl_No = '" & W_supl_No & "' " & _
                    "AND Mtr_No = '" & mtr_no.Text & "' " & _
                    "AND Mtr_Dim = '" & mtr_dim.Text & "' ", G_Con, adOpenDynamic
    If w_rs.EOF Then
        MsgBox "無此質檢單號或該質檢單不是檢驗的該材料,請檢查!", vbExclamation, g_CON_CTitle
        qc_no.SetFocus
        check_ok = False
        Exit Function
    End If
End If

check_ok = True
End Function

Private Sub Form_Unload(Cancel As Integer)
Set Frm401Mx = Nothing
End Sub

Private Sub mtr_amt_Change()
If Trim(Pcs_No.Text) <> "" And Trim(mtr_no.Text) <> "" Then
    Mtr_Amt1.Text = Val(mtr_amt.Text) * W_Trans_Val
End If
End Sub

Private Sub Pcs_No_LostFocus()
If Me.UpdateMode <> 0 Then
    Exit Sub
End If
If Trim(Pcs_No.Text) <> "" Then
    Dim w_rs As New ADODB.Recordset
    w_rs.Open "SELECT Pcs_No FROM mmst205 WHERE Pcs_No='" & Trim(Pcs_No.Text) & "' AND status='2' AND supl_no='" & Trim(W_supl_No) & "'", G_Con, , , adCmdText
    If w_rs.EOF Then
         MsgBox "該廠商" & W_supl_No & "無此採購單號!", vbExclamation, g_CON_CTitle
         Pcs_No.Text = ""
         Pcs_No.SetFocus
         w_rs.Close
         Set w_rs = Nothing
         Exit Sub
    End If
    w_rs.Close
End If
End Sub

Private Sub Mtr_no_DblClick()
If Cmd_brow.Enabled Then
    Call Cmd_Mtr_Brow_Click
End If
End Sub

Private Sub mtr_no_LostFocus()
If Me.UpdateMode = 0 Then
    mtr_name.Text = ""
    mtr_dim.Text = ""
    Unit_Name.Text = ""
    
    If Trim(mtr_no.Text) <> "" Then
        Dim w_rs As New ADODB.Recordset
        If Trim(Pcs_No.Text) <> "" Then
            Supl_Unit.Text = ""
            w_rs.Open "SELECT a.mtr_no, a.trans_val, m.mtr_name, m.mtr_dim, " & _
                             "u1.unit_name, u2.unit_name AS supl_unit,a.mtr_amt " & _
                      "FROM mmst206 a INNER JOIN " & _
                           "mmst611 m ON a.mtr_no = m.mtr_no INNER JOIN " & _
                           "mmst602 u2 ON a.unit_id = u2.unit_id INNER JOIN " & _
                           "mmst602 u1 ON m.unit_id = u1.unit_id " & _
                      "WHERE m.mtr_no='" & Trim(mtr_no.Text) & "' and a.Pcs_No='" & Trim(Pcs_No.Text) & "'", G_Con, , , adCmdText
        Else
            w_rs.Open "select a.mtr_no,a.mtr_name,a.mtr_dim,b.unit_name " & _
                      "from mmst611 a join mmst602 b on a.unit_id=b.unit_id " & _
                      "where a.mtr_no='" & Trim(mtr_no.Text) & "'", G_Con, , , adCmdText
                      
        End If
        If Not w_rs.EOF Then
            mtr_no.Text = w_rs!mtr_no
            mtr_name.Text = NullSetValue(w_rs!mtr_name, "")
            mtr_dim.Text = NullSetValue(w_rs!mtr_dim, "")
            Unit_Name.Text = NullSetValue(w_rs!Unit_Name, "")
            If Trim(Pcs_No.Text) <> "" Then
            
                Supl_Unit.Text = w_rs!Supl_Unit
                W_Trans_Val = w_rs!Trans_Val
                '計算未交量
                If UCase(Me.CallForm.Name) = "MMSS401" Then
                    '採購單訂購數量
                    w_pcs_amt = w_rs!mtr_amt
                    '已交數量
                    Set w_rs = Nothing
                    w_rs.CursorLocation = adUseClient
                    w_rs.Open "SELECT SUM(Case a.inv_type when '1' then 1 else -1 end * mtr_amt) As Inbar_Amt " & _
                              " FROM mmst301 a INNER JOIN " & _
                                    " mmst302 b ON a.Inv_No = b.Inv_No " & _
                              " WHERE b.pcs_No = '" & Pcs_No.Text & "' " & _
                                "AND b.Mtr_No = '" & mtr_no.Text & "' " & _
                                "AND b.mtr_Dim = '" & mtr_dim.Text & "' " & _
                                "AND a.status = '2' " & _
                              " GROUP BY Pcs_no,Mtr_No,Mtr_Dim ", G_Con, adOpenDynamic
                    If w_rs.EOF = False Then
                        w_inbar_amt = Val(NullSetValue(w_rs!inbar_amt, "0"))
                    Else
                        w_inbar_amt = 0
                    End If
                    mtr_amt.Text = w_pcs_amt - w_inbar_amt
                 End If
                 
                 Call mtr_amt_Change
                
            End If
        End If
        w_rs.Close
        Set w_rs = Nothing
        '預設倉別
        Dim w_bar As String
        On Error Resume Next
        w_bar = GetDefaultBar(mtr_no.Text)
        If w_bar <> "" Then
            bar_no.Text = w_bar
        End If
    End If
End If
End Sub

Public Sub ClearFields()
Pcs_No.Text = ""
mtr_no.Text = ""
mtr_name.Text = ""
mtr_dim.Text = ""
mtr_amt.Text = ""
Unit_Name.Text = ""
Supl_Unit.Text = ""
Mtr_Amt1.Text = ""
W_Trans_Val = 1
qc_no.Text = ""
spe_let.Text = " "
Lot_No.Text = ""
note.Text = ""
End Sub

