VERSION 5.00
Begin VB.Form FrmDelivno 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5730
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdOK 
      Height          =   345
      Left            =   1230
      Picture         =   "FrmDelivno.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2790
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancel 
      Height          =   345
      Left            =   3090
      Picture         =   "FrmDelivno.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2790
      Width           =   1140
   End
   Begin VB.TextBox Remark 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   1245
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1155
      Width           =   4350
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2970
      TabIndex        =   1
      Top             =   240
      Width           =   300
   End
   Begin VB.TextBox deliv_no 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      MaxLength       =   12
      TabIndex        =   0
      Top             =   225
      Width           =   2040
   End
   Begin VB.TextBox deliv_date 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   690
      Width           =   4350
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "备    注:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   255
      TabIndex        =   8
      Top             =   1215
      Width           =   810
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出货单号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   255
      TabIndex        =   6
      Top             =   270
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   5670
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "货单日期:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   255
      TabIndex        =   7
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "FrmDelivno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_UpdateMode As Byte '0:add,1:edit
Dim W_CallForm As Form
Dim W_TbName As String
Dim w_money_no As String
Dim W_Cust_No As String
Dim w_ratio As Double
Dim W_Mtr_prs As Double

Public Property Get UpdateMode() As Byte
UpdateMode = W_UpdateMode
End Property
Public Property Let UpdateMode(b As Byte)
W_UpdateMode = b
If b = 0 Then
    Me.Caption = "新增发票明细"
Else
    Me.Caption = "修改发票明细"
    deliv_no.Locked = True
    deliv_no.BackColor = deliv_date.BackColor
    deliv_no.TabStop = False
    Command1.Enabled = False
   
End If
End Property
Public Property Get CallForm() As Form
Set CallForm = W_CallForm
End Property

Public Property Set CallForm(f As Form)
Set W_CallForm = f
End Property

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Me.UpdateMode = 0 Then
    With FrmList
        .G_Sql_Filter = "SELECT Deliv_No as 货单单号,CONVERT(nvarchar(11),Deliv_Date,21) as 货单日期 " & _
                         "FROM mmst501 " & _
                         "WHERE inv_no is null or inv_no='' " & _
                              "AND Deliv_no like '" & Trim(deliv_no.Text) & "%' " & _
                              "AND cust_no='" & Me.CallForm.cust_no.Text & "' and deliv_type='1'"
        .Caption = "货单单号列表"
        .Show vbModal
        If .Col_No1 <> "" Then
            deliv_no.Text = .Col_No1
            deliv_date.Text = Format(.Col_No2, "yyyy-mm-dd")
            Remark.SetFocus
        End If
   End With
End If
End Sub



Private Sub cust_no_Change()

End Sub

Private Sub Form_Load()
Me.KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And ActiveControl.Name <> "Remark" Then
    SendKeys "{TAB}"
End If
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               '确认
       Call CmdOk_Click
    Case vbKeyEscape           '取消
       Call CmdCancel_Click
    End Select
End If
End Sub
Private Sub CmdOk_Click()
If check_ok Then
    If Me.UpdateMode = 0 Then
        G_Con.Execute "Update mmst501 SET inv_no = '" & Me.CallForm.inv_no.Text & "' WHERE deliv_no = '" & deliv_no.Text & "'"
    Else
        G_Con.Execute "Update mmst501 SET Remark = '" & Remark.Text & "' WHERE deliv_no = '" & deliv_no.Text & "'"
    End If
        
    Call Me.CallForm.RefreshGrid
    If Me.UpdateMode = 0 Then
        Call ClearFields
    Else
        Unload Me
    End If
End If
End Sub

Private Function check_ok() As Boolean
    Dim w_deliv_no As String
    Dim W_Mtr_No As String
    w_deliv_no = ""
   
    w_deliv_no = Trim(deliv_no.Text)

    
    Dim w_Rs As New ADODB.Recordset
    
    If Me.UpdateMode = 0 Then
         
        If w_deliv_no = "" And Command1.Visible = True Then
            MsgBox "请输入货单单号.", vbExclamation, g_CON_CTitle
            deliv_no.SetFocus
            Exit Function
       
        Else
            w_Rs.Open "SELECT deliv_no FROM mmst501 WHERE deliv_no='" & w_deliv_no & "'", G_Con
            If w_Rs.EOF = True Then
                w_Rs.Close
                MsgBox "无此货单编号.", vbExclamation, g_CON_CTitle
                deliv_no.SetFocus
                Exit Function
            End If
            w_Rs.Close
            
       End If
       
   End If
    
    
                 
    check_ok = True
End Function

Private Sub Form_Unload(Cancel As Integer)
Set FrmInvMx2 = Nothing
End Sub

Private Sub mtr_no_LostFocus()

End Sub

Sub ClearFields()
deliv_no.Text = ""
deliv_date.Text = ""
Remark.Text = ""
End Sub

