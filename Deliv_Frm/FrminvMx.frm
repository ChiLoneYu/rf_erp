VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrminvMx 
   BackColor       =   &H80000009&
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   FillStyle       =   0  '���
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
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
   ScaleHeight     =   4740
   ScaleWidth      =   5970
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox mtr_amt 
      Appearance      =   0  '����
      Height          =   315
      Left            =   1335
      TabIndex        =   6
      Top             =   1815
      Width           =   1920
   End
   Begin VB.TextBox mtr_dim 
      Appearance      =   0  '����
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1395
      Width           =   4260
   End
   Begin VB.TextBox mtr_name 
      Appearance      =   0  '����
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   990
      Width           =   4260
   End
   Begin VB.TextBox spe_let 
      Appearance      =   0  '����
      Height          =   315
      Left            =   4050
      TabIndex        =   10
      Top             =   2220
      Width           =   1545
   End
   Begin VB.TextBox qc_result 
      Appearance      =   0  '����
      Height          =   315
      Left            =   1335
      TabIndex        =   11
      Top             =   2625
      Width           =   4275
   End
   Begin VB.TextBox note 
      Appearance      =   0  '����
      Height          =   630
      Left            =   1335
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   12
      Top             =   3060
      Width           =   4260
   End
   Begin VB.CommandButton cmd_qc_brow 
      Appearance      =   0  '����
      Caption         =   "..."
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   2235
      Width           =   255
   End
   Begin VB.CommandButton cmd_mtr_brow 
      Caption         =   "..."
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmd_mo_brow 
      Caption         =   "..."
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   195
      Width           =   255
   End
   Begin VB.CommandButton CmdOK 
      Height          =   345
      Left            =   1560
      Picture         =   "FrminvMx.frx":0000
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   13
      Top             =   4125
      Width           =   1110
   End
   Begin VB.CommandButton CmdCancel 
      Height          =   345
      Left            =   3300
      Picture         =   "FrminvMx.frx":15A2
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   14
      Top             =   4125
      Width           =   1110
   End
   Begin VB.TextBox mo_no 
      Appearance      =   0  '����
      Height          =   315
      Left            =   1335
      TabIndex        =   0
      Top             =   180
      Width           =   1920
   End
   Begin VB.TextBox mtr_no 
      Appearance      =   0  '����
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   585
      Width           =   1920
   End
   Begin VB.TextBox qc_no 
      Appearance      =   0  '����
      Height          =   315
      Left            =   1335
      OLEDropMode     =   1  '���
      TabIndex        =   8
      Top             =   2220
      Width           =   1920
   End
   Begin MSForms.ComboBox bar_no 
      Height          =   315
      Left            =   4050
      TabIndex        =   7
      Top             =   1815
      Width           =   1545
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2725;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "�s�ө���"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��        �`:"
      Height          =   195
      Left            =   270
      TabIndex        =   24
      Tag             =   "Remark:"
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��        �q:"
      Height          =   195
      Left            =   270
      TabIndex        =   23
      Tag             =   "Qty:"
      Top             =   1890
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�W        ��:"
      Height          =   195
      Left            =   270
      TabIndex        =   22
      Tag             =   "Standard:"
      Top             =   1470
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�~        �W:"
      Height          =   195
      Left            =   270
      TabIndex        =   21
      Tag             =   "Product Name:"
      Top             =   1065
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���~�N��:"
      Height          =   195
      Left            =   270
      TabIndex        =   20
      Tag             =   "Product Code:"
      Top             =   645
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�Ͳ��渹:"
      Height          =   195
      Left            =   270
      TabIndex        =   19
      Tag             =   "Order No.:"
      Top             =   240
      Width           =   825
   End
   Begin MSForms.Label Label6 
      Height          =   195
      Left            =   3540
      TabIndex        =   18
      Top             =   1890
      Width           =   450
      BackColor       =   -2147483639
      VariousPropertyBits=   276824083
      Caption         =   "�ܧO:"
      Size            =   "794;344"
      FontName        =   "�s�ө���"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label7 
      Height          =   195
      Left            =   270
      TabIndex        =   17
      Top             =   2295
      Width           =   840
      BackColor       =   -2147483639
      VariousPropertyBits=   276824083
      Caption         =   "���˳渹:"
      Size            =   "1482;344"
      FontName        =   "�s�ө���"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label8 
      Height          =   390
      Left            =   270
      TabIndex        =   16
      Top             =   2760
      Width           =   945
      BackColor       =   -2147483639
      VariousPropertyBits=   276824083
      Caption         =   "��  ��  �_:"
      Size            =   "1667;688"
      FontName        =   "�s�ө���"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label10 
      Height          =   195
      Left            =   3540
      TabIndex        =   15
      Top             =   2295
      Width           =   450
      BackColor       =   -2147483639
      VariousPropertyBits=   276824083
      Caption         =   "�S��:"
      Size            =   "794;344"
      FontName        =   "�s�ө���"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "FrminvMx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*�A�Ω󦨫~�J�ܳ�s�W/�ק�J�ܲ��~����
'*Written by OscarChan,
'*at 2001-07-26
'*********************************************************
Dim W_UpdateMode As Byte '0:add,1:edit
Dim W_CallForm As Form
Public Property Get UpdateMode() As Byte
UpdateMode = W_UpdateMode
End Property
Public Property Let UpdateMode(b As Byte)
W_UpdateMode = b
If b = 0 Then
    Me.Caption = "�s�W���~"
Else
    Me.Caption = "�ק令�~"
    mo_no.Locked = True
    mo_no.BackColor = Mtr_Name.BackColor
    mo_no.TabStop = False
    Cmd_brow2.Enabled = False
End If
End Property
Public Property Get CallForm() As Form
Set CallForm = W_CallForm
End Property
Public Property Set CallForm(f As Form)
Set W_CallForm = f
End Property

Private Sub cmd_mo_brow_Click()
With FrmmoList
    .Show vbModal
    If .mo_no <> "" Then
        mo_no.Text = .mo_no
        Mtr_No.SetFocus
    Call mo_No_LostFocus
    End If
End With
End Sub

Private Sub Cmd_Mtr_Brow_Click()
With FrmMtrList
    .Show vbModal
    If .Mtr_No <> "" Then
        Mtr_No.Text = .Mtr_No
        Mtr_Name.Text = .Mtr_Name
        Mtr_Dim.Text = .Mtr_Dim
        
        Call mtr_no_LostFocus
        
        Mtr_Amt.SetFocus
    End If
End With
End Sub

Private Sub Cmd_Qc_Brow_Click()
With FrmIQCList
    
    .Show vbModal
    If Not .Cancel_Click Then
        Qc_No.Text = .Qc_No
        qc_result.Text = .qc_result
        Spe_Let.Text = .qc_spe_let
        
        Call qc_spe_let_LostFocus
        
        Note.SetFocus
    End If
End With
End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
Dim w_temp As New ADODB.Recordset
w_temp.CursorLocation = adUseClient
w_temp.Open "select bar_no ,bar_name from mmst903 ", G_Con, adOpenForwardOnly, adLockPessimistic


While w_temp.EOF <> True
Bar_No.AddItem w_temp!bar_name
w_temp.MoveNext
Wend
w_temp.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And ActiveControl.Name <> "note" Then
    SendKeys "{TAB}"
End If
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               '�T�{
         If CmdOK.Enabled = True Then
             Call CmdOk_Click
         End If
    Case vbKeyEscape           '����
         If CmdCancel.Enabled = True Then
             Call CmdCancel_Click
         End If
    End Select
End If
End Sub


Private Sub CmdOk_Click()
If check_ok Then
    Dim w_rs As New ADODB.Recordset
   
    With w_rs
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        Set .ActiveConnection = G_Con
        If Me.UpdateMode = 0 Then
            .Open "select * from mmst532  where inv_no='' "
            .AddNew
            !inv_no = Trim(mmss601.inv_no.Text)
            !mo_no = Trim(mo_no.Text)
            '!Mtr_Name = Trim(Mtr_Name.Text)
            !Mtr_Dim = Trim(Mtr_Dim.Text)
            !Mtr_Amt = Val(Mtr_Amt.Text)
            !Mtr_No = Trim(Mtr_No.Text)
            !Bar_No = Trim(Bar_No.Text)
            !Qc_No = Trim(Qc_No.Text)
            '!qc_result = Trim(qc_result.Text)
            !Spe_Let = Trim(Spe_Let.Text)
            !Note = Trim(Note.Text)
            .Update
        Else
            .Open "select * from mmst532 where mmst532.inv_no='" & mmss601.inv_no.Text & "'" & _
                  " AND mo_no='" & mo_no.Text & "'" & " AND mtr_no='" & Mtr_No.Text & "'"
            !mo_no = Trim(mo_no.Text)
            !Mtr_Name = Trim(Mtr_Name.Text)
            !Mtr_Dim = Trim(Mtr_Dim.Text)
            !Mtr_Amt = Val(Mtr_Amt.Text)
            !Mtr_No = Trim(Mtr_No.Text)
            !Bar_No = Trim(Bar_No.Text)
            !Qc_No = Trim(Qc_No.Text)
            !qc_result = Trim(qc_result.Text)
            !Spe_Let = Trim(Spe_Let.Text)
            !Note = Trim(Note.Text)
            .Update
        End If
            .Close
    End With
    Set w_rs = Nothing
    Call Me.CallForm.RefreshGrid
    If Me.UpdateMode = 0 Then
        Call ClearFields
        mo_no.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Function check_ok() As Boolean
   Dim w_status As String
   w_status = CheckStatus("inv_no", Me.CallForm.inv_no.Text, "mmst531", "status")
   If w_status <> "0" Then
        If w_status = "9" Then
            MsgBox "���" & Me.CallForm.inv_no.Text & "�w�Q�䥦�Τ�R��.", vbInformation, g_CON_CTitle
        Else
            MsgBox "���" & Me.CallForm.inv_no.Text & "�w�Q�f��,����s�W�έק����.", vbExclamation, g_CON_CTitle
        End If
        check_ok = False
        Unload Me
        Exit Function
    End If
    
    Dim w_rs As New ADODB.Recordset
        
  
    
    If Me.UpdateMode = 0 Then
        If Trim(mo_no.Text) = "" Then
            MsgBox "�п�J�Ͳ���Ƴ渹.", vbExclamation, g_CON_CTitle
            mo_no.SetFocus
            Exit Function
        End If
        w_rs.Open "select mo_no from mmst401 where mo_no='" & mo_no.Text & "'", G_Con
        If w_rs.EOF = True Then
            w_rs.Close
            MsgBox "�L����Ƴ渹.", vbExclamation, g_CON_CTitle
            mo_no.SetFocus
            Exit Function
        End If
        w_rs.Close
        If Trim(Mtr_No) = "" Then
            MsgBox "������J���~�N��.", vbExclamation, g_CON_CTitle
            Mtr_No.SetFocus
            Exit Function
        Else
            w_rs.Open "select *  from mmst401 where mmst401.mo_no='" & Me.mo_no.Text & "' and mtr_no='" & Trim(Mtr_No.Text) & "'", G_Con, , , adCmdText
            If w_rs.EOF Then
                MsgBox "�ӧ�Ƴ�L�����~!", vbExclamation, g_CON_CTitle
                Mtr_No.SetFocus
                Exit Function
            End If
            w_rs.Close
        End If

        w_rs.Open "select *  from mmst532 where mmst532.inv_no='" & Me.CallForm.inv_no.Text & "' and mtr_no='" & Trim(Mtr_No.Text) & "' and mo_no = '" & Me.mo_no.Text & "'", G_Con, , , adCmdText
        If w_rs.EOF = False Then
            MsgBox "��Ƴ渹+���~�N������.", vbExclamation, g_CON_CTitle
            Mtr_No.SetFocus
            Exit Function
        End If
        w_rs.Close
    End If
    
    If Val(Mtr_Amt.Text) <= 0 Or Val(Mtr_Amt.Text) > 1000000 Then
        MsgBox "�п�J���T���ƶq.", vbExclamation, g_CON_CTitle
        Mtr_Amt.SetFocus
        Exit Function
    End If
    
    Set w_rs = Nothing
    check_ok = True
End Function

Private Sub Form_Unload(Cancel As Integer)
Set FrmProdsMx = Nothing
End Sub



Public Sub ClearFields()
mo_no.Text = ""
Mtr_No.Text = ""
Mtr_Name.Text = ""
Mtr_Dim.Text = ""
Mtr_Amt.Text = ""
Bar_No.Text = ""
Qc_No.Text = ""
qc_result.Text = ""
Spe_Let.Text = ""
Note.Text = ""
End Sub

Private Sub mo_No_LostFocus()
If Me.UpdateMode <> 0 Then
    Exit Sub
End If
If Trim(mo_no.Text) <> "" Then
    Dim w_rs As New ADODB.Recordset
    w_rs.Open "SELECT mo_No FROM mmst401 WHERE mo_No='" & Trim(mo_no.Text) & "'", G_Con, , , adCmdText
    If w_rs.EOF Then
         MsgBox "�L���Ͳ��ʳ渹!", vbExclamation, g_CON_CTitle
         mo_no.Text = ""
         mo_no.SetFocus
         w_rs.Close
         Set w_rs = Nothing
         Exit Sub
    End If
    w_rs.Close
End If
End Sub

Private Sub mtr_no_LostFocus()
If Me.UpdateMode <> 0 Then
    Exit Sub
End If
If Trim(Mtr_No.Text) <> "" Then
    Dim w_rs As New ADODB.Recordset
    w_rs.Open "SELECT mtr_No,mtr_name  FROM mmst611 WHERE mtr_No='" & Trim(Mtr_No.Text) & "'", G_Con, , , adCmdText
    If w_rs.EOF Then
         MsgBox "�L�����~�N��!", vbExclamation, g_CON_CTitle
         Mtr_No.Text = ""
         Mtr_Name.Text = ""
         Mtr_No.SetFocus
         w_rs.Close
         Set w_rs = Nothing
         Exit Sub
    End If
    w_rs.Close
End If
End Sub



Private Sub qc_spe_let_LostFocus()
If Me.UpdateMode <> 0 Then
    Exit Sub
End If
If Trim(Spe_Let.Text) <> "" Then
    Dim w_rs As New ADODB.Recordset
    w_rs.Open "SELECT qc_No,qc_result FROM mmsta15 WHERE qc_No='" & Trim(Qc_No.Text) & "'", G_Con, , , adCmdText
    If w_rs.EOF Then
         MsgBox "�L���Ͳ��ʳ渹!", vbExclamation, g_CON_CTitle
         Qc_No.Text = ""
         qc_result.Text = ""
         
        Qc_No.SetFocus
         w_rs.Close
         Set w_rs = Nothing
         Exit Sub
    End If
    w_rs.Close
End If
End Sub
