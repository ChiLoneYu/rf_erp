VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϵͳ����"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "FrmPack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Set company's data(810)"
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�������"
      Height          =   1725
      Left            =   3000
      TabIndex        =   18
      Top             =   3840
      Width           =   7815
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "ȡ��"
         Height          =   345
         Left            =   6420
         TabIndex        =   21
         Top             =   1260
         Width           =   795
      End
      Begin VB.CommandButton Cmd_Yes 
         Caption         =   "ȷ��"
         Height          =   345
         Left            =   5190
         TabIndex        =   20
         Top             =   1260
         Width           =   795
      End
      Begin VB.CommandButton Cmd_Select 
         Caption         =   "ѡ��"
         Height          =   345
         Left            =   6390
         TabIndex        =   19
         Top             =   300
         Width           =   795
      End
      Begin MSComDlg.CommonDialog CmnDlg 
         Left            =   0
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox File_Path 
         Height          =   345
         Left            =   1110
         TabIndex        =   22
         Top             =   270
         Width           =   6075
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�ļ�λ��:"
         Height          =   285
         Left            =   300
         TabIndex        =   23
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "��������������ϵ�"
      Height          =   465
      Left            =   4800
      TabIndex        =   17
      Top             =   4800
      Width           =   2325
   End
   Begin VB.CommandButton Command14 
      Caption         =   "������˿������ϵ�"
      Height          =   465
      Left            =   240
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton Command13 
      Caption         =   "���½᰸Ȩ������"
      Height          =   465
      Left            =   3840
      TabIndex        =   14
      Top             =   4200
      Width           =   3345
   End
   Begin VB.CommandButton Command12 
      Caption         =   "�����ⷢ����BOM"
      Height          =   465
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command11 
      Caption         =   "�ֿ�������"
      Height          =   465
      Left            =   5550
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton Upd0122_Hid_Print 
      Caption         =   "���¶���BOM����������(mmst0122)"
      Height          =   465
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ƽ����Ŀ����Ĳ���"
      Height          =   465
      Left            =   3780
      TabIndex        =   10
      Top             =   3570
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Cmd_Upd0122 
      Caption         =   "���¶���BOM��������λ(P_Line_Name)"
      Height          =   465
      Left            =   240
      TabIndex        =   9
      Top             =   2928
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   465
      Left            =   3780
      TabIndex        =   8
      Top             =   2904
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�����Ƶ�LOSS��Ч�ֶ�"
      Height          =   465
      Left            =   3780
      TabIndex        =   7
      Top             =   2238
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command7 
      Caption         =   "����mmst401_mtr����"
      Height          =   465
      Left            =   3780
      TabIndex        =   6
      Top             =   1572
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command6 
      Caption         =   "���¶�����������"
      Height          =   465
      Left            =   3780
      TabIndex        =   5
      Top             =   906
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�ֿ���������"
      Height          =   465
      Left            =   3780
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ɹ�/��������������"
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   2256
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BOM С��������"
      Enabled         =   0   'False
      Height          =   465
      Left            =   240
      TabIndex        =   2
      Top             =   1584
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����������������"
      Enabled         =   0   'False
      Height          =   465
      Left            =   240
      TabIndex        =   1
      Top             =   912
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���������տͻ�������ˮ���"
      Enabled         =   0   'False
      Height          =   465
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   4800
      Width           =   735
   End
End
Attribute VB_Name = "FrmPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub cmd_quit_Click()
Unload Me
End Sub


Private Sub Cmd_Cancel_Click()
Unload Me
End Sub

Private Sub Cmd_Select_Click()
 
Dim W_File_Name As String
Dim W_txtPath As String

CmnDlg.Filter = "*.txt|*.txt|*.sql|*.sql"
CmnDlg.InitDir = G_Path & "\model"
CmnDlg.Action = 1

W_txtPath = CmnDlg.FileName
File_Path.Text = W_txtPath
W_File_Name = W_txtPath

If W_File_Name <> "" Then
    If FileExists(W_File_Name) = False Then
        MsgBox W_File_Name & "���ļ�������,������ѡ��.", vbOKOnly + vbExclamation, g_CON_CTitle
        W_File_Name = ""
    Else
        
    End If
End If
End Sub

Private Sub Cmd_Upd0122_Click()
G_Con.Execute "UPDATE mmst0122 SET P_Line_Name=mmst612.P_Line_Name " & _
                         "FROM  mmst0122 INNER JOIN mmst612 ON " & _
                         "mmst0122.Mtr_No=mmst612.Mtr_No AND mmst0122.Bom_No=mmst612.Bom_No "

MsgBox "�������"
End Sub

Private Sub Cmd_Yes_Click()
 Call lprodata(File_Path.Text)

End Sub


Private Sub lprodata(p_file_name)
Dim w_handel
Dim temp As New ADODB.Recordset
Dim tmp_Str As String
Dim w_TextLine
'On Error Resume Next

'�����ʱ����
tmp_Str = ""
w_handel = FreeFile

'ProgressBar1.Value = 0

Open p_file_name For Input As w_handel

'W_Rec_Amt = Int(LOF(w_handel) / 36)

W_I = 0

Do While Not EOF(w_handel)
    W_I = W_I + 1
    Line Input #w_handel, w_TextLine
    If UCase(w_TextLine) = "GO" Then
        G_Con.Execute tmp_Str
        tmp_Str = ""
    Else
        tmp_Str = tmp_Str + Chr(10) + w_TextLine

    End If
Loop
'Debug.Print Tmp_str
Close w_handel


MsgBox "������ɣ�"

End Sub
Private Sub Command1_Click()

Dim Tmp_Rb As New ADODB.Recordset
Dim Tmp_Old_NO As String

Dim Tmp_New_NO As String

Set Tmp_Rb = open_RecordSet("select * from mmst501 Where Deliv_type='1' order by deliv_no")

Do Until Tmp_Rb.EOF
    Tmp_Old_NO = Tmp_Rb!deliv_no
    Tmp_New_NO = Creat_No(Tmp_Rb!Cust_No, Tmp_Rb!deliv_date)
    
    G_Con.Execute "update mmst502 set deliv_no='" & Tmp_New_NO & "' Where  deliv_no='" & Tmp_Old_NO & "'"
    G_Con.Execute "update mmst501 set remark='" & "ԭ������:" & Tmp_Old_NO & "' Where  deliv_no='" & Tmp_Old_NO & "'"
    G_Con.Execute "update mmst501 set deliv_no='" & Tmp_New_NO & "' Where  deliv_no='" & Tmp_Old_NO & "'"
    Tmp_Rb.MoveNext
Loop

MsgBox "�������"
End Sub

Private Function Creat_No(P_Cust_NO As String, p_date As Date) As String


    Dim w_tmp As New ADODB.Recordset
    Dim W_Str As String

    Dim W_Deliv_No As String

    W_Deliv_No = "D-" & P_Cust_NO          '����

    W_Deliv_No = W_Deliv_No & Right(CStr(Year(p_date)), 2) & Format(CStr(Month(p_date)), "00") & Format(CStr(Day(p_date)), "00")

    W_Str = "SELECT Max(deliv_no) As deliv_no  FROM mmst501 WHERE deliv_no like '" & W_Deliv_No & "%' "
            
    w_tmp.Open W_Str, G_Con, adOpenForwardOnly, adLockReadOnly, adCmdText

    If w_tmp.EOF = False Then
        If IsNull(w_tmp!deliv_no) Then
            W_Deliv_No = W_Deliv_No & "001"
        Else
            W_Deliv_No = W_Deliv_No & Format(CStr(Val(Right(w_tmp!deliv_no, 3)) + 1), "000")
        End If
    Else
        W_Deliv_No = W_Deliv_No & "001"
    End If

    Creat_No = W_Deliv_No
End Function

Private Sub Command10_Click()
Dim Tmp_Rb As New ADODB.Recordset
Dim tmp_Str As String

Set Tmp_Rb = open_RecordSet("select * from mmst321 where Inv_No='T060331999'")
If Tmp_Rb.EOF Then
    With Tmp_Rb
         .AddNew
         !Inv_No = "T060331999"
         !Inv_date = #3/31/2006#
         !Inv_Type = "�䡡��"
         !form_man = "����Ա"
         !check_man = "����Ա"
         !Status = "3"
         !Remark = "ƽ����Ŀ"
         !Upd_Name = "����Ա"
         !upd_date = #4/4/2006#
         !lock = "NO"
         .Update
    End With
    
    
    tmp_Str = "Insert Into mmst322  (Inv_No,Mtr_No,Mtr_Amt,Mtr_Res,Bar_No,Note,Upd_Name,Upd_Date) values (" & _
                          "'T060331999','21522060010100',5,0,'001','������Ŀ','����Ա','4/4/2006')"
    
    
    G_Con.Execute tmp_Str
    

    
    tmp_Str = "Insert Into mmst322  (Inv_No,Mtr_No,Mtr_Amt,Mtr_Res,Bar_No,Note,Upd_Name,Upd_Date) values (" & _
                          "'T060331999','21547008010105',100,0,'001','������Ŀ','����Ա','4/4/2006')"
    
    
    G_Con.Execute tmp_Str
    
    tmp_Str = "Insert Into mmst322  (Inv_No,Mtr_No,Mtr_Amt,Mtr_Res,Bar_No,Note,Upd_Name,Upd_Date) values (" & _
                          "'T060331999','30140020010100',258,0,'001','������Ŀ','����Ա','4/4/2006')"
    
    
    G_Con.Execute tmp_Str
    
    tmp_Str = "Insert Into mmst322  (Inv_No,Mtr_No,Mtr_Amt,Mtr_Res,Bar_No,Note,Upd_Name,Upd_Date) values (" & _
                          "'T060331999','40412432030251',0,2,'008','������Ŀ','����Ա','4/4/2006')"
    
    
    G_Con.Execute tmp_Str
    
    

End If

Command10.Enabled = False
MsgBox "�������"
End Sub

Private Sub Command11_Click()
Dim tmp_Str As String

G_Con.Execute "Delete From mmii381 "

tmp_Str = "select Mtr_No,Bar_No,sum(mmst381.Mtr_Amt)  as Mtr_Amt From mmst381 Group By Mtr_No,Bar_No"
G_Con.Execute "insert into mmii381 (Mtr_No,Bar_No,mtr_amt) " & tmp_Str

G_Con.Execute "Delete From mmst381 "

tmp_Str = "select Mtr_No,Bar_No,sum(mmii381.Mtr_Amt)  as Mtr_Amt,'" & G_User_Name & "' as Upd_Name,'" & Get_SQLDATE & "' as Upd_Date From mmii381 Group By Mtr_No,Bar_No"
G_Con.Execute "insert into mmst381 (Mtr_No,Bar_No,mtr_amt,Upd_Name,Upd_date) " & tmp_Str

MsgBox "�������"

End Sub

Private Sub Command12_Click()
Dim Tmp_Rb As New ADODB.Recordset

Dim tmp_Str As String


Set Tmp_Rb = open_RecordSet("select Pcs_No from mmst207")
Do Until Tmp_Rb.EOF
    Call Out_Bom(Tmp_Rb!pcs_no)
    Tmp_Rb.MoveNext
Loop
tmp_Str = " Update mmst308 Set Bom_Amt=mmst208_boM.Bom_Amt " & _
          " FROM  mmst208_Bom INNER JOIN  mmst308 ON mmst208_Bom.Pcs_No = mmst308.Pcs_No AND   mmst208_Bom.Mo_No = mmst308.Mo_No AND   mmst208_Bom.Order_No = mmst308.order_no AND   mmst208_Bom.Cust_Order_No = mmst308.Cust_Order_No AND  mmst208_Bom.P_Mtr = mmst308.P_Mtr AND   mmst208_Bom.Cust_Mtr_No = mmst308.Cust_Mtr_No AND  mmst208_Bom.Mtr_No = mmst308.P_Mtr1 AND   mmst208_Bom.Bom_No = mmst308.Mtr_No "
          
G_Con.Execute tmp_Str
    
tmp_Str = " Update mmst308 Set Bom_Amt=1 Where Bom_Amt is null"
G_Con.Execute tmp_Str

MsgBox "�������"

End Sub

Private Sub Command13_Click()
G_Con.Execute "update mmstc02 set prog_type='CRIO' where menu_id='menu_e5_2'"

G_Con.Execute "update mmstc02 set prog_type='CRIO' where menu_id='menu_e5_4'"

MsgBox "�������"

End Sub

Private Sub upd_data()
Dim W_Rs As New ADODB.Recordset
Dim W_SQL As String
Dim W_Inv_No As String
Dim c_reset As Boolean
Dim c_check As Boolean

c_check = True
Dim i As Double

If c_check = True Then    '���
    W_SQL = "UPDATE mmst361 SET Status='1',Lock='No',Check_Man ='" & G_User_Name & "' ,Upd_Name='" & Trim(G_User_Name) & "',Upd_Date='" & Get_SQLDATE & "' WHERE status=0 "
    
ElseIf c_reset = True Then '����
    W_SQL = "UPDATE mmst361 SET Status='0',Lock='No',Check_man ='',Upd_Name='" & Trim(G_User_Name) & "',Upd_Date='" & Get_SQLDATE & "' WHERE status=0 "
End If

W_Rs.CursorLocation = adUseClient
W_Rs.Open "SELECT mmst361.inv_no,mmst362.case_no as order_no,mmst362.Mtr_No," & _
                 "Mtr_Amt," & _
                 "mmst362.Bar_No " & _
            " FROM mmst362,mmst611,mmst361  " & _
            " WHERE  mmst362.mtr_no=mmst611.mtr_no and mmst361.inv_no=mmst362.inv_no AND MMST361.STATUS=0", G_Con, , , adCmdText

''**********************��ʼ��������*********************
On Error GoTo UpdateError
G_Con.BeginTrans
G_Con.Execute W_SQL
If c_reset = True Then
'    Do While Not w_rs.EOF
'        '���¿��
'        Call Sum_StockM(w_rs!order_no, w_rs!mtr_no, w_rs!bar_no, w_rs!Mtr_Amt)
'
'        'д388 ��
'        Call Stock_DelM(W_Inv_No, W_Inv_Date, "�������ϵ�", w_rs!order_no, w_rs!mtr_no, w_rs!bar_no, -w_rs!Mtr_Amt, 0)
'        w_rs.MoveNext
'    Loop
Else
    Do While Not W_Rs.EOF
        '������Ƿ�
'        If Mtr_StockM(w_rs!order_no, w_rs!mtr_no, w_rs!bar_no, w_rs!Mtr_Amt) = False Then
'            G_Con.RollbackTrans
'            MsgBox "��Ʒ" & w_rs!mtr_no & " ��治��.", vbCritical, g_CON_CTitle
'            Call UnLockRecord("mmst361", "Inv_No='" & W_Inv_No & "'")
'            GoTo Endx
'            Exit Sub
'        End If
            Label1.Caption = i
         '���¿��
        Call Sum_StockM(W_Rs!order_no, W_Rs!mtr_no, W_Rs!Bar_No, -W_Rs!Mtr_Amt)
    
        'ɾ����388�е����� ��
        W_Inv_No = W_Rs!Inv_No
        Call Stock_AddM(W_Inv_No, Date, "�������ϵ�", W_Rs!order_no, W_Rs!mtr_no, W_Rs!Bar_No, -W_Rs!Mtr_Amt, 0)
        W_Rs.MoveNext
        i = i + 1
    Loop
End If

'���ĵ���״̬

G_Con.CommitTrans
'
'Help_txt.Caption = IIf(c_check, "���", "����") & "�ɹ�!"
'Help_txt.Refresh

MsgBox "���������ɣ�TKS"
For i = 1 To 8000000
Next i

'���»�ȡ�����ʾ��Ϣ
'Call Erp_Proj.Warn_Check
Endx:
c_check = False
c_reset = False
'Call Inv_No_Click

Exit Sub

UpdateError:
G_Con.RollbackTrans
MsgBox "����ʱ��������!", 64, g_CON_CTitle
'�������
'Call UnLockRecord("mmst361", "inv_no='" & W_Inv_No & "'")
GoTo Endx

End Sub
Private Sub upd_datammst311()
Dim W_Rs As New ADODB.Recordset
Dim Tmp_Rb As New ADODB.Recordset
Dim W_Order_No As String
Dim W_SQL As String
Dim W_Inv_No As String
Dim tmp_Str As String

Dim c_check As Boolean

Dim c_reset As Boolean

c_check = True



Dim i As Double

If c_check = True Then    '���
    W_SQL = "UPDATE mmst311 SET Status='1',Lock='No',Check_Man ='" & G_User_Name & "' ,Upd_Name='" & Trim(G_User_Name) & "',Upd_Date='" & Get_SQLDATE & "' WHERE status=0"
    
ElseIf c_reset = True Then '����
    W_SQL = "UPDATE mmst311 SET Status='0',Lock='No',Check_man ='',Upd_Name='" & Trim(G_User_Name) & "',Upd_Date='" & Get_SQLDATE & "' WHERE status=0"
End If

'If c_check = True Then
'    '������Ƿ�
'    Tmp_Str = " SELECT a.order_no , a.Mtr_No, c.Bar_name , a.Mtr_Amt, isnull(b.mtr_amt,0) as bar_amt   " & _
'              " FROM mmst312 a " & _
'              "     inner join mmst903 c on  a.bar_no=c.bar_no " & _
'              "     Left Join mmst381 b On a.mtr_no=b.mtr_no and a.bar_no=b.bar_no and a.order_no=b.order_no  " & _
'              " WHERE  a.inv_no='" & W_Inv_No & "' and  a.mtr_amt>isnull(b.mtr_amt,0)  " & _
'              " order By a.order_no, a.Mtr_No,a.Bar_No  "
'
'    Set W_RS = open_RecordSet(Tmp_Str)
'
'    If W_RS.RecordCount > 0 And W_inv_date >= "2014-12-24" Then
'        Tmp_Str = "���²ֿ������Ͽ�治��:" & vbCrLf
'        Do While Not W_RS.EOF
'            '������Ƿ�
'            Tmp_Str = Tmp_Str & "      ������" & W_RS!order_no & "���ֿ⡾" & W_RS!Bar_Name & "�������ϡ�" & W_RS!mtr_no & "��������������,��������:" & W_RS!Mtr_Amt - W_RS!bar_amt & "   " & vbCrLf
'            W_RS.MoveNext
'        Loop
'        MsgBox Tmp_Str, vbCritical, g_CON_CTitle
'        Call UnLockRecord("mmst311", "Inv_No='" & W_Inv_No & "'")
'        GoTo Endx
'        Set W_RS = Nothing
'    End If
'
'
'End If

tmp_Str = " SELECT a.inv_no,a.order_no,a.mtr_no,a.mtr_amt,a.bar_no " & _
            " FROM mmst312 a inner join mmsp012_mtr b on a.order_no=b.order_no and a.mtr_no=b.mtr_no" & _
            " inner join mmst311 c on c.inv_no=a.inv_no" & _
            " WHERE c.status=0 and c.inv_date<'2014-12-25' " & _
            " Order By a.order_no,a.mtr_no "

Set W_Rs = open_RecordSet(tmp_Str)


''**********************��ʼ��������*********************
On Error GoTo UpdateError
G_Con.BeginTrans
G_Con.Execute W_SQL

If c_reset = True Then
'    Do While Not W_RS.EOF
'
'        W_Order_No = NullSetValue(W_RS!order_no, "")
'        '���¿����
'        Call Sum_StockM_WM(W_Order_No, W_RS!mtr_no, W_RS!Bar_No, W_RS!Mtr_Amt)
'        'д388
'        Call Stock_DelM2(W_Inv_No, "�������ϵ�", W_RS!mtr_no, W_RS!Bar_No)
'        '�����Ƶ���������
'        G_Con.Execute " exec TsUpdOrderLineamt '" & W_Order_No & "','" & W_RS!mtr_no & "','1' "
'
'        W_RS.MoveNext
'    Loop
Else
'    '�ж��Ƿ񳬹����������
'    Tmp_Str = " select a.order_no,b.mtr_no,b.mtr_name,b.mtr_dim,a.mtr_amt,b.mtr_amt-isnull(b.ling_amt,0)+isnull(b.tui_amt,0) as Ling_amt " & _
'              " from mmst312 a inner join mmsp012_mtr b on a.order_no=b.order_no and a.mtr_no=b.mtr_no " & _
'              " where a.inv_no='" & Trim(Inv_No.Text) & "' and (a.mtr_amt-(b.mtr_amt-isnull(b.ling_amt,0)+isnull(b.tui_amt,0)))>0 "
'    Set Tmp_Rb = open_RecordSet(Tmp_Str)
'
'    If Tmp_Rb.EOF = False Then
'
'          Dim Tmp_amt As Double
'
'          Tmp_amt = Mtr_Base_Amt(Tmp_Rb!mtr_no, Tmp_Rb!ling_amt)
'
'          If Tmp_Rb!Mtr_Amt > Tmp_amt And W_inv_date >= "2014-12-24" Then
'            MsgBox "�Ƶ�Ϊ: " & Tmp_Rb!order_no & ", ����Ʒ," & Tmp_Rb!mtr_no & " �Ѿ�������������,�����벹��.�Ѿ���������Ϊ:" & Tmp_Rb!over_amt, vbCritical, g_CON_CTitle
'
'            G_Con.RollbackTrans
'            Call UnLockRecord("mmst311", "Inv_No='" & W_Inv_No & "'")
'            Set Tmp_Rb = Nothing
'            GoTo Endx
'            Exit Sub
'          End If
'    End If
'    Set Tmp_Rb = Nothing
    '���
    Do While Not W_Rs.EOF
    
        Label1.Caption = i
        Label1.Refresh
        
        W_Order_No = NullSetValue(W_Rs!order_no, "")
        
        W_Inv_No = NullSetValue(W_Rs!Inv_No, "")
        
        '���¿��
        Call Sum_StockM_WM(W_Order_No, W_Rs!mtr_no, W_Rs!Bar_No, -Mtr_Base_Amt(W_Rs!mtr_no, W_Rs!Mtr_Amt))
        'ɾ����388�е�����
        Call Stock_AddM(W_Inv_No, Date, "�������ϵ�", W_Order_No, W_Rs!mtr_no, W_Rs!Bar_No, -W_Rs!Mtr_Amt, 0)
        '�����Ƶ���������
        G_Con.Execute " exec TsUpdOrderLineamt '" & W_Order_No & "','" & W_Rs!mtr_no & "','1' "
        'move next
        i = i + 1
        W_Rs.MoveNext
    Loop
End If
'���ĵ���״̬

G_Con.CommitTrans

For i = 1 To 8000000
Next i
'���»�ȡ�����ʾ��Ϣ
'Call Erp_Bar.Warn_Check
Endx:
c_check = False
c_reset = False
'Call inv_no_Click
MsgBox "���ϵ�����������!", 64, g_CON_CTitle

Exit Sub

UpdateError:
G_Con.RollbackTrans
MsgBox "����ʱ��������!", 64, g_CON_CTitle
'�������
'Call UnLockRecord("mmst311", "inv_no='" & W_Inv_No & "'")
GoTo Endx

End Sub
Private Sub Command14_Click()
Call upd_data
End Sub

Private Sub Command15_Click()
Call upd_datammst311
End Sub

Private Sub Command2_Click()
Dim Tmp_Rb As New ADODB.Recordset


G_Con.Execute "Delete From mmst388 Where Inv_Type='�������ϵ���'"
Set Tmp_Rb = open_RecordSet(" SELECT mmst361.Inv_No, mmst361.Inv_Date, mmst362.Mtr_No,  mmst362.Bar_No , mmst362.Mtr_Amt " & _
                            " FROM   mmst361 INNER JOIN   mmst362 ON mmst361.Inv_No = mmst362.Inv_No " & _
                            " Where mmst361.status<>'0' ")

Do Until Tmp_Rb.EOF
    Call Stock_AddM(Tmp_Rb!Inv_No, Tmp_Rb!Inv_date, "�������ϵ���", "", Tmp_Rb!mtr_no, Tmp_Rb!Bar_No, -Tmp_Rb!Mtr_Amt, 0)
    Tmp_Rb.MoveNext
Loop

MsgBox "�������"
End Sub

Private Sub Command3_Click()

G_Con.Execute "Update  mmst612 set Bom_Amt=Bom_1/Bom_2 Where Bom_2<>0"
G_Con.Execute "Update  mmst0122 set Bom_Amt=Bom_1/Bom_2 Where Bom_2<>0"

MsgBox "�������"
End Sub

Private Sub Command4_Click()
Dim tmp_Str As String

'�ɹ���
tmp_Str = " Update mmst206 " & _
          " Set total_pay = a.total_pay " & _
          " FROM             mmst206,(SELECT  mmst302.Pcs_No, mmst302.Pcs_Need_No, mmst302.Mtr_No, " & _
                                            " SUM(CASE WHEN inv_type = '1' THEN 1 ELSE - 1 END *isnull(mmst302.Pay_Amt,0)) AS total_pay " & _
                                     " FROM  mmst301 INNER JOIN mmst302 ON  mmst301.Inv_No = mmst302.Inv_No " & _
                                     " WHERE mmst301.status <> '0' " & _
                                     " GROUP BY   mmst302.Pcs_No, mmst302.Pcs_Need_No, mmst302.Mtr_No) a " & _
         " WHERE         mmst206.pcs_no = a.pcs_no AND mmst206.pcs_need_no = a.pcs_need_no AND mmst206.Mtr_No = a.Mtr_No"
                           
G_Con.Execute tmp_Str

tmp_Str = " Update mmst206 " & _
         " Set inbar_amt = a.inbar_amt " & _
         " FROM             mmst206,(SELECT  mmst302.Pcs_No, mmst302.Pcs_Need_No, mmst302.Mtr_No, " & _
                                            " SUM(CASE WHEN inv_type = '1' THEN 1 ELSE - 1 END * isnull(mmst302.Mtr_Amt,0)) AS inbar_amt " & _
                                     " FROM  mmst301 INNER JOIN mmst302 ON  mmst301.Inv_No = mmst302.Inv_No " & _
                                     " WHERE mmst301.status <> '0' " & _
                                     " GROUP BY   mmst302.Pcs_No, mmst302.Pcs_Need_No, mmst302.Mtr_No) a " & _
        " WHERE         mmst206.pcs_no = a.pcs_no AND mmst206.pcs_need_no = a.pcs_need_no AND mmst206.Mtr_No = a.Mtr_No"
        
G_Con.Execute tmp_Str


'�й���
tmp_Str = " Update mmst208 " & _
         " Set total_pay = a.total_pay " & _
         " FROM             mmst208,(SELECT  mmst306.Pcs_No,  mmst306.Mtr_No,mmst306.Mo_No,mmst306.Cust_Order_No,mmst306.p_mtr, " & _
                                            " SUM(CASE WHEN inv_type = '1' THEN 1 ELSE - 1 END * isnull(mmst306.Pay_amt,0)) AS total_pay " & _
                                     " FROM  mmst305 INNER JOIN mmst306 ON  mmst305.Inv_No = mmst306.Inv_No " & _
                                     " WHERE mmst305.status <> '0' " & _
                                     " GROUP BY   mmst306.Pcs_No,  mmst306.Mtr_No,mmst306.Mo_No,mmst306.Cust_Order_No,mmst306.p_mtr ) a " & _
        " WHERE         mmst208.pcs_no = a.pcs_no  AND mmst208.Mtr_No = a.Mtr_No  AND mmst208.Mo_No = a.Mo_No  AND mmst208.Cust_Order_No = a.Cust_Order_No  AND mmst208.p_mtr = a.p_mtr "
                           
G_Con.Execute tmp_Str

tmp_Str = " Update mmst208 " & _
         " Set inbar_amt = a.inbar_amt " & _
         " FROM             mmst208,(SELECT  mmst306.Pcs_No,  mmst306.Mtr_No,mmst306.Mo_No,mmst306.Cust_Order_No,mmst306.p_mtr, " & _
                                            " SUM(CASE WHEN inv_type = '1' THEN 1 ELSE - 1 END * isnull(mmst306.mtr_Amt,0)) AS inbar_amt " & _
                                     " FROM  mmst305 INNER JOIN mmst306 ON  mmst305.Inv_No = mmst306.Inv_No " & _
                                     " WHERE mmst305.status <> '0' " & _
                                     " GROUP BY   mmst306.Pcs_No,  mmst306.Mtr_No,mmst306.Mo_No,mmst306.Cust_Order_No,mmst306.p_mtr) a " & _
        " WHERE         mmst208.pcs_no = a.pcs_no  AND mmst208.Mtr_No = a.Mtr_No  AND mmst208.Mo_No = a.Mo_No  AND mmst208.Cust_Order_No = a.Cust_Order_No  AND mmst208.p_mtr = a.p_mtr "
                           
G_Con.Execute tmp_Str

MsgBox "�������"
End Sub

Private Sub Command5_Click()
Dim tmp_Str As String
Dim Tmp_Rb  As New ADODB.Recordset

'��������ʵ��������
G_Con.Execute "Delete From mmii381 "

tmp_Str = "select Mtr_No,Bar_No,sum(mmst381.Mtr_Amt)  as Mtr_Amt From mmst381 Group By Mtr_No,Bar_No"
G_Con.Execute "insert into mmii381 (Mtr_No,Bar_No,mtr_amt) " & tmp_Str

G_Con.Execute "Delete From mmst381 "

tmp_Str = "select Mtr_No,Bar_No,sum(mmii381.Mtr_Amt)  as Mtr_Amt,'" & G_User_Name & "' as Upd_Name,'" & Get_SQLDATE & "' as Upd_Date From mmii381 Group By Mtr_No,Bar_No"
G_Con.Execute "insert into mmst381 (Mtr_No,Bar_No,mtr_amt,Upd_Name,Upd_date) " & tmp_Str



'������������Ŀ
G_Con.Execute "exec ts_Upd_388"

'�����������������Ŀ�Ĳ�������
If MsgBox("��Ҫ����������������?��������ᵱǰ����������Ŀ�����в����,�������ɿ����Ŀ������.", vbYesNo, "��ʾ") = vbYes Then
    Set Tmp_Rb = open_RecordSet("select * from BarDiff_Amt")
    Do Until Tmp_Rb.EOF
        G_Con.Execute "Update mmst381 set mtr_amt=" & Tmp_Rb!Total_Amt & " Where Bar_No='" & Tmp_Rb!Bar_No & "' and mtr_no='" & Tmp_Rb!mtr_no & "'"
        Tmp_Rb.MoveNext
    Loop
    
    tmp_Str = "Select Bar_No,Mtr_No,Total_Amt ,'' as Order_No,'" & G_User_Name & "B' as Upd_Name,'" & Get_SQLDATE & "' as Upd_Date From Diff_381Isnull "
    
    G_Con.Execute "Insert Into mmst381 (Bar_No,Mtr_No,Mtr_Amt ,Order_No,Upd_Name,Upd_Date) " & tmp_Str
End If




MsgBox "�������"
End Sub

Private Sub Command6_Click()
Dim tmp_Str As String
Dim Tmp_Str_1 As String

'Tmp_Str_1 = " SELECT  mmst502.Order_No, mmst502.Cust_Order_No, mmst502.Mtr_No,mmst502.Cust_Mtr_No, " & _
'                               " SUM(CASE WHEN Deliv_Type = '1' THEN 1 ELSE - 1 END *isnull(mmst502.Mtr_Amt,0)) AS Deliv_Amt " & _
'                        " FROM  mmst501 INNER JOIN mmst502 ON  mmst501.Deliv_No = mmst502.Deliv_No " & _
'                        " WHERE mmst501.status <> '0' " & _
'                        " GROUP BY  mmst502.Order_No, mmst502.Cust_Order_No, mmst502.Mtr_No,mmst502.Cust_Mtr_No "
'
'G_Con.Execute "Update mmst012 set Deliv_Amt=0"
'
''������
'Tmp_Str = " Update mmst012 " & _
'                 " Set mmst012.Deliv_Amt = a.Deliv_Amt" & _
'                 " FROM             mmst012,(" & Tmp_Str_1 & ") a " & _
'                " WHERE         mmst012.Order_No = a.Order_No AND mmst012.Cust_Order_No = a.Cust_Order_No AND mmst012.Mtr_No = a.Mtr_No  AND mmst012.Cust_Mtr_No = a.Cust_Mtr_No"
'
'G_Con.Execute Tmp_Str
G_Con.Execute "exec ts_Upd_012_Total"

MsgBox "�������"
End Sub

Private Sub Command7_Click()
Dim tmp_Str As String
tmp_Str = " SELECT  mmst401.Mo_No, mmst401.Plan_No, mmst401.Order_No,  mmst401.Mtr_No AS p_mtr, mmst401.Mtr_No,mmst401.Cust_Order_No , mmst401.mtr_amt " & _
          " FROM    mmst401 LEFT OUTER JOIN   Mmst401_Mtr ON mmst401.Mo_No = Mmst401_Mtr.Mo_No AND  mmst401.Plan_No = Mmst401_Mtr.Plan_No AND  mmst401.Order_No = Mmst401_Mtr.Order_No AND  mmst401.Cust_Order_No = Mmst401_Mtr.Cust_Order_No AND mmst401.mtr_no = Mmst401_Mtr.mtr_no " & _
          " WHERE   (Mmst401_Mtr.P_Mtr IS NULL)"
          
G_Con.Execute "insert into mmst401_mtr (Mo_No, Plan_No, Order_No,  p_mtr, Mtr_No,Cust_Order_No , mtr_amt) " & tmp_Str

'������
tmp_Str = " UPDATE    mmst401_mtr " & _
                    " Set Mmst401_Mtr.mtr_amt = mmst401.mtr_amt  " & _
          " FROM      mmst401 INNER JOIN  " & _
                    " Mmst401_Mtr ON mmst401.Mo_No = Mmst401_Mtr.Mo_No AND  mmst401.Plan_No = Mmst401_Mtr.Plan_No AND  mmst401.Order_No = Mmst401_Mtr.Order_No AND  mmst401.Cust_Order_No = Mmst401_Mtr.Cust_Order_No AND  mmst401.Mtr_No = Mmst401_Mtr.Mtr_No"
                           
G_Con.Execute tmp_Str
              
G_Con.Execute "update mmst401_mtr set mtr_amt= 0 where mtr_amt is null "
G_Con.Execute "update mmst401_mtr set loss_amt= 0 where loss_amt is null "
G_Con.Execute "update mmst401_mtr set ling_amt= 0 where ling_amt is null "
G_Con.Execute "update mmst401_mtr set tui_amt= 0 where tui_amt is null "
G_Con.Execute "update mmst401_mtr set bu_amt= 0 where bu_amt is null "
G_Con.Execute "update mmst401_mtr set bao_amt= 0 where bao_amt is null "
G_Con.Execute "update mmst401_mtr set out_amt= 0 where out_amt is null "
G_Con.Execute "update mmst401_mtr set chao_amt= 0 where chao_amt is null "
G_Con.Execute "update mmst401_mtr set out_p_amt= 0 where out_p_amt is null "

MsgBox "�������"
End Sub

Private Sub Command8_Click()
G_Con.Execute "update mmst615 set NoUse_Loss= 0 where NoUse_Loss is null "
G_Con.Execute "update mmst611 set NoUse_Loss= 0 where NoUse_Loss is null "


MsgBox "�������"
End Sub

Private Sub Command9_Click()
G_Con.Execute " EXEC ts_Upd_401_Mtr_Total"
End Sub

Private Sub Upd0122_Hid_Print_Click()
G_Con.Execute "update mmst0122 set hid_print='0'"
G_Con.Execute "UPDATE mmst0122 SET mmst0122.Hid_Print=mmst612.Hid_print " & _
                         "FROM  mmst0122 INNER JOIN mmst612 ON " & _
                         "mmst0122.Mtr_No=mmst612.Mtr_No AND mmst0122.Bom_No=mmst612.Bom_No " & _
                         " where mmst612.hid_print='1'"
MsgBox "�������"
End Sub
