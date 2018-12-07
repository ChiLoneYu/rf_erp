VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCustMtrSql 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "　　　薛(FrmCustMtrSql)"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrCustMtrSql.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Order Search"
   Begin VB.TextBox Cust_Mtr_No 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   330
      Left            =   1725
      MaxLength       =   23
      TabIndex        =   4
      Top             =   840
      Width           =   2085
   End
   Begin VB.CommandButton CmdOK 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   990
      Picture         =   "FrCustMtrSql.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&OK"
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton CmdCancel 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      Picture         =   "FrCustMtrSql.frx":15AE
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Cancel"
      Top             =   2400
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   330
      Left            =   1725
      TabIndex        =   0
      Top             =   1275
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   151191555
      UpDown          =   -1  'True
      CurrentDate     =   37217
      MaxDate         =   65745
      MinDate         =   32874
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   330
      Left            =   1725
      TabIndex        =   1
      Top             =   1710
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   151191555
      UpDown          =   -1  'True
      CurrentDate     =   37217
      MaxDate         =   65745
      MinDate         =   32874
   End
   Begin VB.TextBox Mtr_No 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   330
      Left            =   1725
      MaxLength       =   23
      TabIndex        =   5
      Top             =   405
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "　　　少:"
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   9
      Top             =   1770
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "获仇聆　:"
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   8
      Top             =   465
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "蒺　　少:"
      Height          =   195
      Index           =   5
      Left            =   570
      TabIndex        =   7
      Top             =   1335
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "　　　　:"
      Height          =   195
      Index           =   9
      Left            =   570
      TabIndex        =   6
      Top             =   900
      Width           =   825
   End
End
Attribute VB_Name = "FrmCustMtrSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_Cancel As Boolean '领领领氮领领"领领"
Dim W_CallForm As Form
Public W_Table_Name As String
Public W_Price_Date As String

Public Property Get CallForm() As Form
    Set CallForm = W_CallForm
End Property

Public Property Set CallForm(f As Form)
    Set W_CallForm = f
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               '
         If CmdOK.Enabled = True Then
             Call CmdOk_Click
         End If
    Case vbKeyEscape           '
         If CmdCancel.Enabled = True Then
             Call CmdCancel_Click
         End If
    End Select
End If
End Sub

Private Sub Form_Load()
date1.Value = Get_SQLDATE - 30
date2.Value = Get_SQLDATE
Set Me.Picture = G_MDIForm.Picture
W_Cancel = False
End Sub

Private Sub CmdOk_Click()
Dim W_SQL As String

'领谊愈胀领少
If Not IsNull(date1.Value) And Not IsNull(date2.Value) Then
    W_SQL = " AND " & W_Table_Name & "." & W_Price_Date & " BETWEEN '" & _
            Format(date1.Value, "yyyy-mm-dd") & "' AND '" & _
            Format(date2.Value, "yyyy-mm-dd") & "' "
ElseIf Not IsNull(date1.Value) And IsNull(date2.Value) Then
    W_SQL = " AND " & W_Table_Name & "." & W_Price_Date & " >='" & Format(date1.Value, "yyyy-mm-dd") & "' "
ElseIf IsNull(date1.Value) And Not IsNull(date2.Value) Then
    W_SQL = " AND " & W_Table_Name & "." & W_Price_Date & " <='" & Format(date2.Value, "yyyy-mm-dd") & "' "
End If

'获仇聆

W_SQL = " AND " & W_Table_Name & ".Mtr_No LIKE '" & Trim(mtr_no.Text) & "%'" & W_SQL


'
W_SQL = " AND " & W_Table_Name & ".Cust_Mtr_No LIKE '" & Trim(Cust_Mtr_No.Text) & "%'" & W_SQL


CallForm.W_SQL_Where = W_SQL

Unload Me
End Sub


Private Sub CmdCancel_Click()
CallForm.W_SQL_Where = ""
W_Cancel = True
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCustMtrSql = Nothing
End Sub
