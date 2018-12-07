VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDptList 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "部門列表"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Tag             =   "Customers List"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5505
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   1
         Left            =   1350
         Picture         =   "FrmDptList.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Tag             =   "&Cancel"
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   90
         Picture         =   "FrmDptList.frx":15A2
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Tag             =   "&OK"
         Top             =   180
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   4725
         Left            =   60
         TabIndex        =   3
         Top             =   690
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   8334
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorFixed  =   -2147483628
         BackColorSel    =   -2147483624
         ForeColorSel    =   -2147483625
         BackColorBkg    =   -2147483628
         GridColor       =   8421504
         GridColorFixed  =   8421376
         FocusRect       =   0
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "FrmDptList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private w_dpt_id As String
Private w_Dpt_Name As String
Private w_Dpt_Right As String
Public g_Dpt_Filter As String
Dim W_CallForm As Form

Public Property Get CallForm() As Form
Set CallForm = W_CallForm
End Property
Public Property Set CallForm(f As Form)
Set W_CallForm = f
End Property

Public Property Get Dpt_Id() As String
    Dpt_Id = w_dpt_id
End Property
Public Property Get Dpt_Name() As String
    Dpt_Name = w_Dpt_Name
End Property

Public Property Get Dpt_Right() As String
    Dpt_Right = w_Dpt_Right
End Property

Private Sub Command1_Click(Index As Integer)
  Dim temp As New ADODB.Recordset
    If Index = 0 Then
        Dim W_Row As Long
        W_Row = Grid1.Row
        If W_Row > 0 Then
            With Grid1
                w_dpt_id = .TextMatrix(W_Row, 0)
                w_Dpt_Name = .TextMatrix(W_Row, 1)
                w_Dpt_Right = .TextMatrix(W_Row, 2)
                Set temp = Nothing
            End With
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               '確認
       Call Command1_Click(0)
    Case vbKeyEscape           '取消
       Call Command1_Click(1)
    End Select
End If
End Sub

Private Sub Form_Load()
    '加載列表
    Dim w_rs As New ADODB.Recordset
  
    'load圖片
    Set Me.Picture1.Picture = Erp_Purc.Picture
    
     w_rs.Open "SELECT Dpt_Id AS 部門代號 , " & _
               " Dpt_name AS 部門名稱,  " & _
               "Dpt_Right AS 部門職責 " & _
               "FROM mmst902 " & _
               "where " & g_Dpt_Filter & "  order BY Dpt_Id ", G_Con
 
    Set Grid1.DataSource = w_rs
    With Grid1
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 1500
    
    End With
End Sub

Private Sub Grid1_DblClick()
    Call Command1_Click(0)
End Sub
