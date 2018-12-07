VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form FrmList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '单线固定工具视窗
   Caption         =   "列表"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所属视窗中央
   Begin VB.CommandButton ComOK 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   60
      Picture         =   "FrmList.frx":0000
      Style           =   1  '图片外观
      TabIndex        =   2
      Tag             =   "&Cancel"
      Top             =   180
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   -15
      ScaleHeight     =   5505
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   -15
      Width           =   5655
      Begin VB.CommandButton ComCancel 
         Height          =   345
         Left            =   1335
         Picture         =   "FrmList.frx":15A2
         Style           =   1  '图片外观
         TabIndex        =   1
         Tag             =   "&Cancel"
         Top             =   180
         Width           =   1155
      End
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   4905
         Left            =   0
         TabIndex        =   3
         Top             =   630
         Width           =   5640
         _cx             =   9948
         _cy             =   8652
         _ConvInfo       =   -1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新细明体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   16711680
         BackColorSel    =   49152
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16711680
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   13
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称: 公用列表
'*编写日期: 2002/07/25
'*制作人员: 毛泽球
'*修改日期:
'*修改人员:
'***********************************************
Private W_Col_Item(15) As String

Private W_Cancel_Click As Boolean

Public G_Sql_Filter As String
Public Property Get Cancel_Click() As Boolean
Cancel_Click = W_Cancel_Click
End Property

Public Property Get Col_No1() As String
    Col_No1 = W_Col_Item(0)
End Property
Public Property Get Col_No2() As String
    Col_No2 = W_Col_Item(1)
End Property
Public Property Get Col_No3() As String
    Col_No3 = W_Col_Item(2)
End Property
Public Property Get Col_No4() As String
    Col_No4 = W_Col_Item(3)
End Property
Public Property Get Col_No5() As String
    Col_No5 = W_Col_Item(4)
End Property
Public Property Get Col_No6() As String
    Col_No6 = W_Col_Item(5)
End Property
Public Property Get Col_No7() As String
    Col_No7 = W_Col_Item(6)
End Property
Public Property Get Col_No8() As String
    Col_No8 = W_Col_Item(7)
End Property
Public Property Get Col_No9() As String
    Col_No9 = W_Col_Item(8)
End Property
Public Property Get Col_No10() As String
    Col_No10 = W_Col_Item(9)
End Property

Private Sub Form_Load()

    Set Me.Picture1.Picture = Erp_Deliv.Picture

    Dim W_Rs As New ADODB.Recordset
    
    W_Rs.Open G_Sql_Filter, G_Con
      
    Set Grid1.DataSource = W_Rs
    
        '设定表格宽度
    With Grid1
        .AutoResize = True
        For i = 1 To .Cols - 1
           .AutoSize (i)
        Next
        For i = 0 To .Rows - 1
            .RowHeight(i) = 350
        Next
    End With
    
End Sub

Private Sub Comok_Click()
    Dim temp As New ADODB.Recordset
    If Index = 0 Then
        Dim W_Row As Long
        Dim w_Col As Long
        W_Row = Grid1.Row
        w_Col = Grid1.Cols
        If W_Row > 0 Then
            With Grid1
                Dim i As Integer
                For i = 0 To w_Col - 1
                    W_Col_Item(i) = NullSetValue(.TextMatrix(W_Row, i), "")
                Next i
            End With
            W_Cancel_Click = False
        Else
            W_Cancel_Click = True
        End If
    ElseIf Index = 1 Then
        W_Cancel_Click = True
        W_Col_Item(0) = ""
    End If
    
        '设定表格宽度
  
    
    Set temp = Nothing
    
    Unload Me
End Sub

Private Sub ComCancel_Click()
    W_Cancel_Click = True
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
            Case vbKeyF5               'F5确定
                Call Comok_Click
            Case vbKeyEscape           'ESC取消
                Call ComCancel_Click
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    G_Sql_Filter = ""
End Sub

Private Sub Grid1_DblClick()
    Call Comok_Click
End Sub
