VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form Frm10FList 
   BackColor       =   &H80000009&
   BorderStyle     =   5  'ñÝñû¡¡ÍÔ¼ñ¡¡¶Ø
   Caption         =   "Â¤¡¡µ¸Í°"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   11580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¡¡õ»¡¡èé
   Tag             =   "Material List"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   8220
      Top             =   3330
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   1440
      Picture         =   "Frm10FList.frx":0000
      Style           =   1  '¡¡¤¨¡¡â¹
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   180
      Picture         =   "Frm10FList.frx":15A2
      Style           =   1  '¡¡¤¨¡¡â¹
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   90
      Width           =   1155
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   6765
      Left            =   135
      TabIndex        =   2
      Top             =   570
      Width           =   11400
      _cx             =   20108
      _cy             =   11933
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Úë¡¡Â¤¼«"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   16711680
      BackColorSel    =   49152
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
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
      MergeCompare    =   1
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
Attribute VB_Name = "Frm10FList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'* Éø¡¡ÃÒ±¸: Â¤¡¡List
'* Éø¡¡÷ÕÙç: Ò«ç»¡¡¡¡
'* Í³¡¡¡¡Â¤: W_SELECT_Data ¡¡ãð¡¡¡¡SQLåÌÀú
'* µî ¡¡ ¡¡: W_List(1-10)  ¡¡¡¡»»SQLåÌÀú ¡¡°ôÆÍñÝ»»¡¡10¶å¡¡(¡¡¡¡)
'* ¡¡    î«: ¡¡¡¡áû¡¡ÉÆ Mtr_No,Mtr_Name,Mtr_Dim ñÝ¡¡»»:
'*           With FrmList
'*              .W_SELECT_Data="SELECT Mtr_No   as [¡¡Ì£²âÑñ]," & _
'*                                    "Mtr_Name as [¡¡Ì£ÃÒ±¸]," & _
'*                                    "Mtr_Dim  as [¡¡Ì£Ä¯¼£] " & _
'*                             "FROM mmst611 " & _
'*                             "WHERE Mtr_No Like '" & Trim(Mtr_No.Text) & "' "
'*              .Show 1
'*              If .List1<>"" Then
'*                  Mtr_No.Text  = .List1
'*                  Mtr_Name.Text= .List2
'*                  Mtr_Dim.Text = .List3
'*              End If
'*           End With
'* Îî¡¡¡¡¡¡: Rain
'* Îî¡¡¡¡¤Ö: 2003/05/10
'* µ³òÛ¡¡¡¡:
'* µ³òÛ¡¡¤Ö:
'* µ³òÛ¡¡Â¤:
'*************************************************************************************************

'Í°µÈ×«Ç»ÓµÖÏ²Ü½²
Private W_List0 As String
Private W_List1 As String
Private W_List2 As String
Private W_List3 As String
Private W_List4 As String
Private W_List5 As String
Private W_List6 As String
Private W_List7 As String
Private W_List8 As String
Private W_List9 As String
Private W_List10 As String
Private W_List11 As String
Private W_List12 As String
Private W_List13 As String

Public W_Select_Data As String

'íÏÊò¡¡¡¡µî¡¡¡¡
Public Property Get List0() As String 'Å¼Á§ÎîÑñ
List0 = W_List0
End Property

Public Property Get List1() As String '¡¡Ññ
List1 = W_List1
End Property

Public Property Get List2() As String 'µÚÌ£²âÑñ
List2 = W_List2
End Property

Public Property Get List3() As String   '¡¡¼£
List3 = W_List3
End Property

Public Property Get List4() As String   '¡¡¼£
List4 = W_List4
End Property

Public Property Get List5() As String '¤³ÃÒ
List5 = W_List5
End Property

Public Property Get List6() As String 'Ä¯¼£
List6 = W_List6
End Property

Public Property Get List7() As String 'µÈ¡¡ÃÒ±¸
List7 = W_List7
End Property

Public Property Get List8() As String '¡¡¼¿ÃÒ±¸
List8 = W_List8
End Property

Public Property Get List9() As String '¼¿Ññ
List9 = W_List9
End Property

Public Property Get List10() As String 'ÀÛ¹£ÃÒ±¸
List10 = W_List10
End Property

Public Property Get List11() As String 'Æô¹£ÃÒ±¸
List11 = W_List11
End Property

Public Property Get List12() As String 'µÈ¡¡
List12 = W_List12
End Property

Public Property Get List13() As String 'µÈ¡¡
List13 = W_List13
End Property

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
    Dim W_Row As Long
    W_Row = Grid1.Row
    If W_Row > 0 Then
        With Grid1
            W_List0 = .TextMatrix(W_Row, 0)
            W_List1 = .TextMatrix(W_Row, 1)
            W_List2 = .TextMatrix(W_Row, 2)
            W_List3 = .TextMatrix(W_Row, 3)
            W_List4 = .TextMatrix(W_Row, 4)
            W_List5 = .TextMatrix(W_Row, 5)
            W_List6 = .TextMatrix(W_Row, 6)
            W_List7 = .TextMatrix(W_Row, 7)
            W_List8 = .TextMatrix(W_Row, 8)
            W_List9 = .TextMatrix(W_Row, 9)
            W_List10 = .TextMatrix(W_Row, 10)
            W_List11 = .TextMatrix(W_Row, 11)
            W_List12 = .TextMatrix(W_Row, 12)
            W_List13 = .TextMatrix(W_Row, 13)
        End With
    End If
End If
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyReturn
       'SendKeys "{tab}"
    Case vbKeyF5               '¡¡¡¡
        Call Command1_Click(0)
    Case vbKeyEscape           '¡¡¡¡
        Call Command1_Click(1)
    End Select
End If
End Sub

Private Sub Form_Load()
'load¡¡¤¨
Set Me.Picture = G_MDIForm.Picture

W_List1 = ""
W_List2 = ""
W_List3 = ""
W_List4 = ""
W_List5 = ""
W_List6 = ""
W_List7 = ""
W_List8 = ""
W_List9 = ""
W_List10 = ""
W_List11 = ""
W_List12 = ""
W_List13 = ""

Me.KeyPreview = True
Call Select_date
End Sub

Private Sub Form_Resize()
On Error Resume Next
'Call ResizeListWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'¡¡¶Ø¼«¡¡¡¡¡¡¢Û¡¡Óç²Î¶Ø¼«Ç»Êó¡¡²Ü½²Ç»¡¡§îÅµ
'¡¡¡¡,¡¡ñû¡¡É­¶Ø¼«¡¡,íÏ¡¡íÑÓçíñÜíÇ»¡¡¡¡,°ÂÏéõúõÓ·ª¡¡¶åÊó¡¡²Ü½²¶­¡¡
W_Select_Data = ""
Set Grid1.DataSource = Nothing
End Sub

Private Sub Grid1_DblClick()
Call Command1_Click(0)
End Sub

Private Sub Timer1_Timer()
If Grid1.Rows = 2 Then
    Call Command1_Click(0)
End If
End Sub

Private Sub Select_date()
Dim W_Rs As New ADODB.Recordset
'On Error GoTo Load_Err
Set W_Rs = Nothing
W_Rs.CursorLocation = adUseClient
W_Rs.Open W_Select_Data, G_Con

Set Grid1.DataSource = W_Rs
'¡¡ÓçÍ°¼£×ñ½ö
With Grid1
    .AutoResize = True
    For I = 1 To .Cols - 1
       .AutoSize (I)
    Next
    For I = 0 To .Rows - 1
        .RowHeight(I) = 350
    Next
End With

If Grid1.Rows = 2 Then
   Timer1.Enabled = True
End If
    
Exit Sub
Load_Err:
End Sub
