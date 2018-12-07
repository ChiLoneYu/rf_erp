VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "AresButtonPro.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form mmss604 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "成品退货单(604)"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15165
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15165
   Tag             =   "Quotations"
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   2085
      ScaleHeight     =   9105
      ScaleWidth      =   13035
      TabIndex        =   21
      Top             =   0
      Width           =   13065
      Begin VB.CheckBox qc_status 
         BackColor       =   &H8000000E&
         Caption         =   "要检验"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   390
         TabIndex        =   39
         Top             =   300
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox remark 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   7380
         Width           =   11625
      End
      Begin MSComCtl2.DTPicker Inv_date 
         Height          =   345
         Left            =   10470
         TabIndex        =   3
         Top             =   870
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
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
         Format          =   140181505
         CurrentDate     =   37240
      End
      Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
         Height          =   5835
         Left            =   270
         TabIndex        =   14
         Top             =   1410
         Width           =   12495
         _cx             =   22040
         _cy             =   10292
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
         BackColorFixed  =   -2147483634
         ForeColorFixed  =   -2147483630
         BackColorSel    =   65280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483641
         GridColorFixed  =   -2147483641
         TreeColor       =   -2147483632
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
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
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
         AllowUserFreezing=   3
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin MSForms.ComboBox Qc_Man 
         Height          =   345
         Left            =   8700
         TabIndex        =   17
         Top             =   8580
         Width           =   1275
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2249;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line_Qc 
         X1              =   8790
         X2              =   10080
         Y1              =   8970
         Y2              =   8970
      End
      Begin VB.Label Lab_Qc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "品检:"
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
         Left            =   8100
         TabIndex        =   38
         Top             =   8730
         Width           =   435
      End
      Begin MSForms.TextBox form_man 
         Height          =   345
         Left            =   11250
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8580
         Width           =   1275
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2249;609"
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Bar_man 
         Height          =   345
         Left            =   5970
         TabIndex        =   18
         Top             =   8580
         Width           =   1275
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2249;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox mag_Man 
         Height          =   345
         Left            =   870
         TabIndex        =   20
         Top             =   8580
         Width           =   1275
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2249;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox check_man 
         Height          =   345
         Left            =   3300
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   8580
         Width           =   1275
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2249;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox inv_style 
         Height          =   345
         Left            =   1440
         TabIndex        =   2
         Top             =   870
         Width           =   1875
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3307;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBLabel 
         Height          =   225
         Left            =   330
         TabIndex        =   35
         Top             =   945
         Width           =   1050
         BackColor       =   -2147483639
         VariousPropertyBits=   8388627
         Caption         =   "入库类别:"
         Size            =   "1852;397"
         BorderColor     =   -2147483643
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox p_line_no 
         Height          =   345
         Left            =   6060
         TabIndex        =   1
         ToolTipText     =   "不能超过12个字符"
         Top             =   870
         Width           =   1875
         VariousPropertyBits=   679495707
         MaxLength       =   12
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3307;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Inv_no 
         Height          =   345
         Left            =   10470
         TabIndex        =   0
         ToolTipText     =   "不能超过12个字符"
         Top             =   450
         Width           =   2025
         VariousPropertyBits=   679495707
         MaxLength       =   12
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3572;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "新细明体"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "生产线:"
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
         Left            =   5160
         TabIndex        =   34
         Top             =   945
         Width           =   630
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   270
         X2              =   12750
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "仓库:"
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
         Left            =   5340
         TabIndex        =   29
         Top             =   8730
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "制表:"
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
         Left            =   10650
         TabIndex        =   28
         Top             =   8730
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "核准:"
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
         Left            =   285
         TabIndex        =   27
         Top             =   8730
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "审核:"
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
         Left            =   2685
         TabIndex        =   26
         Top             =   8730
         Width           =   435
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   5130
         X2              =   8400
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line3 
         X1              =   5130
         X2              =   8400
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
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
         Left            =   9840
         TabIndex        =   25
         Top             =   945
         Width           =   435
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   270
         X2              =   12750
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "备   注:"
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
         Left            =   300
         TabIndex        =   24
         Top             =   7410
         Width           =   810
      End
      Begin VB.Label lbCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "成品出库-退货单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   5040
         TabIndex        =   23
         Tag             =   "Quotations"
         Top             =   180
         Width           =   3465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9960
         TabIndex        =   22
         Top             =   540
         Width           =   315
      End
      Begin VB.Line Line5 
         X1              =   840
         X2              =   2160
         Y1              =   8970
         Y2              =   8970
      End
      Begin VB.Line Line6 
         X1              =   11340
         X2              =   12630
         Y1              =   8970
         Y2              =   8970
      End
      Begin VB.Line Line7 
         X1              =   6060
         X2              =   7350
         Y1              =   8970
         Y2              =   8970
      End
      Begin VB.Line Line8 
         X1              =   3390
         X2              =   4650
         Y1              =   8970
         Y2              =   8970
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5865
      ScaleWidth      =   2010
      TabIndex        =   36
      Top             =   0
      Width           =   2040
      Begin ARESBUTTONLib.AresButton cmd_delete 
         Height          =   420
         Left            =   390
         TabIndex        =   12
         Top             =   4578
         Width           =   495
         _Version        =   327687
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   95969376
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_preview 
         Height          =   420
         Left            =   390
         TabIndex        =   6
         Top             =   1302
         Width           =   495
         _Version        =   327687
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   90747236
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_find 
         Height          =   420
         Left            =   390
         TabIndex        =   4
         Top             =   210
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\billy\xsh_erp\Picture\Norm\Find_Norm.bmp"
         PictureOverURL  =   "Y:\c_sys\billy\xsh_erp\Picture\Over\Find_over.bmp"
         PictureDisableURL=   "Y:\c_sys\billy\xsh_erp\Picture\Dis\Find_Dis.bmp"
         PictureBaseURL  =   "Y:\c_sys\billy\xsh_erp\Picture\Norm\Find_Norm.bmp"
         ToolTipString   =   "查询符合条件的表单"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_print 
         Height          =   420
         Left            =   390
         TabIndex        =   5
         Top             =   756
         Width           =   495
         _Version        =   327687
         ToolTipString   =   "列印报表"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   95977692
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_save 
         Height          =   420
         Left            =   390
         TabIndex        =   7
         Top             =   1848
         Width           =   495
         _Version        =   327687
         ToolTipString   =   "将表单存成文件"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   90853636
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_ok 
         Height          =   420
         Left            =   390
         TabIndex        =   8
         Top             =   2394
         Width           =   495
         _Version        =   327687
         ToolTipString   =   "确认存盘"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   96079780
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_cancel 
         Height          =   420
         Left            =   390
         TabIndex        =   9
         Top             =   2940
         Width           =   495
         _Version        =   327687
         ToolTipString   =   "放弃存盘"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   96113780
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_add 
         Height          =   420
         Left            =   390
         TabIndex        =   10
         Top             =   3486
         Width           =   495
         _Version        =   327687
         ToolTipString   =   "增加一张单据"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   90958684
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_edit 
         Height          =   420
         Left            =   390
         TabIndex        =   11
         Top             =   4032
         Width           =   495
         _Version        =   327687
         ToolTipString   =   "修改该单据"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   96216844
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_quit 
         Height          =   360
         Left            =   390
         TabIndex        =   13
         Top             =   5130
         Width           =   1125
         _Version        =   327687
         ToolTipString   =   "退出该程式"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   91011208
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin MSForms.Label lab_focus 
         Height          =   465
         Left            =   390
         TabIndex        =   37
         Top             =   150
         Width           =   975
         BackColor       =   -2147483643
         Size            =   "1720;820"
         FontName        =   "新细明体"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9255
      Left            =   2010
      TabIndex        =   30
      Top             =   -90
      Width           =   90
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2670
      Top             =   6810
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Frame Frame3 
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   3315
      Left            =   0
      TabIndex        =   31
      Top             =   5820
      Width           =   2070
      Begin ARESBUTTONLib.AresButton AresButton1 
         Height          =   1050
         Left            =   390
         TabIndex        =   32
         Top             =   1110
         Width           =   1050
         _Version        =   327687
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   217092436
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin VB.Label Help_txt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   150
         Width           =   1755
      End
   End
End
Attribute VB_Name = "mmss604"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '***********************************************
'*程序名称: 成品入/出库单 (mmss604)
'*编写日期:
'*制作人员:
'*修改日期:
'*修改人员:
'***********************************************
'定义记录集与命令对象
Dim Comm_533 As ADODB.Command

'指示当前的Inv_no
Dim W_Curr_InvNo As String

'存储生产线代号
Dim W_Line_No As String

'存放状态
Dim W_Status As Boolean

'存放TDBGRID1 的旧字符
Dim W_Old_Str As String

'日期
Dim W_inv_date As Date

Const P_INV_TYPE  As String = 2
'TDBGrid相关
Dim Gridc_604(127) As Grid_Data '存放 Grid 属性值
Dim RoW_Height As Double        'Grid 高度变量

'定义按钮变量
Dim C_Add As Boolean
Dim C_Edit As Boolean
Dim C_Delete As Boolean
Dim C_Print As Boolean
Dim C_View As Boolean
Dim C_Save As Boolean

'权限变量
Dim C_Off_Add As Boolean
Dim C_Off_Edit As Boolean
Dim C_Off_Delete As Boolean
Dim C_Off_Print As Boolean
Dim C_Off_View As Boolean
Dim C_Off_Save As Boolean

Public Sub Form_Activate()
    '当窗口激活时,刷新TDBGrid
    Call GetVSGridSetting("mmss604", "TDBGrid1", Gridc_604, g_CON_IniFile4)
    RoW_Height = Gridc_604(0).Grid_RowHeight
    Call readactive
    Call RefreshGrid
End Sub

Private Sub Form_Load()
    Call load_picture
    '表单接收键值优先
    Me.KeyPreview = True

    '将MDI子窗口置中
    Call CenterWindow(Me, G_MDIForm)
 
    'com_533 产生的单条记录集显示单据的表头内容,它会被反复执行.
    Set Comm_533 = New ADODB.Command
    With Comm_533
        .CommandType = adCmdText
        .CommandText = "SELECT mmst533.*,p_line_name " & _
                        "FROM mmst533,mmst811 " & _
                        "WHERE  mmst533.p_line_no*=mmst811.p_line_no " & _
                              " AND Inv_type='" & P_INV_TYPE & "'   " & _
                              " AND mmst533.inv_No=?"
                   
        .ActiveConnection = G_Con
        .Prepared = True '因为它会多次执行,将它预编绎.
    End With

    '加载 COMBOX 数据
    Call load_combox

    '将按钮变量赋初值
    C_Add = False
    C_Edit = False
    C_Delete = False

    C_Off_Add = False
    C_Off_Edit = False
    C_Off_Delete = False
    C_Off_Print = False
    C_Off_Save = False
    C_Off_View = False

    'MDI子窗口按钮权限设订
    C_Off_Add = lopcheck("A", "604", G_User_ID)
    C_Off_Edit = lopcheck("U", "604", G_User_ID)
    C_Off_Delete = lopcheck("D", "604", G_User_ID)
    C_Off_View = lopcheck("V", "604", G_User_ID)
    C_Off_Print = lopcheck("P", "604", G_User_ID)
    C_Off_Save = lopcheck("S", "604", G_User_ID)

    '调用Inv_no_Click
    Call inv_no_Click

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 34
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Form_KeyDown可照搬
    If KeyCode = vbKeyReturn Then
        If TypeOf Me.ActiveControl Is ComboBox Then
            If ActiveControl.MultiLine = False Then
                SendKeys "{TAB}"
                KeyCode = 0
            End If
        Else
            SendKeys "{TAB}"
            KeyCode = 0
        End If
        Exit Sub
    End If

    If KeyCode = vbKeyM And Shift = 1 Then
        '如果用户手工改动了单据号
         If LCase(Me.ActiveControl.Name) = "remark" Then
             Call inv_no_LostFocus
         End If
        
        '如果没有单据
        If Trim(inv_no.Text) = "" Then
            Exit Sub
        End If
        
        '检查单据状态
        If check_ok = False Then
            Exit Sub
        End If
        '非新增修改状态按下 SHIFT + M 弹出明细菜单
        If Not (C_Add Or C_Edit Or c_delet) Then
            Call TDBGrid1_MouseUp(2, 0, TDBGrid1.Left + 200, TDBGrid1.Top + 50)
        Else
             Exit Sub
        End If
       
    End If

    If Shift = 0 Then
        Select Case KeyCode
        Case vbKeyF2               '新增
             If cmd_add.Enabled = True Then
                 Call vcontrol("A")
                 KeyCode = 0
             End If
        Case vbKeyF3               '编辑
            '如果用户手工改动了单据号
            If LCase(Me.ActiveControl.Name) = "Inv_no" Then
                Call inv_no_LostFocus
            End If
         
            If cmd_edit.Enabled = True Then
                 Call vcontrol("U")
                 KeyCode = 0
            End If
        Case vbKeyF4               '删除
            '如果用户手工改动了单据号
            If LCase(Me.ActiveControl.Name) = "Inv_no" Then
                Call inv_no_LostFocus
            End If
             
            If cmd_delete.Enabled = True Then
                 Call vcontrol("D")
                KeyCode = 0
            End If
        Case vbKeyF5               '确认
             If Cmd_ok.Enabled = True Then
                 Call vcontrol("Y")
                 KeyCode = 0
             End If
        Case vbKeyF6               '退出
             If cmd_quit.Enabled = True Then
                 Call vcontrol("Q")
                 KeyCode = 0
             End If
             
        Case vbKeyF7               '查询
             If cmd_find.Enabled = True Then
                 Call vcontrol("F")
                 KeyCode = 0
             End If
        Case vbKeyF8               '列印
            '如果用户手工改动了单据号
            If LCase(Me.ActiveControl.Name) = "Inv_no" Then
                Call inv_no_LostFocus
            End If
             
            If cmd_print.Enabled = True Then
                 Call vcontrol("P")
                 KeyCode = 0
            End If
        Case vbKeyF9               '预览
            '如果用户手工改动了单据号
            If LCase(Me.ActiveControl.Name) = "Inv_no" Then
                Call inv_no_LostFocus
            End If
    
            If cmd_preview.Enabled = True Then
                 Call vcontrol("V")
                 KeyCode = 0
            End If
        Case vbKeyF12              '存储
            '如果用户手工改动了单据号
            If LCase(Me.ActiveControl.Name) = "Inv_no" Then
                Call inv_no_LostFocus
            End If
             
            If cmd_save.Enabled = True Then
                 Call vcontrol("S")
                 KeyCode = 0
            End If
        Case vbKeyEscape           '取消
             If cmd_cancel.Enabled = True Then
                 Call vcontrol("N")
                 KeyCode = 0
             End If
        End Select
    End If
End Sub

Private Sub readshow()
    '当新增时
    If C_Add = True Then
        Inv_Date.Value = Get_SQLDATE
        Remark.Text = ""
        inv_no.Text = Creat_No
        P_Line_No.Text = ""
        Inv_Style.Text = ""
        
        form_man.Text = Trim(G_User_Name)
        check_man.Text = ""
        Qc_Man.Text = ""
        bar_man.Text = ""
        mag_man.Text = ""
        
        '刷新表格
        Call RefreshGrid
    End If

    '设定各按键的 Enabled 属性
    If C_Add Or C_Edit Or C_Delete Then
        Cmd_ok.Enabled = True
        cmd_cancel.Enabled = True
        
        cmd_add.Enabled = False
        cmd_edit.Enabled = False
        cmd_delete.Enabled = False
        cmd_print.Enabled = False
        cmd_save.Enabled = False
        cmd_preview.Enabled = False
        cmd_find.Enabled = False
    Else
        Cmd_ok.Enabled = False
        cmd_cancel.Enabled = False
        
        cmd_find.Enabled = True
        cmd_add.Enabled = True
        If inv_no.ListCount <= 0 Or inv_no.Text = "" Then
           cmd_edit.Enabled = False
           cmd_delete.Enabled = False
           cmd_print.Enabled = False
           cmd_save.Enabled = False
           cmd_preview.Enabled = False
        Else
           cmd_edit.Enabled = True
           cmd_delete.Enabled = True
           cmd_print.Enabled = True
           cmd_save.Enabled = True
           cmd_preview.Enabled = True
           
           '当已审核时不能修改
           If W_Status = False Then
               cmd_edit.Enabled = False
               cmd_delete.Enabled = False
           End If
        End If
        
        If Adodc1.Recordset.RecordCount < 1 Then
            cmd_print.Enabled = False
            cmd_preview.Enabled = False
        End If

    End If
    
    '通过权限设定按键的 Enabled 属性
    If C_Off_Add = True Then
        cmd_add.Enabled = False
    End If
    
    If C_Off_Edit = True Then
        cmd_edit.Enabled = False
    End If
    
    If C_Off_Delete = True Then
        cmd_delete.Enabled = False
    End If
    
    If C_Off_Print = True Then
        cmd_print.Enabled = False
    End If

    If C_Off_Save = True Then
        cmd_save.Enabled = False
    End If
    
    If C_Off_View = True Then
        cmd_preview.Enabled = False
    End If
    
'    p_line_no.Locked = Not (C_Add Or C_Edit)
    Inv_Style.Locked = Not (C_Add Or C_Edit)
    bar_man.Locked = Not (C_Add Or C_Edit)
    Qc_Man.Locked = Not (C_Add Or C_Edit)
    
    check_man.Locked = True
    mag_man.Locked = True
    form_man.Locked = Not (C_Add Or C_Edit)
    
    If C_Edit Or C_Delete Then
        inv_no.Locked = True
       
    Else
        inv_no.Locked = False
        
    End If
    
    '当新增时qc项目可使用
    If C_Add Then
        qc_status.Enabled = True
    Else
        qc_status.Enabled = False
    End If
    
    '当有明晰内容後不可修改qc选择项目
    If C_Edit Then
        If Adodc1.Recordset.EOF = False Then
            qc_status.Enabled = False
        Else
            qc_status.Enabled = True
        End If
    End If
    
    
    
End Sub

Private Sub readactive()
    Set TDBGrid1.DataSource = Adodc1
    '存储TDBGRID 的属性
    Call SetVSGridSetting(TDBGrid1, Gridc_604)
    
    '刷新全部 ROW 的高度 包括 HEADER
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.RowHeight(i - 1) = RoW_Height
        If i < TDBGrid1.Rows Then
            TDBGrid1.TextMatrix(i, 0) = i
        End If
    Next i
    TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
End Sub

'***********************************************************
'响应弹出菜单项单击的过程
Public Sub menu_add_Click()
    '新增明细先加锁
    If LockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'") Then
        If check_ok = False Then
            Call UnLockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'")
            Exit Sub
        End If
        With Frm604Mx
            Set .CallForm = Me
            .inv_no = Trim(inv_no.Text)
            .UpdateMode = 0 'UpdateMode=0表示新增
            .Show vbModal
        End With
        '新增完毕解锁
        Call UnLockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'")
    
        TDBGrid1.SetFocus
        TDBGrid1.Col = 1
        If TDBGrid1.Rows > 1 Then
            TDBGrid1.Row = 1
        End If
    End If
End Sub

Public Sub menu_edit_Click()
    '修改前加锁
    If LockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'") Then
        If check_ok() = False Then
            Call UnLockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'")
            Exit Sub
        End If
    
        c_row = TDBGrid1.Row
        c_col = TDBGrid1.Col

        With Frm604Mx
            .UpdateMode = 1
             
             Set .CallForm = Me
             .inv_no = Trim(inv_no.Text)
            .order_no = Adodc1.Recordset!order_no
            .mtr_amt.Text = Adodc1.Recordset!mtr_amt
            .Note.Text = NullVal(Adodc1.Recordset!Note, "")
            Call .Order_No_LostFocus
            .Show vbModal
        End With
        Call UnLockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'")
        TDBGrid1.Row = c_row
        TDBGrid1.Col = c_col
    End If
End Sub

Public Sub menu_delete_Click()
    '修改前加锁
    If LockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'") Then
        '删除明细资料
        If MsgBox(g_CON_CDelete, vbYesNo + vbDefaultButton2 + vbInformation, g_CON_CTitle) = vbNo Then
            Call UnLockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'")
            Exit Sub
        End If
    
        '判断当前单据是否已被审核
        If check_ok() = False Then
            Call UnLockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'")
            Exit Sub
        End If
    
        
        '删除明细资料
        G_Con.Execute "DELETE FROM mmst534 WHERE List_No =" & Adodc1.Recordset!list_no
        Call UnLockRecord("mmst533", "Inv_no = '" & Trim(inv_no.Text) & "'")
        Call RefreshGrid
        
        TDBGrid1.SetFocus
        TDBGrid1.Col = 1
        If TDBGrid1.Rows > 1 Then
            TDBGrid1.Row = 1
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '当在新增或修改时 , 提示是否退出
    If C_Add Or C_Edit Or C_Delete Then
        '当有数据改动时.询问是否要退出系统
        If MsgBox(g_CON_CQuit, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
            Cancel = 1
        Else
           ' 当有修改或删除时未解锁时 , 解除锁定
            If C_Edit Or C_Delete Then
                Call UnLockRecord("mmst533", "Inv_no='" & inv_no.Text & "'")
            End If
            Cancel = 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '退出时，存储 TDBGrid 属性
    Call SaveGridSetting("mmss604", "TDBGrid1", Gridc_604, g_CON_IniFile4)
    
    Set comm_534 = Nothing
    
    Set TDBGrid1.DataSource = Nothing
    Set mmss604 = Nothing

End Sub

'各按钮click事件
'************************************************************
'按确认
Sub Cmd_ok_MouseClick()
    Call vcontrol("Y")
End Sub

'按取消
Sub Cmd_cancel_MouseClick()
    Call vcontrol("N")
End Sub

'按新增
Sub cmd_add_MouseClick()
    Call vcontrol("A")
End Sub

'按修改
Sub cmd_edit_MouseClick()
    Call vcontrol("U")
End Sub

'按删除
Sub cmd_delete_MouseClick()
    Call vcontrol("D")
End Sub

'按预览
Private Sub Cmd_print_MouseClick()
    Call vcontrol("P")
End Sub

'按打印
Private Sub Cmd_previeW_MouseClick()
    Call vcontrol("V")
End Sub

'按存档
Private Sub Cmd_save_MouseClick()
    Call vcontrol("S")
End Sub

'按退出
Sub Cmd_quit_MouseClick()
    Call vcontrol("Q")
End Sub
'按查询
Private Sub Cmd_find_MouseClick()
    Call vcontrol("F")
End Sub

'VCONTROL 函数
Private Sub vcontrol(ByVal p_choice As String)
    Dim W_add As Boolean
    
    Select Case p_choice
        Case "Y"            '确定
            If check_ok() Then
                Call upd_data
                TDBGrid1.Enabled = True
            End If
        
        Case "N"            '取消
            '如果新增或修改时取消动作,则要解锁
            If C_Edit Or C_Delete Then
               Call UnLockRecord("mmst533", "Inv_no='" & Trim(inv_no.Text) & "'")
            End If
            
            '当新增时取消动作
            If C_Add = True Then
               W_add = True
            End If
            C_Add = False
            C_Edit = False
            C_Delete = False
            TDBGrid1.Enabled = True
            
            inv_no.Text = W_Curr_InvNo
            
            '调用Inv_no_Click
            Call inv_no_Click
           
        Case "A"            ' 增加
            C_Add = True
            Call readshow
            TDBGrid1.Enabled = False
            inv_no.SetFocus
        
        Case "U"                    '修改
            If LockRecord("mmst533", "Inv_no='" & Trim(inv_no.Text) & "'") Then
                 '检查单据状态
                If check_ok() = False Then
                    Call UnLockRecord("mmst533", "Inv_no='" & inv_no.Text & "'")
                    Exit Sub
                End If
                C_Edit = True
                TDBGrid1.Enabled = False
                Call readshow
            End If
            Inv_Style.SetFocus
            
        Case "D"                 '删除
            '加锁记录
            If LockRecord("mmst533", "Inv_no='" & Trim(inv_no.Text) & "'") = True Then
                '删除当前记录
                If MsgBox(g_CON_CDelete, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
                    Call UnLockRecord("mmst533", "Inv_no='" & inv_no.Text & "'")
                    Exit Sub
                End If
                '判断是否可以删除
                C_Delete = True
                If check_ok = False Then
                    Call UnLockRecord("mmst533", "Inv_no='" & inv_no.Text & "'")
                    Exit Sub
                End If
                
                '错误处理
                err.Clear
                On Error GoTo Del_Err
                '释放质检单号
    
                '事务处理
                G_Con.BeginTrans
                G_Con.Execute "DELETE FROM mmst534 WHERE Inv_no='" & Trim(inv_no.Text) & "'"
                G_Con.Execute "DELETE FROM mmst533 WHERE Inv_no='" & Trim(inv_no.Text) & "'"
                G_Con.CommitTrans
                C_Delete = False
                On Error GoTo 0
                Dim w_index As Integer
                
                '找到对应的 INDEX
                For w_index = 0 To inv_no.ListCount - 1
                     If inv_no.List(w_index) = Trim(inv_no.Text) Then
                        inv_no.RemoveItem (w_index)
                        Exit For
                     End If
                Next w_index
                
                C_Delete = False
                
                If inv_no.ListCount > 0 Then
                    If inv_no.ListCount < w_index + 1 Then
                        w_index = w_index - 1
                    End If
                Else
                    w_index = -1
                End If
                
                If w_index <> -1 Then
                    inv_no.ListIndex = w_index
                Else
                    inv_no.Text = ""
                    Call inv_no_Click
                End If
                Exit Sub
Del_Err:
                On Error Resume Next
                G_Con.RollbackTrans
                MsgBox "删除时出现错误!", 64, g_CON_CTitle
                '当出错时记录未解锁时,解除锁定
                If CheckStatus("Inv_no", Trim(inv_no.Text), "mmst533", "status") = True Then
                    Call UnLockRecord("mmst533", "Inv_no='" & Trim(inv_no.Text) & "'")
                End If
               G_Con.RollbackTrans
    
           End If
        Case "P"    '打印
            C_Print = True
            If check_ok = True Then
                Call sele_data
            End If
       
       Case "V"     '预览
            C_View = True
            If check_ok = True Then
                Call sele_data
            End If
        
        Case "S"
            C_Save = True
            If check_ok = True Then
                Call sele_data
            End If
        
        Case "Q"    '退出
            Unload Me
        
        Case "F"   '查询
            With FrmpoInvSh
                .DefTable = "mmst533"
                .DefField = "inv_no"
                .DefInvDate = "Inv_date"
                .DefInvType = "Inv_type='" & P_INV_TYPE & "'  "
                .Caption = "入库单查询"
            
            Set .CallCoNtrol = inv_no
                .cb_check.AddItem "未审核"
                .cb_check.AddItem "审核"
                .cb_check.AddItem "全部"
                .cb_check.ListIndex = 0
                .G_File = "1"
                .Show vbModal
            
            If .ClickCaNcel = False Then
                If inv_no.ListCount > 0 Then
                    inv_no.ListIndex = 0
                Else
                    inv_no.Text = ""
                    Call inv_no_Click
                End If
            End If
        End With
    End Select
    
End Sub

Private Function check_ok() As Boolean
    Dim W_Rs As New ADODB.Recordset
    Dim W_Check As String

    '当在网络情况更新数据时,先判断单据是否已审核(主档)或删除
    '当状态为'2'时,只是不能异动单据,其它可以,如打印,存档等
    W_Check = CheckStatus("Inv_no", Trim(inv_no.Text), "mmst533", "status")
    If W_Check = "2" Or W_Check = "1" Then
        If Not (C_Add Or C_Save Or C_View) Then
            MsgBox "此单据已审核或结案!", 64, g_CON_CTitle
            If C_Add Or C_Edit Or C_Delete Then
                C_Add = False
                C_Edit = False
                C_Delete = False
            End If
            check_ok = False
            Exit Function
        End If
    ElseIf W_Check = "9" Then
        If C_Edit Or (C_Add = False And C_Edit = False And C_Delete = False) Then
            MsgBox "当前单据已被其它用户删除,不能操作!", 64, g_CON_CTitle
            C_Edit = False
            check_ok = False
            Exit Function
        End If
    End If
    
    '当打印或预览或存档时判断是否有明细资料
    If C_Print Or C_Save Or C_View Then
        If TDBGrid1.Rows <= 0 Then
            MsgBox "此单据没有明细资料,请录入其明细!", vbInformation, g_CON_CTitle
            C_View = False
            C_Print = False
            C_Save = False
            check_ok = False
            Exit Function
        End If
        check_ok = True
        Exit Function
    End If
    
    '新增或修改时检查
    If C_Add = True Then
        If Len(inv_no.Text) > 12 Then
            MsgBox "单据单号不能多於12个字符!", vbInformation, g_CON_CTitle
            check_ok = False
            inv_no.SetFocus
            Exit Function
        End If
        If Trim(inv_no.Text) = "" Then
            MsgBox "必须输入单据单号.", vbExclamation, g_CON_CTitle
            inv_no.SetFocus
            Exit Function
        Else
            W_Rs.CursorLocation = adUseClient
            W_Rs.Open "SELECT Inv_no FROM mmst533 WHERE Inv_no='" & inv_no.Text & "'", G_Con
            If W_Rs.EOF = False Then
                MsgBox "单据单号重复.", vbExclamation, g_CON_CTitle
                inv_no.SetFocus
                Exit Function
            End If
            W_Rs.Close
        End If
    End If
    
    '新增或修改主档时
    If P_Line_No.Text = "" Then
        MsgBox "请选择生产线.", vbExclamation, g_CON_CTitle
        P_Line_No.SetFocus
        check_ok = False
        Exit Function
    Else
        W_Rs.Open "SELECT p_line_no FROM mmst811 WHERE p_line_name='" & P_Line_No.Text & "'", G_Con
        If W_Rs.EOF = True Then
            MsgBox "无此生产线.", vbExclamation, g_CON_CTitle
            P_Line_No.SetFocus
            check_ok = False
            Exit Function
        Else
            W_Line_No = W_Rs!P_Line_No
        End If
        W_Rs.Close
    End If
    check_ok = True
    
    If Inv_Style.Text = "" Then
        MsgBox "请选择入库类型.", vbExclamation, g_CON_CTitle
        Inv_Style.SetFocus
        check_ok = False
        Exit Function
    Else
        If Inv_Style.Text <> "正常入库" And Inv_Style.Text <> "重工入库" Then
            MsgBox "请选择入库类型.", vbExclamation, g_CON_CTitle
            Inv_Style.ListIndex = 0
            Inv_Style.SetFocus
            
            check_ok = False
            Exit Function
        End If
    End If
    
 check_ok = True
End Function

Private Sub upd_data()
    Dim St_533 As New ADODB.Recordset
    With St_533
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .ActiveConnection = G_Con
        .Open "SELECT * FROM mmst533 WHERE Inv_no='" & inv_no.Text & "'", , , , adCmdText
    End With
    
    'upd_data将不再包含删除的过程
    If C_Add = True Then
        With St_533
            .AddNew
            !inv_no = Trim(inv_no.Text)
            !Inv_Date = Inv_Date.Value
            !Inv_type = P_INV_TYPE
            !P_Line_No = W_Line_No
            !Inv_Style = Inv_Style.Text
            !Remark = Remark.Text
            !status = "0"
            !mag_man = mag_man.Text
            !form_man = Trim(form_man.Text)
            !check_man = Trim(check_man.Text)
            !bar_man = bar_man.Text
            !Qc_Man = Qc_Man.Text
            !upd_date = Get_SQLDATE
            !upd_name = Trim(G_User_Name)
            
            !qc_status = IIf(qc_status.Value = 1, "1", "0")
            
            !lock = "No"
            .Update
        End With
        
        '刷新数据
        C_Add = False
        '刷新成品ComboBox
        inv_no.AddItem inv_no.Text
        '用於自动执行一次下拉动作
        inv_no.ListIndex = inv_no.ListCount - 1
    Else
        With St_533
            !Inv_Date = Inv_Date.Value
            !Inv_type = P_INV_TYPE
            !P_Line_No = W_Line_No
            !Inv_Style = Inv_Style.Text
            !Remark = Remark.Text
            !status = "0"
            
            !mag_man = mag_man.Text
            !form_man = Trim(form_man.Text)
            !check_man = Trim(check_man.Text)
            !bar_man = bar_man.Text
            !Qc_Man = Qc_Man.Text
            !upd_date = Get_SQLDATE
            !upd_name = Trim(Trim(G_User_Name))
            
            !qc_status = IIf(qc_status.Value = 1, "1", "0")
            
            !lock = "No"
            .Update
        End With
        C_Edit = False
    End If
    
    Call inv_no_Click
End Sub

Private Sub Inv_Date_DropDown()
    If Not (C_Add Or C_Edit) Then
        SendKeys "{ESCAPE}"
        Exit Sub
    End If
End Sub



Private Sub TDBGrid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        If NewRow >= 0 Then
            TDBGrid1.TextMatrix(OldRow, 0) = W_Old_Str
            W_Old_Str = TDBGrid1.TextMatrix(NewRow, 0)
            TDBGrid1.TextMatrix(NewRow, 0) = "★"
            TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
        End If
      
        '当点击TDBGRID1 cell 时,移动 ADODC1.Recordset 指针
        If Adodc1.Recordset.EOF = False Then
            Adodc1.Recordset.MoveFirst
            Adodc1.Recordset.Move TDBGrid1.Row - 1
'            TDBGrid1.FocusRect = flexFocusRaised
        End If
    End If
    TDBGrid1.TextMatrix(0, 0) = " No"
    TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

End Sub

Private Sub TDBGrid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col > 0 Then
        If ColIndex > Gridc_604(0).Grid_Columns Then
            Cancel = 1
        Else
            If UCase(Mid(Gridc_604(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_604(Col - 1).Grid_Visible = "" Then
                Cancel = 1
            Else
                Gridc_604(Col - 1).Grid_Width = TDBGrid1.ColWidth(Col)
            End If
        End If
    End If

    '移动ROW改变高度
    If Row >= 0 Then
        W_cur_row = TDBGrid1.Row
        RoW_Height = TDBGrid1.RowHeight(Row)
        Gridc_604(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
        For i = 1 To TDBGrid1.Rows
            TDBGrid1.RowHeight(i - 1) = RoW_Height
        Next i
        TDBGrid1.Row = W_cur_row
    End If
End Sub

'弹出菜单,新增/修改或删除从档资料
Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '如果不是左键
    If Button <> 2 Then
        Exit Sub
    End If
    '如果没有Inv_no
    If Trim(inv_no.Text) = "" Then
        Exit Sub
    End If
    '检查单据状态
    If check_ok() = False Then
        Exit Sub
    End If
    '这三个菜单项是整个系统共享的,应在此确保正确设置其enabled
    G_MDIForm.menu_add.Enabled = IIf(C_Off_Add, False, True)
    G_MDIForm.menu_delete.Enabled = IIf(C_Off_Delete, False, Adodc1.Recordset.EOF = False)
    G_MDIForm.menu_edit.Enabled = IIf(C_Off_Edit, False, Adodc1.Recordset.EOF = False)
    PopupMenu G_MDIForm.menu_modify
    '菜单复位
    G_MDIForm.menu_add.Enabled = True
    G_MDIForm.menu_edit.Enabled = True
    G_MDIForm.menu_delete.Enabled = True
End Sub

Private Sub inv_no_LostFocus()
    If Not (C_Add Or C_Edit) Then
        If W_Curr_InvNo <> Trim(inv_no.Text) Then
            Call inv_no_Click
            If inv_no.Text <> "" Then
                Dim i As Long
                For i = 0 To inv_no.ListCount - 1
                    If UCase(inv_no.List(i)) = UCase(inv_no.Text) Then
                        i = -10
                        Exit For
                    End If
                Next
                If i <> -10 Then
                    inv_no.AddItem inv_no.Text
                    Call readshow
                End If
            End If
        End If
    End If
End Sub

Private Sub TDBGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    '鼠标点在HEADER上
    If X > TDBGrid1.Left And Y < RoW_Height Then
       
        '存储 TDBGrid 属性
        Call SaveVSGridSetting("mmss604", "TDBGrid1", Gridc_604, g_CON_IniFile4)
        
        '调用 TDBGrid 属性设定
        With mmss_set
            Set .Parent_form = mmss604
            .Get_FormName = "mmss604"
            .Get_GridName = "TDBGrid1"
            .Gridc_File = g_CON_IniFile4
            .Show vbModal
        End With
        If TDBGrid1.Rows > 1 Then
            TDBGrid1.Col = 1
            TDBGrid1.Row = 1
        End If
    End If
End Sub

Private Sub TDBGrid1_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '不许更改第0行COl的宽度
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Public Sub inv_no_Click()
    Dim W_533 As New ADODB.Recordset  '主档
    If Not (C_Add Or C_Edit) Then
        W_Curr_InvNo = Trim(inv_no.Text)
        Comm_533.Parameters(0).Value = W_Curr_InvNo
        '重新执行原来的sql语句
        
        Set W_533 = Comm_533.Execute
    
        If W_533.EOF = False Then
            inv_no.Text = W_533!inv_no
            Inv_Date.Value = W_533!Inv_Date
            P_Line_No.Text = NullSetValue(W_533!P_Line_Name, "")
            Inv_Style.Text = W_533!Inv_Style
            
            Remark.Text = NullSetValue(W_533!Remark, "")
            
            form_man.Text = NullSetValue(W_533!form_man, "")
            check_man.Text = NullSetValue(W_533!check_man, "")
            bar_man.Text = NullSetValue(W_533!bar_man, "")
            Qc_Man.Text = NullSetValue(W_533!Qc_Man, "")
            mag_man.Text = NullSetValue(W_533!mag_man, "")
            
            qc_status.Value = NullSetValue(W_533!qc_status, "1")
            
            W_Status = IIf(NullSetValue(W_533!status, "0") = "0", True, False)
            
        Else
            inv_no.Text = ""
            Inv_Date.Value = Date
            P_Line_No.Text = ""
            Inv_Style.Text = ""
            
            Remark.Text = ""
            
            form_man.Text = ""
            check_man.Text = ""
            bar_man.Text = ""
            mag_man.Text = ""
            W_Status = False
        End If
        W_533.Close
        Set W_533 = Nothing
    
        '刷新表格
        Call RefreshGrid
        If Adodc1.Recordset.EOF = False Then
            TDBGrid1.Row = 1
        End If
        Call readshow
    End If
End Sub
'刷新TDBGrid1,之所以定为public,是因为还会被表单frmp_linequatmx调用
Public Sub RefreshGrid()
    Dim w_rs534 As New ADODB.Recordset
    
    Set w_rs534 = open_RS(" select *  from  SQL_bar_603 ('" & Trim(inv_no.Text) & "') ")

    
    Set Adodc1.Recordset = w_rs534
    Set TDBGrid1.DataSource = Adodc1
    
    Call readactive
    Set w_rs534 = Nothing
End Sub

Private Sub Inv_Date_Change()
    If Not (C_Add Or C_Edit) Then
        Inv_Date.Value = W_inv_date
    End If
End Sub

'自动生成单号 "前缀字符"I/O"+年份两位+月份+5位流水号
Private Function Creat_No()
    Dim w_tmp As New ADODB.Recordset
    Dim W_Str As String
    
    Dim W_Inv_No As String
        
    W_Inv_No = "O-"        '入库
    
    W_Inv_No = W_Inv_No & Right(CStr(Year(Get_SQLDATE)), 2) & Format(CStr(Month(Get_SQLDATE)), "00") & Format(CStr(Day(Get_SQLDATE)), "00")
    
    W_Str = "SELECT Max(Inv_no) As Inv_no  FROM mmst533  WHERE Inv_no like '" & W_Inv_No & "%' "
                
    w_tmp.Open W_Str, G_Con, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If w_tmp.EOF = False Then
        If IsNull(w_tmp!inv_no) Then
            W_Inv_No = W_Inv_No & "001"
        Else
            W_Inv_No = W_Inv_No & Format(CStr(Val(Right(w_tmp!inv_no, 3)) + 1), "000")
        End If
    Else
        W_Inv_No = W_Inv_No & "001"
    End If
    
    Creat_No = W_Inv_No
End Function

'筛选打印数据并实现列印或预览效果
Private Sub sele_data()
    Dim W_Print As DAO.Recordset
    Dim W_BookMark As Variant
    Dim W_Rs As New ADODB.Recordset
    
    '清除打印数据表
    G_PrintDb.Execute "DELETE * FROM mmsr6041"
    Set W_Print = G_PrintDb.OpenRecordset("SELECT * FROM mmsr6041")
'    On Error Resume Next
    '选取数据
    With W_Print
        If Adodc1.Recordset.AbsolutePosition <> -1 Then
            W_BookMark = Adodc1.Recordset.Bookmark
            Adodc1.Recordset.MoveFirst
            Do Until Adodc1.Recordset.EOF
                .AddNew
                !loc_id = "A"
                !inv_no = NullSetValue(Trim(inv_no.Text), "")
                !Inv_Date = NullSetValue(Inv_Date.Value, "")
                !form_man = NullSetValue(form_man.Text, "")
                !check_man = NullSetValue(check_man.Text, "")
                !bar_man = NullSetValue(bar_man.Text, "")
                !mag_man = NullSetValue(mag_man.Text, "")
                !Qc_Man = NullSetValue(Qc_Man.Text, "")
                If Left(Trim(Inv_Style.Text), 2) = "正常" Then
                    !inv_style1 = "V"
                Else
                    !inv_Style2 = "V"
                End If
                
                !Remark = NullSetValue(Remark.Text, "")
                !P_Line_No = NullSetValue(P_Line_No.Text, "")
    
                !Mo_No = NullSetValue(Trim(Adodc1.Recordset!Mo_No), "")
                !Mtr_No = NullSetValue(Trim(Adodc1.Recordset!Mtr_No), "")
                !Mtr_Name = NullSetValue(Trim(Adodc1.Recordset!Mtr_Name), "") & " / " & NullSetValue(Adodc1.Recordset!Mtr_Dim, "")
                !Mtr_Dim = NullSetValue(Adodc1.Recordset!Mtr_Dim, "")
                !color_name = NullSetValue(Adodc1.Recordset!color_name, "")
                !mtr_amt = NullSetValue(Trim(Adodc1.Recordset!mtr_amt), 0)
                !Mtr_Amt_order = NullSetValue(Trim(Adodc1.Recordset!Mtr_Amt_order), 0)
                !unit_name = NullSetValue(Trim(Adodc1.Recordset!unit_name), "")
                !Qc_No = NullSetValue(Trim(Adodc1.Recordset!Qc_No), "")
                !Bar_Name = NullSetValue(Trim(Adodc1.Recordset!Bar_Name), "")
                If qc_status.Value = 1 Then
                !Spe_Let = NullSetValue(Trim(Adodc1.Recordset!qc_result), "")
                End If
               
                !Note = NullSetValue(Trim(Adodc1.Recordset!Note), "")
                .Update
                Adodc1.Recordset.MoveNext
            Loop
            Adodc1.Recordset.Bookmark = W_BookMark
        End If
    End With
    
    W_Print.Close
    
    If C_Print Then
        C_Print = False
        Call print_rpt(G_MDIForm.Rpt1, "mmsr6041", "P")
    End If
    
    If C_View Then
        C_View = False
        Call print_rpt(G_MDIForm.Rpt1, "mmsr6041", "V")
    End If
    
    If C_Save Then
        C_Save = False
        Set G_Rpt = G_MDIForm.Rpt1
        G_Rpt_Name = "6041"
        mmssave.Show vbModal
    End If
End Sub

Private Sub load_picture()
    cmd_find.PictureURL = App.Path + "\Picture\Norm\find_norm.bmp"
    cmd_find.PictureDisableURL = App.Path + "\Picture\Dis\Find_dis.bmp"
    cmd_find.PictureOverURL = App.Path + "\Picture\Over\Find_Over.bmp"
    
    cmd_print.PictureURL = App.Path + "\Picture\Norm\print_norm.bmp"
    cmd_print.PictureDisableURL = App.Path + "\Picture\Dis\print_dis.bmp"
    cmd_print.PictureOverURL = App.Path + "\Picture\Over\print_Over.bmp"
    
    cmd_preview.PictureURL = App.Path + "\Picture\Norm\preview_norm.bmp"
    cmd_preview.PictureDisableURL = App.Path + "\Picture\Dis\preview_dis.bmp"
    cmd_preview.PictureOverURL = App.Path + "\Picture\Over\preview_Over.bmp"
    
    cmd_save.PictureURL = App.Path + "\Picture\Norm\save_norm.bmp"
    cmd_save.PictureDisableURL = App.Path + "\Picture\Dis\save_dis.bmp"
    cmd_save.PictureOverURL = App.Path + "\Picture\Over\save_Over.bmp"
    
    Cmd_ok.PictureURL = App.Path + "\Picture\Norm\Ok_norm.bmp"
    Cmd_ok.PictureDisableURL = App.Path + "\Picture\Dis\Ok_dis.bmp"
    Cmd_ok.PictureOverURL = App.Path + "\Picture\Over\Ok_Over.bmp"
    
    cmd_cancel.PictureURL = App.Path + "\Picture\Norm\cancel_norm.bmp"
    cmd_cancel.PictureDisableURL = App.Path + "\Picture\Dis\cancel_dis.bmp"
    cmd_cancel.PictureOverURL = App.Path + "\Picture\Over\cancel_Over.bmp"
    
    cmd_add.PictureURL = App.Path + "\Picture\Norm\add_norm.bmp"
    cmd_add.PictureDisableURL = App.Path + "\Picture\Dis\add_dis.bmp"
    cmd_add.PictureOverURL = App.Path + "\Picture\Over\add_Over.bmp"
    
    cmd_edit.PictureURL = App.Path + "\Picture\Norm\edit_norm.bmp"
    cmd_edit.PictureDisableURL = App.Path + "\Picture\Dis\edit_dis.bmp"
    cmd_edit.PictureOverURL = App.Path + "\Picture\Over\edit_Over.bmp"
    
    cmd_delete.PictureURL = App.Path + "\Picture\Norm\delete_norm.bmp"
    cmd_delete.PictureDisableURL = App.Path + "\Picture\Dis\delete_dis.bmp"
    cmd_delete.PictureOverURL = App.Path + "\Picture\Over\delete_Over.bmp"
    
    cmd_quit.PictureURL = App.Path + "\Picture\Norm\Quit_norm.bmp"
    cmd_quit.PictureDisableURL = App.Path + "\Picture\Dis\Quit_dis.bmp"
    cmd_quit.PictureOverURL = App.Path + "\Picture\Over\Quit_Over.bmp"
    
    Set Me.Picture2 = G_MDIForm.Picture
    AresButton1.PictureURL = App.Path + "\Picture\file.gif"
    AresButton1.GifAnimationPlay

End Sub
'**********************************************************************
Private Sub Cmd_find_MouseEnter()
    Help_txt.Caption = cmd_find.ToolTipString
    Help_txt.Refresh
End Sub

Private Sub Cmd_find_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub Cmd_print_MouseEnter()
    Help_txt.Caption = cmd_print.ToolTipString
    Help_txt.Refresh

End Sub

Private Sub Cmd_print_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub Cmd_previeW_MouseEnter()
    Help_txt.Caption = cmd_preview.ToolTipString
    Help_txt.Refresh

End Sub

Private Sub Cmd_previeW_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub Cmd_save_MouseEnter()
    Help_txt.Caption = cmd_save.ToolTipString
    Help_txt.Refresh
End Sub
Private Sub Cmd_save_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub cmd_add_MouseEnter()
    Help_txt.Caption = cmd_add.ToolTipString
    Help_txt.Refresh

End Sub

Private Sub Cmd_Add_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub cmd_edit_MouseEnter()
    Help_txt.Caption = cmd_edit.ToolTipString
    Help_txt.Refresh
End Sub

Private Sub cmd_edit_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub cmd_delete_MouseEnter()
    Help_txt.Caption = cmd_delete.ToolTipString
    Help_txt.Refresh

End Sub

Private Sub cmd_delete_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub cmd_ok_MouseEnter()
    Help_txt.Caption = Cmd_ok.ToolTipString
    Help_txt.Refresh

End Sub

Private Sub cmd_ok_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub cmd_cancel_MouseEnter()
    Help_txt.Caption = cmd_cancel.ToolTipString
    Help_txt.Refresh

End Sub

Private Sub cmd_cancel_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub

Private Sub Cmd_quit_MouseEnter()
    Help_txt.Caption = cmd_quit.ToolTipString
    Help_txt.Refresh

End Sub

Private Sub Cmd_quit_MouseLeave()
    Help_txt.Caption = ""
    Help_txt.Refresh
End Sub
'**********************************************************************
Private Sub Cmd_OK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Cmd_ok_MouseClick
    End If
End Sub


Private Sub Cmd_cancel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Cmd_cancel_MouseClick
    End If
End Sub


Private Sub Cmd_find_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Cmd_find_MouseClick
    End If
End Sub


Private Sub Cmd_print_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Cmd_print_MouseClick
    End If
End Sub


Private Sub Cmd_previeW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Cmd_previeW_MouseClick
    End If
End Sub

Private Sub Cmd_Save_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Cmd_save_MouseClick
    End If
End Sub

Private Sub Cmd_add_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmd_add_MouseClick
    End If
End Sub

Private Sub Cmd_edit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmd_edit_MouseClick
    End If
End Sub

Private Sub Cmd_delete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmd_delete_MouseClick
    End If
End Sub

Private Sub Cmd_quit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Cmd_quit_MouseClick
    End If
End Sub
Private Sub Cmd_find_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_find.Top
End Sub

Private Sub Cmd_find_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub Cmd_print_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_print.Top
End Sub

Private Sub Cmd_print_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub Cmd_previeW_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_preview.Top
End Sub

Private Sub Cmd_previeW_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub Cmd_save_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_save.Top
End Sub

Private Sub Cmd_save_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub cmd_ok_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = Cmd_ok.Top
End Sub

Private Sub cmd_ok_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub cmd_cancel_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_cancel.Top
End Sub

Private Sub cmd_cancel_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub cmd_add_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_add.Top
End Sub

Private Sub cmd_add_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub cmd_edit_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_edit.Top
End Sub

Private Sub cmd_edit_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub cmd_delete_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_delete.Top
End Sub

Private Sub cmd_delete_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub Cmd_quit_SetFocus()
    lab_focus.Visible = True
    lab_focus.Top = cmd_quit.Top
End Sub

Private Sub Cmd_quit_LeaveFocus()
    lab_focus.Visible = False
End Sub

Private Sub load_combox()
    '加载常用人员名单
    Call AddRsToList(bar_man, "SELECT User_name FROM mmst801 order by user_name", , 0)
    'Call AddRsToList(mag_Man, "SELECT User_name FROM mmst801 order by user_name", , 0)
    'Call AddRsToList(check_man, "SELECT User_name FROM mmst801 order by user_name", , 0)
    Call AddRsToList(Qc_Man, "SELECT User_name FROM mmst801 order by user_name", , 0)
'    '生产线
    Call AddRsToList(P_Line_No, "SELECT p_line_name FROM mmst811 order by p_line_name", , 0)
    '出入库种类
    Inv_Style.Clear
    Inv_Style.AddItem "正常入库"
    Inv_Style.AddItem "重工入库"
End Sub
