VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "AresButtonPro.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form mmss904 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "物料料号修改(904)"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15165
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   2490
      ScaleHeight     =   2115
      ScaleWidth      =   12615
      TabIndex        =   15
      Top             =   810
      Width           =   12645
      Begin VB.TextBox Old_Mtr_Type 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   23
         TabIndex        =   27
         Top             =   640
         Width           =   2040
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Caption         =   "物料类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   8070
         TabIndex        =   23
         Top             =   60
         Width           =   3105
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000005&
            Caption         =   "成品"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   270
            TabIndex        =   26
            Top             =   1320
            Width           =   2325
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000005&
            Caption         =   "半成品"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   270
            TabIndex        =   25
            Top             =   810
            Width           =   2325
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000005&
            Caption         =   "原材料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   270
            TabIndex        =   24
            Top             =   330
            Value           =   -1  'True
            Width           =   2325
         End
      End
      Begin VB.TextBox Mtr_Dim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1560
         Width           =   6375
      End
      Begin VB.TextBox Mtr_Name 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1100
         Width           =   6375
      End
      Begin VB.CommandButton Cmd_Mtr_Brow 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3270
         TabIndex        =   7
         Top             =   180
         Width           =   300
      End
      Begin VB.TextBox Old_Mtr_No 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1530
         MaxLength       =   23
         TabIndex        =   6
         Top             =   180
         Width           =   2040
      End
      Begin VB.TextBox New_Mtr_No 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5700
         MaxLength       =   15
         TabIndex        =   8
         Top             =   180
         Width           =   2190
      End
      Begin MSForms.ComboBox New_Mtr_Type 
         Height          =   345
         Left            =   5700
         TabIndex        =   28
         Top             =   640
         Width           =   2190
         VariousPropertyBits=   746604571
         BackColor       =   16776960
         DisplayStyle    =   3
         Size            =   "3863;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   3
         FontName        =   "新细明体"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新物料编号:"
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
         Left            =   4380
         TabIndex        =   22
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新物料类别:"
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
         Left            =   4380
         TabIndex        =   21
         Top             =   735
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "规       格:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "品       名:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "旧物料编号:"
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
         Left            =   225
         TabIndex        =   18
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "旧物料类别:"
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
         Left            =   225
         TabIndex        =   17
         Top             =   735
         Width           =   1185
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
      Bindings        =   "mmss90d.frx":0000
      Height          =   6195
      Left            =   2490
      TabIndex        =   11
      Top             =   2940
      Width           =   12675
      _cx             =   22357
      _cy             =   10927
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      BackColorFixed  =   -2147483634
      ForeColorFixed  =   -2147483630
      BackColorSel    =   65280
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   0
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
      FormatString    =   $"mmss90d.frx":0015
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   5
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8325
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   2475
      TabIndex        =   12
      Top             =   810
      Width           =   2505
      Begin ARESBUTTONLib.AresButton cmd_ok 
         Height          =   360
         Left            =   690
         TabIndex        =   0
         Top             =   420
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
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
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_cancel 
         Height          =   420
         Left            =   690
         TabIndex        =   1
         Top             =   1200
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
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
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_add 
         Height          =   420
         Left            =   690
         TabIndex        =   2
         Top             =   1995
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         ToolTipString   =   "增加一笔记录"
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
         PrevPointer     =   127107100
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_edit 
         Height          =   420
         Left            =   690
         TabIndex        =   3
         Top             =   2775
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         ToolTipString   =   "修改该记录"
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
      Begin ARESBUTTONLib.AresButton cmd_delete 
         Height          =   420
         Left            =   690
         TabIndex        =   4
         Top             =   3570
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         ToolTipString   =   "删除该记录"
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
      Begin ARESBUTTONLib.AresButton cmd_quit 
         Height          =   360
         Left            =   690
         TabIndex        =   5
         Top             =   4350
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
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
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton AresButton1 
         Height          =   825
         Left            =   780
         TabIndex        =   13
         Top             =   6270
         Width           =   1050
         _Version        =   327687
         BackGroundColor =   16777215
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\erp_proj.gif"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\erp_proj.gif"
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
         PrevPointer     =   90897964
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2490
         Y1              =   5070
         Y2              =   5070
      End
      Begin VB.Label Help_txt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   300
         TabIndex        =   14
         Top             =   5400
         Width           =   1755
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   30
      Top             =   6900
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "物料料号修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   465
      Left            =   7380
      TabIndex        =   16
      Top             =   180
      Width           =   2985
   End
End
Attribute VB_Name = "mmss904"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称: 用户资料档(MMSS904)
'*编写日期: 2002年07月29日
'*制作人员: 于
'*修改日期:
'*修改人员:
'***********************************************
'定义欲打开的数据库及数据表名称
Dim St_904 As New ADODB.Recordset

'定义按钮变量
Dim c_add As Boolean
Dim c_edit As Boolean
Dim c_delete As Boolean

'存放TDBGRID1 的旧字符
Dim W_Old_Str As String
Dim W_Mtr_Type As String

Dim c_off_add As Boolean
Dim c_off_edit As Boolean
Dim c_off_delete As Boolean

'纪录当前行列
Dim W_col As Double
Dim W_Row As Double

'定义窗体打开变量
Dim Gridc_904(127) As Grid_Data '存放 Grid 属性值
Dim Row_Height As Double        'Grid 高度变量

Private Sub Old_Mtr_No_LostFocus()
'定位处理
Dim w_tmp As New ADODB.Recordset

'判断代号是否重复
w_tmp.CursorLocation = adUseClient
w_tmp.Open "select * from mmsp611 where Mtr_No = '" & Trim(Old_Mtr_No.Text) & "'", G_Con, adOpenForwardOnly
If w_tmp.EOF = False Then
    Mtr_Name.Text = w_tmp!Mtr_Name
    Mtr_Dim.Text = NullSetValue(w_tmp!Mtr_Dim, "")
    Old_Mtr_Type.Text = NullSetValue(w_tmp!type_name, "")
    
    If w_tmp!Type = "0" Then
        Option1.Value = True
    ElseIf w_tmp!Type = "1" Then
        Option2.Value = True
    ElseIf w_tmp!Type = "2" Then
        Option3.Value = True
    End If
Else

    Mtr_Name.Text = ""
    Mtr_Dim.Text = ""
    Old_Mtr_Type.Text = ""
End If
Set w_tmp = Nothing
End Sub

Private Sub Cmd_Mtr_Brow_Click()
With FrmMtrList
       .G_Type = "012"
       .G_Mtr_No = Trim(Old_Mtr_No.Text)
       
    .Show vbModal
    If .mtr_no <> "" Then
        Old_Mtr_Type.Text = .Mtr_Type_Tmp
        
        Old_Mtr_No.Text = .mtr_no
        Mtr_Name.Text = .Mtr_Name
        Mtr_Dim.Text = .Mtr_Dim
        
        Old_Mtr_Type.Text = .mtr_type
        New_Mtr_Type.SetFocus
    End If
End With
End Sub

Public Sub Form_Activate()
'当窗口激活时,刷新TDBGrid
Call GetVSGridSetting("mmss904", "TDBGrid1", Gridc_904, g_CON_IniFile9)
Row_Height = Gridc_904(0).Grid_RowHeight
Call readactive
'刷新表格
Call RefreshGrid
TDBGrid1.Col = 1
If TDBGrid1.Rows > 1 Then
    TDBGrid1.Row = 1
End If

End Sub

Private Sub Form_Load()
'装载图片
Call load_picture
'将窗口置中
Call CenterWindow(mmss904, sys_main)

'将按钮变量赋初值
c_add = False
c_edit = False
c_delete = False
c_off_add = False
c_off_edit = False
c_off_delete = False

'MDI子窗口按钮权限设订
c_off_add = lopcheck("A", "904", G_User_ID)
c_off_edit = lopcheck("U", "904", G_User_ID)
c_off_delete = lopcheck("D", "904", G_User_ID)

Dim W_902 As New ADODB.Recordset
W_902.Open "select type_name from mmst603 order by type_name", G_Con, adOpenDynamic
Do While W_902.EOF = False
    New_Mtr_Type.AddItem W_902!type_name
    W_902.MoveNext
Loop
W_902.Close

If c_off_add = True Then
    cmd_add.Enabled = False
End If

If c_off_edit = True Then
    cmd_edit.Enabled = False
End If

If c_off_delete = True Then
    cmd_delete.Enabled = False
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Form_KeyDown可照搬
If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "TDBGrid1" Then
    
    If ActiveControl.Name = "note" Then
        If ActiveControl.MultiLine = False Then
            SendKeys "{TAB}"
        End If
    Else
        SendKeys "{TAB}"
    End If
    Exit Sub
End If

If Shift = 0 Then
    Select Case KeyCode
        '按 ESC 时取消操作
        Case vbKeyEscape
            If (c_add Or c_edit Or c_delete) Then
                Call vcontrol("N")
                KeyCode = 0
            End If
        '当按 F2 时,新　记录
        Case vbKeyF2
            If cmd_add.Enabled Then
                Call vcontrol("A")
                KeyCode = 0
            End If
        '当按 F3 时 , 修改记录
        Case vbKeyF3
            If Me.ActiveControl.Name = "Old_Mtr_No" Then
                Call Old_Mtr_No_LostFocus
            End If

            If cmd_edit.Enabled Then
                Call vcontrol("U")
                KeyCode = 0
            End If
        '当按 F4 时 , 删除记录
        Case vbKeyF4
            If Me.ActiveControl.Name = "Old_Mtr_No" Then
                Call Old_Mtr_No_LostFocus
            End If

            If cmd_delete.Enabled Then
                Call vcontrol("D")
                KeyCode = 0
            End If
        '当按 F5 时,保存记录
        Case vbKeyF5
            If cmd_ok.Enabled Then
                Call vcontrol("Y")
                KeyCode = 0
            End If
        '当按 F6 时,退出系统
        Case vbKeyF6
            If Cmd_Quit.Enabled Then
                Call vcontrol("Q")
                KeyCode = 0
            End If
    End Select
End If
End Sub

Sub readshow()
'对控件置值
If c_add = True Or Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then
    Old_Mtr_No.Text = ""
    Mtr_Name.Text = ""
    Mtr_Dim.Text = ""
    Old_Mtr_Type.Text = ""
    New_Mtr_Type.Text = ""
    
    New_Mtr_No.Text = ""
    
Else
    Old_Mtr_No.Text = St_904!Old_Mtr_No
    Mtr_Name.Text = NullSetValue(St_904!Mtr_Name, "")
    Mtr_Dim.Text = NullSetValue(St_904!Mtr_Dim, "")
    New_Mtr_Type.Text = NullSetValue(St_904!New_Mtr_Type, "")
    Old_Mtr_Type.Text = NullSetValue(St_904!Old_Mtr_Type, "")
    
    New_Mtr_No.Text = NullSetValue(St_904!New_Mtr_No, "")
    
End If

'设定按键的 Enabled 属性
If c_add Or c_edit Or c_delete Then
    cmd_add.Enabled = False
    cmd_edit.Enabled = False
    cmd_delete.Enabled = False
    cmd_ok.Enabled = True
    cmd_cancel.Enabled = True
Else
    cmd_add.Enabled = True
    cmd_edit.Enabled = True
    cmd_delete.Enabled = True
    cmd_ok.Enabled = False
    cmd_cancel.Enabled = False
    
    '当数据表无记录时
    If St_904.EOF Then
        cmd_edit.Enabled = False
        cmd_delete.Enabled = False
    Else
        cmd_edit.Enabled = True
        cmd_delete.Enabled = True
    End If
End If

'通过权限设定按键的 Enabled
If c_off_add = True Then
    cmd_add.Enabled = False
End If
    
If c_off_edit = True Then
    cmd_edit.Enabled = False
End If
    
If c_off_delete = True Then
    cmd_delete.Enabled = False
End If

If Not (c_add Or c_edit) Then
    Mtr_Name.Locked = True
    Old_Mtr_Type.Locked = True
    Mtr_Dim.Locked = True
    New_Mtr_Type.Locked = True
Else
    Mtr_Name.Locked = False
    Old_Mtr_Type.Locked = False
    Mtr_Dim.Locked = False
    New_Mtr_Type.Locked = False

End If
If c_edit Then
    Old_Mtr_No.Locked = True
Else
    Old_Mtr_No.Locked = False
End If
End Sub
'刷新表格
Private Sub RefreshGrid()
Call readactive
Call readshow
End Sub

Private Sub readactive()
Set St_904 = Nothing
With St_904
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
    .Open "select a.Old_Mtr_No  ,a.old_mtr_type," & _
                 "a.Mtr_Name,a.Mtr_dim," & _
                 "a.New_Mtr_no,a.New_Mtr_Type," & _
                 "a.upd_name," & _
                 "a.upd_date " & _
            "FROM mmst904 a ORDER BY a.Old_Mtr_No  ,a.old_mtr_type"

End With

'设置tdbgrid1的数据来源
Set Adodc1.Recordset = St_904

Call SetVSGridSetting(TDBGrid1, Gridc_904)

'刷新全部 ROW 的高度 包括 HEADER
For i = 1 To TDBGrid1.Rows
    TDBGrid1.Row = i - 1
    TDBGrid1.RowHeight(i - 1) = Row_Height
    
    If i < TDBGrid1.Rows Then
        TDBGrid1.TextMatrix(i, 0) = i
    End If
Next i
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

End Sub

'命令按键事件
Private Sub Cmd_Add_MouseClick()
Call vcontrol("A")
End Sub

Private Sub Cmd_Edit_MouseClick()
Call vcontrol("U")
End Sub

Private Sub Cmd_Delete_MouseClick()
Call vcontrol("D")
End Sub

Private Sub Cmd_OK_MouseClick()
Call vcontrol("Y")
End Sub

Private Sub cmd_cancel_MouseClick()
Call vcontrol("N")
End Sub

Private Sub cmd_quit_MouseClick()
Call vcontrol("Q")
End Sub

Private Sub vcontrol(p_choice As String)
Select Case p_choice
    Case "Y"            '确定
        If check_ok() = True Then
            Call upd_data
            TDBGrid1.Enabled = True
        End If
        
    Case "N"            '取消
        '解锁处理
        If c_edit Or c_delete Then
            Call UnLockRecord("mmst904", "Old_Mtr_No='" & Old_Mtr_No.Text & "'")
        End If
        c_add = False
        c_edit = False
        c_delete = False
        
        TDBGrid1.Enabled = True
        Call readshow
        
    Case "A"             '增加
        c_add = True
        Call readshow
        Old_Mtr_No.SetFocus
        TDBGrid1.Enabled = False
        
    Case "U"             '修改
        '加锁
        If LockRecord("mmst904", "Old_Mtr_No='" & Old_Mtr_No.Text & "'") Then
            W_Row = TDBGrid1.Row
            W_col = TDBGrid1.Col
            
            c_edit = True
            TDBGrid1.Enabled = False
            Call readshow
'            Mtr_Name.SetFocus
        End If
        
    Case "D"             '删除
        '加锁
        If LockRecord("mmst904", "Old_Mtr_No='" & Old_Mtr_No.Text & "'") = True Then
            If MsgBox(g_CON_CDelete, vbYesNo + vbDefaultButton2 + vbInformation, g_CON_CTitle) = vbNo Then
                Call UnLockRecord("mmst904", "Old_Mtr_No='" & Old_Mtr_No.Text & "'")
                Exit Sub
            End If
            
            '判断是否可以删除
            c_delete = True
            If check_ok = False Then
                Call UnLockRecord("mmst904", "Old_Mtr_No='" & Old_Mtr_No.Text & "'")
                c_delete = False
                Exit Sub
            End If
            
            '删除记录
            G_Con.Execute "DELETE FROM mmst904 WHERE Old_Mtr_No='" & Trim(Old_Mtr_No.Text) & "'"
            c_delete = False
            '刷新数据
            
            Call RefreshGrid
            
            '删除后移动到第一笔记录
            TDBGrid1.Col = 1
            If TDBGrid1.Rows > 1 Then
                TDBGrid1.TopRow = 1
                TDBGrid1.Row = 1
            End If
            
        End If
    Case "Q"            '退出
        Unload Me
End Select
End Sub

'当修改或删除或新增时进行一致性判断
Private Function check_ok() As Boolean
Dim w_tmp As New ADODB.Recordset
'If c_delete Then
'    If Old_Mtr_No = "A001" Then
'        MsgBox "该用户为系统用户,不可删除!", 64, "提示信息"
'        Old_Mtr_No.SetFocus
'        check_ok = False
'        Exit Function
'    End If
'    check_ok = True
'End If
'新增时判断
If c_add = True Then
    If Trim(Old_Mtr_No.Text) = "" Then
        MsgBox "请输入旧物料编号", 64, "提示信息"
        Old_Mtr_No.SetFocus
        check_ok = False
        Exit Function
    Else
        '判断代号是否重复
        w_tmp.CursorLocation = adUseClient
        w_tmp.Open "select Mtr_No from mmst611 where Mtr_No = '" & Trim(Old_Mtr_No.Text) & "'", G_Con, adOpenForwardOnly
        If w_tmp.EOF Then
            MsgBox "旧物料编号不存在,请重新确认!", 64, "提示信息"
            Old_Mtr_No.SetFocus
            check_ok = False
            Set w_tmp = Nothing
            Exit Function
        End If
        Set w_tmp = Nothing
    End If
End If

'If Trim(New_Mtr_Type.Text) = Trim(Old_Mtr_Type.Text) And Trim(New_Mtr_No.Text) = Trim(New_Mtr_No.Text) Then
'    MsgBox "新旧物料类别及物料料号都相同,何必修改阿", 64, "提示信息"
'    New_Mtr_Type.SetFocus
'    check_ok = False
'    Exit Function
'End If


'新增和修改时判断
If New_Mtr_No.Text = "" Then
    MsgBox "请输入新物料编号", 64, "提示信息"
    New_Mtr_No.SetFocus
    check_ok = False
    Exit Function
Else
    If Trim(New_Mtr_Type.Text) = Trim(Old_Mtr_Type.Text) Then

        '判断用户名称是否重复
        w_tmp.CursorLocation = adUseClient
        w_tmp.Open "select Mtr_No from mmst611 where Mtr_No= '" & Trim(New_Mtr_No.Text) & "' and Type='" & IIf(Option1.Value, "0", IIf(Option2.Value, "1", "2")) & "'", G_Con, adOpenForwardOnly
        
        If w_tmp.EOF = False Then
            MsgBox "新物料编号已经存在!", 64, "提示信息"
            New_Mtr_No.SetFocus
            Set w_tmp = Nothing
            check_ok = False
            Exit Function
        End If
        Set w_tmp = Nothing
    End If
End If


If New_Mtr_Type.Text = "" Then
    MsgBox "请输入新物料类别", 64, "提示信息"
    New_Mtr_Type.SetFocus
    check_ok = False
    Exit Function
Else
    w_tmp.Open "select Mtr_Type from mmst603 where type_name='" & New_Mtr_Type.Text & "'", G_Con, , , adCmdText
    If w_tmp.EOF = True Then
        w_tmp.Close
        MsgBox "无此物料类别.", vbExclamation, g_CON_CTitle
        New_Mtr_Type.SetFocus
        Exit Function
    Else
        W_Mtr_Type = w_tmp!mtr_type
    End If
    w_tmp.Close
End If


check_ok = True
End Function

'对数据库进行更新
Private Sub upd_data()
Dim St_904_1 As New ADODB.Recordset
Dim st_905 As New ADODB.Recordset

Dim W_Find As String

W_Find = Old_Mtr_No.Text

On Error GoTo UpdateError
G_Con.BeginTrans

With St_904_1
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
    .Open "select * from mmst904 "
End With

'新增一笔记录到数据库
If c_add = True Then
    With St_904_1
        .AddNew
        !Old_Mtr_No = UCase(Trim(Old_Mtr_No.Text))
        !New_Mtr_No = UCase(Trim(New_Mtr_No.Text))
        
        !Mtr_Name = Trim(Mtr_Name.Text)
        !Mtr_Dim = Trim(Mtr_Dim.Text)
        !Old_Mtr_Type = Old_Mtr_Type.Text
        
        !New_Mtr_Type = New_Mtr_Type.Text
        
        !upd_name = Trim(G_Mtr_Name)
        !upd_date = Get_SQLDATE
        !lock = "No"
        .Update
    End With
    Set St_904_1 = Nothing
    c_add = False
End If
'
''修改记录
'If c_edit = True Then
'    With st_904_1
'        !Mtr_Name = Mtr_Name.Text
'        !Old_Mtr_Type = Old_Mtr_Type.Text
'        !dpt_id = w_dpt_id
'
'        !Mtr_Dim = Trim(Mtr_Dim.Text)
'        !upd_Name = Trim(G_Mtr_Name)
'        !upd_date = Get_SQLDATE
'        !lock = "No"
'        .Update
'    End With
'    Set st_904_1 = Nothing
'    c_edit = False
'End If
If Trim(New_Mtr_No.Text) <> Trim(Old_Mtr_No.Text) Then
    With st_905
        .ActiveConnection = G_Con
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockPessimistic
        .Open "select * from mmst905 "
    End With
    
    G_Con.Execute "update mmst611 set mtr_type='" & W_Mtr_Type & "' ,type='" & IIf(Option1.Value, "0", IIf(Option2.Value, "1", "2")) & "' Where mtr_no='" & Trim(Old_Mtr_No.Text) & "'"
    Do Until st_905.EOF
        
        G_Con.Execute "update " & st_905!Table_Name & " set " & st_905!Mtr_No_Name & "='" & Trim(New_Mtr_No.Text) & "' Where " & st_905!Mtr_No_Name & "='" & Trim(Old_Mtr_No.Text) & "'"
        st_905.MoveNext
    Loop
    Set st_905 = Nothing
Else
    G_Con.Execute "update mmst611 set mtr_type='" & W_Mtr_Type & "',type='" & IIf(Option1.Value, "0", IIf(Option2.Value, "1", "2")) & "' Where mtr_no='" & Trim(Old_Mtr_No.Text) & "'"
End If


'刷新数据表
Call RefreshGrid

TDBGrid1.Row = TDBGrid1.FindRow(W_Find, 0, 1, False)
TDBGrid1.Col = W_col
TDBGrid1.TopRow = TDBGrid1.FindRow(W_Find, 0, 1, False)
G_Con.CommitTrans

Endx:
c_add = False

Call RefreshGrid

Exit Sub

UpdateError:
G_Con.RollbackTrans
MsgBox "更新时发生错误!", 64, g_CON_CTitle
GoTo Endx

End Sub

'表单的 QueryUnload 和 Unload 事件
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If c_add Or c_edit Or c_delete Then
    '当有数据改动时.询问是否要退出系统
    If MsgBox(g_CON_CQuit, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
        Cancel = 1
    Else
        '当有修改或删除时未解锁时,解除锁定
        If c_edit Or c_delete Then
            Call UnLockRecord("mmst904", "Old_Mtr_No='" & Old_Mtr_No.Text & "'")
        End If
        Cancel = 0
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

'退出时，存储 TDBGrid 属性
Call SaveGridSetting("mmss904", "TDBGrid1", Gridc_904, g_CON_IniFile9)

Set TDBGrid1.DataSource = Nothing
Set St_904 = Nothing
Set mmss904 = Nothing
End Sub

'各控件的相关事件
Private Sub Old_Mtr_No_KeyPress(KeyAscii As Integer)
'控制输入为字母或数字
If c_add Then
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
           (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or _
           (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = vbKeyBack Or _
            KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
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
        TDBGrid1.FocusRect = flexFocusRaised
    End If
    Call readshow
End If
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
End Sub

Private Sub TDBGrid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'移动COl改变宽度
If Col > 0 Then
    If Col > Gridc_904(0).Grid_Columns Then
        Cancel = 1
    Else
        If UCase(Mid(Gridc_904(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_904(Col - 1).Grid_Visible = "" Then
            Cancel = 1
        Else
            Gridc_904(Col - 1).Grid_Width = TDBGrid1.ColWidth(Col)
        End If
    End If
End If

'移动ROW改变高度
If Row >= 0 Then
    w_cur_row = TDBGrid1.Row
    Row_Height = TDBGrid1.RowHeight(Row)
    Gridc_904(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.Row = i - 1
        TDBGrid1.RowHeight(i - 1) = Row_Height
    Next i
    TDBGrid1.Row = w_cur_row
End If

End Sub

Private Sub TDBGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'鼠标点在HEADER上
If X > 0 And Y < Row_Height Then
   
    '存储 TDBGrid 属性
    Call SaveVSGridSetting("mmss904", "TDBGrid1", Gridc_904, g_CON_IniFile9)
    
    '调用 TDBGrid 属性设定
    With mmss_set
        Set .Parent_form = mmss904
        .Get_FormName = "mmss904"
        .Get_GridName = "TDBGrid1"
        .Gridc_File = g_CON_IniFile9
        .Show vbModal
    End With
End If
End Sub

Private Sub TDBGrid1_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'不许更改第0行COl的宽度
If Col = 0 Then
    Cancel = True
End If
End Sub


'**********************************************************************
'装载按钮图片
'**********************************************************************
Private Sub load_picture()
cmd_ok.PictureURL = App.Path + "\Picture\Norm\Ok_norm.bmp"
cmd_ok.PictureDisableURL = App.Path + "\Picture\Dis\Ok_dis.bmp"
cmd_ok.PictureOverURL = App.Path + "\Picture\Over\Ok_Over.bmp"

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

Cmd_Quit.PictureURL = App.Path + "\Picture\Norm\Quit_norm.bmp"
Cmd_Quit.PictureDisableURL = App.Path + "\Picture\Dis\Quit_dis.bmp"
Cmd_Quit.PictureOverURL = App.Path + "\Picture\Over\Quit_Over.bmp"

AresButton1.PictureURL = App.Path + "\Picture\file.gif"
AresButton1.GifAnimationPlay
End Sub

'**********************************************************************
'更改提示符
'**********************************************************************

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
Help_txt.Caption = cmd_ok.ToolTipString
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

Private Sub cmd_quit_MouseEnter()
Help_txt.Caption = Cmd_Quit.ToolTipString
Help_txt.Refresh

End Sub

Private Sub Cmd_quit_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

'**********************************************************************
'按 ENTER调用 Mouse_click事件
'**********************************************************************
Private Sub Cmd_OK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_OK_MouseClick
End If
End Sub


Private Sub Cmd_cancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_cancel_MouseClick
End If
End Sub


Private Sub Cmd_add_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Add_MouseClick
End If
End Sub

Private Sub Cmd_edit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Edit_MouseClick
End If
End Sub

Private Sub Cmd_delete_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Delete_MouseClick
End If
End Sub

Private Sub Cmd_quit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_quit_MouseClick
End If
End Sub
