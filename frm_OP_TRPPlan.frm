VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_OP_TRPPlan 
   Caption         =   " 排  車  作  業"
   ClientHeight    =   7920
   ClientLeft      =   225
   ClientTop       =   900
   ClientWidth     =   15405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   15405
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   10800
      TabIndex        =   98
      Top             =   5520
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      StartOfWeek     =   50266113
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTAB1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "一般排車作業"
      TabPicture(0)   =   "frm_OP_TRPPlan.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fam_SrcOrders"
      Tab(0).Control(1)=   "fam_SelectedOrders"
      Tab(0).Control(2)=   "fam_RouteData"
      Tab(0).Control(3)=   "txtReceipt_No"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "路線編號列表"
      TabPicture(1)   =   "frm_OP_TRPPlan.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dg_Tab1_Route"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dg_Tab1_RouteOrders"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "保留訂單"
      TabPicture(2)   =   "frm_OP_TRPPlan.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmd_Tab2_ShowAll"
      Tab(2).Control(1)=   "cmd_Tab2_FilterAndSort"
      Tab(2).Control(2)=   "cmd_Tab2_Delete"
      Tab(2).Control(3)=   "cmd_Tab2_Reset"
      Tab(2).Control(4)=   "cmd_Tab2_Remove"
      Tab(2).Control(5)=   "dg_Tab2_ReservedOrders"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "中南區標貼列印"
      TabPicture(3)   =   "frm_OP_TRPPlan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "fma_Tab3_OrderSum"
      Tab(3).Control(2)=   "dg_RouteData"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame5 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   99
         Top             =   420
         Width           =   12465
         Begin VB.CommandButton cmd_Tab0_Print 
            BackColor       =   &H00C0FFC0&
            Caption         =   "列印BarCode"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   10200
            Picture         =   "frm_OP_TRPPlan.frx":0070
            Style           =   1  '圖片外觀
            TabIndex        =   108
            Top             =   120
            Width           =   1035
         End
         Begin VB.CheckBox Check1 
            Caption         =   "預覽列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   120
            TabIndex        =   107
            Top             =   2040
            Value           =   1  '核取
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "離  開"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Index           =   1
            Left            =   11280
            Picture         =   "frm_OP_TRPPlan.frx":17F2
            Style           =   1  '圖片外觀
            TabIndex        =   106
            Top             =   120
            Width           =   1065
         End
         Begin VB.TextBox DateS 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1155
            MaxLength       =   8
            TabIndex        =   105
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox DateE 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2835
            MaxLength       =   8
            TabIndex        =   104
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton cmd_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "資料查詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   8040
            Picture         =   "frm_OP_TRPPlan.frx":1C34
            Style           =   1  '圖片外觀
            TabIndex        =   103
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Excel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "轉 Excel"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   9120
            Picture         =   "frm_OP_TRPPlan.frx":24FE
            Style           =   1  '圖片外觀
            TabIndex        =   102
            Top             =   120
            Width           =   1065
         End
         Begin VB.ListBox Storerkey 
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   720
            Sorted          =   -1  'True
            Style           =   1  '項目包含核取方塊
            TabIndex        =   101
            Top             =   720
            Width           =   2055
         End
         Begin VB.ListBox Area_Code 
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   3600
            Style           =   1  '項目包含核取方塊
            TabIndex        =   100
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label sumlab 
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   114
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Lab_Storerkey 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   113
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "日期格式：yyyymmdd"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   24
            Left            =   4200
            TabIndex        =   112
            Top             =   360
            Width           =   2010
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "出車日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   25
            Left            =   135
            TabIndex        =   111
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "～"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   26
            Left            =   2565
            TabIndex        =   110
            Top             =   330
            Width           =   240
         End
         Begin VB.Label Lab_Storerkey 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "區碼"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   0
            Left            =   3000
            TabIndex        =   109
            Top             =   720
            Width           =   480
         End
      End
      Begin VB.TextBox txtReceipt_No 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   -64200
         TabIndex        =   95
         Top             =   3240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmd_Tab2_ShowAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "載入全部訂單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":30C0
         Style           =   1  '圖片外觀
         TabIndex        =   94
         Top             =   480
         Width           =   1320
      End
      Begin VB.CommandButton cmd_Tab2_FilterAndSort 
         BackColor       =   &H00C0E0FF&
         Caption         =   "訂單搜尋"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":33CA
         Style           =   1  '圖片外觀
         TabIndex        =   93
         Top             =   1440
         Width           =   1320
      End
      Begin VB.CommandButton cmd_Tab2_Delete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "TMS單號刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":3C94
         Style           =   1  '圖片外觀
         TabIndex        =   91
         ToolTipText     =   "刪除"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Frame fma_Tab3_OrderSum 
         Appearance      =   0  '平面
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74640
         TabIndex        =   80
         Top             =   600
         Visible         =   0   'False
         Width           =   6360
         Begin VB.CommandButton cmd_Tab3_Query 
            BackColor       =   &H00C0E0FF&
            Caption         =   "查  詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3840
            Style           =   1  '圖片外觀
            TabIndex        =   86
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "離  開"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   2
            Left            =   5040
            Style           =   1  '圖片外觀
            TabIndex        =   85
            Top             =   675
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Cancel 
            BackColor       =   &H00C0FFFF&
            Caption         =   "取  消"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3840
            Style           =   1  '圖片外觀
            TabIndex        =   84
            Top             =   675
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_Start 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1020
            TabIndex        =   83
            Top             =   270
            Width           =   1125
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_End 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2415
            TabIndex        =   82
            Top             =   270
            Width           =   1125
         End
         Begin VB.CommandButton cmd_Tab3_Excel 
            BackColor       =   &H00FFFF80&
            Caption         =   "轉 Exccel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   5040
            Style           =   1  '圖片外觀
            TabIndex        =   81
            Top             =   195
            Width           =   1200
         End
         Begin MSDataGridLib.DataGrid gd_Tab3_OrderSum 
            Height          =   2025
            Left            =   120
            TabIndex        =   89
            Top             =   1440
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3572
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "訂單數量分析"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gd_Tab3_Trp02wSum 
            Height          =   2025
            Left            =   120
            TabIndex        =   90
            Top             =   3840
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3572
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "未排車訂單分析"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "送貨日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   88
            Top             =   315
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "～"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   31
            Left            =   2175
            TabIndex        =   87
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.CommandButton cmd_Tab2_Reset 
         Appearance      =   0  '平面
         BackColor       =   &H00808080&
         Caption         =   "訂單排序"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '圖片外觀
         TabIndex        =   72
         Top             =   4800
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton cmd_Tab2_Remove 
         BackColor       =   &H00C0FFC0&
         Caption         =   "移至待排車"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":3F9E
         Style           =   1  '圖片外觀
         TabIndex        =   71
         Top             =   2400
         Width           =   1320
      End
      Begin MSDataGridLib.DataGrid dg_Tab2_ReservedOrders 
         Height          =   6570
         Left            =   -74895
         TabIndex        =   70
         Top             =   390
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   11589
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '平面
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   9375
         TabIndex        =   56
         Top             =   390
         Width           =   1995
         Begin VB.CommandButton cmd_Tab1_RouteNoQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "路線編號查詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   105
            Picture         =   "frm_OP_TRPPlan.frx":42A8
            Style           =   1  '圖片外觀
            TabIndex        =   58
            Top             =   975
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab1_RouteNo 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   180
            MaxLength       =   10
            TabIndex        =   57
            Top             =   525
            Width           =   1605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Left            =   465
            TabIndex        =   59
            Top             =   195
            Width           =   1020
         End
      End
      Begin VB.Frame fam_RouteData 
         Height          =   585
         Left            =   -74895
         TabIndex        =   2
         Top             =   405
         Width           =   12540
         Begin VB.CommandButton cmd_Tab0_Clear 
            BackColor       =   &H008080FF&
            Caption         =   "清除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3210
            Style           =   1  '圖片外觀
            TabIndex        =   78
            Top             =   75
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmd_Tab0_Save 
            BackColor       =   &H00FF8080&
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2595
            Style           =   1  '圖片外觀
            TabIndex        =   77
            Top             =   75
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmd_Tab0_Query 
            Appearance      =   0  '平面
            BackColor       =   &H00808000&
            Caption         =   "查詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1980
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   76
            Top             =   75
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txt_Tab0_RouteNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   540
            TabIndex        =   74
            Top             =   150
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox txt_Tab0_CarCheckInDate 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7380
            TabIndex        =   67
            Top             =   135
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_CarCheckInTime 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9285
            MaxLength       =   4
            TabIndex        =   64
            Top             =   135
            Width           =   750
         End
         Begin VB.TextBox txt_Tab0_DockNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5475
            TabIndex        =   62
            Top             =   135
            Width           =   1155
         End
         Begin VB.CommandButton cmd_Tab0_SelectedRemove_All 
            BackColor       =   &H000080FF&
            Caption         =   "已選訂單移除(全)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3840
            Style           =   1  '圖片外觀
            TabIndex        =   53
            Top             =   60
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_CreateRoute 
            Appearance      =   0  '平面
            BackColor       =   &H00FF8080&
            Caption         =   "建立路線編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10080
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   52
            Top             =   75
            Width           =   1110
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00400040&
            FillStyle       =   0  '實心
            Height          =   540
            Index           =   2
            Left            =   3840
            Top             =   45
            Width           =   1125
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00400040&
            FillStyle       =   0  '實心
            Height          =   540
            Index           =   1
            Left            =   1950
            Top             =   45
            Width           =   1860
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   3
            Left            =   45
            Top             =   120
            Width           =   1890
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   21
            Left            =   105
            TabIndex        =   75
            Top             =   150
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "預計報到日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   20
            Left            =   6720
            TabIndex        =   68
            Top             =   135
            Width           =   675
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   2
            Left            =   6675
            Top             =   105
            Width           =   1875
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   1
            Left            =   8565
            Top             =   105
            Width           =   1500
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   0
            Left            =   4980
            Top             =   105
            Width           =   1680
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "預計報到時間"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   19
            Left            =   8610
            TabIndex        =   63
            Top             =   135
            Width           =   675
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "碼頭暫存"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   18
            Left            =   5040
            TabIndex        =   61
            Top             =   135
            Width           =   435
         End
      End
      Begin VB.Frame fam_SelectedOrders 
         Height          =   3315
         Left            =   -74880
         TabIndex        =   22
         Top             =   900
         Width           =   12585
         Begin VB.CheckBox chk_Tab0_Updatetrpw 
            Caption         =   "更新"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   405
            Left            =   1200
            TabIndex        =   116
            ToolTipText     =   "載入同時更新板箱材重"
            Top             =   150
            Width           =   735
         End
         Begin VB.CommandButton cmd_Tab0_CreateRouteByAds 
            Appearance      =   0  '平面
            BackColor       =   &H00FFFF00&
            Caption         =   "  依地址  組路編"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11520
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   97
            Top             =   120
            Width           =   990
         End
         Begin VB.CheckBox ck_All 
            BackColor       =   &H80000012&
            Caption         =   "全選"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   495
            Left            =   12480
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   96
            Top             =   2880
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H008080FF&
            Caption         =   "離  開"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   10440
            Style           =   1  '圖片外觀
            TabIndex        =   79
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton cmd_Tab0_ImportOrders 
            BackColor       =   &H00C0C0FF&
            Caption         =   "載入待排車訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   30
            Style           =   1  '圖片外觀
            TabIndex        =   73
            Top             =   105
            Width           =   1095
         End
         Begin VB.CommandButton cmd_Tab0_Reserve 
            BackColor       =   &H00FF8080&
            Caption         =   "保留訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7950
            Style           =   1  '圖片外觀
            TabIndex        =   69
            Top             =   2880
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "待選取消"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9000
            Style           =   1  '圖片外觀
            TabIndex        =   24
            Top             =   3000
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.CheckBox chk_Tab0_DriveTimes 
            Caption         =   "顯示車次"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   405
            Left            =   5625
            TabIndex        =   60
            Top             =   150
            Width           =   750
         End
         Begin VB.CommandButton cmd_Tab0_srcOrderReset 
            Appearance      =   0  '平面
            BackColor       =   &H00808080&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11970
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   55
            Top             =   2910
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab0_srcOrderQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "訂單搜尋"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10845
            Style           =   1  '圖片外觀
            TabIndex        =   54
            Top             =   2910
            Width           =   1110
         End
         Begin VB.TextBox txt_Tab0_DeliveryPhone 
            Appearance      =   0  '平面
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   8325
            TabIndex        =   48
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab0_DeliveryCompany 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   6360
            TabIndex        =   46
            Top             =   315
            Width           =   825
         End
         Begin VB.TextBox txt_Tab0_DeliveryDriver 
            Appearance      =   0  '平面
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   7170
            TabIndex        =   44
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarType 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   9480
            TabIndex        =   42
            Top             =   315
            Width           =   945
         End
         Begin VB.CommandButton cmd_Tab0_SelectCar 
            BackColor       =   &H00FFC0C0&
            Caption         =   "？"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5190
            Style           =   1  '圖片外觀
            TabIndex        =   41
            Top             =   150
            Width           =   330
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4080
            TabIndex        =   40
            Top             =   150
            Width           =   1125
         End
         Begin VB.TextBox txt_Tab0_TRPDate 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2445
            TabIndex        =   38
            Top             =   150
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_Selected 
            BackColor       =   &H00FF8080&
            Caption         =   "↑"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7095
            Style           =   1  '圖片外觀
            TabIndex        =   36
            Top             =   2910
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab0_Remove 
            BackColor       =   &H008080FF&
            Caption         =   "↓"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7455
            Style           =   1  '圖片外觀
            TabIndex        =   35
            Top             =   2880
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel_All 
            BackColor       =   &H00FF80FF&
            Caption         =   "待選取消(全)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            Style           =   1  '圖片外觀
            TabIndex        =   23
            Top             =   2880
            Width           =   1530
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_SelectedOrders 
            Height          =   2265
            Left            =   0
            TabIndex        =   25
            Top             =   600
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   3995
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame1 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   26
            Top             =   2820
            Width           =   6915
            Begin VB.TextBox txt_Tab0_Selected_OTqty 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   990
               TabIndex        =   118
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   2310
               TabIndex        =   30
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Pallet 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   3540
               TabIndex        =   29
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Volumn 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   4785
               TabIndex        =   28
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Weight 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   6015
               TabIndex        =   27
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   22
               Left            =   1920
               TabIndex        =   117
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "重量"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   4
               Left            =   5640
               TabIndex        =   34
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "材積"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   5
               Left            =   4395
               TabIndex        =   33
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "板數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   6
               Left            =   3165
               TabIndex        =   32
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "累計：件數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   7
               Left            =   75
               TabIndex        =   31
               Top             =   210
               Width           =   900
            End
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00FF0000&
            FillStyle       =   0  '實心
            Height          =   435
            Left            =   10815
            Top             =   2880
            Width           =   2880
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00400040&
            FillStyle       =   0  '實心
            Height          =   435
            Index           =   0
            Left            =   7920
            Top             =   2880
            Width           =   2790
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '不透明
            Height          =   435
            Left            =   7050
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電  話"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   8685
            TabIndex        =   49
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運輸公司"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   6360
            TabIndex        =   47
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "駕駛人"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   15
            Left            =   7455
            TabIndex        =   45
            Top             =   120
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車   種"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   9630
            TabIndex        =   43
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "車牌號碼"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   13
            Left            =   3630
            TabIndex        =   39
            Top             =   165
            Width           =   420
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "出車日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   12
            Left            =   2010
            TabIndex        =   37
            Top             =   150
            Width           =   435
         End
      End
      Begin VB.Frame fam_SrcOrders 
         Height          =   2955
         Left            =   -74880
         TabIndex        =   1
         Top             =   4200
         Width           =   12540
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   525
            Left            =   6960
            TabIndex        =   13
            Top             =   0
            Width           =   6915
            Begin VB.TextBox txt_Tab0_srcTotal_OTqty 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   990
               TabIndex        =   121
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Weight 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   6000
               TabIndex        =   17
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Volumn 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   4785
               TabIndex        =   16
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Pallet 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   3540
               TabIndex        =   15
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   2295
               TabIndex        =   14
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   27
               Left            =   1920
               TabIndex        =   122
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "總計：件數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   11
               Left            =   75
               TabIndex        =   21
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "板數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   10
               Left            =   3165
               TabIndex        =   20
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "材積"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   9
               Left            =   4395
               TabIndex        =   19
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "重量"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   8
               Left            =   5625
               TabIndex        =   18
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame fam_SelectedSum 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   6915
            Begin VB.TextBox txt_Tab0_srcSelected_OTqty 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   975
               TabIndex        =   119
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   2310
               TabIndex        =   7
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Pallet 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   3540
               TabIndex        =   6
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Volumn 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   4785
               TabIndex        =   5
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Weight 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   6015
               TabIndex        =   4
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   23
               Left            =   1920
               TabIndex        =   120
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "重量"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   3
               Left            =   5640
               TabIndex        =   11
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "材積"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   2
               Left            =   4395
               TabIndex        =   10
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "板數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3165
               TabIndex        =   9
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "選取：件數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   8
               Top             =   210
               Width           =   900
            End
         End
         Begin MSDataGridLib.DataGrid dg_TRP02W 
            Height          =   2280
            Left            =   45
            TabIndex        =   12
            Top             =   525
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   4022
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_RouteOrders 
         Height          =   2640
         Left            =   90
         TabIndex        =   50
         Top             =   4485
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   4657
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_Route 
         Height          =   4065
         Left            =   90
         TabIndex        =   51
         Top             =   390
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   7170
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '平面
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   9405
         TabIndex        =   65
         Top             =   2355
         Width           =   1980
         Begin VB.CommandButton cmdDeliveryDateFix 
            BackColor       =   &H000080FF&
            Caption         =   "到貨時間安排"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   92
            ToolTipText     =   "訂單有備註時"
            Top             =   240
            Width           =   1785
         End
         Begin VB.CommandButton cmd_Tab1_RouteNoDelete 
            BackColor       =   &H00C0C0FF&
            Caption         =   "路線編號刪除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   105
            Picture         =   "frm_OP_TRPPlan.frx":45B2
            Style           =   1  '圖片外觀
            TabIndex        =   66
            ToolTipText     =   "刪除"
            Top             =   1200
            Width           =   1785
         End
      End
      Begin MSDataGridLib.DataGrid dg_RouteData 
         Height          =   4470
         Left            =   -74850
         TabIndex        =   115
         Top             =   2880
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   7885
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_OP_TRPPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬
Private rs_Access As ADODB.Recordset         '報表列印用 >> 轉資料至 Access DB
Private MSAccessAP As access.Application
Private blTRP02WEventEnable As Boolean              '待選取訂單 Event 觸發有效控制
Private blTab0SelectedOrderEventEnable As Boolean   '已選取訂單 Event 觸發有效控制
Private blTab1RouteEventEnable As Boolean           '路線編號列表 Event 觸發有效控制
Private blTab2ReservedEventEnable As Boolean        '保留訂單列表 Event 觸發有效控制

Private blRouteModify As Boolean                    '排車作業 >> 路線編號 查詢：有效路線編號
Private blRouteChange As Boolean                    '排車作業 >> 路線編號 資料異動識別旗標
Private strDispRouteNo As String                    '排車作業 >> 路線編號 查詢：路線編號

Private rs_TRP02W As ADODB.Recordset                '排車作業：匯入之待排車訂單
Private rs_Tab0_SelectedOrders As ADODB.Recordset   '排車作業：已選取之待排車訂單
Private rs_Tab1_Route As ADODB.Recordset            '路編列表：路線編號列表
Private rs_Tab1_RouteOrders As ADODB.Recordset      '路編列表：路線編號所屬之訂單
Private rs_Tab2_ReservedOrders As ADODB.Recordset   '保留訂單
Private rs_Tab3_OrderSum As ADODB.Recordset         '訂單分析：依地區統計
Private rs_Tab3_Trp02wSum As ADODB.Recordset        '未排車訂單分析：依地區統計
Private rs_RouteData As ADODB.Recordset        '一次路線編號標貼


Private strSourceFilter As String        '待排車訂單篩選
Private strSourceOrderBy As String       '待排車訂單排序方式
Private dbsrcSelected_OTqty As Double     '待排車訂單: 選取件數
Private dbsrcSelected_Case As Double     '待排車訂單: 選取箱數
Private dbsrcSelected_Pallet As Double   '待排車訂單: 選取板數
Private dbsrcSelected_Volumn As Double   '待排車訂單: 選取材積
Private dbsrcSelected_Weight As Double   '待排車訂單: 選取重量
Private dbSelectedCount As Double        '選取訂單筆數
Private DelRecord

Private Sub ck_All_Click()
'On Error GoTo errer_handle
'Dim i As Integer
'rs_TRP02W.MoveFirst
'    For i = 0 To rs_TRP02W.RecordCount - 1
'        Call dg_TRP02W_RowColChange(i, 1)
'        rs_TRP02W.MoveNext
'    Next
'errer_handle:
End Sub

Private Sub cmd_Excel_Click()

    If rs_RouteData Is Nothing Then Exit Sub
    If rs_RouteData.RecordCount = 0 Then Exit Sub
    
    Recordset2Excel "一次路編資料", rs_RouteData

    '..在此編輯EXCEL
    With MyXlsApp
    End With
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
    '離開
    Unload Me
End Sub

Private Sub cmd_Tab0_AddressRoute_Click()

End Sub

Private Sub cmd_Query_Click()
    
    Screen.MousePointer = vbHourglass
    Dim str_Where As String, strTmp As String, i As Integer
    str_Where = ""
    
    '出貨日
    If (Len(RTrim(DateS.Text)) > 0 And Len(RTrim(DateE.Text)) = 0) Or (Len(RTrim(DateS.Text)) = 0 And Len(RTrim(DateE.Text)) > 0) Then str_Where = str_Where & " and convert(char(8),t1.delivery_date,112) = '" & RTrim(DateS.Text) & RTrim(DateE.Text) & "' "
    If (Len(RTrim(DateS.Text)) > 0 And Len(RTrim(DateE.Text)) > 0) Then str_Where = " and convert(char(8),t1.delivery_date,112)  between '" & RTrim(DateS.Text) & "'and'" & RTrim(DateE.Text) & "' "

    '貨主
    strTmp = ""
    For i = 0 To Storerkey.ListCount - 1
        If Storerkey.Selected(i) Then
                strTmp = strTmp & "'" & Storerkey.List(i) & "',"
        End If
    Next
    
    If Len(strTmp) > 0 Then str_Where = str_Where & " and o.storerkey in (" & Mid(strTmp, 1, Len(strTmp) - 1) & ")"

    '區碼
    strTmp = ""
    For i = 0 To Area_Code.ListCount - 1
        If Area_Code.Selected(i) Then
                strTmp = strTmp & "'" & Area_Code.List(i) & "',"
        End If
    Next
    
    If Len(strTmp) > 0 Then str_Where = str_Where & " and t1m.area_code in (" & Mid(strTmp, 1, Len(strTmp) - 1) & ")"
    '20160810 修改板數無條件進位
    str_SQL = "select " & _
            "一次駕駛人=isnull(t9.driver,''), " & _
            "出車日=convert(char(8),t1.delivery_date,112), " & _
            "一次路線編號=t2.route_no, " & _
            "客戶簡稱=t1m.short_name, " & _
            " case when (ceiling(sum(t2.pallet_qty)))>=1 then ceiling(sum(t2.pallet_qty)) else 1 end as 板數, " & _
            "箱數=sum(t2.case_cnt) " & _
            "from orders o join trp02t t2 on o.orderkey = t2.c_receipt_no " & _
            "join trp01t t1 on t1.route_no = t2.route_no " & _
            "join trp09m t9 on t9.vehicle_id_no = t2.vehicle_id_no " & _
            "left join trp01m t1m on o.storerkey = t1m.storerkey and t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end " & _
            "where left(rtrim(t1.route_no),1)='F' and 1=1 " & str_Where & _
            "group by isnull(t9.driver,''),convert(char(8),t1.delivery_date,112),t2.route_no,t1m.short_name " & _
            "order by t2.route_no"
      
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無符合之路線編號資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Set dg_RouteData.DataSource = Nothing
       Exit Sub
    End If
    
    Call ReDim_Recordset(rs_RouteData)
    Call Replication_Recordset(tmp_Rs, rs_RouteData)
    tmp_Rs.Close
    
    With dg_RouteData
         .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
         .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
         .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
         .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
    End With
    rs_RouteData.MoveFirst
    Set dg_RouteData.DataSource = rs_RouteData
    With dg_RouteData
        .RowHeight = 250
        .Columns(0).Width = 500        '序號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1200        '一次駕駛人
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1500        '出貨日
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1200        '路線編號
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 2000       '客戶名稱
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000       '板數
        .Columns(5).Alignment = dbgRight
        .Columns(6).Width = 1000        '箱數
        .Columns(6).Alignment = dbgRight
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Clear_Click()
    '排車作業 >> 清除
    If rs_Tab0_SelectedOrders.RecordCount <> 0 And blRouteModify = False Then
        '新增路線編號模式：
        '呼叫 [已選訂單移除(全)] 來處理已被暫時選取之 [待排車訂單] 還原回 [待排車訂單]
        Call cmd_Tab0_SelectedRemove_All_Click
    End If
    If blRouteModify And blRouteChange Then
       '有效路線編號 & 資料已遭異動，要 user 確認是否存檔
        msg_text = "路線編號資料是否存檔？"
        If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
            '呼叫存檔程序
            Call cmd_Tab0_Save_Click
        Else
            '不存檔→必須重新載入 [待排車訂單] 已還原 [選取][移除] 操作對 [待排車訂單] 的影響
            Call cmd_Tab0_ImportOrders_Click
        End If
    End If
    '清除路線編號欄位值，包含已選訂單名細列表
    Call Clear_RouteData
    txt_Tab0_RouteNo.Text = ""
End Sub

Private Sub cmd_Tab0_CreateRoute_Click()
    '排車作業 >> 建立路線編號
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then
        msg_text = "資料錯誤：無裝載資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    'add by Terry 20190614
    Dim str_CheckReceipt_No As String
    str_CheckReceipt_No = ""
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        str_CheckReceipt_No = str_CheckReceipt_No & "'" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "',"
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
    str_CheckReceipt_No = str_CheckReceipt_No & "''"
    rs_Tab0_SelectedOrders.MoveFirst
    str_SQL = "select receipt_no  from trp02t where receipt_no in (" & str_CheckReceipt_No & ")"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        MsgBox ("有訂單已組成一次路編，請重新載入待排車訂單並清空[已選取的一次訂單]"), vbOKOnly + vbCritical
        tmp_Rs.Close
        Exit Sub
    End If
    tmp_Rs.Close
    
    
    
    '檢核路線編號資料是否正確，錯誤將在 Function 直接顯示 MessageBox
    If RouteData_Check() = False Then Exit Sub
    
    On Error GoTo err_Handle
    
    cmd_Tab0_CreateRoute.Enabled = False
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    '檢查可載重量
    Dim intableWT, intableCBM As Long
    str_SQL = "select rtrim(isnull(LOADING_SIZE,0)),rtrim(isnull(MAX_CUBIC_CAPACITY,0)) from dbo.TRP09M where VEHICLE_ID_NO='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intableWT = tmp_Rs.Fields(0).Value
    intableCBM = tmp_Rs.Fields(1).Value
    tmp_Rs.Close
    If intableWT < Val(txt_Tab0_Selected_Weight.Text) Then
        msg_text = "排車重量超過車輛可載重,車輛可載重:" & intableWT
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        cmd_Tab0_CreateRoute.Enabled = True
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    If intableCBM < Val(txt_Tab0_Selected_Volumn.Text) Then
        msg_text = "排車重量超過車輛可載材積,車輛可載材積:" & intableCBM
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        cmd_Tab0_CreateRoute.Enabled = True
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Dim intDriveTimes As Integer    '車次
    Dim strRouteNo As String        '路線編號
    
    '1.產生車次
    str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
              "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    tmp_Rs.Close
    
    '2.產生路線編號
    str_SQL = "Select Isnull(Max(Cast(Right(Route_No,3) as integer))+1,1) as RouteSN " & _
              "From TRP01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'F'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strRouteNo = "F" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
    tmp_Rs.Close
    
    '3.Insert into TRP01T 路線編號主檔
    '  TRP01T.EXE_CONFIRM = '0' 新產生路線編號，尚未回傳過 exe
    str_SQL = "Insert into TRP01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
              strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
              txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '4.insert into TRP05T 車輛進出管理
    str_SQL = "Insert into TRP05T (Route_No,Vehicle_ID_No,Drive_Times,Delivery_Date,Valid_Vehicle,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Dock_No,Expect_Time,Expect_Date) Values ('" & _
              strRouteNo & "','" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",'" & _
              Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "','1'," & _
              txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'" & _
              txt_Tab0_DockNo.Text & "','" & txt_Tab0_CarCheckInTime.Text & "','" & txt_Tab0_CarCheckInDate.Text & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '由車輛主檔更新車輛相關欄位
    str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strRouteNo & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '寫至 SSTab1.Tab 1 [路線編號列表]
    blTab1RouteEventEnable = False
    rs_Tab1_Route.AddNew
    rs_Tab1_Route.Fields("編號").Value = rs_Tab1_Route.RecordCount
    rs_Tab1_Route.Fields("路線編號").Value = strRouteNo
    rs_Tab1_Route.Fields("出車日期").Value = txt_Tab0_TRPDate.Text
    rs_Tab1_Route.Fields("車牌號碼").Value = txt_Tab0_DeliveryCarNo.Text
    rs_Tab1_Route.Fields("車次").Value = intDriveTimes
    rs_Tab1_Route.Fields("駕駛人").Value = txt_Tab0_DeliveryDriver.Text
    rs_Tab1_Route.Fields("件數").Value = txt_Tab0_Selected_OTqty.Text
    rs_Tab1_Route.Fields("箱數").Value = txt_Tab0_Selected_Case.Text
    rs_Tab1_Route.Fields("板數").Value = txt_Tab0_Selected_Pallet.Text
    rs_Tab1_Route.Fields("材積").Value = txt_Tab0_Selected_Volumn.Text
    rs_Tab1_Route.Fields("重量").Value = txt_Tab0_Selected_Weight.Text
    rs_Tab1_Route.Fields("車種").Value = txt_Tab0_DeliveryCarType.Text
    rs_Tab1_Route.Fields("碼頭暫存").Value = txt_Tab0_DockNo.Text
    rs_Tab1_Route.Fields("預計報到日期").Value = txt_Tab0_CarCheckInDate.Text
    rs_Tab1_Route.Fields("預計報到時間").Value = txt_Tab0_CarCheckInTime.Text
    rs_Tab1_Route.Fields("EXE回傳").Value = "新建路編"
    rs_Tab1_Route.Fields("排車者").Value = User_id
    rs_Tab1_Route.Update
    blTab1RouteEventEnable = True
    
    '5.insert into TRP02T [排車訂單檔]
    '  寫至 SSTab1.Tab 1 [路線編號之訂單名細表]
    blTab0SelectedOrderEventEnable = False
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        'insert into TRP02T
        str_SQL = "Insert into TRP02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                  "Select '" & strRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                  "From TRP02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
'        '由Orders更新件數，未開始使用因為有可能遇到拆單
'        str_SQL = "update trp02t set trp02t.otqty = orders.otqty from trp02t join orders on trp02t.receipt_no = orders.orderkey and trp02t.receipt_no = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '寫至 SSTab1.Tab 1 [路線編號之訂單名細表]
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("編號").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("路線編號").Value = strRouteNo
        rs_Tab1_RouteOrders.Fields("送貨日").Value = rs_Tab0_SelectedOrders.Fields("送貨日").Value
        rs_Tab1_RouteOrders.Fields("訂單編號").Value = rs_Tab0_SelectedOrders.Fields("訂單編號").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = rs_Tab0_SelectedOrders.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("客戶名稱").Value = rs_Tab0_SelectedOrders.Fields("客戶名稱").Value
        rs_Tab1_RouteOrders.Fields("件數").Value = rs_Tab0_SelectedOrders.Fields("件數").Value
        rs_Tab1_RouteOrders.Fields("箱數").Value = rs_Tab0_SelectedOrders.Fields("箱數").Value
        rs_Tab1_RouteOrders.Fields("板數").Value = rs_Tab0_SelectedOrders.Fields("板數").Value
        rs_Tab1_RouteOrders.Fields("材積").Value = rs_Tab0_SelectedOrders.Fields("材積").Value
        rs_Tab1_RouteOrders.Fields("重量").Value = rs_Tab0_SelectedOrders.Fields("重量").Value
        rs_Tab1_RouteOrders.Fields("訂單備註").Value = rs_Tab0_SelectedOrders.Fields("訂單備註").Value
        rs_Tab1_RouteOrders.Fields("車種").Value = rs_Tab0_SelectedOrders.Fields("車種").Value
        rs_Tab1_RouteOrders.Fields("特殊需求1").Value = rs_Tab0_SelectedOrders.Fields("特殊需求1").Value
        rs_Tab1_RouteOrders.Fields("特殊需求2").Value = rs_Tab0_SelectedOrders.Fields("特殊需求2").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE回傳").Value = rs_Tab0_SelectedOrders.Fields("EXE回傳").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = rs_Tab0_SelectedOrders.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("型態").Value = rs_Tab0_SelectedOrders.Fields("型態").Value
        rs_Tab1_RouteOrders.Fields("客戶簡稱").Value = rs_Tab0_SelectedOrders.Fields("客戶簡稱").Value
        rs_Tab1_RouteOrders.Update
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
    '確認路線編號的 exe_confirm 狀態
    '主要目的：已回傳之路編刪除後，重新產生之路編，若全部都是已回傳訂單，直接路編設定為 [已回傳]
'Mark by Gemini @20111010
'    str_SQL = "Update TRP01T Set EXE_Confirm = (Select min(EXE_Confirm) From TRP02T Where TRP02T.Route_No = TRP01T.Route_No) " & _
'              "Where TRP01T.Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans
    Tran_Level = 0
    
    If dg_Tab1_Route.SelBookmarks.Count > 0 Then
        dg_Tab1_Route.SelBookmarks.Remove 0
    End If
    dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
    rs_Tab1_RouteOrders.Filter = " 路線編號 = '" & strRouteNo & "' "
    blTab0SelectedOrderEventEnable = True
    
    '4.由 TRP02T Trigger [insert] 進行以下作業
    '   a.寫入 TRP03T -- 排車訂單明細檔
    '   b.刪除 TRP03W -- 待排車訂單明細檔
    '   c.刪除 TRP02W -- 待排車訂單主檔
    
    
    '5.清除 [已選取之待排車訂單列表]
    blTab0SelectedOrderEventEnable = False
    '排車作業：已選取之待排車訂單列表 DBGrid 格式設定-ReSet
    Call CreateRS_Tab0_SelectedOrders
    '重新計算已選取訂單：箱數，板數，材積，重量 + 編號重新產生
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    '6.清除排車作業欄位值
    txt_Tab0_DockNo.Text = ""               '碼頭暫存
    txt_Tab0_CarCheckInDate.Text = ""       '車輛預計報到日期
    txt_Tab0_CarCheckInTime.Text = ""       '車輛預計報到時間
    txt_Tab0_TRPDate.Text = ""              '出車日期
    txt_Tab0_DeliveryCarNo.Text = ""        '車牌號碼
    txt_Tab0_DeliveryCompany.Text = ""      '運輸公司
    txt_Tab0_DeliveryDriver.Text = ""       '駕駛人
    txt_Tab0_DeliveryPhone.Text = ""        '電話
    txt_Tab0_DeliveryCarType.Text = ""      '車種
    
    cmd_Tab0_CreateRoute.Enabled = True
    
    'Call cmd_Tab0_ImportOrders_Click 'edit by Eric 20140729，避免重複篩選特定訂單，雅筑需求
    
    '待排車訂單總計資訊
    Call ReCaculate_OrderSum
    
    SSTab1.Tab = 1
    DoEvents: DoEvents
    '安排到貨時間
    Call cmdDeliveryDateFix_Click
    
    
On Error GoTo err_Handle2
    'Terry 20200212 排車資料轉入BestAPP 觸發推播功能 過度期使用
    cn.Execute "exec Andys_BestTMSOrderImport", RowsAffect, adExecuteNoRecords
    Dim HttpClient As Object

    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/InsertWaybillList", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    
    Exit Sub

err_Handle2:
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
   End If
   
   '遭遇錯誤的話：local 的 Recordset [路線編號列表] 資料必須刪除
   '因為 [路線編號列表] 不受 DB connection.transaction 控制
   blTab1RouteEventEnable = False
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   rs_Tab1_Route.Filter = "路線編號='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   rs_Tab1_RouteOrders.Filter = "路線編號='" & strRouteNo & "'"
   If Not rs_Tab1_RouteOrders.EOF Then
      Do While Not rs_Tab1_RouteOrders.EOF
         rs_Tab1_RouteOrders.Delete
         rs_Tab1_RouteOrders.MoveFirst
      Loop
   End If
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
      
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車作業-建立路線編號", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_CreateRoute.Enabled = True
End Sub

Private Sub cmd_Tab0_CreateRouteByAds_Click()
   '排車作業 >> 建立路線編號
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then
        msg_text = "資料錯誤：無裝載資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    '檢核路線編號資料是否正確，錯誤將在 Function 直接顯示 MessageBox
    If RouteData_Check() = False Then Exit Sub
    
    On Error GoTo err_Handle
    
    cmd_Tab0_CreateRouteByAds.Enabled = False
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    '檢查可載重量,虛排車不檢查
'    Dim intableWT, intableCBM As Long
'    str_SQL = "select rtrim(isnull(LOADING_SIZE,0)),rtrim(isnull(MAX_CUBIC_CAPACITY,0)) from dbo.TRP09M where VEHICLE_ID_NO='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "'"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    intableWT = tmp_Rs.Fields(0).Value
'    intableCBM = tmp_Rs.Fields(1).Value
'    tmp_Rs.Close
'    If intableWT < Val(txt_Tab0_Selected_Weight.Text) Then
'        msg_text = "排車重量超過車輛可載重,車輛可載重:" & intableWT
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        cmd_Tab0_CreateRoute.Enabled = True
'        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
'        txt_Tab0_DeliveryCarNo.SetFocus
'        Exit Sub
'    End If
'    If intableCBM < Val(txt_Tab0_Selected_Volumn.Text) Then
'        msg_text = "排車重量超過車輛可載材積,車輛可載材積:" & intableCBM
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        cmd_Tab0_CreateRoute.Enabled = True
'        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
'        txt_Tab0_DeliveryCarNo.SetFocus
'        Exit Sub
'    End If
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Dim intDriveTimes As Integer    '車次
    Dim strRouteNo As String        '路線編號
    Dim strAddress As String        '不同地址產生新的路線編號
    Dim strRouteNosum As String     '更新TRP01、TRP05
    strAddress = ""
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "Zip,運送地址"
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        
        If Trim(strAddress) <> Trim(rs_Tab0_SelectedOrders.Fields("運送地址").Value) Then '地址不一樣
            
            strAddress = Trim(rs_Tab0_SelectedOrders.Fields("運送地址").Value)
            '1.產生車次
            str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                      "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
            tmp_Rs.Close
            
            '2.產生路線編號
            str_SQL = "Select Isnull(Max(Cast(Right(Route_No,3) as integer))+1,1) as RouteSN " & _
                      "From TRP01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'F'"
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            strRouteNo = "F" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
            tmp_Rs.Close
            
            '紀錄本次產生的路邊
            If Len(strRouteNosum) = 0 Then strRouteNosum = "'" & strRouteNo & "'" Else strRouteNosum = strRouteNosum & ",'" & strRouteNo & "'"
            
            '3.Insert into TRP01T 路線編號主檔
            '  TRP01T.EXE_CONFIRM = '0' 新產生路線編號，尚未回傳過 exe
            str_SQL = "Insert into TRP01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
                      strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
                      txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '4.insert into TRP05T 車輛進出管理
            str_SQL = "Insert into TRP05T (Route_No,Vehicle_ID_No,Drive_Times,Delivery_Date,Valid_Vehicle,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Dock_No,Expect_Time,Expect_Date) Values ('" & _
                      strRouteNo & "','" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",'" & _
                      Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "','1'," & _
                      txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'" & _
                      txt_Tab0_DockNo.Text & "','" & txt_Tab0_CarCheckInTime.Text & "','" & txt_Tab0_CarCheckInDate.Text & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '由車輛主檔更新車輛相關欄位
            str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
                      "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strRouteNo & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
        End If
        
     '5.insert into TRP02T [排車訂單檔]
     '  寫至 SSTab1.Tab 1 [路線編號之訂單名細表]
       blTab0SelectedOrderEventEnable = False
        'insert into TRP02T
        str_SQL = "Insert into TRP02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                  "Select '" & strRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                  "From TRP02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
    '確認路線編號的 exe_confirm 狀態
    '主要目的：已回傳之路編刪除後，重新產生之路編，若全部都是已回傳訂單，直接路編設定為 [已回傳]
    
    '6. update trp01t,trp05t，
    str_SQL = "update TRP01T set WEIGHT=(select sum(TRP02T.WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where   route_no in ( " & strRouteNosum & ") " & _
        "update TRP01T set CASE_CNT=(select sum(TRP02T.CASE_CNT) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where  route_no in ( " & strRouteNosum & ") " & _
        "update TRP01T set Pallet_Qty=(select sum(TRP02T.Pallet_Qty) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") " & _
        "update TRP01T set VOLUMN_WEIGHT=(select sum(TRP02T.VOLUMN_WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "update TRP05T set WEIGHT=(select sum(TRP02T.WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where   route_no in ( " & strRouteNosum & ") " & _
        "update TRP05T set CASE_CNT=(select sum(TRP02T.CASE_CNT) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where  route_no in ( " & strRouteNosum & ") " & _
        "update TRP05T set Pallet_Qty=(select sum(TRP02T.Pallet_Qty) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") " & _
        "update TRP05T set VOLUMN_WEIGHT=(select sum(TRP02T.VOLUMN_WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    
    cn.CommitTrans
    Tran_Level = 0
    
    If dg_Tab1_Route.SelBookmarks.Count > 0 Then
        dg_Tab1_Route.SelBookmarks.Remove 0
    End If
'    dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
'    rs_Tab1_RouteOrders.Filter = " 路線編號 = '" & strRouteNo & "' "
    blTab0SelectedOrderEventEnable = True
    
    '7.由 TRP02T Trigger [insert] 進行以下作業
    '   a.寫入 TRP03T -- 排車訂單明細檔
    '   b.刪除 TRP03W -- 待排車訂單明細檔
    '   c.刪除 TRP02W -- 待排車訂單主檔
    
    '8.清除 [已選取之待排車訂單列表]
    blTab0SelectedOrderEventEnable = False
    '排車作業：已選取之待排車訂單列表 DBGrid 格式設定-ReSet
    Call CreateRS_Tab0_SelectedOrders
    '重新計算已選取訂單：箱數，板數，材積，重量 + 編號重新產生
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    '6.清除排車作業欄位值
    txt_Tab0_DockNo.Text = ""               '碼頭暫存
    txt_Tab0_CarCheckInDate.Text = ""       '車輛預計報到日期
    txt_Tab0_CarCheckInTime.Text = ""       '車輛預計報到時間
    txt_Tab0_TRPDate.Text = ""              '出車日期
    txt_Tab0_DeliveryCarNo.Text = ""        '車牌號碼
    txt_Tab0_DeliveryCompany.Text = ""      '運輸公司
    txt_Tab0_DeliveryDriver.Text = ""       '駕駛人
    txt_Tab0_DeliveryPhone.Text = ""        '電話
    txt_Tab0_DeliveryCarType.Text = ""      '車種
    
    cmd_Tab0_CreateRouteByAds.Enabled = True
    
    'Call cmd_Tab0_ImportOrders_Click
    
    '待排車訂單總計資訊
    Call ReCaculate_OrderSum
    
    SSTab1.Tab = 1
    DoEvents: DoEvents
    '安排到貨時間
    Call cmdDeliveryDateFix_Click
    
    '查詢排車結果
    '設定路線編號列表
    blTab1RouteEventEnable = False
    Call CreateRS_Tab1_Route
    blTab1RouteEventEnable = True
    '設定路線編號之訂單列表
    Call CreateRS_Tab1_RouteOrders
    
    str_SQL = "Select 路線編號,出車日期,車牌號碼,車次,駕駛人,件數,箱數,板數,材積,重量,車種,碼頭暫存,預計報到日期,預計報到時間,EXE回傳,排車者 " & _
              "From TRPPlan_RouteData Where 路線編號 in ( " & strRouteNosum & ") order by 路線編號"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定條件之路線編號資料(TRP01T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTab1RouteEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab1_Route.AddNew
        rs_Tab1_Route.Fields("編號").Value = rs_Tab1_Route.RecordCount
        rs_Tab1_Route.Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
        rs_Tab1_Route.Fields("出車日期").Value = tmp_Rs.Fields("出車日期").Value
        rs_Tab1_Route.Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
        rs_Tab1_Route.Fields("車次").Value = tmp_Rs.Fields("車次").Value
        rs_Tab1_Route.Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
        rs_Tab1_Route.Fields("件數").Value = tmp_Rs.Fields("件數").Value
        rs_Tab1_Route.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
        rs_Tab1_Route.Fields("板數").Value = tmp_Rs.Fields("板數").Value
        rs_Tab1_Route.Fields("材積").Value = tmp_Rs.Fields("材積").Value
        rs_Tab1_Route.Fields("重量").Value = tmp_Rs.Fields("重量").Value
        rs_Tab1_Route.Fields("車種").Value = tmp_Rs.Fields("車種").Value
        rs_Tab1_Route.Fields("碼頭暫存").Value = tmp_Rs.Fields("碼頭暫存").Value
        rs_Tab1_Route.Fields("預計報到日期").Value = tmp_Rs.Fields("預計報到日期").Value
        rs_Tab1_Route.Fields("預計報到時間").Value = tmp_Rs.Fields("預計報到時間").Value
        rs_Tab1_Route.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
        rs_Tab1_Route.Fields("排車者").Value = tmp_Rs.Fields("排車者").Value
        rs_Tab1_Route.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab1_Route.MoveFirst
    blTab1RouteEventEnable = True
    tmp_Rs.Close
    'TRP03W
    str_SQL = "Select 路線編號,送貨日,訂單編號,ZIP,客戶名稱,客戶地址,件數,箱數,板數,材積,重量,Receipt_No,EXE回傳,Area,型態,客戶簡稱,訂單備註 " & _
              "From TRPPlan_RouteOrders " & _
               "Where 路線編號 in ( " & strRouteNosum & ") Order by 路線編號,Receipt_No"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定路線編號之訂單資料(TRP02T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Do While Not tmp_Rs.EOF
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("編號").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
        rs_Tab1_RouteOrders.Fields("送貨日").Value = tmp_Rs.Fields("送貨日").Value
        rs_Tab1_RouteOrders.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("客戶名稱").Value = tmp_Rs.Fields("客戶名稱").Value
        rs_Tab1_RouteOrders.Fields("地址").Value = tmp_Rs.Fields("客戶地址").Value
        rs_Tab1_RouteOrders.Fields("件數").Value = tmp_Rs.Fields("件數").Value
        rs_Tab1_RouteOrders.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
        rs_Tab1_RouteOrders.Fields("板數").Value = tmp_Rs.Fields("板數").Value
        rs_Tab1_RouteOrders.Fields("材積").Value = tmp_Rs.Fields("材積").Value
        rs_Tab1_RouteOrders.Fields("重量").Value = tmp_Rs.Fields("重量").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("型態").Value = tmp_Rs.Fields("型態").Value
        rs_Tab1_RouteOrders.Fields("客戶簡稱").Value = tmp_Rs.Fields("客戶簡稱").Value
        rs_Tab1_RouteOrders.Fields("訂單備註").Value = tmp_Rs.Fields("訂單備註").Value
        rs_Tab1_RouteOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab1_RouteOrders.MoveFirst
    
    Screen.MousePointer = vbDefault
    
    
On Error GoTo err_Handle2
    'Terry 20200212 排車資料轉入BestAPP 觸發推播功能 過度期使用
    cn.Execute "exec Andys_BestTMSOrderImport", RowsAffect, adExecuteNoRecords
    Dim HttpClient As Object

    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/InsertWaybillList", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    Exit Sub
    
err_Handle2:
Exit Sub
    
err_Handle:
   If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
   End If
   
   '遭遇錯誤的話：local 的 Recordset [路線編號列表] 資料必須刪除
   '因為 [路線編號列表] 不受 DB connection.transaction 控制
   blTab1RouteEventEnable = False
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   rs_Tab1_Route.Filter = "路線編號='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   rs_Tab1_RouteOrders.Filter = "路線編號='" & strRouteNo & "'"
   If Not rs_Tab1_RouteOrders.EOF Then
      Do While Not rs_Tab1_RouteOrders.EOF
         rs_Tab1_RouteOrders.Delete
         rs_Tab1_RouteOrders.MoveFirst
      Loop
   End If
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
      
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車作業-依地址建立路線編號", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_CreateRouteByAds.Enabled = True
End Sub

Private Sub cmd_Tab0_ImportOrders_Click()
On Error GoTo err_Handle
Dim strReceiptNo As String
strReceiptNo = ""

    '更新箱板材重資料
    If chk_Tab0_Updatetrpw.Value = 1 Then
        cn.Execute "exec gs_UpdateTRPW", RowsAffect, adExecuteNoRecords
    End If
    
'    '更新Orders件數--疑似導致QTY=0的地方
'    str_SQL = "update trp02w set trp02w.otqty = orders.otqty from trp02w join orders on trp02w.receipt_no = orders.orderkey and trp02w.OTConfirmuser is null and trp02w.OTQTY is null and orders.OTQTY is not null "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '排車作業>>匯入待排車訂單
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_TRP02W.DataSource = Nothing
    '排車作業：待排車訂單
    Call CreateRS_Tab0_TRP02W
    
    strSourceFilter = adFilterNone
    DoEvents
    
    '有已選取訂單者：詢問 user 是否要清除
    If rs_Tab0_SelectedOrders.RecordCount <> 0 Then
       msg_text = "載入待排車訂單：[已選取訂單] 是否進行清除"
       If MsgBox(msg_text, vbYesNo + vbInformation + vbDefaultButton2, msg_title) = vbYes Then
          '清除路線編號欄位值，包含已選訂單名細列表
          Call Clear_RouteData
          txt_Tab0_RouteNo.Text = ""
        Else
            dg_Tab0_SelectedOrders.Enabled = False
            rs_Tab0_SelectedOrders.MoveFirst
            Do While Not rs_Tab0_SelectedOrders.EOF
                strReceiptNo = strReceiptNo & rs_Tab0_SelectedOrders.Fields("Receipt_no") & "','"
                rs_Tab0_SelectedOrders.MoveNext
            Loop
            
            dg_Tab0_SelectedOrders.Enabled = True
            
       End If
    End If
    
    'dg_Tab0_SelectedOrders
    
    '待排車訂單載入：選取小計：歸零
    dbSelectedCount = 0
    dbsrcSelected_OTqty = 0: dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_OTqty.Text = ""
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '取回待排車訂單
    str_SQL = "Select ' ' as '＊',送貨日,訂單編號,件數,箱數,板數,材積,重量,客戶編號,isnull(ZIP,'') as ZIP,isnull(客戶簡稱,'') as 客戶簡稱, " & _
              "isnull(Area,'') as Area ,isnull(型態,'') as 型態,isnull(運送地址,'') as 運送地址,訂單備註,啤酒允收期,配送倉別,isnull(車種,'') as 車種, " & _
              "isnull(特殊需求1,'') as 特殊需求1,isnull(特殊需求2,'') as 特殊需求2,急單,專車,冷藏,Receipt_No,C_Receipt_No,貨主單號,EXE回傳,isnull(客戶名稱,'') as 客戶名稱 " & _
              " From TRPPlan_SourceOrder " & _
              " where Receipt_No not in ( '" & strReceiptNo & " ') " & _
              " Order by 訂單編號 "
    strSourceOrderBy = " 訂單編號 "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定條件之待排車訂單資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    Do While Not tmp_Rs.EOF
        rs_TRP02W.AddNew
        rs_TRP02W.Fields("編號").Value = rs_TRP02W.RecordCount
        rs_TRP02W.Fields("送貨日").Value = tmp_Rs.Fields("送貨日").Value
        rs_TRP02W.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
        rs_TRP02W.Fields("件數").Value = tmp_Rs.Fields("件數").Value
        rs_TRP02W.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
        rs_TRP02W.Fields("板數").Value = tmp_Rs.Fields("板數").Value
        rs_TRP02W.Fields("材積").Value = tmp_Rs.Fields("材積").Value
        rs_TRP02W.Fields("重量").Value = tmp_Rs.Fields("重量").Value
        rs_TRP02W.Fields("客戶編號").Value = tmp_Rs.Fields("客戶編號").Value
        rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_TRP02W.Fields("zip").Value = tmp_Rs.Fields("zip").Value
        rs_TRP02W.Fields("客戶名稱").Value = tmp_Rs.Fields("客戶名稱").Value
        rs_TRP02W.Fields("運送地址").Value = tmp_Rs.Fields("運送地址").Value
        rs_TRP02W.Fields("訂單備註").Value = tmp_Rs.Fields("訂單備註").Value & ""
        rs_TRP02W.Fields("啤酒允收期").Value = RTrim(tmp_Rs.Fields("啤酒允收期").Value)
        rs_TRP02W.Fields("配送倉別").Value = RTrim(tmp_Rs.Fields("配送倉別").Value)
        rs_TRP02W.Fields("車種").Value = tmp_Rs.Fields("車種").Value
        rs_TRP02W.Fields("特殊需求1").Value = tmp_Rs.Fields("特殊需求1").Value
        rs_TRP02W.Fields("特殊需求2").Value = tmp_Rs.Fields("特殊需求2").Value
        rs_TRP02W.Fields("急單").Value = tmp_Rs.Fields("急單").Value
        rs_TRP02W.Fields("專車").Value = tmp_Rs.Fields("專車").Value
        rs_TRP02W.Fields("冷藏").Value = tmp_Rs.Fields("冷藏").Value
        rs_TRP02W.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_TRP02W.Fields("C_Receipt_No").Value = tmp_Rs.Fields("C_Receipt_No").Value
        rs_TRP02W.Fields("貨主單號").Value = tmp_Rs.Fields("貨主單號").Value
        rs_TRP02W.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
        rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_TRP02W.Fields("客戶簡稱").Value = tmp_Rs.Fields("客戶簡稱").Value
        rs_TRP02W.Fields("型態").Value = tmp_Rs.Fields("型態").Value
        rs_TRP02W.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_TRP02W.MoveFirst
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
    
    '待排車訂單總計資訊
    Call Retrive_OrderSum
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車列表-匯入待排車訂單", Me.Caption, "cmd_Tab0_ImportOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Print_Click()
'花王分貨表>Barcode列印

On Error GoTo err_Handle
'1. 資料寫出 Access 資料庫 >> 一次路編標貼
Dim i As Integer
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 一次路編標貼"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "一次路編標貼", cnAccess, adOpenStatic, adLockOptimistic

    rs_RouteData.MoveFirst
    Do While Not rs_RouteData.EOF
            For i = 1 To Val(rs_RouteData.Fields("板數"))
                str_SQL = "Insert into 一次路編標貼 (line,出車日,一次路線編號,一次駕駛人,客戶簡稱,總板數) " & _
                          "Values ('" & i & "','" & Left(Trim(rs_RouteData.Fields("出車日")), 4) & "/" & Mid(Trim(rs_RouteData.Fields("出車日")), 5, 2) & "/" & Right(Trim(rs_RouteData.Fields("出車日")), 2) & "','" & rs_RouteData.Fields("一次路線編號") & "','" & Trim(rs_RouteData.Fields("一次駕駛人").Value) & "','" & rs_RouteData.Fields("客戶簡稱") & "','" & i & " / " & Trim(rs_RouteData.Fields("板數").Value) & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Next
            rs_RouteData.MoveNext
    Loop
    cnAccess.CommitTrans
    Tran_Level = 0
    Call DB_Disconnect(cnAccess)
    
    '2. call Access 列印報表
    strAccessDBFileName_FullPath = GetAccessDBFileName
    Set MSAccessAP = New access.Application
    MSAccessAP.Visible = False
    MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)
    
    '[報表列印] 命令鈕 -- 利用 Access 報表
    'If chk_Tab2_PreView.Value = vbChecked Then
    '預覽列印
       MSAccessAP.Visible = True
       MSAccessAP.DoCmd.OpenReport "一次路編標貼", acViewPreview
       Call Unload_RunLogForm
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Call Unload_RunLogForm
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--TIHI_LABEL", Me.Caption, "cmd_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab0_Query_Click()
    '排車作業 >> 查詢
    If Len(txt_Tab0_RouteNo.Text) = 0 Then Exit Sub
    
    '欲修改之路編：是否已回傳WMS
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "Select EXE_CONFIRM From TRP01T Where Route_No = '" & txt_Tab0_RouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields("EXE_CONFIRM").Value = "2" Then
        msg_text = "注意：此路線編號已回傳WMS，無法修改或刪除!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
       tmp_Rs.Close
       
    '欲修改之路編：是否已出車確認
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "Select sdnstatus From TRP05T Where Route_No = '" & txt_Tab0_RouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields("sdnstatus").Value = "1" Then
        msg_text = "注意：此路線編號出車確認，無法修改或刪除!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
       tmp_Rs.Close
    
    If rs_Tab0_SelectedOrders.RecordCount <> 0 And blRouteModify = False Then
        '新增路線編號模式：
        '呼叫 [已選訂單移除(全)] 來處理已被暫時選取之 [待排車訂單] 還原回 [待排車訂單]
        Call cmd_Tab0_SelectedRemove_All_Click
    End If
    If blRouteModify And blRouteChange Then
        '有效路線編號 & 資料已遭異動，要 user 確認是否存檔
        msg_text = "路線編號資料是否存檔？"
        If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
            '呼叫存檔程序
            Call cmd_Tab0_Save_Click
        Else
            '不存檔→必須重新載入 [待排車訂單] 已還原 [選取][移除] 操作對 [待排車訂單] 的影響
            Call cmd_Tab0_ImportOrders_Click
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    '清除路線編號欄位值，包含已選訂單名細列表
    Call Clear_RouteData
    
    '取得路編資料
    str_SQL = "Select 出車日期,車牌號碼,碼頭暫存,預計報到日期,預計報到時間,運輸公司,駕駛人,駕駛電話,車種,箱數,板數,材積,重量 " & _
              "From TRPPlan_RouteQuery Where 路線編號 = '" & txt_Tab0_RouteNo.Text & "'"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定條件之排車資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
        '清除路線編號欄位值，包含已選訂單名細列表
        Call Clear_RouteData
        
        Screen.MousePointer = vbDefault
        txt_Tab0_RouteNo.SelStart = 0: txt_Tab0_RouteNo.SelLength = Len(txt_Tab0_RouteNo.Text)
        txt_Tab0_RouteNo.SetFocus
        Exit Sub
    End If
    txt_Tab0_TRPDate.Text = tmp_Rs.Fields("出車日期").Value
    txt_Tab0_DeliveryCarNo.Text = tmp_Rs.Fields("車牌號碼").Value
    txt_Tab0_DockNo.Text = tmp_Rs.Fields("碼頭暫存").Value
    txt_Tab0_CarCheckInDate.Text = tmp_Rs.Fields("預計報到日期").Value
    txt_Tab0_CarCheckInTime.Text = tmp_Rs.Fields("預計報到時間").Value
    txt_Tab0_DeliveryCompany.Text = tmp_Rs.Fields("運輸公司").Value
    txt_Tab0_DeliveryDriver.Text = tmp_Rs.Fields("駕駛人").Value
    txt_Tab0_DeliveryPhone.Text = tmp_Rs.Fields("駕駛電話").Value
    txt_Tab0_DeliveryCarType.Text = tmp_Rs.Fields("車種").Value
    txt_Tab0_Selected_Case.Text = tmp_Rs.Fields("箱數").Value
    txt_Tab0_Selected_Pallet.Text = tmp_Rs.Fields("板數").Value
    txt_Tab0_Selected_Volumn.Text = tmp_Rs.Fields("材積").Value
    txt_Tab0_Selected_Weight.Text = tmp_Rs.Fields("重量").Value
    tmp_Rs.Close
    
    '取得路編訂單
    str_SQL = "Select 送貨日,訂單編號,ZIP,Area,型態,客戶簡稱,箱數,板數,材積,重量,車種,訂單備註,特殊需求1,特殊需求2,Receipt_No,EXE回傳,客戶名稱,運送地址 " & _
              "From TRPPlan_RouteQueryOrders Where 路線編號 = '" & txt_Tab0_RouteNo.Text & "' Order by Receipt_No "
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定條件之訂單名細資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
        '清除路線編號欄位值，包含已選訂單名細列表
        Call Clear_RouteData
        
        Screen.MousePointer = vbDefault
        txt_Tab0_RouteNo.SelStart = 0: txt_Tab0_RouteNo.SelLength = Len(txt_Tab0_RouteNo.Text)
        txt_Tab0_RouteNo.SetFocus
        Exit Sub
    End If
    blTab0SelectedOrderEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab0_SelectedOrders.AddNew
        rs_Tab0_SelectedOrders.Fields("編號").Value = rs_Tab0_SelectedOrders.RecordCount
        rs_Tab0_SelectedOrders.Fields("送貨日").Value = tmp_Rs.Fields("送貨日").Value
        rs_Tab0_SelectedOrders.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
        rs_Tab0_SelectedOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
        rs_Tab0_SelectedOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value & ""
        rs_Tab0_SelectedOrders.Fields("型態").Value = tmp_Rs.Fields("型態").Value
        rs_Tab0_SelectedOrders.Fields("客戶簡稱").Value = tmp_Rs.Fields("客戶簡稱").Value
        rs_Tab0_SelectedOrders.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
        rs_Tab0_SelectedOrders.Fields("板數").Value = tmp_Rs.Fields("板數").Value
        rs_Tab0_SelectedOrders.Fields("材積").Value = tmp_Rs.Fields("材積").Value
        rs_Tab0_SelectedOrders.Fields("重量").Value = tmp_Rs.Fields("重量").Value
        rs_Tab0_SelectedOrders.Fields("車種").Value = tmp_Rs.Fields("車種").Value
        rs_Tab0_SelectedOrders.Fields("訂單備註").Value = tmp_Rs.Fields("訂單備註").Value
        rs_Tab0_SelectedOrders.Fields("特殊需求1").Value = tmp_Rs.Fields("特殊需求1").Value
        rs_Tab0_SelectedOrders.Fields("特殊需求2").Value = tmp_Rs.Fields("特殊需求2").Value
        rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab0_SelectedOrders.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
        rs_Tab0_SelectedOrders.Fields("客戶名稱").Value = tmp_Rs.Fields("客戶名稱").Value
        rs_Tab0_SelectedOrders.Fields("運送地址").Value = tmp_Rs.Fields("運送地址").Value
        rs_Tab0_SelectedOrders.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab0_SelectedOrders.MoveFirst
    rs_Tab0_SelectedOrders.Sort = " 編號 asc "
    blTab0SelectedOrderEventEnable = True
    tmp_Rs.Close
    blRouteModify = True
    blRouteChange = False
    strDispRouteNo = Trim(txt_Tab0_RouteNo.Text)
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車列表-查詢", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Remove_Click()
    '排車作業 >> ↓ 已選取訂單取消
    If rs_TRP02W Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Sub
    '已選取訂單若無反白選取：Disable 已選取消的動作，防止誤刪
    If dg_Tab0_SelectedOrders.SelBookmarks.Count = 0 Then Exit Sub
    
    blTab0SelectedOrderEventEnable = False
    
    '欲移除之訂單編號 Receipt_No
    Dim strReceiptNo As String
    strReceiptNo = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
       
    '將欲刪除之 [已選取訂單] 加入 [待排車訂單]
    Call SelectedOrders_Removeto_TRP02W(strReceiptNo)
    '重新產生 [待排車訂單] 之 [編號] 欄位值
    Call ReSet_TRP02W_SeqNo
    
    '刪除反白選取之訂單：已選取訂單部分
    rs_Tab0_SelectedOrders.Delete
    If Not rs_Tab0_SelectedOrders.EOF Then rs_Tab0_SelectedOrders.MoveFirst
    If dg_Tab0_SelectedOrders.SelBookmarks.Count > 0 Then dg_Tab0_SelectedOrders.SelBookmarks.Remove 0
    '重新計算已選取訂單：箱數，板數，材積，重量 + 編號重新產生
    Call Calculate_SelectedOrders
    
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
       strSourceFilter = adFilterNone
       rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '原始排序，一般資料序號由小至大
    blTRP02WEventEnable = True
    blTab0SelectedOrderEventEnable = True
    
    '重新計算 [待排車列表] 的總計資訊
    Call ReCaculate_OrderSum

End Sub

Private Sub cmd_Tab0_Reserve_Click()
    '待排車訂單：保留訂單
    cmd_Tab0_Reserve.Enabled = False
    
    '待排車訂單：選取小計：歸零
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '還原所有篩選設定，並以預設 [編號] 排列
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    
    Dim strRouteNo As String, intDriveTimes As Integer, dbOrderCnt As Double, iLoop As Double
    strRouteNo = "D"   '特殊路線編號，統管所有保留訂單
    intDriveTimes = 1
    dbOrderCnt = 0
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    blTab2ReservedEventEnable = False
    '篩選已選取者
    rs_TRP02W.Filter = "＊='V'"
    If Not rs_TRP02W.EOF Then
        Do While Not rs_TRP02W.EOF
            rs_Tab2_ReservedOrders.AddNew
            For iLoop = 0 To rs_TRP02W.Fields.Count - 1
                rs_Tab2_ReservedOrders.Fields(iLoop).Value = rs_TRP02W.Fields(iLoop).Value
            Next iLoop
            rs_Tab2_ReservedOrders.Fields(1).Value = " "
            rs_Tab2_ReservedOrders.Update
            
            'insert into TRP02T
            str_SQL = "Insert into TRP02T (Route_No,StorerKey,C_Receipt_No,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                      "Select '" & strRouteNo & "',StorerKey,C_Receipt_No,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " 'D'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                      "From TRP02W Where Receipt_No = '" & rs_TRP02W.Fields("Receipt_No").Value & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '後續由 TRP02T Trigger [insert] 進行以下作業
            '   a.寫入 TRP03T -- 排車訂單明細檔
            '   b.刪除 TRP03W -- 待排車訂單明細檔
            '   c.刪除 TRP02W -- 待排車訂單主檔
            
            rs_TRP02W.MoveNext
        Loop
        '[待選取訂單] 中，刪除已選取之訂單
        rs_TRP02W.MoveFirst
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Delete
            rs_TRP02W.MoveFirst
        Loop
    End If
    
    '更新 trp01t & trp05t 的 [箱數] [板數] [重量] [材積]
    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    blTab2ReservedEventEnable = True
    
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        strSourceFilter = adFilterNone
        rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '原始排序，一般資料序號由小至大
    '取消反白選取狀態
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    blTRP02WEventEnable = True
    cmd_Tab0_Reserve.Enabled = True
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車作業-建立路線編號", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_Reserve.Enabled = True
End Sub

Private Sub cmd_Tab0_Save_Click()
    '排車作業 >> 路線編號修改模式存檔
    If blRouteModify = False Then
        msg_text = "非經 [查詢] 程序所顯示之有效 [路線編號]"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    End If
    If blRouteChange = False Then
        msg_text = "[路線編號] 的資料並未異動，不須執行 [存檔] 程序"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    Else
       '訂單資料有異動，且全部都被移除，等同刪除
        If rs_Tab0_SelectedOrders.RecordCount = 0 Then
            msg_text = "此路線編號目前已無訂單，是否刪除此路編？"
            If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
                Call Delete_RouteNo(strDispRouteNo)
                Call Clear_RouteData
                txt_Tab0_RouteNo.Text = ""
                Exit Sub
            End If
        End If
    End If
    '檢核路線編號資料是否正確輸入
    If RouteData_Check = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    cmd_Tab0_Save.Enabled = False
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Dim intDriveTimes As Integer    '車次
    '1.確認 [出車日期] 與 [車牌號碼] [修改權限] & [資料是否遭異動]
    '  若異動則必須重新計算車次
    str_SQL = "Select Rtrim(t05t.Vehicle_ID_No) as 車牌號碼,Convert(varchar(8),t01t.Delivery_Date,112) as 出車日期,Rtrim(Isnull(t01t.AddWho,'')) as AddWho,t05t.Drive_Times as 車次 " & _
              "From TRP05T t05t inner join TRP01T t01t on t01t.Route_No = t05t.Route_No " & _
              "Where t05t.Route_No = '" & strDispRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "路線編號 [" & strDispRouteNo & "] 已找不到資料了"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl Then
        tmp_Rs.Close
        msg_text = "權限控管：路線編號之修改只允許由原排車者執行"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault: cmd_Tab0_Save.Enabled = True
        Exit Sub
    End If
    intDriveTimes = tmp_Rs.Fields("車次").Value
    
    If tmp_Rs.Fields("出車日期").Value <> txt_Tab0_TRPDate.Text Or UCase(tmp_Rs.Fields("車牌號碼").Value) <> txt_Tab0_DeliveryCarNo.Text Then
        '出車日期 or 車牌號碼遭異動：必須重新計算車次
        tmp_Rs.Close
        str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                  "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    End If
    tmp_Rs.Close
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '2.更新 TRP05T & TRP01T & TRP03T add TRP02T add By Gemini @20080313
    str_SQL = "Update TRP01T Set Delivery_Date = '" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "' " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP05T Set Delivery_Date = '" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "', " & _
              "   Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & ",Dock_No = '" & txt_Tab0_DockNo.Text & "',Expect_Date = '" & txt_Tab0_CarCheckInDate.Text & "'," & _
              "   Expect_Time = '" & txt_Tab0_CarCheckInTime.Text & "' " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP03T Set Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP02T Set Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '3.由車輛主檔更新 TRP05T 車輛相關欄位
    str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strDispRouteNo & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '4..將 TRP02T 全部更新標示為 [更新旗標] DeleteFlag = '1'
    str_SQL = "Update TRP02T Set DeleteFlag='1' Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '5.將訂單更新旗標 DeleteFalg 還原回 0
    '  找不到的，表示是新加入的，進行新增程序
    blTab0SelectedOrderEventEnable = False
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        str_SQL = "Update TRP02T Set DeleteFlag='0' Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        If RowsAffect = 0 Then
            '新增訂單
            str_SQL = "Insert into TRP02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OPQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                      "Select '" & strDispRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OTQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                      "From TRP02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    blTab0SelectedOrderEventEnable = True
    
    '6.將移除訂單還原回 TRP02W & TRP03W
    '(1).將 TRP03T 寫回 TRP03W >> 刪除 TRP03T
    str_SQL = "Insert into TRP03W(" & _
              " STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From TRP03T A INNER JOIN TRP02T B ON B.Receipt_No = a.Receipt_No and b.DeleteFlag = '1' and b.Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).將 TRP02T 寫回 TRP02W >> 刪除 TRP02T
    str_SQL = "Insert into TRP02W(" & _
              " RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              " WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
              "Select RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
              "From TRP02T Where DeleteFlag='1' and Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(3).刪除 TRP02T & TRP03T
    str_SQL = "Delete TRP03T FROM TRP02T Where TRP02T.Receipt_No = TRP03T.Receipt_No and TRP02T.DeleteFlag='1' and TRP02T.Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Delete From TRP02T Where DeleteFlag='1' and Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '7.更新 TRP01T & TRP05T 的統計欄位值
    str_SQL = "exec  ReservedOrders_Recalculate " & strDispRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
       cn.CommitTrans
       Tran_Level = 0
    End If
    
    '清除螢幕欄位值
    Call Clear_RouteData
    txt_Tab0_RouteNo.Text = ""
    cmd_Tab0_Save.Enabled = True
    
    '待排車訂單總計資訊
    Call Retrive_OrderSum
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車作業-路線編號修改存檔", Me.Caption, "cmd_Tab0_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_SelectCar_Click()
    '排車作業 >> 司機選取
    If Len(txt_Tab0_TRPDate.Text) = 0 Then
        msg_text = "請先輸入：出車日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SetFocus
        Exit Sub
    Else
        If chk_Tab0_DriveTimes.Value = vbChecked Then
            '顯示運送車輛待選清單--包含已排定之車次資料顯示
            Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar.Name & "1")
        Else
            '顯示運送車輛待選清單--不顯示車次資料
            Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar.Name & "2")
        End If
    End If
End Sub

Private Sub cmd_Tab0_Selected_Click()
    '待排車訂單：選取
    
    '待排車訂單：選取小計：歸零
    dbSelectedCount = 0
    dbsrcSelected_OTqty = 0: dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_OTqty.Text = ""
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '還原所有篩選設定，並以預設 [編號] 排列
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    
    '篩選已選取者
    rs_TRP02W.Filter = "＊='V'"
    If Not rs_TRP02W.EOF Then
        dg_Tab0_SelectedOrders.Visible = False
        blTab0SelectedOrderEventEnable = False
        Do While Not rs_TRP02W.EOF
                
            '判斷是否已經選取過
            rs_Tab0_SelectedOrders.Filter = adFilterNone
            rs_Tab0_SelectedOrders.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
            rs_Tab0_SelectedOrders.Filter = "Receipt_No = '" & rs_TRP02W.Fields("Receipt_No").Value & "'"
            '如果是查詢所顯示之有效路編，設定路編異動識別旗標
            If blRouteModify Then blRouteChange = True
            If rs_Tab0_SelectedOrders.EOF Then
                '新增選取之訂單
                rs_Tab0_SelectedOrders.AddNew
                rs_Tab0_SelectedOrders.Fields("編號").Value = 999
                rs_Tab0_SelectedOrders.Fields("送貨日").Value = rs_TRP02W.Fields("送貨日").Value
                rs_Tab0_SelectedOrders.Fields("訂單編號").Value = rs_TRP02W.Fields("訂單編號").Value
                rs_Tab0_SelectedOrders.Fields("ZIP").Value = rs_TRP02W.Fields("ZIP").Value
                rs_Tab0_SelectedOrders.Fields("Area").Value = rs_TRP02W.Fields("Area").Value
                rs_Tab0_SelectedOrders.Fields("急單").Value = rs_TRP02W.Fields("急單").Value
                rs_Tab0_SelectedOrders.Fields("型態").Value = rs_TRP02W.Fields("型態").Value
                rs_Tab0_SelectedOrders.Fields("客戶簡稱").Value = rs_TRP02W.Fields("客戶簡稱").Value
                rs_Tab0_SelectedOrders.Fields("件數").Value = rs_TRP02W.Fields("件數").Value
                rs_Tab0_SelectedOrders.Fields("箱數").Value = rs_TRP02W.Fields("箱數").Value
                rs_Tab0_SelectedOrders.Fields("板數").Value = rs_TRP02W.Fields("板數").Value
                rs_Tab0_SelectedOrders.Fields("材積").Value = rs_TRP02W.Fields("材積").Value
                rs_Tab0_SelectedOrders.Fields("重量").Value = rs_TRP02W.Fields("重量").Value
                rs_Tab0_SelectedOrders.Fields("車種").Value = rs_TRP02W.Fields("車種").Value
                rs_Tab0_SelectedOrders.Fields("訂單備註").Value = rs_TRP02W.Fields("訂單備註").Value
                rs_Tab0_SelectedOrders.Fields("特殊需求1").Value = rs_TRP02W.Fields("特殊需求1").Value
                rs_Tab0_SelectedOrders.Fields("特殊需求2").Value = rs_TRP02W.Fields("特殊需求2").Value
                rs_Tab0_SelectedOrders.Fields("訂單備註").Value = rs_TRP02W.Fields("訂單備註").Value
                rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = rs_TRP02W.Fields("Receipt_No").Value
                rs_Tab0_SelectedOrders.Fields("EXE回傳").Value = rs_TRP02W.Fields("EXE回傳").Value
                rs_Tab0_SelectedOrders.Fields("客戶名稱").Value = rs_TRP02W.Fields("客戶名稱").Value
                rs_Tab0_SelectedOrders.Fields("運送地址").Value = rs_TRP02W.Fields("運送地址").Value
                rs_Tab0_SelectedOrders.Update
            Else
                '更新選取之訂單資料
                rs_Tab0_SelectedOrders.Fields("送貨日").Value = rs_TRP02W.Fields("送貨日").Value
                rs_Tab0_SelectedOrders.Fields("訂單編號").Value = rs_TRP02W.Fields("訂單編號").Value
                rs_Tab0_SelectedOrders.Fields("ZIP").Value = rs_TRP02W.Fields("ZIP").Value
                rs_Tab0_SelectedOrders.Fields("Area").Value = rs_TRP02W.Fields("Area").Value
                rs_Tab0_SelectedOrders.Fields("型態").Value = rs_TRP02W.Fields("型態").Value
                rs_Tab0_SelectedOrders.Fields("客戶簡稱").Value = rs_TRP02W.Fields("客戶簡稱").Value
                rs_Tab0_SelectedOrders.Fields("件數").Value = rs_TRP02W.Fields("件數").Value
                rs_Tab0_SelectedOrders.Fields("箱數").Value = rs_TRP02W.Fields("箱數").Value
                rs_Tab0_SelectedOrders.Fields("板數").Value = rs_TRP02W.Fields("板數").Value
                rs_Tab0_SelectedOrders.Fields("材積").Value = rs_TRP02W.Fields("材積").Value
                rs_Tab0_SelectedOrders.Fields("重量").Value = rs_TRP02W.Fields("重量").Value
                rs_Tab0_SelectedOrders.Fields("車種").Value = rs_TRP02W.Fields("車種").Value
                rs_Tab0_SelectedOrders.Fields("訂單備註").Value = rs_TRP02W.Fields("訂單備註").Value
                rs_Tab0_SelectedOrders.Fields("特殊需求1").Value = rs_TRP02W.Fields("特殊需求1").Value
                rs_Tab0_SelectedOrders.Fields("特殊需求2").Value = rs_TRP02W.Fields("特殊需求2").Value
                rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = rs_TRP02W.Fields("Receipt_No").Value
                rs_Tab0_SelectedOrders.Fields("EXE回傳").Value = rs_TRP02W.Fields("EXE回傳").Value
                rs_Tab0_SelectedOrders.Fields("客戶名稱").Value = rs_TRP02W.Fields("客戶名稱").Value
                rs_Tab0_SelectedOrders.Fields("運送地址").Value = rs_TRP02W.Fields("運送地址").Value
            End If
            rs_TRP02W.MoveNext
        Loop
        '重新對 [已選取訂單] 產生 [編號] 與相關資料統計：箱數，板數，材積，重量
        Call Calculate_SelectedOrders
        dg_Tab0_SelectedOrders.Visible = True
        blTab0SelectedOrderEventEnable = True
        
        '[待選取訂單] 中，刪除已選取之訂單
        rs_TRP02W.MoveFirst
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Delete
            rs_TRP02W.MoveFirst
        Loop
    End If
    
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        rs_TRP02W.Filter = adFilterNone
        strSourceFilter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '原始排序，一般資料序號由小至大
    '取消反白選取狀態
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    
    '重新計算 [待排車列表] 的總計資訊
    Call ReCaculate_OrderSum
    
    blTRP02WEventEnable = True

End Sub

Private Sub cmd_Tab0_SelectedCancel_All_Click()
    '排車作業 >> X待選全部取消
    
    '待排車訂單：選取小計：歸零
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '還原所有篩選設定，並以預設 [編號] 排列
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    
    '篩選已選取者
    rs_TRP02W.Filter = "＊='V'"
    If Not rs_TRP02W.EOF Then
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Fields("＊").Value = " "
            rs_TRP02W.MoveNext
        Loop
    End If
    
    blTRP02WEventEnable = False
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        strSourceFilter = adFilterNone
        rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '原始排序，一般資料序號由小至大
    blTRP02WEventEnable = True
    
    '取消反白選取狀態
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    '還原 [待排車訂單] 排序設定
    blTRP02WEventEnable = True
End Sub

Private Sub cmd_Tab0_SelectedCancel_Click()
    '排車作業 >> X待選取消
    If rs_TRP02W Is Nothing Then Exit Sub
        '待選取訂單若無反白選取：Disable 待選取消，防止誤刪
        If dg_TRP02W.SelBookmarks.Count = 0 Then Exit Sub
        
        If Trim(rs_TRP02W.Fields(1).Value) = "V" Then
        rs_TRP02W.Fields(1).Value = " "
        dbSelectedCount = dbSelectedCount - 1
        '待選定單：選取小計更新
        If dbSelectedCount = 0 Then
            dbsrcSelected_Case = 0
            dbsrcSelected_Pallet = 0
            dbsrcSelected_Volumn = 0
            dbsrcSelected_Weight = 0
            txt_Tab0_srcSelected_Case.Text = 0: txt_Tab0_srcSelected_Pallet.Text = 0
            txt_Tab0_srcSelected_Volumn.Text = 0: txt_Tab0_srcSelected_Weight.Text = 0
        Else
            dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP02W.Fields("箱數").Value
            dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP02W.Fields("板數").Value
            dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP02W.Fields("材積").Value
            dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP02W.Fields("重量").Value
            txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
            txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
        End If
        '取消反白選取狀態
        If dg_TRP02W.SelBookmarks.Count > 0 Then
            dg_TRP02W.SelBookmarks.Remove 0
        End If
    End If

End Sub


Private Sub cmd_Tab0_SelectedRemove_All_Click()
    '排車作業 >> ↓ 已選取訂單取消-全部
    If rs_TRP02W Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then Exit Sub
    '路線編號查詢：有效路線編號
    '按下 [已選訂單移除(全) 等同於刪除路線編號
    If blRouteModify Then
        msg_text = "確定要刪除此路線編號 [" & txt_Tab0_RouteNo.Text & "]"
        If MsgBox(msg_text, vbCritical + vbOKCancel, msg_title) = vbOK Then
            '刪除指定路線編號
            Call Delete_RouteNo(strDispRouteNo)
            '清除路線編號欄位值，包含已選訂單名細列表
            Call Clear_RouteData
            txt_Tab0_RouteNo.Text = ""
        End If
        Exit Sub
    End If
    
    blTab0SelectedOrderEventEnable = False
    
    '欲移除之訂單編號 Receipt_No
    Dim strReceiptNo As String
    '逐筆寫回 [待排車訂單 TRP02W]
    rs_Tab0_SelectedOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        strReceiptNo = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
        '將欲刪除之 [已選取訂單] 加入 [待排車訂單]
        Call SelectedOrders_Removeto_TRP02W(strReceiptNo)
        rs_Tab0_SelectedOrders.MoveNext
    Loop
       
    '重新產生 [待排車訂單] 之 [編號] 欄位值
    Call ReSet_TRP02W_SeqNo
    
    '排車作業：已選取之待排車訂單列表 DBGrid 格式設定-ReSet
    Call CreateRS_Tab0_SelectedOrders
    
    '重新計算已選取訂單：箱數，板數，材積，重量 + 編號重新產生
    Call Calculate_SelectedOrders
    '排序方式
    
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        strSourceFilter = adFilterNone
        rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '原始排序，一般資料序號由小至大
    blTRP02WEventEnable = True
    
    '重新計算 [待排車列表] 的總計資訊
    Call ReCaculate_OrderSum
    
    blTab0SelectedOrderEventEnable = True
End Sub

Private Sub cmd_Tab0_srcOrderReset_Click()
    '排車作業 >> 取消待排車訂單篩選排序
    If rs_TRP02W Is Nothing Then Exit Sub
    '移除篩選條件，重設排序依據
     blTRP02WEventEnable = False
    '篩選已選取者：取消選取
    rs_TRP02W.Filter = "＊='V'"
    If Not rs_TRP02W.EOF Then
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Fields(1).Value = " "
            rs_TRP02W.MoveNext
        Loop
    End If
    rs_TRP02W.Filter = adFilterNone
    strSourceFilter = adFilterNone
     'rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    rs_TRP02W.Sort = strSourceOrderBy
    blTRP02WEventEnable = True
    
    '重新計算 [待排車列表] 的總計資訊
    Call ReCaculate_OrderSum

End Sub

Private Sub cmd_Tab1_RouteNoDelete_Click()
    '路線編號列表 >> 路線編號刪除
    If rs_Tab1_Route.RecordCount = 0 Then Exit Sub
    If dg_Tab1_Route.SelBookmarks.Count = 0 Then Exit Sub
    
    Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
    strDeleteRouteNo = Trim(rs_Tab1_Route.Fields("路線編號").Value)
    strCarno = Trim(rs_Tab1_Route.Fields("車牌號碼").Value)
    dbDriveTimes = Trim(rs_Tab1_Route.Fields("車次").Value)
     
    '欲刪除之路編：是否已回傳WMS
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "Select isnull(Route,'') From " & strWMSDB & "..orders Where Route = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        msg_text = "注意：WMS系統有此路線編號時，無法修改或刪除!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
       tmp_Rs.Close

    '欲刪除之路編：是否已出車確認
    'str_SQL = "Select c_Route_No  From SDN01T Where c_Route_No = '" & strDeleteRouteNo & "'"
    'Terry 20191127 改為檢查出車狀態
    str_SQL = "Select Route_No  From TRP05T Where Route_No = '" & strDeleteRouteNo & "' and sdnstatus = '1' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "注意：此路線編號已出車確認，無法刪除! "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    tmp_Rs.Close

    '欲刪除之路編：是否已打散重組 Add by Terry 20191127
    str_SQL = "Select Route_No  From SDN02W Where Route_No = '" & strDeleteRouteNo & "' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "注意：此路線編號已打散重組，無法刪除! "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    tmp_Rs.Close
    

    msg_text = "確認刪除路線編號：" & strDeleteRouteNo
    If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub

    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)

    '驗證欲刪除之路編，排車者是否為此時登入之使用者
    str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "資料異常：找不到欲刪除之路線編號!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    Else
        If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl = True Then
            tmp_Rs.Close
            msg_text = "權限控管：路線編號之刪除只允許由原排車者執行"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Exit Sub
        End If
    End If
    tmp_Rs.Close

    '刪除路編
    Call Delete_RouteNo(strDeleteRouteNo)
    
    '刪除查詢結果中該筆路線編號--rs_Tab1_RouteOrders
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab1_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    rs_Tab1_RouteOrders.Filter = "路線編號='" & strDeleteRouteNo & "'"
    If Not rs_Tab1_RouteOrders.EOF Then
        Do While Not rs_Tab1_RouteOrders.EOF
            rs_Tab1_RouteOrders.Delete
            rs_Tab1_RouteOrders.MoveFirst
        Loop
    End If
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab1_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    
    '(7).刪除查詢結果中該筆路線編號--rs_Tab1_Route
    rs_Tab1_Route.Delete
    If Not rs_Tab1_Route.EOF Then rs_Tab1_Route.MoveFirst
    
    blTab1RouteEventEnable = True
    Screen.MousePointer = vbDefault
    
    
On Error GoTo err_Handle2
    'Terry 20200220 同步刪除BestAPP上該路編資料
    Dim HttpClient As Object
    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/DeleteRouteNoByWareHouse?Route_NO=" & strDeleteRouteNo & "&WareHouse=GYDC_BEST", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    Exit Sub

err_Handle2:
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-路線編號列表-路線編號刪除", Me.Caption, "cmd_Tab1_RouteNoDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_RouteNoQuery_Click()
    '路線編號列表 >> 路線編號查詢
    If Len(Trim(txt_Tab1_RouteNo.Text)) = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    
    '設定路線編號列表
    blTab1RouteEventEnable = False
    Call CreateRS_Tab1_Route
    blTab1RouteEventEnable = True
    '設定路線編號之訂單列表
    Call CreateRS_Tab1_RouteOrders
    
    str_SQL = "Select  Rtrim(a1.Route_No) as 路線編號 " & _
            ",Convert(varchar(8),a1.Delivery_Date,112) as 出車日期 " & _
            ",Rtrim(a2.Vehicle_ID_No) as 車牌號碼 " & _
            ",a2.Drive_Times as 車次 " & _
            ",Rtrim(Isnull(a2.Driver,'')) as 駕駛人 " & _
            ",Round((select isnull(sum(t2.otqty),0) from trp02t t2 where a1.route_no = t2.route_no),2) as 件數 " & _
            ",Round(a1.Case_cnt,2) as 箱數 " & _
            ",Round(a1.Pallet_Qty,2) as 板數 " & _
            ",Round(a1.Volumn_Weight,2) as 材積 " & _
            ",Round(a1.Weight,2) as 重量 " & _
            ",Rtrim(Isnull(b1.VEHICLE_TYPE,'')) as 車種 " & _
            ",Rtrim(Isnull(a2.Dock_No,'')) as 碼頭暫存 " & _
            ",Rtrim(Isnull(a2.Expect_Date,'')) as 預計報到日期 " & _
            ",Rtrim(Isnull(a2.Expect_time,'')) as 預計報到時間 " & _
            ",Case a1.EXE_Confirm When '0' Then '新建路編' When '1' Then '設定回傳' When '2' Then '已回傳' When '9' Then '預先揀貨' else '未知狀態' End as EXE回傳 " & _
            ",Rtrim(Isnull(a1.AddWho,'')) as 排車者 " & _
            "From TRP01T a1 " & _
            "join TRP05T a2 on a2.Route_No = a1.Route_No " & _
            "join TRP09M b1 on b1.Vehicle_ID_No = a2.Vehicle_ID_No " & _
            "Left join TRP15M b2 on b2.Vehicle_Type = b1.Vehicle_Type " & _
            "Where Left(a1.Route_No,1) = 'F' and Rtrim(a1.Route_No) like '%" & RTrim(txt_Tab1_RouteNo.Text) & "%'  order by Rtrim(a1.Route_No)"
    
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定條件之路線編號資料(TRP01T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTab1RouteEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab1_Route.AddNew
        rs_Tab1_Route.Fields("編號").Value = rs_Tab1_Route.RecordCount
        rs_Tab1_Route.Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
        rs_Tab1_Route.Fields("出車日期").Value = tmp_Rs.Fields("出車日期").Value
        rs_Tab1_Route.Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
        rs_Tab1_Route.Fields("車次").Value = tmp_Rs.Fields("車次").Value
        rs_Tab1_Route.Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
        rs_Tab1_Route.Fields("件數").Value = tmp_Rs.Fields("件數").Value
        rs_Tab1_Route.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
        rs_Tab1_Route.Fields("板數").Value = tmp_Rs.Fields("板數").Value
        rs_Tab1_Route.Fields("材積").Value = tmp_Rs.Fields("材積").Value
        rs_Tab1_Route.Fields("重量").Value = tmp_Rs.Fields("重量").Value
        rs_Tab1_Route.Fields("車種").Value = tmp_Rs.Fields("車種").Value
        rs_Tab1_Route.Fields("碼頭暫存").Value = tmp_Rs.Fields("碼頭暫存").Value
        rs_Tab1_Route.Fields("預計報到日期").Value = tmp_Rs.Fields("預計報到日期").Value
        rs_Tab1_Route.Fields("預計報到時間").Value = tmp_Rs.Fields("預計報到時間").Value
        rs_Tab1_Route.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
        rs_Tab1_Route.Fields("排車者").Value = tmp_Rs.Fields("排車者").Value
        rs_Tab1_Route.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab1_Route.MoveFirst
    blTab1RouteEventEnable = True
    tmp_Rs.Close
    'TRP03W
    str_SQL = "Select  Rtrim(a1.Route_No) as 路線編號 " & _
            ",Convert(varchar(8),a1.Arrive_Date,112) as 送貨日 " & _
            ",Rtrim(a1.Receipt_No) + '(' + Rtrim(a1.StorerKey)+'-'+Rtrim(a1.Extern)+')' as 訂單編號 " & _
            ",Rtrim(a2.ZIP) as ZIP " & _
            ",Rtrim(a2.Full_Name) as 客戶名稱 " & _
            ",Rtrim(a2.address) as 客戶地址 " & _
            ",件數 =  isnull(a1.otqty,(select sum(isnull(o.otqty,0)) from orders o where o.orderkey = a1.c_receipt_no)) " & _
            ",Round(isnull(a1.Case_cnt,0),2) as 箱數 " & _
            ",Round(isnull(a1.Pallet_Qty,0),2) as 板數 " & _
            ",Round(isnull(a1.Volumn_Weight,0),2) as 材積 " & _
            ",Round(isnull(a1.Weight,0),2) as 重量 " & _
            ",Rtrim(a1.Receipt_No) as Receipt_No " & _
            ",Case a1.EXE_Confirm When '0' Then '新建路編' When '1' Then '設定回傳' When '2' Then '已回傳' When '9' Then '預先揀貨' else '未知狀態' End  AS EXE回傳 " & _
            ",Rtrim(Isnull(a2.Area_Code,'')) as Area " & _
            ",Rtrim(Isnull(a1.urgent_mark,'')) as 急單 " & _
            ",Rtrim(Isnull(a1.Priority,'')) as 型態 " & _
            ",Rtrim(Isnull(a2.Short_Name,'')) as 客戶簡稱 " & _
            ",Rtrim(Isnull(a1.Priority,'')) as 型態 " & _
            ",訂單備註 = Rtrim(Isnull(a1.description,'')) " & _
            "From TRP02T a1 " & _
            "inner join TRP01M a2 on a2.ConsigneeKey = a1.ConsigneeKey and a2.storerkey = a1.storerkey " & _
            "Where Rtrim(a1.Route_No) like '%" & txt_Tab1_RouteNo.Text & "%' order by Rtrim(a1.Route_No),Rtrim(a1.Receipt_No)"
            
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定路線編號之訂單資料(TRP02T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Do While Not tmp_Rs.EOF
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("編號").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
        rs_Tab1_RouteOrders.Fields("送貨日").Value = tmp_Rs.Fields("送貨日").Value
        rs_Tab1_RouteOrders.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("客戶名稱").Value = tmp_Rs.Fields("客戶名稱").Value
        rs_Tab1_RouteOrders.Fields("地址").Value = tmp_Rs.Fields("客戶地址").Value
        rs_Tab1_RouteOrders.Fields("件數").Value = tmp_Rs.Fields("件數").Value
        rs_Tab1_RouteOrders.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
        rs_Tab1_RouteOrders.Fields("板數").Value = tmp_Rs.Fields("板數").Value
        rs_Tab1_RouteOrders.Fields("材積").Value = tmp_Rs.Fields("材積").Value
        rs_Tab1_RouteOrders.Fields("重量").Value = tmp_Rs.Fields("重量").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("型態").Value = tmp_Rs.Fields("型態").Value
        rs_Tab1_RouteOrders.Fields("客戶簡稱").Value = tmp_Rs.Fields("客戶簡稱").Value
        rs_Tab1_RouteOrders.Fields("訂單備註").Value = tmp_Rs.Fields("訂單備註").Value
        rs_Tab1_RouteOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab1_RouteOrders.MoveFirst
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-路線編號列表-路線編號查詢", Me.Caption, "cmd_Tab1_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Delete_Click()
    '保留訂單 >> 移至 [待排車訂單]
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    DelRecord = MsgBox("刪除後資料無法復原,確定要刪除? ", vbQuestion + vbYesNo, "刪除")
    If DelRecord = vbNo Then
        Exit Sub
    End If
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    blTab2ReservedEventEnable = False
    blTRP02WEventEnable = False
    cmd_Tab2_Delete.Enabled = False
    
    '篩選已選取者
    rs_Tab2_ReservedOrders.Filter = "＊='V'"
    If rs_Tab2_ReservedOrders.EOF Then
        rs_Tab2_ReservedOrders.Filter = adFilterNone
        rs_Tab2_ReservedOrders.Sort = "編號 ASC"
        blTab2ReservedEventEnable = True
        blTRP02WEventEnable = True
        cmd_Tab2_Remove.Enabled = True
        Exit Sub
    End If
    
    Dim strRouteNo As String, iLoop As Double
    strRouteNo = "D"
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        str_SQL = "delete TRP02T where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP03T where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP02W where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP03W where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP02W_TEMP where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        cn.Execute "delete status where orderkey ='" & rs_Tab2_ReservedOrders("TMS單號") & "'", RowsAffect, adExecuteNoRecords
        
        str_SQL = "update orders set B_PHONE2='00',trafficCop=null,type='刪單' ,editdate = getdate()  where orderkey='" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "' and priority = '" & rs_Tab2_ReservedOrders.Fields("型態").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        rs_Tab2_ReservedOrders.MoveNext
    Loop
    
    '[待選取訂單] 中，刪除已選取之訂單
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        rs_Tab2_ReservedOrders.Delete
        rs_Tab2_ReservedOrders.MoveFirst
    Loop
    
'    '更新 trp01t & trp05t 的 [箱數] [板數] [重量] [材積]
'    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    
    If Not (rs_TRP02W Is Nothing) Then
        If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
        If rs_TRP02W.EOF Then
            strSourceFilter = adFilterNone
            rs_TRP02W.Filter = adFilterNone
        End If
        rs_TRP02W.Sort = strSourceOrderBy  '原始排序，一般資料序號由小至大
    End If
    
    rs_Tab2_ReservedOrders.Filter = adFilterNone
    rs_Tab2_ReservedOrders.Sort = "編號 ASC"
    blTab2ReservedEventEnable = True
    blTRP02WEventEnable = True
    
    cmd_Tab2_Delete.Enabled = True
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-保留訂單-移至待排車訂單列表", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   blTab2ReservedEventEnable = True
   blTRP02WEventEnable = True
   cmd_Tab2_Delete.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Remove_Click()

    '保留訂單 >> 移至 [待排車訂單]
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    blTab2ReservedEventEnable = False
    blTRP02WEventEnable = False
    cmd_Tab2_Remove.Enabled = False
    
    '篩選已選取者
    rs_Tab2_ReservedOrders.Filter = "＊='V'"
    If rs_Tab2_ReservedOrders.EOF Then
        rs_Tab2_ReservedOrders.Filter = adFilterNone
        rs_Tab2_ReservedOrders.Sort = "編號 ASC"
        blTab2ReservedEventEnable = True
        blTRP02WEventEnable = True
        cmd_Tab2_Remove.Enabled = True
        Exit Sub
    End If
    
    Dim strRouteNo As String, iLoop As Double
    strRouteNo = "D"
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    Call ReDim_Recordset(tmp_Rs)
    
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
    
    '檢查訂單是否刪除
    tmp_Rs.Open "select receipt_no from trp02t where route_no = 'D' and receipt_no = '" & rs_Tab2_ReservedOrders("TMS單號") & "' ", cn
    If Not tmp_Rs.EOF Then

       '(1).將 TRP03T 寫回 TRP03W >> 刪除 TRP03T
       str_SQL = "Insert into TRP03W(" & _
                 "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
                 "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
                 "From TRP03T A Where a.Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

       If Not (rs_TRP02W Is Nothing) Then
            rs_TRP02W.AddNew
            For iLoop = 0 To rs_Tab2_ReservedOrders.Fields.Count - 1
                rs_TRP02W.Fields(iLoop).Value = rs_Tab2_ReservedOrders.Fields(iLoop).Value
            Next iLoop
            rs_TRP02W.Fields(0).Value = rs_TRP02W.RecordCount
            rs_TRP02W.Fields(1).Value = " "
            rs_TRP02W.Update
       End If
    
       '(2).將 TRP02T 寫回 TRP02W >> 刪除 TRP02T
       str_SQL = "Insert into TRP02W(" & _
                 "   RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
                 "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                 "Select RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
                 "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                 "From TRP02T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
       '(3).刪除 TRP02T & TRP03T
       str_SQL = "Delete From TRP03T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
       
       str_SQL = "Delete From TRP02T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS單號").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
       
    End If
    tmp_Rs.Close
    
       rs_Tab2_ReservedOrders.MoveNext
    Loop
    
    '[待選取訂單] 中，刪除已選取之訂單
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        rs_Tab2_ReservedOrders.Delete
        rs_Tab2_ReservedOrders.MoveFirst
    Loop
    
    '更新 trp01t & trp05t 的 [箱數] [板數] [重量] [材積]
    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    
    If Not (rs_TRP02W Is Nothing) Then
        If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
        If rs_TRP02W.EOF Then
            strSourceFilter = adFilterNone
            rs_TRP02W.Filter = adFilterNone
        End If
        rs_TRP02W.Sort = strSourceOrderBy  '原始排序，一般資料序號由小至大
    End If
    
    rs_Tab2_ReservedOrders.Filter = adFilterNone
    rs_Tab2_ReservedOrders.Sort = "編號 ASC"
    blTab2ReservedEventEnable = True
    blTRP02WEventEnable = True
    
    cmd_Tab2_Remove.Enabled = True
    
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-保留訂單-移至待排車訂單列表", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   blTab2ReservedEventEnable = True
   blTRP02WEventEnable = True
   cmd_Tab2_Remove.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_ShowAll_Click()
    '排車作業>>顯示所有保留訂單資料
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    
    '保留訂單列表
    blTab2ReservedEventEnable = False
    Call CreateRS_Tab2_ReservedOrders
    DoEvents
    
    '取回保留訂單資料
    str_SQL = "Select ' ' as '＊',送貨日,訂單編號,件數,箱數,板數,材積,重量,客戶編號,ZIP,客戶簡稱,Area,型態,運送地址,訂單備註,啤酒允收期,配送倉別,車種,特殊需求1,特殊需求2,急單,專車,冷藏,TMS單號,C_RECEIPT_NO,貨主單號,EXE回傳,客戶名稱 " & _
              "From TRPPlan_ReservedOrder Order by 訂單編號 "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定條件之保留訂單資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Dim iLoop As Double
    Do While Not tmp_Rs.EOF
        rs_Tab2_ReservedOrders.AddNew
        For iLoop = 1 To rs_Tab2_ReservedOrders.Fields.Count - 1
            rs_Tab2_ReservedOrders.Fields(iLoop).Value = tmp_Rs.Fields(iLoop - 1).Value & ""
        Next iLoop
        rs_Tab2_ReservedOrders.Fields(0).Value = rs_Tab2_ReservedOrders.RecordCount
        rs_Tab2_ReservedOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    blTab2ReservedEventEnable = True
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-保留訂單-顯示全部訂單", Me.Caption, "cmd_Tab2_ShowAll_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_FilterAndSort_Click()
    '排車作業 >> 保留訂單搜尋
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    
    strFormName_FilterAndSort = Me.Name
    strRSName_FilterAndSort = "rs_Tab2_ReservedOrders"
    
    If ShowForm_RS_FilterAndSort(rs_Tab2_ReservedOrders, "保留訂單", Me.Tag) = False Then
        MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Me.WindowState = 2

End Sub

Private Sub cmd_Tab2_Reset_Click()
    '排車作業 >> 取消保留訂單篩選排序
    '移除篩選條件，重設排序依據
     blTab2ReservedEventEnable = False
     rs_Tab2_ReservedOrders.Filter = adFilterNone
     rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
     blTab2ReservedEventEnable = True
End Sub

Private Sub cmd_Tab3_Cancel_Click()
    txt_Tab3_DeliveryDate_Start.Text = ""
    txt_Tab3_DeliveryDate_End.Text = ""
    Set gd_Tab3_OrderSum.DataSource = Nothing
    cmd_Tab3_Cancel.Enabled = False
    cmd_Tab3_Excel.Enabled = False
End Sub

Private Sub cmd_Tab3_Excel_Click()
    If gd_Tab3_OrderSum Is Nothing Then Exit Sub
    If rs_Tab3_OrderSum.RecordCount = 0 Then Exit Sub
    On Error GoTo err_Handle
    '將資料寫入excel檔
    Dim MyXlsApp As Excel.Application   '開啟excel檔
    Dim objFld As Field
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '新增Wookbooks
    MyXlsApp.Workbooks.Add
    '新增Sheets
'    MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "訂單分析"
    MyXlsApp.ActiveSheet.Name = "訂單分析"
    
    Dim i As Integer
    i = 1
    ''出車日,路編,車號,司機,領款人,單號,單位,數量,應收單價,應付單價,其他金額,原因,實收,實付,起點,迄點,備註
    MyXlsApp.Cells(i, 1).Value = "地區"
    MyXlsApp.Cells(i, 2).Value = "才積"
    MyXlsApp.Cells(i, 3).Value = "重量"
    
    i = i + 1
    rs_Tab3_OrderSum.MoveFirst
    '日期,車號,單號,班別,借出,還入
    Do While Not rs_Tab3_OrderSum.EOF
        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab3_OrderSum.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab3_OrderSum.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab3_OrderSum.Fields(3))
        rs_Tab3_OrderSum.MoveNext
        i = i + 1
    Loop
    
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A1:C" & i - 1).Select
    MyXlsApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    MyXlsApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With MyXlsApp.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    'Exit Sub
    
    '未排車訂單
    '將資料寫入excel檔
'    MyXlsApp.Sheets("Sheet2").Select
'    MyXlsApp.Sheets("Sheet2").Name = "未排車訂單分析"
    MyXlsApp.ActiveSheet.Name = "未排車訂單分析"
    
    'Dim i As Integer
    i = 1
    ''出車日,路編,車號,司機,領款人,單號,單位,數量,應收單價,應付單價,其他金額,原因,實收,實付,起點,迄點,備註
    MyXlsApp.Cells(i, 1).Value = "地區"
    MyXlsApp.Cells(i, 2).Value = "才積"
    MyXlsApp.Cells(i, 3).Value = "重量"
    
    i = i + 1
    rs_Tab3_Trp02wSum.MoveFirst
    '日期,車號,單號,班別,借出,還入
    Do While Not rs_Tab3_Trp02wSum.EOF
        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab3_Trp02wSum.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab3_Trp02wSum.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab3_Trp02wSum.Fields(3))
        rs_Tab3_Trp02wSum.MoveNext
        i = i + 1
    Loop
    
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A1:C" & i - 1).Select
    MyXlsApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    MyXlsApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With MyXlsApp.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "排車作業-訂單分析", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_Query_Click()
    '訂單分析
    Dim strSubwhere, str_Where, str_SQl1 As String
    strSubwhere = ""
    '送貨日期
    txt_Tab3_DeliveryDate_Start.Text = Trim(txt_Tab3_DeliveryDate_Start.Text)
    txt_Tab3_DeliveryDate_End.Text = Trim(txt_Tab3_DeliveryDate_End.Text)
    strSubwhere = ""
    If Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
        strSubwhere = "  Between '" & txt_Tab3_DeliveryDate_Start.Text & "' and '" & txt_Tab3_DeliveryDate_End.Text & "' "
    ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) = 0 Then
        strSubwhere = "  = '" & txt_Tab3_DeliveryDate_Start.Text & "' "
    ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) = 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
        strSubwhere = "  = '" & txt_Tab3_DeliveryDate_End.Text & "' "
    End If
    If Len(strSubwhere) > 0 Then
        If Len(str_Where) = 0 Then
           str_Where = strSubwhere
        Else
           str_Where = str_Where & " and " & strSubwhere
        End If
    End If

    If Len(str_Where) = 0 Then
        msg_text = "請輸入日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        str_SQL = "select case when t1m.AREA_CODE='C' then '台中物流'  when t1m.AREA_CODE='S' then '高雄' else '台南' end as 地區, " & _
                "round(sum(cast(s.Busr4 as float)*od.OriginalQty),0) as 才積,round(sum(cast(s.Stdgrosswgt as float)*od.OriginalQty),0) as 重量 " & _
                "from orderdetail od " & _
                "inner join gv_skuxpack s on od.sku=s.sku and s.storerkey = od.storerkey " & _
                "inner join orders o on od.orderkey=o.orderkey " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('C','S','S1')  and convert(char(8),o.DeliveryDate,112) " & str_Where & " " & _
                "group by t1m.AREA_CODE " & _
                "Union " & _
                "select  case when  left(t1m.SHORT_NAME,2)  in ('明達','黑百','金藏')  then '屏東經銷'  when  left(t1m.SHORT_NAME,2)  in ('宏晉','宏縉','厚陞')  then '台南經銷' else '中區經銷' end as 地區, " & _
                "round(sum(cast(s.Busr4 as float)*od.OriginalQty),0) as 才積,round(sum(cast(s.Stdgrosswgt as float)*od.OriginalQty),0) as 重量 " & _
                "from orderdetail od " & _
                "inner join gv_skuxpack s on od.sku=s.sku and s.storerkey = od.storerkey " & _
                "inner join orders o on od.orderkey=o.orderkey " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('w') " & _
                "and  left(t1m.SHORT_NAME,2)  in ('生葆','宏晉','宏縉','明達','黑百','金藏','展城','厚陞','大德','力客') " & _
                "and convert(char(8),o.DeliveryDate,112) " & str_Where & " " & _
                "group by case when  left(t1m.SHORT_NAME,2)  in ('明達','黑百','金藏')  then '屏東經銷'  when  left(t1m.SHORT_NAME,2)  in ('宏晉','宏縉','厚陞')  then '台南經銷' else '中區經銷' end "
                
        str_SQl1 = "select case when t1m.AREA_CODE='C' then '台中物流'  when t1m.AREA_CODE='S' then '高雄' else '台南' end as 地區, " & _
                "round(sum(cast(s.Busr4 as float)*od.ORDER_QTY),0) as 才積,round(sum(cast(s.Stdgrosswgt as float)*od.ORDER_QTY),0) as 重量 " & _
                "from trp03w od " & _
                "inner join gv_skuxpack s on od.PRODUCT_NO=s.sku and s.storerkey = od.storerkey " & _
                "inner join trp02w o on od.RECEIPT_NO=o.RECEIPT_NO " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('C','S','S1')  and  convert(char(8),o.ARRIVE_DATE,112) " & str_Where & "  " & _
                "group by t1m.AREA_CODE " & _
                "Union " & _
                "select  case when  left(t1m.SHORT_NAME,2)  in ('明達','黑百','金藏')  then '屏東經銷'  when  left(t1m.SHORT_NAME,2)  in ('宏晉','宏縉','厚陞')  then '台南經銷' else '中區經銷' end as 地區, " & _
                "round(sum(cast(s.Busr4 as float)*od.ORDER_QTY),0) as 才積,round(sum(cast(s.Stdgrosswgt as float)*od.ORDER_QTY),0) as 重量 " & _
                "from trp03w od " & _
                "inner join gv_skuxpack s on od.PRODUCT_NO=s.sku and s.storerkey = od.storerkey " & _
                "inner join trp02w o on od.RECEIPT_NO=o.RECEIPT_NO " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('w') " & _
                "and  left(t1m.SHORT_NAME,2)  in ('生葆','宏晉','宏縉','明達','黑百','金藏','展城','厚陞','大德','力客') " & _
                "and  convert(char(8),o.ARRIVE_DATE,112) " & str_Where & "  " & _
                "group by case when  left(t1m.SHORT_NAME,2)  in ('明達','黑百','金藏')  then '屏東經銷'  when  left(t1m.SHORT_NAME,2)  in ('宏晉','宏縉','厚陞')  then '台南經銷' else '中區經銷' end"
    End If
    On Error GoTo err_Handle
    
    '訂單分析
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "查詢結果：無符合搜尋條件之排車資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Set gd_Tab3_OrderSum.DataSource = Nothing
        Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_OrderSum)
    Set gd_Tab3_OrderSum.DataSource = rs_Tab3_OrderSum
    tmp_Rs.Close
    With gd_Tab3_OrderSum
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000       '地區
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1200       '才積
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 1200      '重量
        .Columns(3).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    '未排車訂單分析
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQl1, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "查詢結果：無未排車之訂單"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Set gd_Tab3_Trp02wSum.DataSource = Nothing
        Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_Trp02wSum)
    Set gd_Tab3_Trp02wSum.DataSource = rs_Tab3_Trp02wSum
    tmp_Rs.Close
    With gd_Tab3_Trp02wSum
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000       '地區
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1200       '才積
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 1200      '重量
        .Columns(3).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    
    cmd_Tab3_Cancel.Enabled = True
    cmd_Tab3_Excel.Enabled = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "排車作業-定單分析", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDeliveryDateFix_Click()
    If rs_Tab1_Route.EOF Then Exit Sub
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select * from trp02t t2 where len(rtrim(t2.description)) > 0 and t2.route_no = '" & rs_Tab1_Route("路線編號") & "' "
    tmp_Rs.Open str_SQL, cn
    If tmp_Rs.EOF Then tmp_Rs.Close: Exit Sub
    tmp_Rs.Close
    strDeliveryDateFiRouteNo = rs_Tab1_Route.Fields("路線編號").Value
    frm_DeliveryDateFix.Show vbModal
    
End Sub



Private Sub DateS_Click()
    
    If Trim(DateS.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(DateS.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(DateS.Text, 4) & "/" & Mid(DateS.Text, 5, 2) & "/" & Right(DateS.Text, 2))
        End If
    End If
    mvDate.Left = fam_SelectedOrders.Left + DateS.Left
    mvDate.Top = fam_SelectedOrders.Top + DateS.Top + DateS.Height
    mvDate.Visible = True
End Sub

Private Sub dg_Tab0_SelectedOrders_HeadClick(ByVal ColIndex As Integer)
    '以滑鼠點選 [已選取訂單] dg_Tab0_SelectedOrders 欄位標題區：排序欄位選取
    Dim OrderFieldName As String
    If TypeName(rs_Tab0_SelectedOrders) <> "Nothing" Then
        OrderFieldName = "[" & dg_Tab0_SelectedOrders.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_Tab0_SelectedOrders.Sort = OrderFieldName & " DESC "
        Else
            strOrder = "ASC"
            rs_Tab0_SelectedOrders.Sort = OrderFieldName & " ASC "
        End If
    End If
End Sub

Private Sub dg_Tab0_SelectedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '排車作業 >> 已選取訂單 DBGrid
    If blTab0SelectedOrderEventEnable Then
        With dg_Tab0_SelectedOrders
            '反白顯示選取之資料列
            If Not rs_Tab0_SelectedOrders.EOF Then
                dg_Tab0_SelectedOrders.SelBookmarks.Add rs_Tab0_SelectedOrders.Bookmark
            End If
        End With
    End If
    
'If Not rs_Tab0_SelectedOrders.EOF Then
'    txt_Tab0_DockNo = rs_Tab0_SelectedOrders("area")
'Else
'    txt_Tab0_DockNo = ""
'End If

End Sub

Private Sub dg_Tab1_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '路線編號列表：整行選取
    If blTab1RouteEventEnable Then
        If Not rs_Tab1_Route.EOF Then
            dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
            rs_Tab1_RouteOrders.Filter = " 路線編號 = '" & rs_Tab1_Route.Fields("路線編號").Value & "' "
        End If
    End If
End Sub

Private Sub dg_Tab2_ReservedOrders_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim objDataGrid As Object: Set objDataGrid = dg_Tab2_ReservedOrders
If Len(objDataGrid.Columns(ColIndex).DataField) = 0 Or objDataGrid.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, "保留訂單" & objDataGrid.Name, objDataGrid.Columns(ColIndex).DataField, objDataGrid.Columns(ColIndex).Width
End Sub

Private Sub dg_Tab2_ReservedOrders_HeadClick(ByVal ColIndex As Integer)
    '以滑鼠點選 [保留訂單] dg_Tab2_ReservedOrder 欄位標題區：排序欄位選取
    Dim OrderFieldName As String
    If TypeName(rs_Tab2_ReservedOrders) <> "Nothing" Then
        OrderFieldName = "[" & dg_Tab2_ReservedOrders.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_Tab2_ReservedOrders.Sort = OrderFieldName & " DESC "
        Else
            strOrder = "ASC"
            rs_Tab2_ReservedOrders.Sort = OrderFieldName & " ASC "
        End If
    End If
End Sub

Private Sub dg_Tab2_ReservedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '排車作業 >> 保留訂單 DBGrid
    If rs_Tab2_ReservedOrders.EOF Then Exit Sub
    If blTab2ReservedEventEnable Then
        With dg_Tab2_ReservedOrders
            '點一下選取，續點則 [取消]
            If Trim(rs_Tab2_ReservedOrders.Fields(1).Value) = "" Then
                rs_Tab2_ReservedOrders.Fields(1).Value = "V"
            Else
                rs_Tab2_ReservedOrders.Fields(1).Value = " "
            End If
            '反白顯示選取之資料列
            If Not rs_Tab2_ReservedOrders.EOF Then
                dg_Tab2_ReservedOrders.SelBookmarks.Add rs_Tab2_ReservedOrders.Bookmark
            End If
        End With
    End If
End Sub

Private Sub dg_TRP02W_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim objDataGrid As Object: Set objDataGrid = dg_TRP02W
If Len(objDataGrid.Columns(ColIndex).DataField) = 0 Or objDataGrid.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, "待排車訂單" & objDataGrid.Name, objDataGrid.Columns(ColIndex).DataField, objDataGrid.Columns(ColIndex).Width
End Sub

Private Sub dg_TRP02W_DblClick()
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub
If Len(RTrim(txtReceipt_no)) = 0 Then Exit Sub
frm_OrderDetail.Show vbModal
End Sub

Private Sub dg_TRP02W_HeadClick(ByVal ColIndex As Integer)
    '以滑鼠點選 [待排車訂單] dg_TRP02W 欄位標題區：排序欄位選取
    Dim OrderFieldName As String
    If TypeName(rs_TRP02W) <> "Nothing" Then
        '避免產生 [選取] 的動作
        blTRP02WEventEnable = False
        OrderFieldName = "[" & dg_TRP02W.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_TRP02W.Sort = OrderFieldName & " DESC "
            strSourceOrderBy = OrderFieldName & " desc "
        Else
            strOrder = "ASC"
            rs_TRP02W.Sort = OrderFieldName & " ASC "
            strSourceOrderBy = OrderFieldName & " asc "
        End If
        blTRP02WEventEnable = True
    End If
End Sub

Private Sub dg_TRP02W_RowColChange(LastRow As Variant, ByVal LastCkmol As Integer)
    On Error GoTo err_Handle
    '排車作業 >> 待排車訂單 DBGrid
    If blTRP02WEventEnable Then
    txtReceipt_no = rs_TRP02W("receipt_no")
        With dg_TRP02W
            '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
            If Trim(rs_TRP02W.Fields(1).Value) = "" Then
                            
'                '未回傳WMS訂單無法加入已回傳路編
'                If rs_Tab0_SelectedOrders.Fields("EXE回傳") = "已回傳" And rs_TRP02W("EXE回傳") = "新建路編" Then MsgBox "未回傳WMS訂單，無法加入已回傳路編!!", 64, "訂單選取": Exit Sub

                rs_TRP02W.Fields(1).Value = "V"
                dbSelectedCount = dbSelectedCount + 1
                '選取小計更新
                dbsrcSelected_OTqty = dbsrcSelected_OTqty + rs_TRP02W.Fields("件數").Value
                dbsrcSelected_Case = dbsrcSelected_Case + rs_TRP02W.Fields("箱數").Value
                dbsrcSelected_Pallet = dbsrcSelected_Pallet + rs_TRP02W.Fields("板數").Value
                dbsrcSelected_Volumn = dbsrcSelected_Volumn + rs_TRP02W.Fields("材積").Value
                dbsrcSelected_Weight = dbsrcSelected_Weight + rs_TRP02W.Fields("重量").Value
                txt_Tab0_srcSelected_OTqty.Text = dbsrcSelected_OTqty
                txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
                txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
            Else
                dbSelectedCount = dbSelectedCount - 1
                rs_TRP02W.Fields(1).Value = " "
                '選取小計更新
                If dbSelectedCount = 0 Then
                    dbsrcSelected_OTqty = 0
                    dbsrcSelected_Case = 0
                    dbsrcSelected_Pallet = 0
                    dbsrcSelected_Volumn = 0
                    dbsrcSelected_Weight = 0
                    txt_Tab0_srcSelected_OTqty.Text = 0
                    txt_Tab0_srcSelected_Case.Text = 0: txt_Tab0_srcSelected_Pallet.Text = 0
                    txt_Tab0_srcSelected_Volumn.Text = 0: txt_Tab0_srcSelected_Weight.Text = 0
                Else
                    dbsrcSelected_OTqty = dbsrcSelected_OTqty - rs_TRP02W.Fields("件數").Value
                    dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP02W.Fields("箱數").Value
                    dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP02W.Fields("板數").Value
                    dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP02W.Fields("材積").Value
                    dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP02W.Fields("重量").Value
                    txt_Tab0_srcSelected_OTqty.Text = dbsrcSelected_OTqty
                    txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
                    txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
                End If
            End If
            '反白顯示選取之資料列
            If Not rs_TRP02W.EOF Then
                dg_TRP02W.SelBookmarks.Add rs_TRP02W.Bookmark
            End If
        End With
    End If
    Exit Sub
err_Handle:
End Sub

Private Sub cmd_Tab0_srcOrderQuery_Click()
    '排車作業 >> 待排車訂單搜尋
    If rs_TRP02W Is Nothing Then Exit Sub
    If rs_TRP02W.RecordCount = 0 Then Exit Sub
    
    strFormName_FilterAndSort = Me.Name
    strRSName_FilterAndSort = "rs_TRP02W"
    
    If ShowForm_RS_FilterAndSort(rs_TRP02W, "待排車訂單", Me.Tag) = False Then
        MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Me.WindowState = 2

End Sub

Private Sub Form_Activate()
    '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
    Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
    msg_title = "排車作業"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '攔截整個表單鍵盤按鍵事件
    '用途：使用者按下 Esc 則不傳回任何資料，且關閉日期選取視窗
    If KeyCode = vbKeyEscape Then
        mvDate.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '設定 Form 大小、位置
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 11475
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    
    '排車作業：待排車訂單
    Call CreateRS_Tab0_TRP02W
    strSourceFilter = adFilterNone
    strSourceOrderBy = " 編號 asc "
    
    '排車作業：已選取之待排車訂單列表 DBGrid 格式設定
    Call CreateRS_Tab0_SelectedOrders
    
    '已產生之路線編號列表
    Call CreateRS_Tab1_Route
    Call CreateRS_Tab1_RouteOrders
    
    '保留訂單列表
    Call CreateRS_Tab2_ReservedOrders
    blTab2ReservedEventEnable = True
    SSTab1.Tab = 0
    
    Dim rsTmp As New ADODB.Recordset
    With rsTmp
        .CursorLocation = 3
        '貨主
        .Open "select distinct storerkey from " & strWMSDB & "..storer (nolock) where type = 1 order by storerkey ", cn
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Storerkey.AddItem RTrim(rsTmp("storerkey"))
            rsTmp.MoveNext
        Loop
        .Close
        
        '區碼
        .Open "select distinct area_code from trp01m(nolock) where len(isnull(area_code,'')) > 0 order by area_code ", cn
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Area_Code.AddItem RTrim(rsTmp("area_code"))
            rsTmp.MoveNext
        Loop
        .Close
    End With
    
    DateS = Format(Now() + 1, "YYYYMMDD")
    DateE = Format(Now() + 1, "YYYYMMDD")
    
End Sub

Private Sub Form_Resize()
    '視窗大小變動
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    If Me.ScaleHeight < dbsrcFormHeight Then
        '變小
        SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
        SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
        
        fam_SelectedOrders.Width = fam_SelectedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dg_Tab0_SelectedOrders.Width = dg_Tab0_SelectedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        fam_SrcOrders.Height = fam_SrcOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        fam_SrcOrders.Width = fam_SrcOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dg_TRP02W.Height = dg_TRP02W.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_TRP02W.Width = dg_TRP02W.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        Frame3.Left = Frame3.Left - (dbsrcFormWidth - Me.ScaleWidth)
        Frame4.Left = Frame4.Left - (dbsrcFormWidth - Me.ScaleWidth)
        dg_Tab1_Route.Width = dg_Tab1_Route.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        dg_Tab1_RouteOrders.Height = dg_Tab1_RouteOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_Tab1_RouteOrders.Width = dg_Tab1_RouteOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        dg_Tab2_ReservedOrders.Height = dg_Tab2_ReservedOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_Tab2_ReservedOrders.Width = dg_Tab2_ReservedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Remove.Left = cmd_Tab2_Remove.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_FilterAndSort.Left = cmd_Tab2_FilterAndSort.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Reset.Left = cmd_Tab2_Reset.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_ShowAll.Left = cmd_Tab2_ShowAll.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Delete.Left = cmd_Tab2_Delete.Left - (dbsrcFormWidth - Me.ScaleWidth)
        dbsrcFormHeight = Me.ScaleHeight
        dbsrcFormWidth = Me.ScaleWidth
    Else
       SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
       SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
       
       fam_SelectedOrders.Width = fam_SelectedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       dg_Tab0_SelectedOrders.Width = dg_Tab0_SelectedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       fam_SrcOrders.Width = fam_SrcOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       fam_SrcOrders.Height = fam_SrcOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_TRP02W.Height = dg_TRP02W.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_TRP02W.Width = dg_TRP02W.Width + (Me.ScaleWidth - dbsrcFormWidth)
       
       Frame3.Left = Frame3.Left + (Me.ScaleWidth - dbsrcFormWidth)
       dg_Tab1_Route.Width = dg_Tab1_Route.Width + (Me.ScaleWidth - dbsrcFormWidth)
       Frame4.Left = Frame4.Left + (Me.ScaleWidth - dbsrcFormWidth)
       
       dg_Tab1_RouteOrders.Height = dg_Tab1_RouteOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_Tab1_RouteOrders.Width = dg_Tab1_RouteOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       
       dg_Tab2_ReservedOrders.Height = dg_Tab2_ReservedOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_Tab2_ReservedOrders.Width = dg_Tab2_ReservedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Remove.Left = cmd_Tab2_Remove.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_FilterAndSort.Left = cmd_Tab2_FilterAndSort.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Reset.Left = cmd_Tab2_Reset.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_ShowAll.Left = cmd_Tab2_ShowAll.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Delete.Left = cmd_Tab2_Delete.Left + (Me.ScaleWidth - dbsrcFormWidth)
       
       dbsrcFormHeight = Me.ScaleHeight
       dbsrcFormWidth = Me.ScaleWidth
    End If
End Sub

Private Sub Form_Terminate()
    '更新 Menu [視窗]→[已開視窗清單]
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '從記憶體中移除表單，藉此引起 [Terminate] 事件
    Set frm_OP_TRPPlan = Nothing
End Sub

Private Sub CreateRS_Tab0_TRP02W()
    '排車作業：待排車訂單
    Call ReDim_Recordset(rs_TRP02W)
    With rs_TRP02W
        .Fields.Append "編號", adDouble
        .Fields.Append "＊", adVarChar, 2
        .Fields.Append "送貨日", adVarChar, 10
        .Fields.Append "訂單編號", adVarChar, 60
        .Fields.Append "件數", adDouble
        .Fields.Append "箱數", adDouble
        .Fields.Append "板數", adDouble
        .Fields.Append "材積", adDouble
        .Fields.Append "重量", adDouble
        .Fields.Append "客戶編號", adVarChar, 30
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "客戶簡稱", adVarChar, 60
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "型態", adVarChar, 10
        .Fields.Append "運送地址", adVarChar, 120
        .Fields.Append "訂單備註", adVarChar, 300
        .Fields.Append "啤酒允收期", adVarChar, 4
        .Fields.Append "配送倉別", adVarChar, 20
        .Fields.Append "車種", adVarChar, 10
        .Fields.Append "特殊需求1", adVarChar, 60
        .Fields.Append "特殊需求2", adVarChar, 60
        .Fields.Append "急單", adVarChar, 10
        .Fields.Append "專車", adVarChar, 10
        .Fields.Append "冷藏", adVarChar, 10
        .Fields.Append "Receipt_No", adVarChar, 10
        .Fields.Append "C_Receipt_No", adVarChar, 10
        .Fields.Append "貨主單號", adVarChar, 40
        .Fields.Append "EXE回傳", adVarChar, 20
        .Fields.Append "客戶名稱", adVarChar, 120
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '不需連接物件
    End With
    Set dg_TRP02W.DataSource = rs_TRP02W
    '設定顯示欄位
    With dg_TRP02W
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .RowHeight = 250
'        .Columns(0).Width = 500         '序號
'        .Columns(0).Alignment = dbgCenter
'        .Columns(1).Width = 300         '選取識別欄位
'        .Columns(1).Alignment = dbgCenter
'        .Columns(2).Width = 800         '送貨日
'        .Columns(2).Alignment = dbgCenter
'        .Columns(3).Width = 2500        '訂單編號：訂單編號+貨主單號+貨主
'        .Columns(3).Alignment = dbgLeft
'        .Columns(4).Width = 600         '箱數
'        .Columns(4).Alignment = dbgRight
'        .Columns(5).Width = 600         '板數
'        .Columns(5).Alignment = dbgRight
'        .Columns(6).Width = 600         '材積
'        .Columns(6).Alignment = dbgRight
'        .Columns(7).Width = 600         '重量
'        .Columns(7).Alignment = dbgRight
'        .Columns(8).Width = 1100        '客戶編號
'        .Columns(8).Alignment = dbgLeft
'        .Columns(9).Width = 400         'zip
'        .Columns(9).Alignment = dbgCenter
'        .Columns(10).Width = 1000       '客戶簡稱
'        .Columns(10).Alignment = dbgLeft
'        .Columns(11).Width = 450        'Area_Code
'        .Columns(11).Alignment = dbgCenter
'        .Columns(12).Width = 450        '型態：Priority
'        .Columns(12).Alignment = dbgCenter
'        .Columns(13).Width = 2500       '運送地址
'        .Columns(13).Alignment = dbgLeft
'        .Columns(14).Width = 1400       '訂單備註
'        .Columns(14).Alignment = dbgLeft
'        .Columns(15).Width = 1000       '啤酒允收期
'        .Columns(15).Alignment = dbgLeft
'        .Columns(16).Width = 500        '車種
'        .Columns(16).Alignment = dbgCenter
'        .Columns(17).Width = 1500       '特殊需求1
'        .Columns(17).Alignment = dbgLeft
'        .Columns(18).Width = 1500       '特殊需求2
'        .Columns(18).Alignment = dbgLeft
'        .Columns(19).Width = 500        '急單
'        .Columns(19).Alignment = dbgCenter
'        .Columns(20).Width = 500        '專車
'        .Columns(20).Alignment = dbgCenter
'        .Columns(21).Width = 500        '冷藏
'        .Columns(21).Alignment = dbgCenter
'        .Columns(22).Width = 1100       'Receipt_No
'        .Columns(22).Alignment = dbgLeft
'        .Columns(23).Width = 1100       'C_Receipt_No
'        .Columns(23).Alignment = dbgLeft
'        .Columns(24).Width = 900        '貨主單號
'        .Columns(24).Alignment = dbgLeft
'        .Columns(25).Width = 1500       'EXE回傳
'        .Columns(25).Alignment = dbgLeft
'        .Columns(26).Width = 1500       '客戶名稱
'        .Columns(26).Alignment = dbgLeft
    End With
    SetDataGridColWidth "待排車訂單", dg_TRP02W
End Sub

Private Sub CreateRS_Tab0_SelectedOrders()
    '排車作業：已選取之待排車訂單列表
    Call ReDim_Recordset(rs_Tab0_SelectedOrders)
    With rs_Tab0_SelectedOrders
        .Fields.Append "編號", adDouble
        .Fields.Append "送貨日", adVarChar, 20
        .Fields.Append "訂單編號", adVarChar, 60
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "急單", adVarChar, 10
        .Fields.Append "型態", adVarChar, 20
        .Fields.Append "客戶簡稱", adVarChar, 120
        .Fields.Append "件數", adDouble
        .Fields.Append "箱數", adDouble
        .Fields.Append "板數", adDouble
        .Fields.Append "材積", adDouble
        .Fields.Append "重量", adDouble
        .Fields.Append "車種", adVarChar, 10
        .Fields.Append "訂單備註", adVarChar, 300
        .Fields.Append "特殊需求1", adVarChar, 60
        .Fields.Append "特殊需求2", adVarChar, 60
        .Fields.Append "Receipt_No", adVarChar, 10
        .Fields.Append "EXE回傳", adVarChar, 20
        .Fields.Append "客戶名稱", adVarChar, 120
        .Fields.Append "運送地址", adVarChar, 120
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '不需連接物件
    End With
    Set dg_Tab0_SelectedOrders.DataSource = rs_Tab0_SelectedOrders
    '設定顯示欄位
    With dg_Tab0_SelectedOrders
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 500        '編號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 800         '送貨日
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 2500        '訂單編號
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 400         'ZIP
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 450         'Area
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 450         '型態：Orders.Priority
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 450         '型態：Orders.Priority
        .Columns(6).Alignment = dbgCenter
        .Columns(7).Width = 1000        '客戶簡稱
        .Columns(7).Alignment = dbgLeft
        .Columns(8).Width = 600         '件數
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 600         '箱數
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 600         '板數
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 600         '材積
        .Columns(11).Alignment = dbgRight
        .Columns(12).Width = 600        '重量
        .Columns(12).Alignment = dbgRight
        .Columns(13).Width = 450        '車種
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1200       '訂單備註
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 1500       '特殊需求-1
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1500       '特殊需求-2
        .Columns(16).Alignment = dbgLeft
        .Columns(17).Width = 1000       'Receipt_No
        .Columns(17).Alignment = dbgLeft
        .Columns(18).Width = 1000       'EXE回傳
        .Columns(18).Alignment = dbgLeft
        .Columns(19).Width = 1500       '客戶名稱
        .Columns(19).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab1_Route()
    '排車作業：已編妥之路線編號列表
    Call ReDim_Recordset(rs_Tab1_Route)
    With rs_Tab1_Route
        .Fields.Append "編號", adDouble
        .Fields.Append "路線編號", adVarChar, 10
        .Fields.Append "出車日期", adVarChar, 8
        .Fields.Append "車牌號碼", adVarChar, 10
        .Fields.Append "車次", adDouble
        .Fields.Append "駕駛人", adVarChar, 20
        .Fields.Append "件數", adDouble
        .Fields.Append "箱數", adDouble
        .Fields.Append "板數", adDouble
        .Fields.Append "材積", adDouble
        .Fields.Append "重量", adDouble
        .Fields.Append "車種", adVarChar, 10
        .Fields.Append "碼頭暫存", adVarChar, 10
        .Fields.Append "預計報到日期", adVarChar, 8
        .Fields.Append "預計報到時間", adVarChar, 4
        .Fields.Append "EXE回傳", adVarChar, 20
        .Fields.Append "排車者", adVarChar, 30
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '不需連接物件
    End With
    Set dg_Tab1_Route.DataSource = rs_Tab1_Route
    '設定顯示欄位
    With dg_Tab1_Route
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 500         '編號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000        '路線編號
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 900         '出車日期
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 850         '車牌號碼
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 500         '車次
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 900         '駕駛人
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 700         '件數
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 700         '箱數
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 700         '板數
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 700         '材積
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 700         '重量
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 450        '車種
        .Columns(11).Alignment = dbgCenter
        .Columns(12).Width = 1000       '碼頭暫存
        .Columns(12).Alignment = dbgLeft
        .Columns(13).Width = 1400       '預計車輛報到日期
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1400       '預計車輛報到時間
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 900        'EXE 回傳狀態
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1200       '排車者
        .Columns(16).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab1_RouteOrders()
    '排車作業：已編妥路編之訂單列表
    Call ReDim_Recordset(rs_Tab1_RouteOrders)
    With rs_Tab1_RouteOrders
        .Fields.Append "編號", adDouble
        .Fields.Append "路線編號", adVarChar, 10
        .Fields.Append "送貨日", adVarChar, 20
        .Fields.Append "訂單編號", adVarChar, 60
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "客戶名稱", adVarChar, 120
        .Fields.Append "地址", adVarChar, 300
        .Fields.Append "件數", adDouble
        .Fields.Append "箱數", adDouble
        .Fields.Append "板數", adDouble
        .Fields.Append "材積", adDouble
        .Fields.Append "重量", adDouble
        .Fields.Append "訂單備註", adVarChar, 300
        .Fields.Append "車種", adVarChar, 10
        .Fields.Append "特殊需求1", adVarChar, 60
        .Fields.Append "特殊需求2", adVarChar, 60
        .Fields.Append "Receipt_No", adVarChar, 60
        .Fields.Append "EXE回傳", adVarChar, 20
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "型態", adVarChar, 10
        .Fields.Append "客戶簡稱", adVarChar, 255
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '不需連接物件
    End With
    Set dg_Tab1_RouteOrders.DataSource = rs_Tab1_RouteOrders
    '設定顯示欄位
    With dg_Tab1_RouteOrders
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 500         '編號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1050        '路線編號
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 900         '送貨日
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2150        '訂單編號
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 400         'ZIP
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 1500        '客戶名稱
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1500        '客戶名稱
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 700         '件數
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 700         '箱數
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 700         '板數
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 700         '材積
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 700         '重量
        .Columns(11).Alignment = dbgRight
        .Columns(12).Width = 1500       '訂單備註
        .Columns(12).Alignment = dbgLeft
        .Columns(13).Width = 1200       '車種
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1500       '特殊需求-1
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 1500       '特殊需求-2
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1100       'Receipt_No
        .Columns(16).Alignment = dbgLeft
        .Columns(17).Width = 1100       'EXE回傳
        .Columns(17).Alignment = dbgLeft
        .Columns(18).Width = 450        'Area
        .Columns(18).Alignment = dbgCenter
        .Columns(19).Width = 450        '型態
        .Columns(19).Alignment = dbgCenter
        .Columns(20).Width = 1100       '客戶簡稱
        .Columns(20).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab2_ReservedOrders()
    '排車作業：保留訂單
    Call ReDim_Recordset(rs_Tab2_ReservedOrders)
    With rs_Tab2_ReservedOrders
         .Fields.Append "編號", adDouble
         .Fields.Append "＊", adVarChar, 2
         .Fields.Append "送貨日", adVarChar, 10
         .Fields.Append "訂單編號", adVarChar, 60
         .Fields.Append "件數", adDouble
         .Fields.Append "箱數", adDouble
         .Fields.Append "板數", adDouble
         .Fields.Append "材積", adDouble
         .Fields.Append "重量", adDouble
         .Fields.Append "客戶編號", adVarChar, 30
         .Fields.Append "ZIP", adVarChar, 10
         .Fields.Append "客戶簡稱", adVarChar, 60
         .Fields.Append "Area", adVarChar, 10
         .Fields.Append "型態", adVarChar, 10
         .Fields.Append "運送地址", adVarChar, 120
         .Fields.Append "訂單備註", adVarChar, 300
         .Fields.Append "啤酒允收期", adVarChar, 4
         .Fields.Append "配送倉別", adVarChar, 20
         .Fields.Append "車種", adVarChar, 10
         .Fields.Append "特殊需求1", adVarChar, 60
         .Fields.Append "特殊需求2", adVarChar, 60
         .Fields.Append "急單", adVarChar, 10
         .Fields.Append "專車", adVarChar, 10
         .Fields.Append "冷藏", adVarChar, 10
         .Fields.Append "TMS單號", adVarChar, 10
         .Fields.Append "C_RECEIPT_NO", adVarChar, 10
         .Fields.Append "貨主單號", adVarChar, 40
         .Fields.Append "EXE回傳", adVarChar, 20
         .Fields.Append "客戶名稱", adVarChar, 120
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '不需連接物件
    End With
    Set dg_Tab2_ReservedOrders.DataSource = rs_Tab2_ReservedOrders
    '設定顯示欄位
    With dg_Tab2_ReservedOrders
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .RowHeight = 250
'        .Columns(0).Width = 500         '序號
'        .Columns(0).Alignment = dbgCenter
'        .Columns(1).Width = 300         '選取識別欄位
'        .Columns(1).Alignment = dbgCenter
'        .Columns(2).Width = 800         '送貨日
'        .Columns(2).Alignment = dbgCenter
'        .Columns(3).Width = 2100        '訂單編號：訂單編號+貨主單號+貨主
'        .Columns(3).Alignment = dbgLeft
'        .Columns(4).Width = 600         '箱數
'        .Columns(4).Alignment = dbgRight
'        .Columns(5).Width = 600         '板數
'        .Columns(5).Alignment = dbgRight
'        .Columns(6).Width = 600         '材積
'        .Columns(6).Alignment = dbgRight
'        .Columns(7).Width = 600         '重量
'        .Columns(7).Alignment = dbgRight
'        .Columns(8).Width = 1100        '客戶編號
'        .Columns(8).Alignment = dbgLeft
'        .Columns(9).Width = 400         'zip
'        .Columns(9).Alignment = dbgCenter
'        .Columns(10).Width = 1000       '客戶簡稱
'        .Columns(10).Alignment = dbgLeft
'        .Columns(11).Width = 450        'Area_Code
'        .Columns(11).Alignment = dbgCenter
'        .Columns(12).Width = 450        '型態：Priority
'        .Columns(12).Alignment = dbgCenter
'        .Columns(13).Width = 3000       '運送地址
'        .Columns(13).Alignment = dbgLeft
'        .Columns(14).Width = 1400       '訂單備註
'        .Columns(14).Alignment = dbgLeft
'        .Columns(15).Width = 1000       '啤酒允收期
'        .Columns(15).Alignment = dbgLeft
'        .Columns(16).Width = 500       '車種
'        .Columns(16).Alignment = dbgLeft
'        .Columns(17).Width = 1500       '特殊需求1
'        .Columns(17).Alignment = dbgLeft
'        .Columns(18).Width = 1500       '特殊需求2
'        .Columns(18).Alignment = dbgLeft
'        .Columns(19).Width = 500        '急單
'        .Columns(19).Alignment = dbgCenter
'        .Columns(20).Width = 500        '專車
'        .Columns(20).Alignment = dbgCenter
'        .Columns(21).Width = 500        '冷藏
'        .Columns(21).Alignment = dbgCenter
'        .Columns(22).Width = 1100       'Receipt_No
'        .Columns(22).Alignment = dbgLeft
'        .Columns(23).Width = 1100       'C_Receipt_No
'        .Columns(23).Alignment = dbgLeft
'        .Columns(24).Width = 900        '貨主單號
'        .Columns(24).Alignment = dbgLeft
'        .Columns(25).Width = 1500       'EXE回傳
'        .Columns(25).Alignment = dbgLeft
'        .Columns(26).Width = 1500       '客戶名稱
'        .Columns(26).Alignment = dbgLeft
    End With
        SetDataGridColWidth "保留訂單", dg_Tab2_ReservedOrders
End Sub

Private Sub Calculate_SelectedOrders()
    '作業內容：
    '1.針對已選取訂單列表，依訂單編號重新產生 [編號] 欄位值
    '2.計算已選取訂單之累計資料
    Dim dbSeqNo As Double
    dbSeqNo = 0
    txt_Tab0_Selected_OTqty.Text = ""
    txt_Tab0_Selected_Case.Text = ""
    txt_Tab0_Selected_Pallet.Text = ""
    txt_Tab0_Selected_Volumn.Text = ""
    txt_Tab0_Selected_Weight.Text = ""
    
    rs_Tab0_SelectedOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "Receipt_No asc"  '原始排序，一般資料序號由小至大
    If Not rs_Tab0_SelectedOrders.EOF Then
       rs_Tab0_SelectedOrders.MoveFirst
    Else
        '清出篩選條件，仍無資料者，結束 SubProgram 執行
        Exit Sub
    End If
    Do While Not rs_Tab0_SelectedOrders.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_SelectedOrders.Fields("編號").Value = dbSeqNo
        txt_Tab0_Selected_OTqty.Text = Val(txt_Tab0_Selected_OTqty.Text) + rs_Tab0_SelectedOrders.Fields("件數").Value
        txt_Tab0_Selected_Case.Text = Val(txt_Tab0_Selected_Case.Text) + rs_Tab0_SelectedOrders.Fields("箱數").Value
        txt_Tab0_Selected_Pallet.Text = Val(txt_Tab0_Selected_Pallet.Text) + rs_Tab0_SelectedOrders.Fields("板數").Value
        txt_Tab0_Selected_Volumn.Text = Val(txt_Tab0_Selected_Volumn.Text) + rs_Tab0_SelectedOrders.Fields("材積").Value
        txt_Tab0_Selected_Weight.Text = Val(txt_Tab0_Selected_Weight.Text) + rs_Tab0_SelectedOrders.Fields("重量").Value
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    rs_Tab0_SelectedOrders.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_Tab0_SelectedOrders.EOF Then rs_Tab0_SelectedOrders.MoveFirst
End Sub

Private Sub SelectedOrders_Removeto_TRP02W(ByVal strReceiptNo As String)
    '將指定之 [訂單編號] 加入 [待排車訂單]
    blTRP02WEventEnable = False
    
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
    
    If rs_TRP02W.RecordCount > 0 Then
        rs_TRP02W.Filter = "Receipt_No = '" & strReceiptNo & "'"
        If Not rs_TRP02W.EOF Then
            '訂單編號已存在的話，不進行新增，也不更新
            rs_TRP02W.Filter = adFilterNone
            rs_TRP02W.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
            blTRP02WEventEnable = True
            Exit Sub
        End If
    End If
    
    '取回待排車訂單
    If blRouteModify Then
        '如果是查詢所顯示之有效路編，設定路編異動識別旗標
        blRouteChange = True
        '經由查詢路線編號所得之訂單資料
        str_SQL = "Select 送貨日,訂單編號,件數,箱數,板數,材積,重量,客戶編號,ZIP,客戶名稱,運送地址,訂單備註,車種,特殊需求1,特殊需求2,急單,專車,冷藏,Receipt_No,c_receipt_no,貨主單號,EXE回傳,Area,客戶簡稱,啤酒允收期,配送倉別,型態 " & _
                  "From TRPPlan_RouteQueryOrdersRemove Where Receipt_No = '" & strReceiptNo & "' Order by 訂單編號 "
    Else
        str_SQL = "Select 送貨日,訂單編號,件數,箱數,板數,材積,重量,客戶編號,ZIP,客戶名稱,運送地址,訂單備註,車種,特殊需求1,特殊需求2,急單,專車,冷藏,Receipt_No,c_receipt_no,貨主單號,EXE回傳,Area,客戶簡稱,型態,啤酒允收期,配送倉別 " & _
                  "From TRPPlan_SourceOrder Where Receipt_No = '" & strReceiptNo & "' Order by 訂單編號 "
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：無符合設定條件之待排車訂單資料可以還原回 [待選取訂單]"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        blTRP02WEventEnable = True
        Exit Sub
    End If
    
    rs_TRP02W.AddNew
    rs_TRP02W.Fields("編號").Value = 999
    rs_TRP02W.Fields("送貨日").Value = tmp_Rs.Fields("送貨日").Value
    rs_TRP02W.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
    rs_TRP02W.Fields("件數").Value = tmp_Rs.Fields("件數").Value
    rs_TRP02W.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
    rs_TRP02W.Fields("板數").Value = tmp_Rs.Fields("板數").Value
    rs_TRP02W.Fields("材積").Value = tmp_Rs.Fields("材積").Value
    rs_TRP02W.Fields("重量").Value = tmp_Rs.Fields("重量").Value
    rs_TRP02W.Fields("客戶編號").Value = tmp_Rs.Fields("客戶編號").Value
    rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
    rs_TRP02W.Fields("zip").Value = tmp_Rs.Fields("zip").Value
    rs_TRP02W.Fields("客戶名稱").Value = tmp_Rs.Fields("客戶名稱").Value
    rs_TRP02W.Fields("運送地址").Value = tmp_Rs.Fields("運送地址").Value
    rs_TRP02W.Fields("訂單備註").Value = tmp_Rs.Fields("訂單備註").Value
    rs_TRP02W.Fields("車種").Value = tmp_Rs.Fields("車種").Value
    rs_TRP02W.Fields("特殊需求1").Value = tmp_Rs.Fields("特殊需求1").Value
    rs_TRP02W.Fields("特殊需求2").Value = tmp_Rs.Fields("特殊需求2").Value
    rs_TRP02W.Fields("急單").Value = tmp_Rs.Fields("急單").Value
    rs_TRP02W.Fields("專車").Value = tmp_Rs.Fields("專車").Value
    rs_TRP02W.Fields("冷藏").Value = tmp_Rs.Fields("冷藏").Value
    rs_TRP02W.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
    rs_TRP02W.Fields("C_Receipt_No").Value = tmp_Rs.Fields("C_Receipt_No").Value
    rs_TRP02W.Fields("貨主單號").Value = tmp_Rs.Fields("貨主單號").Value
    rs_TRP02W.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
    rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
    rs_TRP02W.Fields("客戶簡稱").Value = tmp_Rs.Fields("客戶簡稱").Value & ""
    rs_TRP02W.Fields("型態").Value = tmp_Rs.Fields("型態").Value
        rs_TRP02W.Fields("啤酒允收期").Value = tmp_Rs.Fields("啤酒允收期").Value & ""
    rs_TRP02W.Fields("配送倉別").Value = tmp_Rs.Fields("配送倉別").Value
    rs_TRP02W.Update
    tmp_Rs.Close
    
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "訂單編號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_TRP02W.EOF Then rs_TRP02W.MoveFirst
    blTRP02WEventEnable = True
End Sub

Private Sub ReSet_TRP02W_SeqNo()
    '重新產生 [待排車訂單] 之 [編號] 欄位值
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "訂單編號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_TRP02W.EOF Then rs_TRP02W.MoveFirst
    
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_TRP02W.EOF
        dbSeqNo = dbSeqNo + 1
        rs_TRP02W.Fields("編號").Value = dbSeqNo
        rs_TRP02W.MoveNext
    Loop
    If rs_TRP02W.RecordCount > 0 Then rs_TRP02W.MoveFirst
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
End Sub



Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    '日期選取
    Select Case mvDate.Tag
           Case "出車日期"
                txt_Tab0_TRPDate.Text = Format(mvDate.Value, "yyyymmdd")
           Case "預計報到日期"
                txt_Tab0_CarCheckInDate.Text = Format(mvDate.Value, "yyyymmdd")
           Case "訂單分析起"
                txt_Tab3_DeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
           Case "訂單分析迄"
                txt_Tab3_DeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
    End Select
    mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.mvDate.Visible = False
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_Tab0_CarCheckInDate_Click()
    '排車作業 >> 預計報到日期
    If Trim(txt_Tab0_CarCheckInDate.Text) = "" Then
        mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab0_CarCheckInDate.Text, 4) & "/" & Mid(txt_Tab0_CarCheckInDate.Text, 5, 2) & "/" & Right(txt_Tab0_CarCheckInDate.Text, 2))
        End If
    End If
    mvDate.Left = fam_RouteData.Left + txt_Tab0_CarCheckInDate.Left
    mvDate.Top = fam_RouteData.Top + txt_Tab0_CarCheckInDate.Top + txt_Tab0_CarCheckInDate.Height
    mvDate.Tag = "預計報到日期"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab0_CarCheckInDate_KeyPress(KeyAscii As Integer)
    '預計報到日期
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '不允許輸入字元
              KeyAscii = 0
         Case vbKeyReturn
              KeyAscii = 0
              txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
              txt_Tab0_CarCheckInTime.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_CarCheckInTime_KeyPress(KeyAscii As Integer)
    '預計報到時間
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '不允許輸入字元
              KeyAscii = 0
    End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_KeyPress(KeyAscii As Integer)
    '排車作業 >> 車牌號碼
    Select Case KeyAscii
           Case 97 To 122   '轉換為大寫字元
                KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_LostFocus()  'daniel--20040928<防只user輸入錯誤之車號>
    If Len(txt_Tab0_DeliveryCarNo.Text) = 0 Then Exit Sub
    str_SQL = "Select Vehicle_ID_No from trp09m where Vehicle_ID_No='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "' "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        'tmp_rs.Close
        msg_text = "無此車號資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SetFocus
    End If
    tmp_Rs.Close
End Sub

Private Sub txt_Tab0_DockNo_KeyPress(KeyAscii As Integer)
    '碼頭暫存
    Select Case KeyAscii
           Case 97 To 122   '轉換為大寫字元
                KeyAscii = KeyAscii - 32
           Case vbKeyReturn
                KeyAscii = 0
                txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
                txt_Tab0_CarCheckInDate.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_RouteNo_KeyPress(KeyAscii As Integer)
    '排車作業 >> 路線編號
    Select Case KeyAscii
        Case 97 To 122     '小寫字元改為大寫字元
             KeyAscii = KeyAscii - 32
        Case vbKeyReturn
             cmd_Tab0_Query.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_TRPDate_Click()
    '排車作業 >> 出車日期
    If Trim(txt_Tab0_TRPDate.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2))
        End If
    End If
    mvDate.Left = fam_SelectedOrders.Left + txt_Tab0_TRPDate.Left
    mvDate.Top = fam_SelectedOrders.Top + txt_Tab0_TRPDate.Top + txt_Tab0_TRPDate.Height
    mvDate.Tag = "出車日期"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab0_TRPDate_KeyPress(KeyAscii As Integer)
    '排車作業 > [出車日期] 資料格式：yyyymmdd
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '不允許輸入字元
              KeyAscii = 0
         Case vbKeyReturn
              If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
                 msg_text = "出車日期：" & funRtn_msg
                 MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                 txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
                 Exit Sub
              Else
                 cmd_Tab0_SelectCar.SetFocus
              End If
    End Select
End Sub

Public Sub frm_OP_TRPPlan_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
    '表單公用副程式，由 frm_RS_FilterAndSort 表單呼叫
    '傳入值：strCode      動作識別碼
    '                     [FILTER] 自訂篩選    [SORT] 排序
    '        strReturn    篩選 or 排序 之設定字串
    
    Select Case strCode
           Case "FILTER"  '自訂篩選
                Select Case UCase(strRSName_FilterAndSort)
                       Case "RS_TRP02W"                '待排車訂單資料
                            blTRP02WEventEnable = False
                            '篩選已選取者：取消選取
                            rs_TRP02W.Filter = "＊='V'"
                            If Not rs_TRP02W.EOF Then
                               Do While Not rs_TRP02W.EOF
                                  rs_TRP02W.Fields(1).Value = " "
                                  rs_TRP02W.MoveNext
                               Loop
                            End If
                            rs_TRP02W.Filter = adFilterNone
                            rs_TRP02W.Filter = strReturn
                            strSourceFilter = strReturn
                            If rs_TRP02W.RecordCount = 0 Then
                               msg_text = "抱歉ㄟ，找不到符合條件的訂單喔"
                               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                               rs_TRP02W.Filter = adFilterNone
                               strSourceFilter = adFilterNone
                               rs_TRP02W.Sort = strSourceOrderBy   '還原排序方式
                               blTRP02WEventEnable = True
                               Exit Sub
                            End If
                            '重新計算 [待排車列表] 的總計資訊
                            Call ReCaculate_OrderSum
                            blTRP02WEventEnable = True
                       Case "RS_TAB2_RESERVEDORDERS"   '保留訂單
                            blTab2ReservedEventEnable = False
                            rs_Tab2_ReservedOrders.Filter = adFilterNone
                            rs_Tab2_ReservedOrders.Filter = strReturn
                            If rs_Tab2_ReservedOrders.RecordCount = 0 Then
                               msg_text = "抱歉ㄟ，找不到符合條件的保留訂單喔"
                               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                               rs_Tab2_ReservedOrders.Filter = adFilterNone
                               rs_Tab2_ReservedOrders.Sort = strSourceOrderBy   '還原排序方式
                               blTRP02WEventEnable = True
                               Exit Sub
                            End If
                            blTab2ReservedEventEnable = True
                       
                End Select
           Case "SORT"    '排序
                Select Case UCase(strRSName_FilterAndSort)
                       Case "RS_TRP02W"               '待排車訂單資料
                            If rs_TRP02W.EOF Then Exit Sub
                            blTRP02WEventEnable = False
                            rs_TRP02W.Sort = strReturn
                            strSourceOrderBy = strReturn
                            blTRP02WEventEnable = True
                       Case "RS_TAB2_RESERVEDORDERS"   '保留訂單
                            If rs_Tab2_ReservedOrders.EOF Then Exit Sub
                            blTab2ReservedEventEnable = False
                            rs_Tab2_ReservedOrders.Sort = strReturn
                            strSourceOrderBy = strReturn
                            blTab2ReservedEventEnable = True
                End Select
    End Select
End Sub

Private Sub txt_Tab1_RouteNo_KeyPress(KeyAscii As Integer)
    '路線編號列表 >> 路線編號
    Select Case KeyAscii
         Case 97 To 122   '轉換大寫字元
              KeyAscii = KeyAscii - 32
         Case vbKeyReturn
              cmd_Tab1_RouteNoQuery.SetFocus
    End Select
End Sub

Private Sub Clear_RouteData()
    '排車作業：清除路線編號資料欄位
    blRouteModify = False
    strDispRouteNo = ""
    blRouteChange = False
    
    blTab0SelectedOrderEventEnable = False
    '排車作業：已選取之待排車訂單列表 DBGrid 格式設定
    Call CreateRS_Tab0_SelectedOrders
    '重新計算已選取訂單：箱數，板數，材積，重量 + 編號重新產生
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    txt_Tab0_TRPDate.Text = ""
    txt_Tab0_DeliveryCarNo.Text = ""
    txt_Tab0_DockNo.Text = ""
    txt_Tab0_CarCheckInDate.Text = ""
    txt_Tab0_CarCheckInTime.Text = ""
    txt_Tab0_DeliveryCompany.Text = ""
    txt_Tab0_DeliveryDriver.Text = ""
    txt_Tab0_DeliveryPhone.Text = ""
    txt_Tab0_DeliveryCarType.Text = ""
    txt_Tab0_Selected_Case.Text = ""
    txt_Tab0_Selected_Pallet.Text = ""
    txt_Tab0_Selected_Volumn.Text = ""
    txt_Tab0_Selected_Weight.Text = ""
End Sub

Private Function RouteData_Check() As Boolean
    Dim Str_D_Orderkey As String
    Str_D_Orderkey = ""
    '檢核路線編號資料是否正確
    RouteData_Check = False
    
    If Len(Trim(txt_Tab0_TRPDate.Text)) = 0 Then
        msg_text = "資料錯誤：未輸入出車日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SetFocus
        Exit Function
    End If
    If Len(Trim(txt_Tab0_DeliveryCarNo.Text)) = 0 Then
        msg_text = "資料錯誤：未輸入車牌號碼"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Function
    End If
    
    '資料檢核
    'a1.出車日期：格式 yyyymmdd
    txt_Tab0_TRPDate.Text = Trim(txt_Tab0_TRPDate.Text)
    If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
        msg_text = "出車日期：" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
        Exit Function
    End If
'    'a2.出車日期 >= 今天
'    If txt_Tab0_TRPDate.Text < Format(Now, "yyyymmdd") Then
'        msg_text = "出車日期不得小於今天"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
'        Exit Function
'    End If
    
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    'b.檢核 [車牌號碼] 是否有效
    txt_Tab0_DeliveryCarNo.Text = Trim(txt_Tab0_DeliveryCarNo.Text)
    str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "資料錯誤：車牌號碼 " & txt_Tab0_DeliveryCarNo.Text & " 未建檔"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Function
    End If
    tmp_Rs.Close
'    '指定碼頭暫存：必須輸入
'    txt_Tab0_DockNo.Text = Trim(txt_Tab0_DockNo.Text)
'    If Len(Trim(txt_Tab0_DockNo.Text)) = 0 Then
'        msg_text = "資料錯誤：[碼頭暫存] 必須輸入"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        txt_Tab0_DockNo.SetFocus
'        Exit Function
'    End If
    '預計報到日期
    txt_Tab0_CarCheckInDate.Text = Trim(txt_Tab0_CarCheckInDate.Text)
    If Len(Trim(txt_Tab0_CarCheckInDate.Text)) <> 8 Then
        msg_text = "預計報到日期：資料格式 yyyymmdd "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
        txt_Tab0_CarCheckInDate.SetFocus
        Exit Function
    End If
    If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
        msg_text = "預計報到日期：資料錯誤 yyyymmdd，" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
        txt_Tab0_CarCheckInDate.SetFocus
        Exit Function
    End If
'    'a2.預計報到日期 >= 今天
'    If txt_Tab0_CarCheckInDate.Text < Format(Now, "yyyymmdd") Then
'       msg_text = "預計報到日期不得小於今天"
'       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'       txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text): txt_Tab0_CarCheckInDate.SetFocus
'       Exit Function
'    End If
    
    '預計報到時間
    txt_Tab0_CarCheckInTime.Text = Trim(txt_Tab0_CarCheckInTime.Text)
    If Len(Trim(txt_Tab0_CarCheckInTime.Text)) <> 4 Then
        msg_text = "預計報到時間：資料格式 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
        txt_Tab0_CarCheckInTime.SetFocus
        Exit Function
    End If
    Select Case Left(txt_Tab0_CarCheckInTime.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "預計報到時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
                txt_Tab0_CarCheckInTime.SetFocus
                Exit Function
    End Select
    Select Case Right(txt_Tab0_CarCheckInTime.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "預計報到時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
                txt_Tab0_CarCheckInTime.SetFocus
                Exit Function
    End Select
    
    
    '檢查是否有被移至保留訂單的資料by Eric 20141215
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Function
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then Exit Function
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        Str_D_Orderkey = Str_D_Orderkey & "'" & rs_Tab0_SelectedOrders.Fields("Receipt_no") & "',"
        rs_Tab0_SelectedOrders.MoveNext
    Loop
        Call DB_CheckConnectStatus
        Call ReDim_Recordset(tmp_Rs)
        str_SQL = "select receipt_no from trp02t where route_no = 'D' and receipt_no in (" & Mid(Str_D_Orderkey, 1, Len(Str_D_Orderkey) - 1) & ")"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
                tmp_Rs.MoveFirst
                msg_text = "發現欲排車資料被移至保留訂單" & Chr(13) + Chr(10) & "請重新載入待排車資料再進行排車作業。"
                MsgBox msg_text, vbOKOnly + vbCritical, msg_title
                Str_D_Orderkey = ""
                Do While Not tmp_Rs.EOF
                    Str_D_Orderkey = Str_D_Orderkey & "'" & tmp_Rs.Fields("receipt_no") & "',"
                    tmp_Rs.MoveNext
                Loop
                msg_text = "被移至保留訂單的TMS訂單:" & Chr(13) + Chr(10) & Mid(Str_D_Orderkey, 1, Len(Str_D_Orderkey) - 1)
                MsgBox msg_text, vbOKOnly + vbCritical, "保留訂單資料檢查"
                tmp_Rs.Close
                Exit Function
    End If
    tmp_Rs.Close
    
    RouteData_Check = True
End Function

Private Sub Delete_RouteNo(strRouteNo As String)
    Screen.MousePointer = vbHourglass
    blTab1RouteEventEnable = False
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '刪除 TRP01T 路線編號主檔
    Call DB_CheckConnectStatus
    
    '(1).將 TRP03T 寫回 TRP03W >> 刪除 TRP03T
    str_SQL = "Insert into TRP03W(" & _
              " STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From TRP03T A Where a.Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).將 TRP02T 寫回 TRP02W >> 刪除 TRP02T
    
    'EXE_CONFIRM=9時(預先揀貨)，寫回TRP02W時狀態不變
    cn.Execute "update TRP02T set EXE_CONFIRM = 0 Where Route_No = '" & strRouteNo & "' and exe_confirm <> 9 ", RowsAffect, adExecuteNoRecords
    
    str_SQL = "Insert into TRP02W(" & _
              " RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              " WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
              "Select RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
              "From TRP02T Where Route_No = '" & strRouteNo & "'" 'exe_confirm 設定為 0 by gemini 20080106
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'    '(3).刪除 TRP02T & TRP03T
'    str_SQL = "Delete From TRP03T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    str_SQL = "Delete From TRP02T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    '(4).刪除 TRP05T
'    str_SQL = "Delete From TRP05T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    '(5).刪除 TRP01T
'    str_SQL = "Delete From TRP01T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords


    '刪除Trp1,2,3,5t 使用分號edit by Eric 20141215
    str_SQL = "Delete From TRP03T Where Route_No = '" & strRouteNo & "';" & _
              "Delete From TRP02T Where Route_No = '" & strRouteNo & "';" & _
              "Delete From TRP05T Where Route_No = '" & strRouteNo & "';" & _
              "Delete From TRP01T Where Route_No = '" & strRouteNo & "';"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    
    '(6)資料庫異動確認
    cn.CommitTrans
    Tran_Level = 0
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車作業-路線編號刪除", Me.Caption, "Form 內部 SubProgram Delete_RouteNo", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub Retrive_OrderSum()
    '取的待排車訂單：總計資料值
    txt_Tab0_srcTotal_OTqty.Text = ""
    txt_Tab0_srcTotal_Case.Text = ""
    txt_Tab0_srcTotal_Pallet.Text = ""
    txt_Tab0_srcTotal_Volumn.Text = ""
    txt_Tab0_srcTotal_Weight.Text = ""
    str_SQL = "Select Isnull(Round(sum(箱數),3),0) as 總箱數,Isnull(Round(sum(重量),3),0) as 總重量," & _
              "       Isnull(Round(sum(材積),3),0) as 總材積,Isnull(Round(sum(板數),3),0) as 總板數," & _
              "       Isnull(Round(sum(件數),3),0) as 總件數 " & _
              "From CutOrders_SourceOrder  "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        txt_Tab0_srcTotal_OTqty.Text = tmp_Rs.Fields("總件數").Value
        txt_Tab0_srcTotal_Case.Text = tmp_Rs.Fields("總箱數").Value
        txt_Tab0_srcTotal_Pallet.Text = tmp_Rs.Fields("總板數").Value
        txt_Tab0_srcTotal_Volumn.Text = tmp_Rs.Fields("總材積").Value
        txt_Tab0_srcTotal_Weight.Text = tmp_Rs.Fields("總重量").Value
    End If
    tmp_Rs.Close
End Sub

Private Sub ReCaculate_OrderSum()
    '取的待排車訂單：總計資料值  >>  目前待選列表的總計
    txt_Tab0_srcTotal_OTqty.Text = ""
    txt_Tab0_srcTotal_Case.Text = ""
    txt_Tab0_srcTotal_Pallet.Text = ""
    txt_Tab0_srcTotal_Volumn.Text = ""
    txt_Tab0_srcTotal_Weight.Text = ""
    
    If rs_TRP02W.RecordCount = 0 Then Exit Sub
    Dim dbTotalOTqty As Double
    Dim dbTotalCase As Double
    Dim dbTotalPallet As Double
    Dim dbTotalWeight As Double
    Dim dbTotalVolumn As Double
    dbTotalOTqty = 0: dbTotalCase = 0: dbTotalPallet = 0: dbTotalVolumn = 0: dbTotalWeight = 0
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    rs_TRP02W.MoveFirst
    Do While Not rs_TRP02W.EOF
        dbTotalOTqty = dbTotalOTqty + rs_TRP02W.Fields("件數").Value
        dbTotalCase = dbTotalCase + rs_TRP02W.Fields("箱數").Value
        dbTotalPallet = dbTotalPallet + rs_TRP02W.Fields("板數").Value
        dbTotalVolumn = dbTotalVolumn + rs_TRP02W.Fields("材積").Value
        dbTotalWeight = dbTotalWeight + rs_TRP02W.Fields("重量").Value
        rs_TRP02W.MoveNext
    Loop
    rs_TRP02W.MoveFirst
    txt_Tab0_srcTotal_OTqty.Text = dbTotalOTqty
    txt_Tab0_srcTotal_Case.Text = dbTotalCase
    txt_Tab0_srcTotal_Pallet.Text = dbTotalPallet
    txt_Tab0_srcTotal_Volumn.Text = dbTotalVolumn
    txt_Tab0_srcTotal_Weight.Text = dbTotalWeight
    
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
End Sub


Private Sub txt_Tab3_DeliveryDate_End_Click()
    '排車作業 >> 訂單分析
    If Trim(txt_Tab3_DeliveryDate_End.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_End.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_End.Text, 2))
        End If
    End If
    mvDate.Left = fma_Tab3_OrderSum.Left + txt_Tab3_DeliveryDate_End.Left
    mvDate.Top = fma_Tab3_OrderSum.Top + txt_Tab3_DeliveryDate_End.Top + txt_Tab3_DeliveryDate_End.Height
    mvDate.Tag = "訂單分析迄"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab3_DeliveryDate_Start_Click()
    '排車作業 >> 訂單分析
    If Trim(txt_Tab3_DeliveryDate_Start.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_Start.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_Start.Text, 2))
        End If
    End If
    mvDate.Left = fma_Tab3_OrderSum.Left + txt_Tab3_DeliveryDate_Start.Left
    mvDate.Top = fma_Tab3_OrderSum.Top + txt_Tab3_DeliveryDate_Start.Top + txt_Tab3_DeliveryDate_Start.Height
    mvDate.Tag = "訂單分析起"
    mvDate.Visible = True
End Sub
