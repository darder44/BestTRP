VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_OP_CarControl 
   Caption         =   "車輛進出管制作業"
   ClientHeight    =   7140
   ClientLeft      =   180
   ClientTop       =   795
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11475
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   4680
      TabIndex        =   90
      Top             =   4440
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
      StartOfWeek     =   92667905
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "車輛報到"
      TabPicture(0)   =   "frm_OP_CarControl.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fam_Tab0_CarData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dg_Tab0_CarCheckin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd_Exit(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ado_CarCheckin"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd_Tab0_CarList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fam_Tab0_CarCheckin"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fam_Tab0_Query"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_Tab0_ShowQuery"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "車輛離倉"
      TabPicture(1)   =   "frm_OP_CarControl.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fam_Tab1_Query"
      Tab(1).Control(1)=   "cmd_Tab1_ShowQuery"
      Tab(1).Control(2)=   "fam_Tab1_CarData"
      Tab(1).Control(3)=   "cmd_Exit(1)"
      Tab(1).Control(4)=   "cmd_Tab1_CarList"
      Tab(1).Control(5)=   "fam_Tab1_CarCheckout"
      Tab(1).Control(6)=   "ado_CarCheckout"
      Tab(1).Control(7)=   "dg_Tab1_CarCheckout"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "卡鐘資料整理"
      TabPicture(2)   =   "frm_OP_CarControl.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fam_SelectedOrders"
      Tab(2).Control(1)=   "fam_SrcOrders"
      Tab(2).ControlCount=   2
      Begin VB.Frame fam_SrcOrders 
         Caption         =   "待確認路編"
         Height          =   2835
         Left            =   -74880
         TabIndex        =   99
         Top             =   4080
         Width           =   11220
         Begin MSDataGridLib.DataGrid dg_Tab2_Route 
            Height          =   2520
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   4445
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
      Begin VB.Frame fam_SelectedOrders 
         Caption         =   "待確認卡鐘資料"
         Height          =   3735
         Left            =   -74880
         TabIndex        =   91
         Top             =   360
         Width           =   11220
         Begin VB.TextBox txt_Tab2_OutTime 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6780
            TabIndex        =   105
            Top             =   3315
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab2_InTime 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3900
            TabIndex        =   103
            Top             =   3315
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab2_Route_NO 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1050
            TabIndex        =   101
            Top             =   3315
            Width           =   1470
         End
         Begin VB.CommandButton cmd_Tab2_Selected 
            BackColor       =   &H00FF8080&
            Caption         =   "V 存檔"
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
            Left            =   240
            Style           =   1  '圖片外觀
            TabIndex        =   97
            Top             =   2790
            Width           =   945
         End
         Begin VB.CommandButton cmd_Tab2_srcOrderReset 
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
            Left            =   7770
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   95
            Top             =   2790
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab2_SelectedCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "取消"
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
            Left            =   3280
            Style           =   1  '圖片外觀
            TabIndex        =   94
            Top             =   2790
            Width           =   1260
         End
         Begin VB.CommandButton cmd_Tab2_ImportCard 
            BackColor       =   &H00C0C0FF&
            Caption         =   "載入待整理資料"
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
            Left            =   1205
            Style           =   1  '圖片外觀
            TabIndex        =   93
            Top             =   2790
            Width           =   2055
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
            Height          =   375
            Index           =   2
            Left            =   4560
            Style           =   1  '圖片外觀
            TabIndex        =   92
            Top             =   2790
            Width           =   1110
         End
         Begin MSDataGridLib.DataGrid dg_Tab2_CardIn 
            Height          =   2505
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   4419
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
         Begin MSDataGridLib.DataGrid dg_Tab2_CardOut 
            Height          =   2505
            Left            =   5640
            TabIndex        =   107
            Top             =   240
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   4419
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
         Begin VB.CommandButton cmd_Tab2_RouteQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "路編搜尋"
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
            Left            =   6645
            Style           =   1  '圖片外觀
            TabIndex        =   96
            Top             =   2790
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "離廠時間日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   18
            Left            =   5520
            TabIndex        =   106
            Top             =   3360
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "報到時間日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   17
            Left            =   2640
            TabIndex        =   104
            Top             =   3360
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   16
            Left            =   150
            TabIndex        =   102
            Top             =   3360
            Width           =   840
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00400040&
            FillStyle       =   0  '實心
            Height          =   435
            Index           =   1
            Left            =   120
            Top             =   3240
            Width           =   8175
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00400040&
            FillStyle       =   0  '實心
            Height          =   435
            Index           =   0
            Left            =   120
            Top             =   2760
            Width           =   5655
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00FF0000&
            FillStyle       =   0  '實心
            Height          =   435
            Left            =   6615
            Top             =   2760
            Width           =   1680
         End
      End
      Begin VB.Frame fam_Tab1_Query 
         BackColor       =   &H00404000&
         Caption         =   "篩選條件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   -74610
         TabIndex        =   82
         Top             =   1185
         Visible         =   0   'False
         Width           =   2910
         Begin VB.CommandButton cmd_Tab1_Default 
            BackColor       =   &H00FFC0FF&
            Caption         =   "預設"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2250
            Style           =   1  '圖片外觀
            TabIndex        =   88
            Top             =   195
            Width           =   570
         End
         Begin VB.CheckBox chk_Tab1_Checkin 
            Caption         =   "篩選已報到車輛"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   300
            TabIndex        =   85
            Top             =   330
            Value           =   1  '核取
            Width           =   1800
         End
         Begin VB.TextBox txt_Tab1_QueryDate 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1170
            TabIndex        =   84
            Top             =   600
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab1_QueryCarID 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1155
            TabIndex        =   83
            Top             =   960
            Width           =   1470
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   15
            Left            =   255
            TabIndex        =   87
            Top             =   645
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   86
            Top             =   1020
            Width           =   840
         End
      End
      Begin VB.CommandButton cmd_Tab1_ShowQuery 
         BackColor       =   &H00FFC0C0&
         Caption         =   "？"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72720
         Style           =   1  '圖片外觀
         TabIndex        =   81
         Top             =   825
         Width           =   345
      End
      Begin VB.CommandButton cmd_Tab0_ShowQuery 
         BackColor       =   &H00FFC0C0&
         Caption         =   "？"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         Style           =   1  '圖片外觀
         TabIndex        =   80
         Top             =   825
         Width           =   360
      End
      Begin VB.Frame fam_Tab0_Query 
         BackColor       =   &H00404000&
         Caption         =   "篩選條件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1455
         Left            =   525
         TabIndex        =   74
         Top             =   1185
         Visible         =   0   'False
         Width           =   2910
         Begin VB.CommandButton cmd_Tab0_Default 
            BackColor       =   &H00FFC0FF&
            Caption         =   "預設"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2205
            Style           =   1  '圖片外觀
            TabIndex        =   89
            Top             =   210
            Width           =   570
         End
         Begin VB.TextBox txt_Tab0_QueryCarID 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1155
            TabIndex        =   78
            Top             =   960
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab0_QueryDate 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1155
            TabIndex        =   76
            Top             =   600
            Width           =   1470
         End
         Begin VB.CheckBox chk_Tab0_Checkin 
            Caption         =   "篩選待報到車輛"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   270
            TabIndex        =   75
            Top             =   270
            Value           =   1  '核取
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   79
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   14
            Left            =   255
            TabIndex        =   77
            Top             =   645
            Width           =   840
         End
      End
      Begin VB.Frame fam_Tab1_CarData 
         Appearance      =   0  '平面
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   -74760
         TabIndex        =   47
         Top             =   1380
         Width           =   10920
         Begin VB.TextBox txt_Tab1_RouteNo 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9750
            TabIndex        =   61
            Top             =   495
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_CarID 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   60
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Driver 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   59
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   58
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_DriveTimes 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   57
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab1_Phone 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5190
            TabIndex        =   56
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Checkin 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   55
            Top             =   495
            Width           =   2355
         End
         Begin VB.TextBox txt_Tab1_Company 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5010
            TabIndex        =   54
            Top             =   840
            Width           =   2745
         End
         Begin VB.TextBox txt_Tab1_VehicleType 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7770
            TabIndex        =   53
            Top             =   840
            Width           =   3075
         End
         Begin VB.TextBox txt_Tab1_CaseQty 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   52
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Palletin 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   51
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab1_PalletQty 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   50
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Volumn 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   49
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Weight 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   48
            Top             =   495
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車牌號碼"
            Height          =   180
            Index           =   1
            Left            =   195
            TabIndex        =   71
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "駕駛人"
            Height          =   180
            Index           =   13
            Left            =   2490
            TabIndex        =   70
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "出車日期"
            Height          =   180
            Index           =   7
            Left            =   195
            TabIndex        =   69
            Top             =   555
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車次"
            Height          =   180
            Index           =   12
            Left            =   525
            TabIndex        =   68
            Top             =   900
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電話"
            Height          =   180
            Index           =   11
            Left            =   4755
            TabIndex        =   67
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "報到時間"
            Height          =   180
            Index           =   6
            Left            =   2310
            TabIndex        =   66
            Top             =   555
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運輸公司/車種"
            Height          =   180
            Index           =   10
            Left            =   3840
            TabIndex        =   65
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "箱數 / 板數"
            Height          =   180
            Index           =   9
            Left            =   7065
            TabIndex        =   64
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "攜入棧板數"
            Height          =   180
            Index           =   5
            Left            =   2160
            TabIndex        =   63
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材積 / 重量"
            Height          =   180
            Index           =   8
            Left            =   7065
            TabIndex        =   62
            Top             =   555
            Width           =   855
         End
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
         Height          =   495
         Index           =   1
         Left            =   -65025
         Style           =   1  '圖片外觀
         TabIndex        =   45
         Top             =   705
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Tab1_CarList 
         BackColor       =   &H8000000A&
         Caption         =   "載入已報到車輛"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -74625
         Style           =   1  '圖片外觀
         TabIndex        =   44
         Top             =   585
         Width           =   1905
      End
      Begin VB.Frame fam_Tab1_CarCheckout 
         Height          =   930
         Left            =   -72165
         TabIndex        =   37
         Top             =   405
         Width           =   6120
         Begin VB.CommandButton cmd_Tab1_CheckOutSave 
            BackColor       =   &H00FF8080&
            Caption         =   "存  檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4875
            Style           =   1  '圖片外觀
            TabIndex        =   73
            Top             =   135
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab1_PalletOut 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1290
            TabIndex        =   41
            Top             =   510
            Width           =   660
         End
         Begin VB.TextBox txt_Tab1_Checkout 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   360
            Left            =   1290
            TabIndex        =   40
            Top             =   135
            Width           =   2340
         End
         Begin VB.CommandButton cmd_Tab1_Checkout 
            BackColor       =   &H008080FF&
            Caption         =   "車輛離倉"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3660
            Style           =   1  '圖片外觀
            TabIndex        =   39
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab1_ClearCheckin 
            BackColor       =   &H00FF8080&
            Caption         =   "取  消"
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
            Left            =   2445
            Style           =   1  '圖片外觀
            TabIndex        =   38
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "離倉日期/時間"
            Height          =   180
            Index           =   4
            Left            =   105
            TabIndex        =   43
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "攜出棧板數"
            Height          =   180
            Index           =   7
            Left            =   330
            TabIndex        =   42
            Top             =   600
            Width           =   900
         End
      End
      Begin VB.Frame fam_Tab0_CarCheckin 
         BackColor       =   &H00400000&
         Height          =   930
         Left            =   2835
         TabIndex        =   18
         Top             =   420
         Width           =   6135
         Begin VB.CommandButton cmd_Tab0_CheckinSave 
            BackColor       =   &H00FF8080&
            Caption         =   "存  檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4875
            Style           =   1  '圖片外觀
            TabIndex        =   72
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab0_ClearCheckin 
            BackColor       =   &H00FF8080&
            Caption         =   "取  消"
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
            Left            =   2430
            Style           =   1  '圖片外觀
            TabIndex        =   35
            Top             =   495
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab0_Checkin 
            BackColor       =   &H008080FF&
            Caption         =   "車輛報到"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3660
            Style           =   1  '圖片外觀
            TabIndex        =   34
            Top             =   135
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab0_Checkin 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   360
            Left            =   1290
            TabIndex        =   22
            Top             =   135
            Width           =   2340
         End
         Begin VB.TextBox txt_Tab0_PalletIN 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1290
            TabIndex        =   21
            Top             =   510
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "攜入棧板數"
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   3
            Left            =   330
            TabIndex        =   20
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "報到日期/時間"
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   2
            Left            =   105
            TabIndex        =   19
            Top             =   225
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmd_Tab0_CarList 
         BackColor       =   &H8000000A&
         Caption         =   "載入待報到車輛"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   525
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   570
         Width           =   1905
      End
      Begin MSAdodcLib.Adodc ado_CarCheckin 
         Height          =   405
         Left            =   330
         Top             =   6420
         Visible         =   0   'False
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   714
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
         Caption         =   "ado_CarCheckin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
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
         Height          =   495
         Index           =   0
         Left            =   9960
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   720
         Width           =   1035
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_CarCheckin 
         Height          =   4245
         Left            =   225
         TabIndex        =   1
         Top             =   2595
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   7488
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
      Begin VB.Frame fam_Tab0_CarData 
         Appearance      =   0  '平面
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   225
         TabIndex        =   4
         Top             =   1395
         Width           =   10920
         Begin VB.TextBox txt_Tab0_Weight 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   33
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_Volumn 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   31
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_PalletQty 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   30
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_DockNo 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   29
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab0_CaseQty 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   26
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_VehicleType 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7770
            TabIndex        =   25
            Top             =   840
            Width           =   3075
         End
         Begin VB.TextBox txt_Tab0_Company 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5010
            TabIndex        =   24
            Top             =   840
            Width           =   2745
         End
         Begin VB.TextBox txt_Tab0_ExpectTime 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5190
            TabIndex        =   17
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_ExpectDate 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4035
            TabIndex        =   16
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_Phone 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5190
            TabIndex        =   14
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_DriveTimes 
            Alignment       =   2  '置中對齊
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   12
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab0_DeliveryDate 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   10
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_Driver 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   8
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_CarID 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   6
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_RouteNo 
            Appearance      =   0  '平面
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9750
            TabIndex        =   36
            Top             =   495
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材積 / 重量"
            Height          =   180
            Index           =   6
            Left            =   7065
            TabIndex        =   32
            Top             =   555
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "碼頭暫存"
            Height          =   180
            Index           =   3
            Left            =   2340
            TabIndex        =   28
            Top             =   900
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "箱數 / 板數"
            Height          =   180
            Index           =   5
            Left            =   7065
            TabIndex        =   27
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運輸公司/車種"
            Height          =   180
            Index           =   4
            Left            =   3840
            TabIndex        =   23
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "預計報到日期/時間"
            Height          =   180
            Index           =   1
            Left            =   2490
            TabIndex        =   15
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電話"
            Height          =   180
            Index           =   2
            Left            =   4755
            TabIndex        =   13
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車次"
            Height          =   180
            Index           =   1
            Left            =   525
            TabIndex        =   11
            Top             =   900
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "出車日期"
            Height          =   180
            Index           =   0
            Left            =   195
            TabIndex        =   9
            Top             =   555
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "駕駛人"
            Height          =   180
            Index           =   0
            Left            =   2490
            TabIndex        =   7
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車牌號碼"
            Height          =   180
            Index           =   0
            Left            =   195
            TabIndex        =   5
            Top             =   240
            Width           =   720
         End
      End
      Begin MSAdodcLib.Adodc ado_CarCheckout 
         Height          =   405
         Left            =   -74655
         Top             =   6405
         Visible         =   0   'False
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   714
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
         Caption         =   "ado_CarCheckOut"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_CarCheckout 
         Height          =   4245
         Left            =   -74760
         TabIndex        =   46
         Top             =   2595
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   7488
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
End
Attribute VB_Name = "frm_OP_CarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬

Private blTab0CarListEvent As Boolean    '車輛報到：事件控制旗標
Private blCancelChangeRecord As Boolean  '車輛報到：是否可以跳到下ㄧ筆記錄
Private blTab1CarListEvent As Boolean    '車輛報到：事件控制旗標
Private CardChange As Boolean            '事件控制旗標

Private rs_Tab0_CarCheckin As ADODB.Recordset    '待報到之車輛列表
Private rs_Tab1_CarCheckOut As ADODB.Recordset   '待報到之車輛列表
Private rs_Tab2_CardIn As ADODB.Recordset
Private rs_Tab2_CardOut As ADODB.Recordset
Private rs_Tab2_Route As ADODB.Recordset


Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub

Private Sub cmd_Tab0_CarList_Click()
'車輛報到 >> 載入待報到車輛
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
fam_Tab0_Query.Visible = False
Set dg_Tab0_CarCheckin.DataSource = Nothing
Set rs_Tab0_CarCheckin = Nothing
txt_Tab0_Checkin.Text = ""          '報到日期/時間
txt_Tab0_PalletIN.Text = ""         '攜入棧板數

str_SQL = "Select 出車日期,車牌號碼,車次,ㄧ單多車,報到時間,攜入棧板數,駕駛人,電話,箱數,板數,材積,重量,運輸公司," & _
          "    車種,預計報到日期,預計報到時間,碼頭暫存,路線編號,離倉時間 " & _
          "From CarControl_srcCheckin "

Dim strWhere As String, strTmp As String, tmp_data() As String, intloop As Integer
strWhere = ""
'篩選報到時間
strTmp = ""
If chk_Tab0_Checkin.Value = vbChecked Then
   strTmp = " 報到時間 = '' "
   If Len(strTmp) > 0 Then
      If Len(strWhere) = 0 Then
         strWhere = strTmp
      Else
         strWhere = strWhere & " and " & strTmp
      End If
   End If
End If
'出車日期
strTmp = ""
If Len(txt_Tab0_QueryDate.Text) > 0 Then
   strTmp = " 出車日期 = '" & strTmp & "' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'車牌號碼
strTmp = ""
If Len(txt_Tab0_QueryCarID.Text) > 0 Then
   tmp_data = Split(txt_Tab0_QueryCarID.Text, ",", -1, vbTextCompare)
   For intloop = LBound(tmp_data) To UBound(tmp_data)
       If Len(strTmp) = 0 Then
          strTmp = "'" & tmp_data(intloop) & "'"
       Else
          strTmp = strTmp & ",'" & tmp_data(intloop) & "'"
       End If
   Next intloop
   strTmp = " 車牌號碼 in (" & strTmp & ") "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
If Len(strWhere) > 0 Then
   str_SQL = str_SQL & " Where " & strWhere
End If
str_SQL = str_SQL & " Order by 出車日期,車牌號碼 "

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   msg_text = "查詢結果：無符合設定條件之待報到資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call CreateRS_Tab0_Checkin
Do While Not tmp_Rs.EOF
   With rs_Tab0_CarCheckin
     .AddNew
     .Fields("編號") = .RecordCount
     .Fields("出車日期").Value = tmp_Rs.Fields("出車日期").Value
     .Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
     .Fields("車次").Value = tmp_Rs.Fields("車次").Value
     .Fields("ㄧ單多車").Value = tmp_Rs.Fields("ㄧ單多車").Value
     .Fields("報到時間").Value = tmp_Rs.Fields("報到時間").Value
     .Fields("攜入棧板").Value = tmp_Rs.Fields("攜入棧板數").Value
     .Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
     .Fields("電話").Value = tmp_Rs.Fields("電話").Value
     .Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
     .Fields("板數").Value = tmp_Rs.Fields("板數").Value
     .Fields("材積").Value = tmp_Rs.Fields("材積").Value
     .Fields("重量").Value = tmp_Rs.Fields("重量").Value
     .Fields("運輸公司").Value = tmp_Rs.Fields("運輸公司").Value
     .Fields("車種").Value = tmp_Rs.Fields("車種").Value
     .Fields("預計報到日期").Value = tmp_Rs.Fields("預計報到日期").Value
     .Fields("預計報到時間").Value = tmp_Rs.Fields("預計報到時間").Value
     .Fields("碼頭暫存").Value = tmp_Rs.Fields("碼頭暫存").Value
     .Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
     .Fields("離倉時間").Value = tmp_Rs.Fields("離倉時間").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab0_CarCheckin.MoveFirst
Set dg_Tab0_CarCheckin.DataSource = rs_Tab0_CarCheckin
With dg_Tab0_CarCheckin
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
blTab0CarListEvent = False
Set dg_Tab0_CarCheckin.DataSource = rs_Tab0_CarCheckin
'設定顯示欄位
With dg_Tab0_CarCheckin
    .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
    .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
    .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
    .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
    .Columns(0).Width = 500         '編號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000        '出車日期
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900         '車牌號碼
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 450         '車次
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 850         'ㄧ單多車
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1700        '報到時間
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800         '攜入棧板
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 700         '駕駛人
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1000        '電話
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 800         '箱數
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800        '板數
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 800        '材積
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800        '重量
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 1500       '運輸公司
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1500       '車種
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1300       '預計報到日期
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1300       '預計報到時間
    .Columns(16).Alignment = dbgLeft
    .Columns(17).Width = 900        '碼頭暫存
    .Columns(17).Alignment = dbgLeft
    .Columns(18).Width = 1200       '路線編號
    .Columns(18).Alignment = dbgLeft

End With
Set ado_CarCheckin.Recordset = rs_Tab0_CarCheckin
txt_Tab0_CarID.DataField = "車牌號碼"
txt_Tab0_Driver.DataField = "駕駛人"
txt_Tab0_Phone.DataField = "電話"
txt_Tab0_DeliveryDate.DataField = "出車日期"
txt_Tab0_ExpectDate.DataField = "預計報到日期"
txt_Tab0_ExpectTime.DataField = "預計報到時間"
txt_Tab0_DriveTimes.DataField = "車次"
txt_Tab0_DockNo.DataField = "碼頭暫存"
txt_Tab0_Company.DataField = "運輸公司"
txt_Tab0_VehicleType.DataField = "車種"
txt_Tab0_CaseQty.DataField = "箱數"
txt_Tab0_PalletQty.DataField = "板數"
txt_Tab0_Volumn.DataField = "材積"
txt_Tab0_Weight.DataField = "重量"
txt_Tab0_Checkin.DataField = "報到時間"
txt_Tab0_PalletIN.DataField = "攜入棧板"
txt_Tab0_RouteNo.DataField = "路線編號"

'預設選取
If Not rs_Tab0_CarCheckin.EOF Then
   dg_Tab0_CarCheckin.SelBookmarks.Add dg_Tab0_CarCheckin.Bookmark
   txt_Tab0_Checkin.Text = rs_Tab0_CarCheckin.Fields("報到時間").Value
   txt_Tab0_PalletIN.Text = rs_Tab0_CarCheckin.Fields("攜入棧板").Value
End If

blTab0CarListEvent = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛報到-載入待報到車輛", Me.Caption, "cmd_Tab0_CarList_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Checkin_Click()
'車輛報到 >> 車輛報到
If rs_Tab0_CarCheckin Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab0_CarCheckin.SelBookmarks.Count = 0 Then
   msg_text = "未選取待報到車輛"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
'取得 sql serer 系統時間
str_SQL = "Select Convert(varchar,Getdate(),111) as CheckinDate , Convert(varchar,Getdate(),108) as CheckinTime "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
txt_Tab0_Checkin.Text = tmp_Rs.Fields("CheckinDate").Value & " " & tmp_Rs.Fields("CheckinTime").Value
tmp_Rs.Close
'按 [存檔] 才將時間、棧板數 顯示於 [待報到車輛列表]
'rs_Tab0_CarCheckin.Fields("報到時間").Value = txt_Tab0_Checkin.Text
'rs_Tab0_CarCheckin.Fields("攜入棧板").Value = Val(txt_Tab0_PalletIN.Text)
txt_Tab0_PalletIN.SelStart = 0: txt_Tab0_PalletIN.SelLength = Len(txt_Tab0_PalletIN.Text)
txt_Tab0_PalletIN.SetFocus
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛報到-車輛報到", Me.Caption, "cmd_Tab0_Checkin_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_CheckinSave_Click()
'車輛報到 >> 存檔
If rs_Tab0_CarCheckin Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab0_CarCheckin.SelBookmarks.Count = 0 Then
   msg_text = "未選取待報到車輛"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

   rs_Tab0_CarCheckin.Fields("報到時間").Value = txt_Tab0_Checkin.Text
   rs_Tab0_CarCheckin.Fields("攜入棧板").Value = Val(txt_Tab0_PalletIN.Text)

   If Len(Trim(txt_Tab0_Checkin.Text)) > 0 Then
      '檢查日期部份
      If Fun_ChkDateFormat2(Left(txt_Tab0_Checkin.Text, 10)) = 1 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛報到時間資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛報到時間作為輸入參考-[日期]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '檢查時間部份：時
      If Val(Mid(txt_Tab0_Checkin.Text, 12, 2)) < 0 Or Val(Mid(txt_Tab0_Checkin.Text, 12, 2)) >= 24 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛報到時間資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛報到時間作為輸入參考-[時]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '檢查時間部份：分
      If Val(Mid(txt_Tab0_Checkin.Text, 15, 2)) < 0 Or Val(Mid(txt_Tab0_Checkin.Text, 15, 2)) > 59 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛報到時間資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛報到時間作為輸入參考-[分]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '檢查時間部份：秒
      If Val(Mid(txt_Tab0_Checkin.Text, 18, 2)) < 0 Or Val(Mid(txt_Tab0_Checkin.Text, 18, 2)) > 59 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛報到時間：資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛報到時間作為輸入參考-[秒]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '將報到資料寫回 TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If

      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_IN"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab0_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab0_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab0_DriveTimes.Text)
      'VEHICLE_CHECK_IN
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_CHECK_IN", adChar, adParamInput, 20)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_CHECK_IN").Value = txt_Tab0_Checkin.Text
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("PALLET_IN", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("PALLET_IN").Value = Val(txt_Tab0_PalletIN.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '非同步執行
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
      Loop
      Set tmp_Cmd = Nothing
   Else
      msg_text = "資料錯誤：未輸入車輛報到時間"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      Exit Sub
   End If
Exit Sub

err_Handle:
    If Not (tmp_Cmd Is Nothing) Then
       Set tmp_Cmd = Nothing
    End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛報到-存檔", Me.Caption, "cmd_Tab0_CheckinSave_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Default_Click()
'載入待報到車輛 >> 篩選條件設定 >> 預設
chk_Tab0_Checkin.Value = vbChecked
txt_Tab0_QueryDate.Text = ""
txt_Tab0_QueryCarID.Text = ""
End Sub

Private Sub cmd_Tab0_ImportOrders_Click()
    
End Sub

Private Sub cmd_Tab0_ShowQuery_Click()
'載入待報到車輛 >> 顯示篩選條件設定
fam_Tab0_Query.Visible = Not fam_Tab0_Query.Visible
End Sub

Private Sub cmd_Tab1_CheckOutSave_Click()
'車輛離倉 >> 存檔
If rs_Tab1_CarCheckOut Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab1_CarCheckout.SelBookmarks.Count = 0 Then
   msg_text = "未選取欲離倉車輛"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
   rs_Tab1_CarCheckOut.Fields("離倉時間").Value = txt_Tab1_Checkout.Text
   rs_Tab1_CarCheckOut.Fields("攜出棧板").Value = Val(txt_Tab1_PalletOut.Text)

   If Len(Trim(txt_Tab1_Checkout.Text)) > 0 Then
      '檢查日期部份
      If Fun_ChkDateFormat2(Left(txt_Tab1_Checkout.Text, 10)) = 1 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛離倉時間資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛離倉時間作為輸入參考-[日期]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '檢查時間部份：時
      If Val(Mid(txt_Tab1_Checkout.Text, 12, 2)) < 0 Or Val(Mid(txt_Tab1_Checkout.Text, 12, 2)) >= 24 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛離倉時間資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛離倉時間作為輸入參考-[時]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '檢查時間部份：分
      If Val(Mid(txt_Tab1_Checkout.Text, 15, 2)) < 0 Or Val(Mid(txt_Tab1_Checkout.Text, 15, 2)) > 59 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛離倉時間資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛離倉時間作為輸入參考-[分]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '檢查時間部份：秒
      If Val(Mid(txt_Tab1_Checkout.Text, 18, 2)) < 0 Or Val(Mid(txt_Tab1_Checkout.Text, 18, 2)) > 59 Then
         msg_text = "資料錯誤：" & vbCrLf & "  車輛離倉時間：資料格式應為 yyyy/mm/dd hh:nn:ss，若有不解之處" & vbCrLf & _
                    "  請按 [車輛報到] 鈕自動產生車輛時間作為輸入參考-[秒]" & vbCrLf & _
                    "  注意：[日期] 與 [時間] 間隔應為ㄧ空格"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '將報到資料寫回 TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If

      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_OUT"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab1_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab1_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab1_DriveTimes.Text)
      'VEHICLE_CHECK_OUT
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_CHECK_OUT", adChar, adParamInput, 20)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_CHECK_OUT").Value = txt_Tab1_Checkout.Text
      'PALLET_OUT
      Set tmp_para = tmp_Cmd.CreateParameter("PALLET_OUT", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("PALLET_OUT").Value = Val(txt_Tab1_PalletOut.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '非同步執行
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
      Loop
      Set tmp_Cmd = Nothing
   Else
      msg_text = "資料錯誤：未輸入車輛離倉時間"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
Exit Sub

err_Handle:
    If Not (tmp_Cmd Is Nothing) Then
       Set tmp_Cmd = Nothing
    End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛離倉-存檔", Me.Caption, "cmd_Tab1_CheckinSave_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_ClearCheckin_Click()
'車輛報到 >> 清除車輛報到時間
On Error GoTo err_Handle
If rs_Tab0_CarCheckin Is Nothing Then Exit Sub
If dg_Tab0_CarCheckin.SelBookmarks.Count = 0 Then
   msg_text = "未選取欲進行報到取消之車輛"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
txt_Tab0_Checkin.Text = " "
txt_Tab0_PalletIN.Text = 0

      '將報到資料清除 TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If
      
      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_IN_Cancel"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab0_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab0_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab0_DriveTimes.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '非同步執行
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
      Loop
      Set tmp_Cmd = Nothing
      rs_Tab0_CarCheckin.Fields("報到時間").Value = " "
      rs_Tab0_CarCheckin.Fields("攜入棧板").Value = 0
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛報到-取消", Me.Caption, "cmd_Tab0_ClearCheckin_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_CarList_Click()
'車輛離倉 >> 載入已報到車輛
On Error GoTo err_Handle
fam_Tab1_Query.Visible = False
Screen.MousePointer = vbHourglass
Set dg_Tab1_CarCheckout.DataSource = Nothing
Set rs_Tab1_CarCheckOut = Nothing

str_SQL = "Select 出車日期,車牌號碼,車次,ㄧ單多車,離倉時間,攜出棧板數,報到時間,攜入棧板數,駕駛人,電話,箱數,板數,材積,重量,運輸公司,車種,路線編號 " & _
          "From CarControl_srcCheckOut "
Dim strWhere As String, strTmp As String, tmp_data() As String, intloop As Integer
strWhere = ""
'篩選報到時間
strTmp = ""
If chk_Tab1_Checkin.Value = vbChecked Then
   strTmp = " 離倉時間 = '' "
   If Len(strTmp) > 0 Then
      If Len(strWhere) = 0 Then
         strWhere = strTmp
      Else
         strWhere = strWhere & " and " & strTmp
      End If
   End If
End If
'出車日期
strTmp = ""
If Len(txt_Tab1_QueryDate.Text) > 0 Then
   tmp_data = Split(txt_Tab1_QueryDate.Text, ",", -1, vbTextCompare)
   For intloop = LBound(tmp_data) To UBound(tmp_data)
       If Len(strTmp) = 0 Then
          strTmp = "'" & tmp_data(intloop) & "'"
       Else
          strTmp = strTmp & ",'" & tmp_data(intloop) & "'"
       End If
   Next intloop
   strTmp = " 出車日期 in (" & strTmp & ") "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'車牌號碼
strTmp = ""
If Len(txt_Tab1_QueryCarID.Text) > 0 Then
   tmp_data = Split(txt_Tab1_QueryCarID.Text, ",", -1, vbTextCompare)
   For intloop = LBound(tmp_data) To UBound(tmp_data)
       If Len(strTmp) = 0 Then
          strTmp = "'" & tmp_data(intloop) & "'"
       Else
          strTmp = strTmp & ",'" & tmp_data(intloop) & "'"
       End If
   Next intloop
   strTmp = " 車牌號碼 in (" & strTmp & ") "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
If Len(strWhere) > 0 Then
   str_SQL = str_SQL & " Where " & strWhere
End If
str_SQL = str_SQL & " Order by 出車日期,車牌號碼 "

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   msg_text = "查詢結果：無符合設定條件之已報到車輛資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   
   txt_Tab1_CarID.Text = ""
   txt_Tab1_Driver.Text = ""
   txt_Tab1_Phone.Text = ""
   txt_Tab1_DeliveryDate.Text = ""
   txt_Tab1_DriveTimes.Text = ""
   txt_Tab1_Checkin.Text = ""
   txt_Tab1_Palletin.Text = ""
   txt_Tab1_Company.Text = ""
   txt_Tab1_VehicleType.Text = ""
   txt_Tab1_CaseQty.Text = ""
   txt_Tab1_PalletQty.Text = ""
   txt_Tab1_Volumn.Text = ""
   txt_Tab1_Weight.Text = ""
   txt_Tab1_Checkout.Text = ""
   txt_Tab1_PalletOut.Text = ""
   txt_Tab1_RouteNo.Text = ""
   
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call CreateRS_Tab1_CheckOut
Do While Not tmp_Rs.EOF
   With rs_Tab1_CarCheckOut
     .AddNew
     .Fields("編號") = .RecordCount
     .Fields("出車日期").Value = tmp_Rs.Fields("出車日期").Value
     .Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
     .Fields("車次").Value = tmp_Rs.Fields("車次").Value
     .Fields("ㄧ單多車").Value = tmp_Rs.Fields("ㄧ單多車").Value
     .Fields("離倉時間").Value = tmp_Rs.Fields("離倉時間").Value
     .Fields("攜出棧板").Value = tmp_Rs.Fields("攜出棧板數").Value
     .Fields("報到時間").Value = tmp_Rs.Fields("報到時間").Value
     .Fields("攜入棧板").Value = tmp_Rs.Fields("攜入棧板數").Value
     .Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
     .Fields("電話").Value = tmp_Rs.Fields("電話").Value
     .Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
     .Fields("板數").Value = tmp_Rs.Fields("板數").Value
     .Fields("材積").Value = tmp_Rs.Fields("材積").Value
     .Fields("重量").Value = tmp_Rs.Fields("重量").Value
     .Fields("運輸公司").Value = tmp_Rs.Fields("運輸公司").Value
     .Fields("車種").Value = tmp_Rs.Fields("車種").Value
     .Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab1_CarCheckOut.MoveFirst
Set dg_Tab1_CarCheckout.DataSource = rs_Tab1_CarCheckOut
With dg_Tab1_CarCheckout
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
blTab1CarListEvent = False
Set dg_Tab1_CarCheckout.DataSource = rs_Tab1_CarCheckOut
'設定顯示欄位
With dg_Tab1_CarCheckout
    .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
    .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
    .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
    .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
    .Columns(0).Width = 500         '編號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000        '出車日期
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900         '車牌號碼
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 450         '車次
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 850         'ㄧ單多車
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1700        '離倉時間
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800         '攜出棧板數
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 1700        '報到時間
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 800         '攜入棧板數
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700         '駕駛人
    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 1000       '電話
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 800        '箱數
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800        '板數
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 800        '材積
    .Columns(13).Alignment = dbgRight
    .Columns(14).Width = 800        '重量
    .Columns(14).Alignment = dbgRight
    .Columns(15).Width = 1500       '運輸公司
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1500       '車種
    .Columns(16).Alignment = dbgLeft
    .Columns(17).Width = 1200       '路線編號
    .Columns(17).Alignment = dbgLeft
End With

Set ado_CarCheckout.Recordset = rs_Tab1_CarCheckOut
txt_Tab1_CarID.DataField = "車牌號碼"
txt_Tab1_Driver.DataField = "駕駛人"
txt_Tab1_Phone.DataField = "電話"
txt_Tab1_DeliveryDate.DataField = "出車日期"
txt_Tab1_DriveTimes.DataField = "車次"
txt_Tab1_Checkin.DataField = "報到時間"
txt_Tab1_Palletin.DataField = "攜入棧板"
txt_Tab1_Company.DataField = "運輸公司"
txt_Tab1_VehicleType.DataField = "車種"
txt_Tab1_CaseQty.DataField = "箱數"
txt_Tab1_PalletQty.DataField = "板數"
txt_Tab1_Volumn.DataField = "材積"
txt_Tab1_Weight.DataField = "重量"
txt_Tab1_Checkout.DataField = "離倉時間"
txt_Tab1_PalletOut.DataField = "攜出棧板"
txt_Tab1_RouteNo.DataField = "路線編號"

'反白顯示第一筆資料
If Not rs_Tab1_CarCheckOut.EOF Then
   dg_Tab1_CarCheckout.SelBookmarks.Add dg_Tab1_CarCheckout.Bookmark
   txt_Tab1_Checkout.Text = rs_Tab1_CarCheckOut.Fields("離倉時間").Value
   txt_Tab1_PalletOut.Text = rs_Tab1_CarCheckOut.Fields("攜出棧板").Value
End If
blTab1CarListEvent = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛離倉-載入已報到車輛", Me.Caption, "cmd_Tab1_CarList_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Checkout_Click()
'車輛離倉 >> 車輛離倉
If rs_Tab1_CarCheckOut Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab1_CarCheckout.SelBookmarks.Count = 0 Then
   msg_text = "未選取欲離倉車輛"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'以 DB Server 時間為準
str_SQL = "Select Convert(varchar,Getdate(),111) as CheckoutDate , Convert(varchar,Getdate(),108) as CheckoutTime "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
txt_Tab1_Checkout.Text = tmp_Rs.Fields("CheckoutDate").Value & " " & tmp_Rs.Fields("CheckoutTime").Value
tmp_Rs.Close
'按 [存檔] 才將時間、棧板數 顯示於 [已報到車輛列表]
'rs_Tab1_CarCheckOut.Fields("離倉時間").Value = txt_Tab1_Checkout.Text
'rs_Tab1_CarCheckOut.Fields("攜出棧板").Value = txt_Tab1_PalletOut.Text
txt_Tab1_PalletOut.SelStart = 0: txt_Tab1_PalletOut.SelLength = Len(txt_Tab1_PalletOut.Text)
txt_Tab1_PalletOut.SetFocus
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛離倉-車輛離倉", Me.Caption, "cmd_Tab1_Checkout_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_ClearCheckin_Click()
'車輛離倉 >> 清除車輛離倉時間
On Error GoTo err_Handle
If rs_Tab1_CarCheckOut Is Nothing Then Exit Sub
If dg_Tab1_CarCheckout.SelBookmarks.Count = 0 Then
   msg_text = "未選取欲離倉車輛"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
txt_Tab1_Checkout.Text = " "
txt_Tab1_PalletOut.Text = 0

      '將報到資料清除 TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If
      
      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_OUT_Cancel"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab1_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab1_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab1_DriveTimes.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '非同步執行
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
      Loop
      Set tmp_Cmd = Nothing
      rs_Tab1_CarCheckOut.Fields("離倉時間").Value = " "
      rs_Tab1_CarCheckOut.Fields("攜出棧板").Value = 0
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-車輛離倉-取消", Me.Caption, "cmd_Tab1_ClearCheckin_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Default_Click()
'載入已報到車輛 >> 篩選條件設定 >> 預設
chk_Tab1_Checkin.Value = vbChecked
txt_Tab1_QueryDate.Text = ""
txt_Tab1_QueryCarID.Text = ""
End Sub

Private Sub cmd_Tab1_ShowQuery_Click()
'載入已報到車輛 >> 顯示篩選條件設定
fam_Tab1_Query.Visible = Not fam_Tab1_Query.Visible

End Sub

Private Sub cmd_Tab2_ImportCard_Click()
    '車輛進出作業>>匯入待整理卡鐘
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab2_CardIn.DataSource = Nothing
    Set dg_Tab2_CardOut.DataSource = Nothing
    Set dg_Tab2_Route.DataSource = Nothing
    
    '排車作業：待排車訂單
    Call CreateRS_Tab2_CardIn
    Call CreateRS_Tab2_CardOut
    Call CreateRS_Tab2_Route
    CardChange = False
    DoEvents
    
    '取回待排車訂單
    str_SQL = "select YMD,HM,Door,isnull(Port,'') as Port,Number,Username,Nickmane,CardNo,CardKey from gt_door where Status='0' and left(Port,3)='Car'"
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
    dg_Tab2_CardIn.Visible = False
    dg_Tab2_CardOut.Visible = False
    Do While Not tmp_Rs.EOF
        If Trim(tmp_Rs.Fields("Door").Value) = "1" Then
            rs_Tab2_CardIn.AddNew
            rs_Tab2_CardIn.Fields("日期").Value = tmp_Rs.Fields("YMD").Value
            rs_Tab2_CardIn.Fields("時間").Value = tmp_Rs.Fields("HM").Value
            rs_Tab2_CardIn.Fields("機台").Value = tmp_Rs.Fields("Door").Value
            rs_Tab2_CardIn.Fields("部門").Value = Trim(tmp_Rs.Fields("Port").Value)
            rs_Tab2_CardIn.Fields("系統編號").Value = tmp_Rs.Fields("CardKey").Value
            rs_Tab2_CardIn.Fields("使用者").Value = tmp_Rs.Fields("Username").Value
            rs_Tab2_CardIn.Fields("別名").Value = tmp_Rs.Fields("Nickmane").Value
            rs_Tab2_CardIn.Fields("卡號").Value = tmp_Rs.Fields("CardNo").Value
            rs_Tab2_CardIn.Update
        Else
            rs_Tab2_CardOut.AddNew
            rs_Tab2_CardOut.Fields("日期").Value = tmp_Rs.Fields("YMD").Value
            rs_Tab2_CardOut.Fields("時間").Value = tmp_Rs.Fields("HM").Value
            rs_Tab2_CardOut.Fields("機台").Value = tmp_Rs.Fields("Door").Value
            rs_Tab2_CardOut.Fields("部門").Value = tmp_Rs.Fields("Port").Value
            rs_Tab2_CardOut.Fields("系統編號").Value = tmp_Rs.Fields("CardKey").Value
            rs_Tab2_CardOut.Fields("使用者").Value = tmp_Rs.Fields("Username").Value
            rs_Tab2_CardOut.Fields("別名").Value = tmp_Rs.Fields("Nickmane").Value
            rs_Tab2_CardOut.Fields("卡號").Value = tmp_Rs.Fields("CardNo").Value
            rs_Tab2_CardOut.Update
        End If
       tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab2_CardIn.MoveFirst
    rs_Tab2_CardOut.MoveFirst
    dg_Tab2_CardIn.Visible = True
    dg_Tab2_CardOut.Visible = True
    
    '匯入待確認路編
    str_SQL = "select 送貨日,路線編號,車號,駕駛人,預計報到日期,預計報到時間,報到時間,離廠時間 from CarControl_Card where 報到時間='' or 離廠時間=''"
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
    dg_Tab2_Route.Visible = False
    Do While Not tmp_Rs.EOF
       rs_Tab2_Route.AddNew
       rs_Tab2_Route.Fields("送貨日").Value = tmp_Rs.Fields("送貨日").Value
       rs_Tab2_Route.Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
       rs_Tab2_Route.Fields("車號").Value = tmp_Rs.Fields("車號").Value
       rs_Tab2_Route.Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
       rs_Tab2_Route.Fields("預計報到日期").Value = tmp_Rs.Fields("預計報到日期").Value
       rs_Tab2_Route.Fields("預計報到時間").Value = tmp_Rs.Fields("預計報到時間").Value
       rs_Tab2_Route.Fields("報到時間").Value = tmp_Rs.Fields("報到時間").Value
       rs_Tab2_Route.Fields("離廠時間").Value = tmp_Rs.Fields("離廠時間").Value
       rs_Tab2_Route.Update
       tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab2_Route.MoveFirst
    'Set dg_Tab2_Route.DataSource = rs_Tab2_Route
    txt_Tab2_Route_NO.Text = rs_Tab2_Route.Fields("路線編號")
    dg_Tab2_Route.Visible = True
    
    CardChange = True
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

Private Sub cmd_Tab2_Selected_Click()
    If Len(Trim(txt_Tab2_InTime.Text)) = 0 And Len(Trim(txt_Tab2_OutTime.Text)) = 0 Then Exit Sub
    cn.BeginTrans
        str_SQL = "update trp05t set VEHICLE_CHECK_IN='" & Left(Trim(txt_Tab2_InTime.Text), 6) & " " & Mid(Trim(txt_Tab2_InTime.Text), 7, 2) & ":" & Right(Trim(txt_Tab2_InTime.Text), 2) & "'" & _
                ",VEHICLE_CHECK_OUT='" & Left(Trim(txt_Tab2_OutTime.Text), 6) & " " & Mid(Trim(txt_Tab2_OutTime.Text), 7, 2) & ":" & Right(Trim(txt_Tab2_OutTime.Text), 2) & "' where ROUTE_NO='" & Trim(txt_Tab2_Route_NO.Text) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        If Len(Trim(txt_Tab2_InTime.Text)) > 0 Then
            str_SQL = "update gt_door set Status='1' where CardKey='" & Trim(rs_Tab2_CardIn.Fields(4)) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            rs_Tab2_CardIn.Delete
            rs_Tab2_CardIn.Update
        End If
        If Len(Trim(txt_Tab2_OutTime.Text)) > 0 Then
            str_SQL = "update gt_door set Status='1' where CardKey='" & Trim(rs_Tab2_CardOut.Fields(4)) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            rs_Tab2_CardOut.Delete
            rs_Tab2_CardOut.Update
        End If
        rs_Tab2_Route.Delete
        rs_Tab2_Route.Update
        txt_Tab2_InTime.Text = ""
        txt_Tab2_OutTime.Text = ""
    cn.CommitTrans
End Sub

Private Sub cmd_Tab2_SelectedCancel_Click()
    txt_Tab2_InTime.Text = ""
    txt_Tab2_OutTime.Text = ""
End Sub

Private Sub cmd_Tab2_srcOrderReset_Click()
    If rs_Tab2_CardIn Is Nothing Then Exit Sub
    rs_Tab2_CardIn.Filter = adFilterNone
    rs_Tab2_CardIn.Sort = "日期 ASC"  '原始排序，一般資料序號由小至大
    rs_Tab2_CardOut.Filter = adFilterNone
    rs_Tab2_CardOut.Sort = "日期 ASC"  '原始排序，一般資料序號由小至大
    txt_Tab2_InTime.Text = ""
    txt_Tab2_OutTime.Text = ""
End Sub

Private Sub dg_Tab0_CarCheckin_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'車輛報到 >> 待報到車輛列表 >> 選取
If blTab0CarListEvent Then
   With dg_Tab0_CarCheckin
        '反白顯示選取之資料列
        If Not rs_Tab0_CarCheckin.EOF Then
           dg_Tab0_CarCheckin.SelBookmarks.Add dg_Tab0_CarCheckin.Bookmark
           txt_Tab0_Checkin.Text = rs_Tab0_CarCheckin.Fields("報到時間").Value
           txt_Tab0_PalletIN.Text = rs_Tab0_CarCheckin.Fields("攜入棧板").Value
        End If
   End With
End If
End Sub

Private Sub dg_Tab1_CarCheckout_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'車輛離倉 >> 已報到車輛列表
If blTab1CarListEvent Then
    With dg_Tab1_CarCheckout
        '反白顯示選取之資料列
        If Not rs_Tab1_CarCheckOut.EOF Then
            dg_Tab1_CarCheckout.SelBookmarks.Add dg_Tab1_CarCheckout.Bookmark
            txt_Tab1_Checkout.Text = rs_Tab1_CarCheckOut.Fields("離倉時間").Value
            txt_Tab1_PalletOut.Text = rs_Tab1_CarCheckOut.Fields("攜出棧板").Value
        End If
    End With
End If
End Sub

Private Sub dg_Tab1_CarCheckout_SelChange(Cancel As Integer)
'車輛離倉 >> 已報到車輛列表
If blCancelChangeRecord Then
   Cancel = True
End If
End Sub

Private Sub dg_Tab2_CardIn_Click()
    If CardChange = False Then Exit Sub
    txt_Tab2_InTime.Text = rs_Tab2_CardIn.Fields("日期").Value & rs_Tab2_CardIn.Fields("時間").Value
End Sub

Private Sub dg_Tab2_CardOut_Click()
    If CardChange = False Then Exit Sub
    Me.txt_Tab2_OutTime.Text = rs_Tab2_CardOut.Fields("日期").Value & rs_Tab2_CardOut.Fields("時間").Value
End Sub


Private Sub dg_Tab2_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If CardChange = False Then Exit Sub
    txt_Tab2_Route_NO.Text = rs_Tab2_Route.Fields("路線編號")
    str_SQL = "(使用者 LIKE '" & rs_Tab2_Route.Fields(3).Value & "' or 別名 LIKE '" & rs_Tab2_Route.Fields(2).Value & "')"
    rs_Tab2_CardIn.Filter = str_SQL
    If rs_Tab2_CardIn.RecordCount = 0 Then
         rs_Tab2_CardIn.Filter = adFilterNone
         rs_Tab2_CardIn.Sort = "日期 ASC"  '原始排序，一定要有這行資料才會重新顯示
    End If
    rs_Tab2_CardOut.Filter = str_SQL
    If rs_Tab2_CardOut.RecordCount = 0 Then
        rs_Tab2_CardOut.Filter = adFilterNone
        rs_Tab2_CardOut.Sort = "日期 ASC"
    End If
    txt_Tab2_InTime.Text = ""
    txt_Tab2_OutTime.Text = ""
End Sub


Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "車輛進出管制作業"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'車輛報到
Call CreateRS_Tab0_Checkin
'車輛離倉
Call CreateRS_Tab0_Checkin

Call CreateRS_Tab2_CardIn
Call CreateRS_Tab2_CardOut
Call CreateRS_Tab2_Route
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'攔截整個表單鍵盤按鍵事件
'用途：使用者按下 Esc 則不傳回任何資料，且關閉日期選取視窗
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
   fam_Tab0_Query.Visible = False
End If
End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '變小
   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   
   cmd_Tab0_CarList.Left = cmd_Tab0_CarList.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab0_CarData.Left = fam_Tab0_CarData.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   cmd_Exit(0).Left = cmd_Exit(0).Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab0_CarCheckin.Left = fam_Tab0_CarCheckin.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab0_CarCheckin.Height = dg_Tab0_CarCheckin.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Tab0_CarCheckin.Width = dg_Tab0_CarCheckin.Width - (dbsrcFormWidth - Me.ScaleWidth)
   cmd_Tab0_ShowQuery.Left = cmd_Tab0_ShowQuery.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab0_Query.Left = fam_Tab0_Query.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
      
   cmd_Tab1_CarList.Left = cmd_Tab1_CarList.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_CarData.Left = fam_Tab1_CarData.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   cmd_Exit(1).Left = cmd_Exit(1).Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_CarCheckout.Left = fam_Tab1_CarCheckout.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab1_CarCheckout.Height = dg_Tab1_CarCheckout.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Tab1_CarCheckout.Width = dg_Tab1_CarCheckout.Width - (dbsrcFormWidth - Me.ScaleWidth)
   cmd_Tab1_ShowQuery.Left = cmd_Tab1_ShowQuery.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_Query.Left = fam_Tab1_Query.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)

   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   
   cmd_Tab0_CarList.Left = cmd_Tab0_CarList.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab0_CarData.Left = fam_Tab0_CarData.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   cmd_Exit(0).Left = cmd_Exit(0).Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab0_CarCheckin.Left = fam_Tab0_CarCheckin.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab0_CarCheckin.Height = dg_Tab0_CarCheckin.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Tab0_CarCheckin.Width = dg_Tab0_CarCheckin.Width + (Me.ScaleWidth - dbsrcFormWidth)
   cmd_Tab0_ShowQuery.Left = cmd_Tab0_ShowQuery.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab0_Query.Left = fam_Tab0_Query.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   
   cmd_Tab1_CarList.Left = cmd_Tab1_CarList.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_CarData.Left = fam_Tab1_CarData.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   cmd_Exit(1).Left = cmd_Exit(1).Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_CarCheckout.Left = fam_Tab1_CarCheckout.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab1_CarCheckout.Height = dg_Tab1_CarCheckout.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Tab1_CarCheckout.Width = dg_Tab1_CarCheckout.Width + (Me.ScaleWidth - dbsrcFormWidth)
   cmd_Tab1_ShowQuery.Left = cmd_Tab1_ShowQuery.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_Query.Left = fam_Tab1_Query.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
      
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
Set frm_OP_CarControl = Nothing
End Sub

Private Sub CreateRS_Tab0_Checkin()
'車輛報到：待報到車輛列表
Call ReDim_Recordset(rs_Tab0_CarCheckin)
With rs_Tab0_CarCheckin
     .Fields.Append "編號", adVarChar, 10
     .Fields.Append "出車日期", adVarChar, 10
     .Fields.Append "車牌號碼", adVarChar, 10
     .Fields.Append "車次", adDouble
     .Fields.Append "ㄧ單多車", adVarChar, 10
     .Fields.Append "報到時間", adVarChar, 20
     .Fields.Append "攜入棧板", adDouble
     .Fields.Append "駕駛人", adVarChar, 20
     .Fields.Append "電話", adVarChar, 20
     .Fields.Append "箱數", adDouble
     .Fields.Append "板數", adDouble
     .Fields.Append "材積", adDouble
     .Fields.Append "重量", adDouble
     .Fields.Append "運輸公司", adVarChar, 60
     .Fields.Append "車種", adVarChar, 60
     .Fields.Append "預計報到日期", adVarChar, 10
     .Fields.Append "預計報到時間", adVarChar, 10
     .Fields.Append "碼頭暫存", adVarChar, 10
     .Fields.Append "路線編號", adVarChar, 10
     .Fields.Append "離倉時間", adVarChar, 20
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
End Sub

Private Sub CreateRS_Tab1_CheckOut()
'車輛離倉：已報到車輛列表
Call ReDim_Recordset(rs_Tab1_CarCheckOut)
With rs_Tab1_CarCheckOut
     .Fields.Append "編號", adVarChar, 10
     .Fields.Append "出車日期", adVarChar, 10
     .Fields.Append "車牌號碼", adVarChar, 10
     .Fields.Append "車次", adDouble
     .Fields.Append "ㄧ單多車", adVarChar, 10
     .Fields.Append "離倉時間", adVarChar, 20
     .Fields.Append "攜出棧板", adDouble
     .Fields.Append "報到時間", adVarChar, 20
     .Fields.Append "攜入棧板", adDouble
     .Fields.Append "駕駛人", adVarChar, 20
     .Fields.Append "電話", adVarChar, 20
     .Fields.Append "箱數", adDouble
     .Fields.Append "板數", adDouble
     .Fields.Append "材積", adDouble
     .Fields.Append "重量", adDouble
     .Fields.Append "運輸公司", adVarChar, 60
     .Fields.Append "車種", adVarChar, 60
     .Fields.Append "路線編號", adVarChar, 10
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'日期選取
Select Case mvDate.Tag
   Case "車輛報到.載入待報到車輛.篩選條件.出車日期"
        txt_Tab0_QueryDate.Text = Format(mvDate.Value, "yyyymmdd")
   Case "車輛離倉.載入已報到車輛.篩選條件.出車日期"
        txt_Tab1_QueryDate.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_Tab0_PalletIN_KeyPress(KeyAscii As Integer)
'車輛報到 >> 攜入棧板數
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   cmd_Tab0_CheckinSave.SetFocus
End If
End Sub

Private Sub txt_Tab0_QueryDate_Click()
'車輛報到 >> 載入待報到車輛 >> 篩選條件 >> 出車日期
If Trim(txt_Tab0_QueryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_QueryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_QueryDate.Text, 4) & "/" & Mid(txt_Tab0_QueryDate.Text, 5, 2) & "/" & Right(txt_Tab0_QueryDate.Text, 2))
   End If
End If
mvDate.Tag = "車輛報到.載入待報到車輛.篩選條件.出車日期"
mvDate.Top = SSTab1.Top + fam_Tab0_Query.Top + txt_Tab0_QueryDate.Top + txt_Tab0_QueryDate.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Query.Left + txt_Tab0_QueryDate.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_QueryDate_Click()
'車輛離倉 >> 載入已報到車輛 >> 篩選條件 >> 出車日期
If Trim(txt_Tab1_QueryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_QueryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_QueryDate.Text, 4) & "/" & Mid(txt_Tab1_QueryDate.Text, 5, 2) & "/" & Right(txt_Tab1_QueryDate.Text, 2))
   End If
End If
mvDate.Tag = "車輛離倉.載入已報到車輛.篩選條件.出車日期"
mvDate.Top = SSTab1.Top + fam_Tab1_Query.Top + txt_Tab1_QueryDate.Top + txt_Tab1_QueryDate.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Query.Left + txt_Tab1_QueryDate.Left
mvDate.Visible = True

End Sub

Private Sub CreateRS_Tab2_Route()
    Call ReDim_Recordset(rs_Tab2_Route)
    With rs_Tab2_Route
         .Fields.Append "送貨日", adVarChar, 10
         .Fields.Append "路線編號", adVarChar, 10
         .Fields.Append "車號", adVarChar, 10
         .Fields.Append "駕駛人", adVarChar, 20
         .Fields.Append "預計報到日期", adVarChar, 8
         .Fields.Append "預計報到時間", adVarChar, 4
         .Fields.Append "報到時間", adVarChar, 20
         .Fields.Append "離廠時間", adVarChar, 20
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '不需連接物件
    End With
    Set dg_Tab2_Route.DataSource = rs_Tab2_Route
    '設定顯示欄位
    With dg_Tab2_Route
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 1000
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1000
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1200
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1200
        .Columns(7).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab2_CardIn()
    Call ReDim_Recordset(rs_Tab2_CardIn)
    With rs_Tab2_CardIn
         .Fields.Append "日期", adVarChar, 6
         .Fields.Append "時間", adVarChar, 4
         .Fields.Append "機台", adVarChar, 10
         .Fields.Append "部門", adVarChar, 20
         .Fields.Append "系統編號", adVarChar, 6
         .Fields.Append "使用者", adVarChar, 20
         .Fields.Append "別名", adVarChar, 20
         .Fields.Append "卡號", adVarChar, 10
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '不需連接物件
    End With
    Set dg_Tab2_CardIn.DataSource = rs_Tab2_CardIn
    '設定顯示欄位
    With dg_Tab2_CardIn
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 700
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 500
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 400
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1000
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1000
        .Columns(7).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab2_CardOut()
    Call ReDim_Recordset(rs_Tab2_CardOut)
    With rs_Tab2_CardOut
         .Fields.Append "日期", adVarChar, 6
         .Fields.Append "時間", adVarChar, 4
         .Fields.Append "機台", adVarChar, 10
         .Fields.Append "部門", adVarChar, 20
         .Fields.Append "系統編號", adVarChar, 6
         .Fields.Append "使用者", adVarChar, 20
         .Fields.Append "別名", adVarChar, 20
         .Fields.Append "卡號", adVarChar, 10
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '不需連接物件
    End With
    Set dg_Tab2_CardOut.DataSource = rs_Tab2_CardOut
    '設定顯示欄位
    With dg_Tab2_CardOut
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 700
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 500
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 400
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1000
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1000
        .Columns(7).Alignment = dbgLeft
    End With
End Sub
