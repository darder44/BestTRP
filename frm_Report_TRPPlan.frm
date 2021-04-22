VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_Report_TRPPlan 
   Caption         =   " 排車作業報表"
   ClientHeight    =   7140
   ClientLeft      =   135
   ClientTop       =   1020
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3600
      TabIndex        =   87
      Top             =   4200
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
      StartOfWeek     =   122355713
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7800
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   13758
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "VLL 裝載"
      TabPicture(0)   =   "frm_Report_TRPPlan.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CmnDialog"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fam_Tab0_Header"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dg_Tab0_VLL"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "車輛裝載彙總表"
      TabPicture(1)   =   "frm_Report_TRPPlan.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fam_Tab1_Header"
      Tab(1).Control(1)=   "dg_Tab1_VLLSum"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "訂單總表"
      TabPicture(2)   =   "frm_Report_TRPPlan.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fam_Tab2"
      Tab(2).Control(1)=   "dg_Tab2_OrdersSum"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "揀貨裝載稽核表"
      TabPicture(3)   =   "frm_Report_TRPPlan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fam_Tab3"
      Tab(3).Control(1)=   "dg_Tab3_PickLoadCheck"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "轉運站路線匯總表"
      TabPicture(4)   =   "frm_Report_TRPPlan.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTab2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "排車一覽表"
      TabPicture(5)   =   "frm_Report_TRPPlan.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dg_Tab5_PlanList"
      Tab(5).Control(1)=   "fam_Tab5_Header"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "棧板維護by路編"
      TabPicture(6)   =   "frm_Report_TRPPlan.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "dg_Tab6_PlanList"
      Tab(6).Control(1)=   "fam_Tab6_Header"
      Tab(6).ControlCount=   2
      Begin VB.Frame fam_Tab6_Header 
         Height          =   1500
         Left            =   -74880
         TabIndex        =   137
         Top             =   720
         Width           =   10185
         Begin VB.TextBox txt_Tab6_route_End 
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
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   144
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab6_route_Start 
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
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   143
            Top             =   720
            Width           =   1365
         End
         Begin VB.CommandButton cmd_Tab6_SaveToExcel 
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
            Left            =   7725
            Picture         =   "frm_Report_TRPPlan.frx":00C4
            Style           =   1  '圖片外觀
            TabIndex        =   142
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab6_Query 
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
            Left            =   6570
            Picture         =   "frm_Report_TRPPlan.frx":0C86
            Style           =   1  '圖片外觀
            TabIndex        =   141
            Top             =   240
            Width           =   1065
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
            Index           =   7
            Left            =   8880
            Picture         =   "frm_Report_TRPPlan.frx":1550
            Style           =   1  '圖片外觀
            TabIndex        =   140
            Top             =   210
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab6_DeliveryDate_End 
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
            Left            =   2700
            MaxLength       =   8
            TabIndex        =   139
            Top             =   270
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab6_DeliveryDate_Start 
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
            Left            =   1140
            MaxLength       =   8
            TabIndex        =   138
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Index           =   37
            Left            =   105
            TabIndex        =   149
            Top             =   720
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
            Index           =   38
            Left            =   2550
            TabIndex        =   148
            Top             =   840
            Width           =   240
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00008080&
            BorderWidth     =   2
            Height          =   1260
            Index           =   2
            Left            =   6480
            Top             =   120
            Width           =   3570
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
            Index           =   40
            Left            =   2430
            TabIndex        =   147
            Top             =   360
            Width           =   240
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
            Index           =   41
            Left            =   120
            TabIndex        =   146
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "僅查30天 且已出車確認"
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
            Index           =   42
            Left            =   4080
            TabIndex        =   145
            Top             =   360
            Width           =   2160
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab5_PlanList 
         Height          =   4890
         Left            =   -74805
         TabIndex        =   102
         Top             =   2280
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   8625
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
      Begin VB.Frame fam_Tab5_Header 
         Height          =   1500
         Left            =   -74805
         TabIndex        =   88
         Top             =   720
         Width           =   12105
         Begin VB.CommandButton cmd_Tab5_PrintReport1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "分貨表"
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
            Left            =   8520
            Picture         =   "frm_Report_TRPPlan.frx":1992
            Style           =   1  '圖片外觀
            TabIndex        =   134
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab5_SaveToExcel_NEW 
            BackColor       =   &H00FFC0C0&
            Caption         =   "轉 Excel NEW"
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
            Left            =   6840
            Picture         =   "frm_Report_TRPPlan.frx":1C9C
            Style           =   1  '圖片外觀
            TabIndex        =   129
            ToolTipText     =   "由於包含運費試算，請勿一次查詢過多天數資料"
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab5_route_End 
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
            MaxLength       =   10
            TabIndex        =   113
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab5_route_Start 
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
            MaxLength       =   10
            TabIndex        =   112
            Top             =   1080
            Width           =   1365
         End
         Begin VB.ComboBox cmb_Tab5_AreaCode 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   315
            Left            =   1170
            Style           =   2  '單純下拉式
            TabIndex        =   100
            Top             =   225
            Width           =   3960
         End
         Begin VB.CommandButton cmd_Tab5_ReSet 
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
            Height          =   360
            Left            =   5085
            Style           =   1  '圖片外觀
            TabIndex        =   99
            Top             =   195
            Width           =   765
         End
         Begin VB.CommandButton cmd_Tab5_SaveToExcel 
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
            Left            =   7365
            Picture         =   "frm_Report_TRPPlan.frx":285E
            Style           =   1  '圖片外觀
            TabIndex        =   98
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab5_Query 
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
            Left            =   6210
            Picture         =   "frm_Report_TRPPlan.frx":3420
            Style           =   1  '圖片外觀
            TabIndex        =   97
            Top             =   240
            Width           =   1065
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
            Index           =   6
            Left            =   10860
            Picture         =   "frm_Report_TRPPlan.frx":3CEA
            Style           =   1  '圖片外觀
            TabIndex        =   96
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab5_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "排車管制表列印"
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
            Left            =   9705
            Picture         =   "frm_Report_TRPPlan.frx":412C
            Style           =   1  '圖片外觀
            TabIndex        =   95
            Top             =   225
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab5_DeliveryDate_End 
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
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   91
            Top             =   630
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab5_DeliveryDate_Start 
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
            TabIndex        =   90
            Top             =   615
            Width           =   1245
         End
         Begin VB.CheckBox chk_Tab5_PreView 
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
            Left            =   4680
            TabIndex        =   89
            Top             =   1080
            Value           =   1  '核取
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Index           =   28
            Left            =   120
            TabIndex        =   115
            Top             =   1080
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
            Index           =   27
            Left            =   2565
            TabIndex        =   114
            Top             =   1080
            Width           =   240
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00008080&
            BorderWidth     =   2
            Height          =   1140
            Index           =   1
            Left            =   6135
            Top             =   165
            Width           =   5865
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區碼"
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
            Left            =   120
            TabIndex        =   101
            Top             =   270
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
            Index           =   25
            Left            =   2445
            TabIndex        =   94
            Top             =   690
            Width           =   240
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
            Index           =   24
            Left            =   135
            TabIndex        =   93
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "僅查7天"
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
            Index           =   17
            Left            =   4200
            TabIndex        =   92
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame fam_Tab3 
         Height          =   2100
         Left            =   -74805
         TabIndex        =   66
         Top             =   840
         Width           =   11070
         Begin VB.TextBox txt_Tab3_route_Start 
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
            Left            =   1875
            MaxLength       =   10
            TabIndex        =   121
            Top             =   1680
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab3_route_End 
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
            Left            =   3555
            MaxLength       =   10
            TabIndex        =   120
            Top             =   1680
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_Start 
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
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   104
            Top             =   1230
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_End 
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
            Left            =   3495
            MaxLength       =   8
            TabIndex        =   103
            Top             =   1230
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab3_UploadMinute_End 
            Alignment       =   1  '靠右對齊
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
            Left            =   3450
            MaxLength       =   2
            TabIndex        =   79
            Top             =   870
            Width           =   375
         End
         Begin VB.TextBox txt_Tab3_UploadMinute_Start 
            Alignment       =   1  '靠右對齊
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
            Left            =   3450
            MaxLength       =   2
            TabIndex        =   78
            Top             =   525
            Width           =   375
         End
         Begin VB.CheckBox chk_Tab3_PreView 
            Caption         =   "預覽列印"
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
            Left            =   8160
            TabIndex        =   77
            Top             =   1320
            Value           =   1  '核取
            Width           =   1155
         End
         Begin VB.TextBox txt_Tab3_UploadHour_End 
            Alignment       =   1  '靠右對齊
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
            Left            =   3075
            MaxLength       =   2
            TabIndex        =   76
            Top             =   870
            Width           =   375
         End
         Begin VB.TextBox txt_Tab3_UploadHour_Start 
            Alignment       =   1  '靠右對齊
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
            Left            =   3075
            MaxLength       =   2
            TabIndex        =   75
            Top             =   525
            Width           =   375
         End
         Begin VB.CommandButton cmd_Tab3_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "報表列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   8685
            Picture         =   "frm_Report_TRPPlan.frx":4436
            Style           =   1  '圖片外觀
            TabIndex        =   74
            Top             =   240
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Tab3_Reset 
            BackColor       =   &H00FFC0FF&
            Caption         =   "重設"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5235
            Style           =   1  '圖片外觀
            TabIndex        =   73
            Top             =   570
            Width           =   795
         End
         Begin VB.TextBox txt_Tab3_UploadDate_Start 
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
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   72
            Top             =   525
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab3_UploadDate_End 
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
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   71
            Top             =   870
            Width           =   1185
         End
         Begin VB.ComboBox cmb_Tab3_AreaCode 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            ItemData        =   "frm_Report_TRPPlan.frx":4740
            Left            =   1110
            List            =   "frm_Report_TRPPlan.frx":4742
            Style           =   2  '單純下拉式
            TabIndex        =   70
            Top             =   165
            Width           =   4500
         End
         Begin VB.CommandButton cmd_Tab3_Query 
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
            Height          =   870
            Left            =   6420
            Picture         =   "frm_Report_TRPPlan.frx":4744
            Style           =   1  '圖片外觀
            TabIndex        =   69
            Top             =   255
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "離  開"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Index           =   3
            Left            =   9795
            Picture         =   "frm_Report_TRPPlan.frx":500E
            Style           =   1  '圖片外觀
            TabIndex        =   68
            Top             =   255
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Tab3_SaveToExcel 
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
            Height          =   870
            Left            =   7515
            Picture         =   "frm_Report_TRPPlan.frx":5450
            Style           =   1  '圖片外觀
            TabIndex        =   67
            Top             =   255
            Width           =   1065
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
            Index           =   32
            Left            =   3285
            TabIndex        =   123
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Left            =   900
            TabIndex        =   122
            Top             =   1680
            Width           =   960
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
            Index           =   7
            Left            =   900
            TabIndex        =   106
            Top             =   1290
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
            Index           =   6
            Left            =   3210
            TabIndex        =   105
            Top             =   1290
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "mm：0 ~ 59"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   12
            Left            =   4065
            TabIndex        =   86
            Top             =   825
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "hh：0 ~ 23"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   14
            Left            =   4155
            TabIndex        =   85
            Top             =   615
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "yyyymmdd   hh   mm"
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
            Height          =   195
            Index           =   15
            Left            =   6000
            TabIndex        =   84
            Top             =   1320
            Width           =   1860
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "資料格式："
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
            Index           =   16
            Left            =   4920
            TabIndex        =   83
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "迄"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   4
            Left            =   1605
            TabIndex        =   82
            Top             =   915
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區域"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   81
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "回傳日期：起"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   2
            Left            =   405
            TabIndex        =   80
            Top             =   600
            Width           =   1440
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00008080&
            BorderWidth     =   2
            Height          =   1080
            Index           =   0
            Left            =   6300
            Top             =   150
            Width           =   4620
         End
      End
      Begin VB.Frame fam_Tab2 
         Height          =   1320
         Left            =   -74850
         TabIndex        =   30
         Top             =   720
         Width           =   11145
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
            Index           =   2
            Left            =   9930
            Picture         =   "frm_Report_TRPPlan.frx":6012
            Style           =   1  '圖片外觀
            TabIndex        =   40
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab2_Query 
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
            Left            =   6420
            Picture         =   "frm_Report_TRPPlan.frx":6454
            Style           =   1  '圖片外觀
            TabIndex        =   39
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab2_SaveToExcel 
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
            Left            =   7560
            Picture         =   "frm_Report_TRPPlan.frx":6D1E
            Style           =   1  '圖片外觀
            TabIndex        =   38
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab2_ReSet 
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
            Height          =   375
            Left            =   4035
            Style           =   1  '圖片外觀
            TabIndex        =   37
            Top             =   885
            Width           =   630
         End
         Begin VB.CheckBox chk_Tab2_PreView 
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
            Left            =   1695
            TabIndex        =   36
            Top             =   960
            Value           =   1  '核取
            Width           =   1425
         End
         Begin VB.TextBox txt_Tab2_DeliveryDate_End 
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
            Left            =   2730
            MaxLength       =   8
            TabIndex        =   35
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab2_DeliveryDate_Start 
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
            Left            =   1125
            MaxLength       =   8
            TabIndex        =   34
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab2_RouteNo_End 
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
            Left            =   3075
            MaxLength       =   10
            TabIndex        =   33
            Top             =   525
            Width           =   1605
         End
         Begin VB.TextBox txt_Tab2_RouteNo_Start 
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
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   32
            Top             =   525
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab2_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "報表列印"
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
            Left            =   8760
            Picture         =   "frm_Report_TRPPlan.frx":78E0
            Style           =   1  '圖片外觀
            TabIndex        =   31
            Top             =   195
            Width           =   1065
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
            Index           =   11
            Left            =   2445
            TabIndex        =   45
            Top             =   240
            Width           =   240
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
            Index           =   10
            Left            =   120
            TabIndex        =   44
            Top             =   240
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
            Index           =   9
            Left            =   2790
            TabIndex        =   43
            Top             =   585
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Index           =   8
            Left            =   120
            TabIndex        =   42
            Top             =   585
            Width           =   960
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
            Index           =   5
            Left            =   4020
            TabIndex        =   41
            Top             =   225
            Width           =   2010
         End
      End
      Begin VB.Frame fam_Tab1_Header 
         Height          =   1530
         Left            =   -74850
         TabIndex        =   3
         Top             =   840
         Width           =   11145
         Begin VB.TextBox txt_Tab1_route_Start 
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
            MaxLength       =   10
            TabIndex        =   117
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab1_route_End 
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
            MaxLength       =   10
            TabIndex        =   116
            Top             =   1080
            Width           =   1365
         End
         Begin VB.CheckBox chk_Tab1_PreView 
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
            Left            =   4680
            TabIndex        =   28
            Top             =   1080
            Value           =   1  '核取
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Tab1_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "報表列印"
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
            Left            =   8745
            Picture         =   "frm_Report_TRPPlan.frx":7BEA
            Style           =   1  '圖片外觀
            TabIndex        =   27
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab1_Reset 
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
            Height          =   360
            Left            =   5100
            Style           =   1  '圖片外觀
            TabIndex        =   15
            Top             =   180
            Width           =   765
         End
         Begin VB.ComboBox cmb_Tab1_AreaCode 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   315
            Left            =   1155
            Style           =   2  '單純下拉式
            TabIndex        =   13
            Top             =   210
            Width           =   3960
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
            Left            =   9975
            Picture         =   "frm_Report_TRPPlan.frx":7EF4
            Style           =   1  '圖片外觀
            TabIndex        =   8
            Top             =   180
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate_Start 
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
            TabIndex        =   7
            Top             =   615
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate_End 
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
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   6
            Top             =   630
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab1_Query 
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
            Left            =   6315
            Picture         =   "frm_Report_TRPPlan.frx":8336
            Style           =   1  '圖片外觀
            TabIndex        =   5
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab1_SaveToExcel 
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
            Left            =   7545
            Picture         =   "frm_Report_TRPPlan.frx":8C00
            Style           =   1  '圖片外觀
            TabIndex        =   4
            Top             =   195
            Width           =   1065
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
            Index           =   30
            Left            =   2565
            TabIndex        =   119
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Index           =   29
            Left            =   120
            TabIndex        =   118
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區碼"
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
            Index           =   4
            Left            =   135
            TabIndex        =   12
            Top             =   255
            Width           =   960
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
            Index           =   3
            Left            =   4200
            TabIndex        =   11
            Top             =   600
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
            Index           =   2
            Left            =   135
            TabIndex        =   10
            Top             =   660
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
            Index           =   1
            Left            =   2445
            TabIndex        =   9
            Top             =   690
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_VLL 
         Height          =   4545
         Left            =   150
         TabIndex        =   2
         Top             =   2160
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8017
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
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
      Begin VB.Frame fam_Tab0_Header 
         Height          =   1320
         Left            =   150
         TabIndex        =   1
         Top             =   720
         Width           =   13425
         Begin VB.CheckBox chkVllPallet 
            Caption         =   "只印棧板管制表"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   4680
            TabIndex        =   152
            Top             =   600
            Value           =   1  '核取
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkDetail 
            Caption         =   "明細列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   5520
            TabIndex        =   151
            Top             =   960
            Width           =   1185
         End
         Begin VB.CheckBox chkLFAShipList 
            Caption         =   "含利豐出貨單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   136
            Top             =   960
            Width           =   1545
         End
         Begin VB.CheckBox chkKAOShipList 
            Caption         =   "含花王出貨單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   135
            Top             =   240
            Width           =   1545
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
            Index           =   0
            Left            =   12240
            Picture         =   "frm_Report_TRPPlan.frx":97C2
            Style           =   1  '圖片外觀
            TabIndex        =   133
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "報表列印"
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
            Left            =   11040
            Picture         =   "frm_Report_TRPPlan.frx":9C04
            Style           =   1  '圖片外觀
            TabIndex        =   132
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_SaveToExcel1 
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
            Left            =   9840
            Picture         =   "frm_Report_TRPPlan.frx":9F0E
            Style           =   1  '圖片外觀
            TabIndex        =   131
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_Query 
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
            Left            =   8640
            Picture         =   "frm_Report_TRPPlan.frx":AAD0
            Style           =   1  '圖片外觀
            TabIndex        =   130
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkNSLShipList 
            Caption         =   "含亞培出貨單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   127
            Top             =   720
            Width           =   1545
         End
         Begin VB.CheckBox chkTHLShipList 
            Caption         =   "含百事出貨單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   126
            Top             =   480
            Width           =   1545
         End
         Begin VB.CommandButton cmdLTHL01ShipDate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "THL出貨資料"
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
            Left            =   8640
            Style           =   1  '圖片外觀
            TabIndex        =   125
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkSUMDetail 
            Caption         =   "明細匯總列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   3840
            TabIndex        =   124
            Top             =   1000
            Width           =   1665
         End
         Begin VB.TextBox txt_Tab0_RouteNo_Start 
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
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   21
            Top             =   555
            Width           =   1605
         End
         Begin VB.TextBox txt_Tab0_RouteNo_End 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   20
            Top             =   555
            Width           =   1605
         End
         Begin VB.TextBox txt_Tab0_DeliveryDate_Start 
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
            Left            =   1125
            MaxLength       =   8
            TabIndex        =   19
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab0_DeliveryDate_End 
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
            Left            =   2730
            MaxLength       =   8
            TabIndex        =   18
            Top             =   180
            Width           =   1245
         End
         Begin VB.CheckBox chk_Tab0_PreView 
            Caption         =   "預覽列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   2625
            TabIndex        =   17
            Top             =   1000
            Value           =   1  '核取
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Tab0_ReSet 
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
            Height          =   375
            Left            =   5040
            Style           =   1  '圖片外觀
            TabIndex        =   16
            Top             =   120
            Width           =   630
         End
         Begin VB.CheckBox chk_Tab0_PrintedRoute 
            Caption         =   "含已列印過的路線編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Left            =   150
            TabIndex        =   22
            Top             =   1000
            Width           =   2625
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "僅查7天"
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
            Left            =   4080
            TabIndex        =   128
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Index           =   21
            Left            =   120
            TabIndex        =   26
            Top             =   615
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
            Index           =   20
            Left            =   2790
            TabIndex        =   25
            Top             =   615
            Width           =   240
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
            Index           =   19
            Left            =   120
            TabIndex        =   24
            Top             =   240
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
            Index           =   18
            Left            =   2445
            TabIndex        =   23
            Top             =   240
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_VLLSum 
         Height          =   4950
         Left            =   -74850
         TabIndex        =   14
         Top             =   2400
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
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
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dg_Tab2_OrdersSum 
         Height          =   5250
         Left            =   -74850
         TabIndex        =   29
         Top             =   2040
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   9260
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
      Begin MSDataGridLib.DataGrid dg_Tab3_PickLoadCheck 
         Height          =   4365
         Left            =   -74805
         TabIndex        =   46
         Top             =   3000
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7699
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   6570
         Left            =   -74880
         TabIndex        =   47
         Top             =   720
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   11589
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "資料篩選"
         TabPicture(0)   =   "frm_Report_TRPPlan.frx":B39A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(4)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label2(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label2(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(13)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(22)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label1(23)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dg_Tab4_RouteList"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmd_Tab4_Query_RouteDetail"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmd_Tab4_QueryBySRouteNo"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmd_Exit(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt_Tab4_SecondRouteNo"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt_Tab4_DeliveryDate_Start"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt_Tab4_DeliveryDate_End"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chk_Tab4_Selected"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "列印資料"
         TabPicture(1)   =   "frm_Report_TRPPlan.frx":B3B6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmd_Tab4_PrintReport"
         Tab(1).Control(1)=   "chk_Tab4_PreView"
         Tab(1).Control(2)=   "dg_Tab4_OrderDetail"
         Tab(1).ControlCount=   3
         Begin VB.CheckBox chk_Tab4_Selected 
            Caption         =   "查詢結果全選"
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
            Left            =   225
            TabIndex        =   111
            Top             =   1080
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab4_DeliveryDate_End 
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
            Left            =   2790
            MaxLength       =   8
            TabIndex        =   108
            Top             =   1395
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab4_DeliveryDate_Start 
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
            Left            =   1215
            MaxLength       =   8
            TabIndex        =   107
            Top             =   1395
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab4_SecondRouteNo 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   240
            MaxLength       =   10
            TabIndex        =   55
            Top             =   645
            Width           =   1650
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
            Height          =   945
            Index           =   5
            Left            =   9825
            Picture         =   "frm_Report_TRPPlan.frx":B3D2
            Style           =   1  '圖片外觀
            TabIndex        =   54
            Top             =   735
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab4_QueryBySRouteNo 
            BackColor       =   &H008080FF&
            Caption         =   "路編篩選"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2370
            Picture         =   "frm_Report_TRPPlan.frx":B814
            Style           =   1  '圖片外觀
            TabIndex        =   53
            Top             =   405
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab4_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "報表列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Left            =   -68535
            Picture         =   "frm_Report_TRPPlan.frx":C0DE
            Style           =   1  '圖片外觀
            TabIndex        =   52
            Top             =   555
            Width           =   2100
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
            Height          =   870
            Index           =   4
            Left            =   -65040
            Picture         =   "frm_Report_TRPPlan.frx":C3E8
            Style           =   1  '圖片外觀
            TabIndex        =   51
            Top             =   645
            Width           =   1065
         End
         Begin VB.CheckBox chk_Tab4_PreView 
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
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -71070
            TabIndex        =   50
            Top             =   1170
            Value           =   1  '核取
            Width           =   1380
         End
         Begin VB.CommandButton cmd_Tab4_Query_RouteDetail 
            BackColor       =   &H00FF8080&
            Caption         =   "路線匯總"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   8565
            Picture         =   "frm_Report_TRPPlan.frx":C82A
            Style           =   1  '圖片外觀
            TabIndex        =   49
            Top             =   735
            Width           =   1035
         End
         Begin MSDataGridLib.DataGrid dg_Tab4_OrderDetail 
            Height          =   4740
            Left            =   -74850
            TabIndex        =   48
            Top             =   1650
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   8361
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483624
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab4_RouteList 
            Height          =   4530
            Left            =   120
            TabIndex        =   56
            Top             =   1845
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   7990
            _Version        =   393216
            BackColor       =   -2147483624
            Cols            =   11
            TextStyleFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   11
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
            Index           =   23
            Left            =   2535
            TabIndex        =   110
            Top             =   1455
            Width           =   240
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
            Index           =   22
            Left            =   225
            TabIndex        =   109
            Top             =   1455
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "排車路線編號"
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
            Height          =   240
            Index           =   13
            Left            =   225
            TabIndex        =   65
            Top             =   345
            Width           =   1530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "操作步驟："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   195
            Index           =   3
            Left            =   4140
            TabIndex        =   64
            Top             =   555
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "報表已指定由雷射印表機"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   195
            Index           =   0
            Left            =   -74400
            TabIndex        =   63
            Top             =   1050
            Width           =   2310
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "Fujitsu 16 ADV 列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   180
            Index           =   0
            Left            =   -73845
            TabIndex        =   62
            Top             =   1365
            Width           =   1755
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "A4 直印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   0
            Left            =   -73200
            TabIndex        =   61
            Top             =   585
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "報表格式："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   225
            Index           =   3
            Left            =   -74415
            TabIndex        =   60
            Top             =   585
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "1. 輸入 [排車路線編號]，執行 [路編篩選]"
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
            Index           =   1
            Left            =   4425
            TabIndex        =   59
            Top             =   840
            Width           =   3795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "2. 確認欲列印的路線編號資料"
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
            Index           =   2
            Left            =   4425
            TabIndex        =   58
            Top             =   1125
            Width           =   2745
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "3. 執行 [路線匯總]，取出列印資料"
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
            Index           =   4
            Left            =   4425
            TabIndex        =   57
            Top             =   1410
            Width           =   3165
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab6_PlanList 
         Height          =   4410
         Left            =   -74880
         TabIndex        =   150
         Top             =   2280
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   7779
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
   End
End
Attribute VB_Name = "frm_Report_TRPPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬
Private blVLLReportEventEnable As Boolean   'VLL裝載
Private arAreaCode() As String
Private rs_Tab0_VLL As ADODB.Recordset           'VLL 裝載表：待選路編清單
Private rs_Tab0_VLLSum As ADODB.Recordset        'VLL 裝載總表
Private rs_Tab0_VLLDetail As ADODB.Recordset     'VLL 裝載明細表
Private rs_Tab0_VLLSUMDetail As ADODB.Recordset     'VLL 裝載明細表
Private rs_Tab0_VLLOrder As ADODB.Recordset      'VLL 出貨單
Private rs_Tab1_VLLSum As ADODB.Recordset        '車輛裝載彙總表
Private rs_Tab2_OrdersSum As ADODB.Recordset     '訂單總表
Private rs_Tab3_PickLoadCheck As ADODB.Recordset '揀貨裝載稽核表
Private rs_Tab4_OrderDetail As ADODB.Recordset   '轉運站路線匯總表：訂單名戲
Private rs_Tab5_PlanList As ADODB.Recordset      '排車一覽表
Private rs_Tab5_TRPPlanList As ADODB.Recordset   '排車一覽表_新
Private rs_Tab6_PlanList As ADODB.Recordset
Private str_SQL_Excel As String
Private strAccessDBFileName_FullPath As String
Private MSAccessAP As access.Application
Private rs_Access As ADODB.Recordset
Private rs_Access1 As ADODB.Recordset            '轉運站路線匯總表
Private rs_Access2 As ADODB.Recordset            '轉運站路線匯總表A
Private rs_Tab0_VLLDetailxSection As ADODB.Recordset

Private Sub cmd_Exit_Click(Index As Integer)
'離開
Set rs_Tab0_VLL = Nothing
Set rs_Tab0_VLLSum = Nothing
Set rs_Tab0_VLLDetail = Nothing
Set rs_Tab0_VLLSUMDetail = Nothing
Set rs_Tab0_VLLOrder = Nothing
Set rs_Tab1_VLLSum = Nothing
Set rs_Tab2_OrdersSum = Nothing
Set rs_Tab3_PickLoadCheck = Nothing
Set rs_Tab4_OrderDetail = Nothing
Set rs_Tab5_PlanList = Nothing
Set rs_Tab5_TRPPlanList = Nothing
Set rs_Access1 = Nothing
Set rs_Access2 = Nothing
Set rs_Tab0_VLLDetailxSection = Nothing

Unload Me
End Sub

Private Sub cmd_Tab0_SaveToExcel1_Click()

Recordset2Excel "VLL裝載", rs_Tab0_VLL
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd_Tab2_SaveToExcel_Click()
'訂單總表 >> 轉 EXCEL
Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel 檔案名稱
CmnDialog.DialogTitle = "轉存 Excel 檔"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "訂單總表_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel檔案(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
   msg_text = "選擇 [取消] 按鈕，必須於 Excel 中自行存檔"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab2_OrdersSum) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab2_OrdersSum.MoveFirst
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單總表-轉 EXCEL", Me.Caption, "cmd_Tab2_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_PrintReport_Click()
'揀貨裝載稽核表 >> 報表列印
If rs_Tab3_PickLoadCheck Is Nothing Then Exit Sub
If rs_Tab3_PickLoadCheck.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. 資料寫出 Access 資料庫 >> 訂單總表
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 揀貨裝載稽核表"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "揀貨裝載稽核表", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab3_PickLoadCheck.MoveFirst
Do While Not rs_Tab3_PickLoadCheck.EOF
   rs_Access.AddNew
   rs_Access.Fields("序號").Value = rs_Tab3_PickLoadCheck.Fields("編號").Value
   rs_Access.Fields("區域").Value = rs_Tab3_PickLoadCheck.Fields("運送區域").Value
   rs_Access.Fields("上傳日期").Value = rs_Tab3_PickLoadCheck.Fields("回傳日期").Value
   rs_Access.Fields("路線編號").Value = rs_Tab3_PickLoadCheck.Fields("路線編號").Value
   rs_Access.Fields("出車日期").Value = rs_Tab3_PickLoadCheck.Fields("出車日期").Value
   rs_Access.Fields("訂單張數").Value = rs_Tab3_PickLoadCheck.Fields("訂單數").Value
   rs_Access.Fields("送貨點").Value = rs_Tab3_PickLoadCheck.Fields("送貨點").Value
   rs_Access.Fields("客戶簡稱").Value = rs_Tab3_PickLoadCheck.Fields("客戶簡稱").Value
   rs_Access.Fields("箱數").Value = rs_Tab3_PickLoadCheck.Fields("箱數").Value
   rs_Access.Fields("板數").Value = rs_Tab3_PickLoadCheck.Fields("板數").Value
   rs_Access.Fields("重量").Value = rs_Tab3_PickLoadCheck.Fields("重量").Value
   rs_Access.Fields("材積").Value = rs_Tab3_PickLoadCheck.Fields("材積").Value
   rs_Access.Fields("貨運公司").Value = rs_Tab3_PickLoadCheck.Fields("貨運公司").Value
   rs_Access.Fields("車號").Value = rs_Tab3_PickLoadCheck.Fields("車號").Value
   rs_Access.Fields("車次").Value = rs_Tab3_PickLoadCheck.Fields("車次").Value
   rs_Access.Fields("預計報到時間").Value = rs_Tab3_PickLoadCheck.Fields("預計報到時間").Value
   rs_Access.Fields("碼頭").Value = rs_Tab3_PickLoadCheck.Fields("碼頭").Value
   rs_Access.Update
   rs_Tab3_PickLoadCheck.MoveNext
Loop
rs_Tab3_PickLoadCheck.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab3_PreView.Value = vbChecked Then
   '預覽列印
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "揀貨裝載稽核表", acViewPreview
Else
   '直接列印至印表機
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "揀貨裝載稽核表", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cnAccess.RollbackTrans
      Tran_Level = 0
   End If
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-揀貨裝載稽核表-列印", Me.Caption, "cmd_Tab3_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab0_PrintReport_Click()
'報表列印
If rs_Tab0_VLL Is Nothing Then Exit Sub
If rs_Tab0_VLL.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle
Dim strPrintDate As String       '列印時間
Dim strUserName As String        '列印者
Dim strRouteNo As String         '路線編號
Dim iLoop As Double
Dim strTmp As String, strRoute_No As String '所有路編與路編暫存資料
Dim i As Integer, strCompany As String

blVLLReportEventEnable = False
Dim strSelectedRouteNo As String    '選取之路線編號

strSelectedRouteNo = ""
rs_Tab0_VLL.MoveFirst
Do While Not rs_Tab0_VLL.EOF
   If Len(Trim(rs_Tab0_VLL.Fields(1).Value)) > 0 Then
      If strSelectedRouteNo = "" Then
         strSelectedRouteNo = "'" & rs_Tab0_VLL.Fields("路線編號").Value & "'"
      Else
         strSelectedRouteNo = strSelectedRouteNo & ",'" & rs_Tab0_VLL.Fields("路線編號").Value & "'"
      End If
   End If
   rs_Tab0_VLL.MoveNext
Loop

blVLLReportEventEnable = True
If strSelectedRouteNo = "" Then
   msg_text = "資料錯誤：未選取欲列印之 [路線編號]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
cmd_Tab0_PrintReport.Enabled = False    'daniel

Screen.MousePointer = 11

'一、VLL 裝載總表
str_SQL = "Select Distinct ' ' as '＊',出車日期,路線編號,車牌號碼,車次,駕駛人,運輸公司,貨主單號,客戶編號,客戶名稱," & _
          "   送貨地址,訂單備註,箱數,個數,板數,材積,重量,訂單日,預計報到日期,預計報到時間,碼頭暫存,列印次數," & _
          "   列印時間,指定到期日註記,排車者 as LoginUserID,車數," & _
          "   '                          ' as 車數註記,Receipt_No,二次排車路編,訂單類型,件數 " & _
          "From Report_VLL Where 二次排車路編 IN (" & strSelectedRouteNo & ") or 路線編號 in (" & strSelectedRouteNo & ") order by 二次排車路編 "
          
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料或出車日期已超過查詢限制!!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   blVLLReportEventEnable = True
   cmd_Tab0_PrintReport.Enabled = True     'daniel
   Screen.MousePointer = 0
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLSum)
tmp_Rs.Close

'1.1 ㄧ單多車，車數計數
Dim strExtern As String
rs_Tab0_VLLSum.Filter = " 車數 > 1 "
rs_Tab0_VLLSum.Sort = " 貨主單號 desc "
If Not rs_Tab0_VLLSum.EOF Then
   Do While Not rs_Tab0_VLLSum.EOF
   
      If rs_Tab0_VLLSum.Fields("貨主單號").Value <> strExtern Then
         '更新車數計數欄位資料
         str_SQL = "exec VLL_Extern_CarCount '" & rs_Tab0_VLLSum.Fields("貨主單號").Value & "' "
         cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
         
         strExtern = rs_Tab0_VLLSum.Fields("貨主單號").Value
      End If
      
      '取得車數計數欄位資料
      str_SQL = "Select Rtrim(isnull(Car_Notes,' ')) as 車數註記 From TRP02T Where Receipt_No = '" & rs_Tab0_VLLSum.Fields("Receipt_No").Value & "'"
      tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
      If Not tmp_Rs.EOF Then
         rs_Tab0_VLLSum.Fields("車數註記").Value = tmp_Rs.Fields("車數註記").Value
      End If
      tmp_Rs.Close
      rs_Tab0_VLLSum.MoveNext
   Loop
End If

rs_Tab0_VLLSum.Filter = adFilterNone
rs_Tab0_VLLSum.Sort = " 二次排車路編 desc "

'1-2. 取得 DB Server 時間
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select Convert(varchar,GetDate(),111) + ' ' + convert(varchar,GetDate(),108) as '列印時間' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
strPrintDate = tmp_Rs.Fields("列印時間").Value
tmp_Rs.Close

'1-3. 更新 TRP01T 欄位資料
strRouteNo = ""
rs_Tab0_VLLSum.MoveFirst
Do While Not rs_Tab0_VLLSum.EOF
   If strRouteNo <> rs_Tab0_VLLSum.Fields("二次排車路編").Value Then
      rs_Tab0_VLLSum.Fields("列印次數").Value = rs_Tab0_VLLSum.Fields("列印次數").Value + 1
      rs_Tab0_VLLSum.Fields("列印時間").Value = strPrintDate
      
      '以 路線編號 為資料處理依據
      strRouteNo = rs_Tab0_VLLSum.Fields("二次排車路編").Value
      str_SQL = "Update TRP01T Set VLListCount = " & rs_Tab0_VLLSum.Fields("列印次數").Value & ",VLListPrintDate = '" & strPrintDate & "' " & _
                "Where Route_No = '" & strRouteNo & "' or C_Route_No = '" & strRouteNo & "'"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0_VLLSum.MoveNext
Loop
rs_Tab0_VLLSum.MoveFirst

If chkSUMDetail.Value = vbChecked Then

        '二、VLL 加總裝載明細表
'        str_SQL = "Select 路線編號,出車日期,預計報到日期,預計報到時間,碼頭暫存,車牌號碼,車次,駕駛人,運輸公司," & _
'                  "  排車者,貨主單號,TMS單號,倉別,冷藏,貨號,品名,出貨箱數=isnull(出貨箱數,0),出貨個數=isnull(出貨個數,0),揀貨重量,揀貨材積,列印次數,列印時間,二次排車路編,排車箱數,製造日,到期日 " & _
'                  "From Report_VLLSUMDetail Where 二次排車路編 IN (" & strSelectedRouteNo & ") or 路線編號 in (" & strSelectedRouteNo & ")"
                    
         str_SQL = "Select 路線編號 = a1.Route_No,出車日期 = Case When Isnull(a1.C_Route_No,'') = '' Then Convert(varchar,a1.Delivery_Date,112) else Convert(varchar,t01t2.Delivery_Date,112) End " & _
                    ",預計報到日期 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Date,'')) else Rtrim(t05t2.Expect_Date) End " & _
                    ",預計報到時間 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Time,'')) else Rtrim(t05t2.Expect_Time) End " & _
                    ",碼頭暫存 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Dock_No,'')) else Rtrim(t05t2.Dock_No) End " & _
                    ",車牌號碼 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Vehicle_ID_No) else Rtrim(t05t2.Vehicle_ID_No) End " & _
                    ",車次 = Case When Isnull(a1.C_Route_No,'') = '' Then Round(Cast(a2.Drive_Times as float),2) else Round(Cast(t05t2.Drive_Times as float),2) End " & _
                    ",駕駛人 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Driver) else Rtrim(t05t2.Driver) End " & _
                    ",運輸公司 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(c2.C_Name,'')) else Rtrim(Isnull(t08m2.C_Name,'')) End " & _
                    ",排車者 = Case When Isnull(a1.C_Route_No,'') = '' Then Isnull(Rtrim(a1.AddWho),'') else Rtrim(t01t2.AddWho) End " & _
                    ",貨主單號 = ' ',TMS單號 = ' ',倉別 = rtrim(l.lottable06),冷藏 = substring(sp.skugroup,7,1),地段 = case when t02t.priority = 'C' then '' else rtrim(loc.sectionkey) end,貨號 = Rtrim(t03t.Product_No),品名 = Isnull(Rtrim(sp.Descr),'') " & _
                    ",出貨箱數 = case when sp.casecnt = 0 then 0 else Isnull(floor(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) /(sp.Casecnt)),0) end " & _
                    ",出貨個數 = case when sp.casecnt = 0 then sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) else cast(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) as int) % cast(sp.Casecnt as int) end " & _
                    ",揀貨重量 = Isnull(Round((sp.STDGrossWGT * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0),揀貨材積 = Isnull(Round((sp.STDCUBE * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0) " & _
                    ",列印次數 = Isnull(a1.VLListCount,0),列印時間 = Isnull((Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()),111) + ' ' + Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()) , 108)),'') " & _
                    ",二次排車路編 = Case When Isnull(a1.C_Route_No,'') = '' Then a1.Route_No Else a1.C_Route_No End,排車箱數 = 0,製造日 = ' ',到期日 = ' ' " & _
                    "From TRP01T a1 inner join trp02t t02t on t02t.Route_No = a1.Route_No join TRP03T t03t on t03t.receipt_No = t02t.receipt_No and a1.Route_No <> 'D' inner join TRP05T a2 on a2.Route_No = a1.Route_No " & _
                    "inner join TRP09M c1 on c1.Vehicle_ID_No = a2.Vehicle_ID_No inner join gv_SKUxpack sp on sp.StorerKey = t03t.StorerKey and sp.SKU = t03t.Product_No " & _
                    "Left outer join TRP01T t01t2 on t01t2.Route_No = a1.C_Route_No " & _
                    "Left outer join TRP05T t05t2 on t05t2.Route_No = a1.C_Route_No Left outer join TRP09M t09m2 on t09m2.Vehicle_ID_No = t05t2.Vehicle_ID_No " & _
                    "Left outer join TRP08M t08m2 on t08m2.Company_Code = t09m2.TRP_Company_Code Left outer join TRP08M c2 on c2.Company_Code = c1.TRP_Company_Code " & _
                    "Left join " & strWMSDB & "..orders o (nolock) on o.updatesource = t03t.receipt_no Left join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey and od.externlineno = t03t.seq_no  " & _
                    "Left join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = od.orderkey and od.orderlinenumber  = p.orderlinenumber Left join " & strWMSDB & "..lotattribute l (nolock) on p.lot = l.lot and p.sku = l.sku Left join " & strWMSDB & "..loc loc (nolock) on p.loc = loc.loc " & _
                    "where a1.C_Route_No IN (" & strSelectedRouteNo & ") or a1.Route_No in (" & strSelectedRouteNo & ") " & _
                    "Group by t02t.priority,a1.Route_No , loc.sectionkey,a1.Delivery_Date , a2.Expect_Date , a2.Expect_Time , a2.Dock_No , a2.Vehicle_ID_No,a2.Drive_Times , a2.Driver , c2.C_Name , a1.C_Route_No , a1.AddWho,t03t.Product_No,sp.Descr,a1.VLListCount , a1.VLListPrintDate ,sp.CaseCnt,sp.STDGROSSWGT,sp.STDCUBE,t01t2.Delivery_Date , t05t2.Expect_Date , t05t2.Expect_Time , t05t2.Dock_No , t05t2.Vehicle_ID_No , t05t2.Driver , " & _
                    "t08m2.C_Name , t01t2.AddWho , t05t2.Drive_Times,l.lottable06,substring(sp.skugroup,7,1) "

        Call DB_CheckConnectStatus
        Call Confirm_Recordset_Closed(tmp_Rs)
        cn.CommandTimeout = 600
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
           tmp_Rs.Close
           cmd_Tab0_PrintReport.Enabled = True
           msg_text = "查詢結果：無符合搜尋條件之訂單明細資料"
           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
           Screen.MousePointer = vbDefault
           blVLLReportEventEnable = True
           Exit Sub
        End If
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLSUMDetail)
        tmp_Rs.Close
End If
If chkDetail.Value = vbChecked Then

        '二、VLL 裝載明細表
'        str_SQL = "Select 路線編號,出車日期,預計報到日期,預計報到時間,碼頭暫存,車牌號碼,車次,駕駛人,運輸公司," & _
'                  "  排車者,貨主單號,TMS單號,倉別,冷藏,貨號,品名,出貨箱數=isnull(出貨箱數,0),出貨個數=isnull(出貨個數,0),揀貨重量,揀貨材積,列印次數,列印時間,二次排車路編,排車箱數,製造日,到期日 " & _
'                  "From Report_VLLDetail Where 二次排車路編 IN (" & strSelectedRouteNo & ") or 路線編號 in (" & strSelectedRouteNo & ")"
        
        str_SQL = "Select 路線編號 = a1.Route_No,出車日期 = Case When Isnull(a1.C_Route_No,'') = '' Then Convert(varchar,a1.Delivery_Date,112) else Convert(varchar,t01t2.Delivery_Date,112) End " & _
                    ",預計報到日期 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Date,'')) else Rtrim(t05t2.Expect_Date) End " & _
                    ",預計報到時間 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Time,'')) else Rtrim(t05t2.Expect_Time) End " & _
                    ",碼頭暫存 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Dock_No,'')) else Rtrim(t05t2.Dock_No) End " & _
                    ",車牌號碼 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Vehicle_ID_No) else Rtrim(t05t2.Vehicle_ID_No) End " & _
                    ",車次 = Case When Isnull(a1.C_Route_No,'') = '' Then Round(Cast(a2.Drive_Times as float),2) else Round(Cast(t05t2.Drive_Times as float),2) End " & _
                    ",駕駛人 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Driver) else Rtrim(t05t2.Driver) End " & _
                    ",運輸公司 = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(c2.C_Name,'')) else Rtrim(Isnull(t08m2.C_Name,'')) End " & _
                    ",排車者 = Case When Isnull(a1.C_Route_No,'') = '' Then Isnull(Rtrim(a1.AddWho),'') else Rtrim(t01t2.AddWho) End " & _
                    ",貨主單號 = t03t.extern,TMS單號 = t03t.receipt_no,倉別 = rtrim(l.lottable06),冷藏 = substring(sp.skugroup,7,1),地段 = case when t02t.priority = 'C' then '' else rtrim(loc.sectionkey) end,貨號 = Rtrim(t03t.Product_No),品名 = Isnull(Rtrim(sp.Descr),'') " & _
                    ",出貨箱數 = case when sp.casecnt = 0 then 0 else Isnull(floor(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) /(sp.Casecnt)),0) end " & _
                    ",出貨個數 = case when sp.casecnt = 0 then sum(isnull(p.qty,0)) else cast(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) as int) % cast(sp.Casecnt as int) end " & _
                    ",揀貨重量 = Isnull(Round((sp.STDGrossWGT * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0),揀貨材積 = Isnull(Round((sp.STDCUBE * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0) " & _
                    ",列印次數 = Isnull(a1.VLListCount,0),列印時間 = Isnull((Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()),111) + ' ' + Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()) , 108)),'') " & _
                    ",二次排車路編 = Case When Isnull(a1.C_Route_No,'') = '' Then a1.Route_No Else a1.C_Route_No End,排車箱數 = 0,製造日 = isnull(convert(char(10),l.lottable04,111),''),到期日 = isnull(convert(char(10),l.lottable05,111),'') " & _
                    "From TRP01T a1 inner join trp02t t02t on t02t.Route_No = a1.Route_No join TRP03T t03t on t03t.receipt_No = t02t.receipt_No and a1.Route_No <> 'D' inner join TRP05T a2 on a2.Route_No = a1.Route_No " & _
                    "inner join TRP09M c1 on c1.Vehicle_ID_No = a2.Vehicle_ID_No inner join gv_SKUxpack sp on sp.StorerKey = t03t.StorerKey and sp.SKU = t03t.Product_No " & _
                    "Left outer join TRP01T t01t2 on t01t2.Route_No = a1.C_Route_No " & _
                    "Left outer join TRP05T t05t2 on t05t2.Route_No = a1.C_Route_No Left outer join TRP09M t09m2 on t09m2.Vehicle_ID_No = t05t2.Vehicle_ID_No " & _
                    "Left outer join TRP08M t08m2 on t08m2.Company_Code = t09m2.TRP_Company_Code Left outer join TRP08M c2 on c2.Company_Code = c1.TRP_Company_Code " & _
                    "Left join " & strWMSDB & "..orders o (nolock) on o.updatesource = t03t.receipt_no Left join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey and od.externlineno = t03t.seq_no  " & _
                    "Left join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = od.orderkey and od.orderlinenumber  = p.orderlinenumber Left join " & strWMSDB & "..lotattribute l (nolock) on p.lot = l.lot and p.sku = l.sku Left join " & strWMSDB & "..loc loc (nolock) on p.loc = loc.loc " & _
                    "where a1.C_Route_No IN (" & strSelectedRouteNo & ") or a1.Route_No in (" & strSelectedRouteNo & ") " & _
                    "Group by t02t.priority,a1.Route_No , a1.Delivery_Date , a2.Expect_Date , a2.Expect_Time , a2.Dock_No , a2.Vehicle_ID_No,a2.Drive_Times , a2.Driver , c2.C_Name , a1.C_Route_No , a1.AddWho , t03t.Product_No ,sp.Descr,a1.VLListCount , a1.VLListPrintDate ,sp.CaseCnt,sp.STDGROSSWGT,sp.STDCUBE,t01t2.Delivery_Date , t05t2.Expect_Date , t05t2.Expect_Time , t05t2.Dock_No , t05t2.Vehicle_ID_No , t05t2.Driver , " & _
                    "t08m2.C_Name , t01t2.AddWho , t05t2.Drive_Times,l.lottable04,l.lottable05,t03t.extern,t03t.receipt_no,l.lottable06,substring(sp.skugroup,7,1),loc.sectionkey "
        
        Call DB_CheckConnectStatus
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
           tmp_Rs.Close
           cmd_Tab0_PrintReport.Enabled = True
           msg_text = "查詢結果：無符合搜尋條件之訂單明細資料"
           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
           Screen.MousePointer = vbDefault
           blVLLReportEventEnable = True
           cmd_Tab0_PrintReport.Enabled = True
           Screen.MousePointer = 0
           Exit Sub
        End If
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLDetail)
        tmp_Rs.Close
        
End If
        
        '4. 資料寫出 Access 資料庫 >> VLL上貨表
        Call AccessDB_Connect
        Tran_Level = 0
        Tran_Level = cnAccess.BeginTrans
        str_SQL = "Delete From VLL裝載總表"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL裝載總表", cnAccess, adOpenStatic, adLockOptimistic
            
        '取配置資料
        Dim rsTmp As New ADODB.Recordset, strSectionKey As String, lngSectionCS As Long, lngSectionEA As Long, strSectionQty As String
        
        str_SQL = "select distinct sectionkey ,o.updatesource " & _
        "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = o.orderkey " & _
        "join " & strWMSDB & "..loc l (nolock) on l.loc = p.loc " & _
        "left join trp02t t2 on t2.receipt_no = o.updatesource " & _
        "left join trp01t t1 on t2.route_no = t1.route_no " & _
        "where isnull(t1.c_route_no,t1.route_no) in (" & strSelectedRouteNo & ") " & _
        "order by sectionkey ,o.updatesource "
        
        tmp_Rs.Open str_SQL, cn
        Call Replication_Recordset(tmp_Rs, rsTmp)
        tmp_Rs.Close
        
        '取地段數量
        str_SQL = "select route_no = isnull(t1.c_route_no,t1.route_no),sectionkey " & _
                ",CS = case when s.casecnt = 0 then 0 else Isnull(floor(sum(isnull(p.qty,0)) /(s.Casecnt)),0) end " & _
                ",EA = case when s.casecnt = 0 then sum(isnull(p.qty,0)) else cast(sum(isnull(p.qty,0)) as int) % cast(s.Casecnt as int) end " & _
                "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = o.orderkey " & _
                "join " & strWMSDB & "..loc l (nolock) on l.loc = p.loc " & _
                "join gv_skuxpack s on s.storerkey = p.storerkey and s.sku = p.sku " & _
                "left join trp02t t2 on t2.receipt_no = o.updatesource " & _
                "left join trp01t t1 on t2.route_no = t1.route_no " & _
                "where isnull(t1.c_route_no,t1.route_no) in (" & strSelectedRouteNo & ") " & _
                "group by sectionkey ,isnull(t1.c_route_no,t1.route_no),p.orderkey,p.orderlinenumber,s.Casecnt " & _
                "union all select '          ','                                                                                                         ',0,0 " & _
                "order by isnull(t1.c_route_no,t1.route_no),sectionkey  "
        tmp_Rs.Open str_SQL, cn
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLDetailxSection)
        tmp_Rs.Close
        
        rs_Tab0_VLLSum.MoveFirst
        Do While Not rs_Tab0_VLLSum.EOF
        
            '統計路編x地段出貨箱數個數
            rs_Tab0_VLLDetailxSection.MoveFirst
            strRouteNo = rs_Tab0_VLLSum("二次排車路編")
            strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey")
            strSectionQty = "": lngSectionCS = 0: lngSectionEA = 0
            
            '此路編無配貨則下一個路編
            rs_Tab0_VLLDetailxSection.Filter = "(route_no = '" & rs_Tab0_VLLSum("二次排車路編") & "')"
            If rs_Tab0_VLLDetailxSection.EOF Then rs_Tab0_VLLDetailxSection.Filter = "": GoTo nestRoute
            strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey")
    
            Do While Not rs_Tab0_VLLDetailxSection.EOF
            
                If strRouteNo = rs_Tab0_VLLDetailxSection("route_no") Then '同路編
                    If strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey") Then '同一地段
                        lngSectionCS = lngSectionCS + rs_Tab0_VLLDetailxSection("cs")
                        lngSectionEA = lngSectionEA + rs_Tab0_VLLDetailxSection("ea")
                    Else
                        strSectionQty = strSectionQty & RTrim(strSectionKey) & " 共 " & lngSectionCS & "CS / " & lngSectionEA & "EA;"
                        strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey")
                        lngSectionCS = rs_Tab0_VLLDetailxSection("cs")
                        lngSectionEA = rs_Tab0_VLLDetailxSection("ea")
                    End If
                End If
                rs_Tab0_VLLDetailxSection.MoveNext
            Loop
            strSectionQty = strSectionQty & RTrim(strSectionKey) & " 共 " & lngSectionCS & "CS / " & lngSectionEA & "EA ;"
            
            rs_Tab0_VLLDetailxSection.Filter = ""
nestRoute:
        
           rs_Access.AddNew
           rs_Access.Fields("序號").Value = rs_Tab0_VLLSum.Fields("編號").Value
           rs_Access.Fields("路線編號").Value = rs_Tab0_VLLSum.Fields("路線編號").Value
           rs_Access.Fields("出車日期").Value = rs_Tab0_VLLSum.Fields("出車日期").Value
           rs_Access.Fields("車號").Value = rs_Tab0_VLLSum.Fields("車牌號碼").Value
           rs_Access.Fields("司機").Value = rs_Tab0_VLLSum.Fields("駕駛人").Value
           rs_Access.Fields("貨運行").Value = rs_Tab0_VLLSum.Fields("運輸公司").Value
           rs_Access.Fields("訂單編號").Value = rs_Tab0_VLLSum.Fields("貨主單號").Value
           rs_Access.Fields("Receipt_no").Value = rs_Tab0_VLLSum.Fields("Receipt_no").Value
           rs_Access.Fields("客戶編號").Value = rs_Tab0_VLLSum.Fields("客戶編號").Value
           rs_Access.Fields("客戶名稱").Value = rs_Tab0_VLLSum.Fields("客戶名稱").Value
           rs_Access.Fields("送貨地址").Value = rs_Tab0_VLLSum.Fields("送貨地址").Value
           rs_Access.Fields("送貨備註").Value = rs_Tab0_VLLSum.Fields("訂單備註").Value
           rs_Access.Fields("送貨備註").Value = rs_Tab0_VLLSum.Fields("訂單備註").Value
           rs_Access.Fields("出貨板數").Value = rs_Tab0_VLLSum.Fields("板數").Value
           rs_Access.Fields("出貨箱數").Value = rs_Tab0_VLLSum.Fields("箱數").Value
           rs_Access.Fields("出貨個數").Value = rs_Tab0_VLLSum.Fields("個數").Value
           rs_Access.Fields("材積").Value = rs_Tab0_VLLSum.Fields("材積").Value
           rs_Access.Fields("重量").Value = rs_Tab0_VLLSum.Fields("重量").Value
           rs_Access.Fields("列印次數").Value = rs_Tab0_VLLSum.Fields("列印次數").Value + 1 'daniel<第一次列印應為1>
           rs_Access.Fields("列印時間").Value = rs_Tab0_VLLSum.Fields("列印時間").Value
           rs_Access.Fields("指定到期日註記").Value = rs_Tab0_VLLSum.Fields("指定到期日註記").Value
           rs_Access.Fields("LoginUserID").Value = rs_Tab0_VLLSum.Fields("LoginUserID").Value
           rs_Access.Fields("預計報到日期").Value = rs_Tab0_VLLSum.Fields("預計報到日期").Value
           rs_Access.Fields("預計報到時間").Value = rs_Tab0_VLLSum.Fields("預計報到時間").Value
           rs_Access.Fields("碼頭暫存").Value = rs_Tab0_VLLSum.Fields("碼頭暫存").Value
           rs_Access.Fields("車數註記").Value = rs_Tab0_VLLSum.Fields("車數註記").Value
           rs_Access.Fields("二次排車路編").Value = rs_Tab0_VLLSum.Fields("二次排車路編").Value
           rs_Access.Fields("地段數量").Value = IIf(Len(RTrim(strSectionQty)) = 0, "未配置", strSectionQty)

            rsTmp.Filter = "(updatesource = '" & rs_Tab0_VLLSum.Fields("Receipt_no") & "')"
        
            strSectionKey = ""
        
            If rsTmp.EOF Then
                rs_Access.Fields("訂單類型") = "未配置"
            Else
                Do While Not rsTmp.EOF
                    strSectionKey = strSectionKey & RTrim(rsTmp("sectionkey")) & ";"
                    rsTmp.MoveNext
                Loop
                rs_Access.Fields("訂單類型") = strSectionKey
            End If
            
            rsTmp.Filter = ""
           
           rs_Access.Fields("件數").Value = rs_Tab0_VLLSum.Fields("件數").Value
           
           '取所有路編
           If strTmp <> rs_Tab0_VLLSum.Fields("二次排車路編") Then
               strTmp = rs_Tab0_VLLSum.Fields("二次排車路編")
               
               str_SQL = "select distinct route_no from trp01t where isnull(c_route_no,route_no) = '" & rs_Tab0_VLLSum.Fields("二次排車路編") & "'and left(route_no,1) <> 'S' order by route_no "
               Call Confirm_Recordset_Closed(tmp_Rs)
               tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
               strRoute_No = ""
               tmp_Rs.MoveFirst
               Do While Not tmp_Rs.EOF
    
                strRoute_No = strRoute_No & RTrim(tmp_Rs("route_no")) & "; "
                tmp_Rs.MoveNext
    
               Loop
               tmp_Rs.Close
           
           End If
           
           rs_Access.Fields("路編數").Value = strRoute_No & ""
           rs_Access.Update
           rs_Tab0_VLLSum.MoveNext
        Loop
        rs_Tab0_VLLDetailxSection.Close
        rs_Tab0_VLLSum.MoveFirst
        rs_Access.Close

'VLL裝載明細表
If chkDetail = vbChecked Then
    str_SQL = "Delete From VLL裝載明細表"
    cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
    Call ReDim_Recordset(rs_Access)
    rs_Access.Open "VLL裝載明細表", cnAccess, adOpenStatic, adLockOptimistic
    rs_Tab0_VLLDetail.MoveFirst
    Do While Not rs_Tab0_VLLDetail.EOF
       rs_Access.AddNew
       For iLoop = 0 To rs_Tab0_VLLDetail.Fields.Count - 1
           rs_Access.Fields(iLoop).Value = rs_Tab0_VLLDetail.Fields(iLoop).Value
       Next iLoop
       rs_Access.Update
       rs_Tab0_VLLDetail.MoveNext
    Loop
rs_Tab0_VLLDetail.MoveFirst
End If

If chkSUMDetail = vbChecked Then
    str_SQL = "Delete From VLL裝載明細表"
    cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
    Call ReDim_Recordset(rs_Access)
    rs_Access.Open "VLL裝載明細表", cnAccess, adOpenStatic, adLockOptimistic
    rs_Tab0_VLLSUMDetail.MoveFirst
    Do While Not rs_Tab0_VLLSUMDetail.EOF
       rs_Access.AddNew
       For iLoop = 0 To rs_Tab0_VLLSUMDetail.Fields.Count - 1
           rs_Access.Fields(iLoop).Value = rs_Tab0_VLLSUMDetail.Fields(iLoop).Value
       Next iLoop
       rs_Access.Update
       rs_Tab0_VLLSUMDetail.MoveNext
    Loop
rs_Tab0_VLLSUMDetail.MoveFirst
End If


'VLL出貨單
Dim blnPrintVLLorder As Boolean
blnPrintVLLorder = True

str_SQL = "Select * From gv_Report_VLLOrder where 佰事達單號 in (select trp02t.receipt_no from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & "))"

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then '無資料時無須列印
        blnPrintVLLorder = False
Else
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLOrder)
        tmp_Rs.Close
        str_SQL = "Delete From VLL出貨單"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL出貨單", cnAccess, adOpenStatic, adLockOptimistic
        With rs_Tab0_VLLOrder
            .MoveFirst
            Do While Not .EOF
            
            If .Fields("貨主") = "LPSI01" And chkTHLShipList = 0 Then GoTo NextRow
            If .Fields("貨主") = "LKAO01" And chkKAOShipList = 0 Then GoTo NextRow
            If .Fields("貨主") = "LABT01" And chkNSLShipList = 0 Then GoTo NextRow
            If .Fields("貨主") = "LLFA01" And chkLFAShipList = 0 Then GoTo NextRow
            
'            If .Fields("貨主") = "LNSL01" Then
'                If chkNSLShipList = 0 And Left(.Fields("貨主單號"), 1) = "8" Then GoTo NextRow
'            End If
                   rs_Access.AddNew
                   rs_Access.Fields("編號").Value = .Fields("編號").Value
                   rs_Access.Fields("貨主名稱").Value = .Fields("貨主名稱").Value
                   rs_Access.Fields("路線編號").Value = .Fields("路線編號").Value
                   rs_Access.Fields("出車日期").Value = .Fields("出車日期").Value
                   rs_Access.Fields("需求日期").Value = .Fields("需求日期").Value
                   rs_Access.Fields("貨主單號").Value = .Fields("貨主單號").Value
                   rs_Access.Fields("佰事達單號").Value = .Fields("佰事達單號").Value
                   rs_Access.Fields("採購編號").Value = .Fields("採購編號").Value & ""
                   rs_Access.Fields("客戶名稱").Value = .Fields("客戶名稱").Value
                   rs_Access.Fields("客戶地址").Value = .Fields("客戶地址").Value
                   rs_Access.Fields("電話").Value = .Fields("電話").Value
                   rs_Access.Fields("客戶需求").Value = .Fields("客戶需求").Value
                   rs_Access.Fields("備註").Value = .Fields("備註").Value
                   rs_Access.Fields("駕駛").Value = .Fields("駕駛").Value
                   rs_Access.Fields("車號").Value = .Fields("車號").Value
                   rs_Access.Fields("項次").Value = .Fields("項次").Value
                   rs_Access.Fields("貨號").Value = .Fields("貨號").Value
                   rs_Access.Fields("品名").Value = .Fields("品名").Value
                   rs_Access.Fields("箱數").Value = .Fields("出貨箱數").Value
                   rs_Access.Fields("大包裝").Value = .Fields("大包裝").Value
                   rs_Access.Fields("個數").Value = .Fields("出貨個數").Value
                   rs_Access.Fields("小包裝").Value = .Fields("小包裝").Value
                   rs_Access.Fields("總個數").Value = .Fields("總個數").Value
                   rs_Access.Fields("倉別").Value = .Fields("倉別").Value
                   rs_Access.Fields("二次排車路編").Value = .Fields("二次排車路編").Value
                   rs_Access.Fields("件數").Value = .Fields("件數").Value
                '   rs_Access.Fields("製造日").Value = .Fields("製造日").Value
                '   rs_Access.Fields("到期日").Value = .Fields("到期日").Value
                    rs_Access.Fields("USER").Value = User_Name
                   rs_Access.Update
              i = i + 1
               
NextRow:
               .MoveNext
            Loop
        
        .MoveFirst
        End With
End If

'LCHF出貨單 add by Terry 20180724
Dim blnPrintLCHForder As Boolean
blnPrintLCHForder = True

str_SQL = "Select * From Xv_Report_VLLOrder_LCHF01 where 佰事達單號 in (select trp02t.receipt_no from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & "))"

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then '無資料時無須列印
        blnPrintLCHForder = False
Else
        str_SQL = "Delete From VLL出貨單_LCHF01"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL出貨單_LCHF01", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            Do While Not .EOF
                   rs_Access.AddNew
                   'rs_Access.Fields("編號").Value = .Fields("編號").Value
                   rs_Access.Fields("貨主名稱").Value = .Fields("貨主名稱").Value
                   rs_Access.Fields("路線編號").Value = .Fields("路線編號").Value
                   rs_Access.Fields("出車日期").Value = .Fields("出車日期").Value
                   rs_Access.Fields("需求日期").Value = .Fields("需求日期").Value
                   rs_Access.Fields("貨主單號").Value = .Fields("貨主單號").Value
                   rs_Access.Fields("佰事達單號").Value = .Fields("佰事達單號").Value
                   rs_Access.Fields("採購編號").Value = .Fields("採購編號").Value & ""
                   rs_Access.Fields("客戶名稱").Value = .Fields("客戶名稱").Value
                   rs_Access.Fields("客戶地址").Value = .Fields("客戶地址").Value
                   rs_Access.Fields("電話").Value = .Fields("電話").Value
                   rs_Access.Fields("客戶需求").Value = .Fields("客戶需求").Value
                   rs_Access.Fields("備註").Value = .Fields("備註").Value
                   rs_Access.Fields("駕駛").Value = .Fields("駕駛").Value
                   rs_Access.Fields("車號").Value = .Fields("車號").Value
                   rs_Access.Fields("項次").Value = .Fields("項次").Value
                   rs_Access.Fields("貨號").Value = .Fields("貨號").Value
                   rs_Access.Fields("品名").Value = .Fields("品名").Value & " Exp." & .Fields("到期日").Value
                   rs_Access.Fields("箱數").Value = .Fields("出貨箱數").Value
                   rs_Access.Fields("大包裝").Value = .Fields("大包裝").Value
                   rs_Access.Fields("個數").Value = .Fields("出貨個數").Value
                   rs_Access.Fields("小包裝").Value = .Fields("小包裝").Value
                   rs_Access.Fields("總個數").Value = .Fields("總個數").Value
                   rs_Access.Fields("倉別").Value = .Fields("倉別").Value
                   rs_Access.Fields("二次排車路編").Value = .Fields("二次排車路編").Value
                   rs_Access.Fields("件數").Value = .Fields("件數").Value
                   rs_Access.Fields("USER").Value = User_Name
                   rs_Access.Fields("中祥備註").Value = .Fields("中祥備註").Value
                   rs_Access.Update
              i = i + 1
               
               .MoveNext
            Loop
        
        .MoveFirst
        End With
End If

'LNVA出貨單 add by Terry 20190422
Dim blnPrintLNVAorder As Boolean
blnPrintLNVAorder = True

'str_SQL = "Select * From Xv_Report_VLLOrder_LNVA01 where 佰事達單號 in (select trp02t.receipt_no from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & "))"
str_SQL = "delete from codelkup where listname = 'VLLReport_LNVA01'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "insert into CODELKUP (LISTNAME,Code,Description,AddDate,AddWho,EditDate,EditWho) select 'VLLReport_LNVA01',trp02t.receipt_no,'LNVA01出貨單',GETDATE(),'" & User_Name & "',getdate(),'" & User_Name & "' from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & ")"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "exec Xs_VLLReport_LNVA01"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then '無資料時無須列印
        blnPrintLNVAorder = False
Else
        str_SQL = "Delete From VLL出貨單_LNVA01"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL出貨單_LNVA01", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            Do While Not .EOF
                   rs_Access.AddNew
                   'rs_Access.Fields("編號").Value = .Fields("編號").Value
                   rs_Access.Fields("貨主名稱").Value = .Fields("貨主名稱").Value
                   rs_Access.Fields("路線編號").Value = .Fields("路線編號").Value
                   rs_Access.Fields("出車日期").Value = .Fields("出車日期").Value
                   rs_Access.Fields("需求日期").Value = .Fields("需求日期").Value
                   rs_Access.Fields("貨主單號").Value = .Fields("貨主單號").Value
                   rs_Access.Fields("佰事達單號").Value = .Fields("佰事達單號").Value
                   rs_Access.Fields("採購編號").Value = .Fields("採購編號").Value & ""
                   rs_Access.Fields("客戶名稱").Value = .Fields("客戶名稱").Value
                   rs_Access.Fields("客戶地址").Value = .Fields("客戶地址").Value
                   rs_Access.Fields("電話").Value = .Fields("電話").Value
                   rs_Access.Fields("客戶需求").Value = .Fields("客戶需求").Value
                   rs_Access.Fields("備註").Value = .Fields("備註").Value
                   rs_Access.Fields("駕駛").Value = .Fields("駕駛").Value
                   rs_Access.Fields("車號").Value = .Fields("車號").Value
                   rs_Access.Fields("項次").Value = .Fields("項次").Value
                   rs_Access.Fields("貨號").Value = .Fields("貨號").Value
                   rs_Access.Fields("品名").Value = .Fields("品名").Value & "  " & .Fields("批號").Value
                   rs_Access.Fields("箱數").Value = .Fields("出貨箱數").Value
                   rs_Access.Fields("大包裝").Value = .Fields("大包裝").Value
                   rs_Access.Fields("個數").Value = .Fields("出貨個數").Value
                   rs_Access.Fields("小包裝").Value = .Fields("小包裝").Value
                   rs_Access.Fields("總個數").Value = .Fields("總個數").Value
                   rs_Access.Fields("倉別").Value = .Fields("倉別").Value
                   rs_Access.Fields("二次排車路編").Value = .Fields("二次排車路編").Value
                   rs_Access.Fields("件數").Value = .Fields("件數").Value
                    rs_Access.Fields("USER").Value = User_Name
                   rs_Access.Update
              i = i + 1

               .MoveNext
            Loop

        .MoveFirst
        End With
End If

'VTL出貨單
Dim blnPrintVTLorder As Boolean: blnPrintVTLorder = True
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from gv_Report_VTLOrder Where 二次排車路編 IN (" & strSelectedRouteNo & ") or 路線編號 in (" & strSelectedRouteNo & ") "
tmp_Rs.Open str_SQL, cn
If tmp_Rs.EOF Then '無資料時無須列印

        blnPrintVTLorder = False

Else
        str_SQL = "Delete From VTL出貨單"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VTL出貨單", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            
            Do While Not .EOF
            
                '判斷維他露公司別
                If UCase(Left(.Fields("承運商代號"), 1)) = "P" Then
                strCompany = "源興"
                ElseIf UCase(Left(.Fields("出貨單號碼"), 1)) = "V" Then strCompany = "維他露"
                ElseIf UCase(Left(.Fields("出貨單號碼"), 1)) = "C" Then strCompany = "源慶"
                ElseIf UCase(Left(.Fields("出貨單號碼"), 1)) = "E" Then strCompany = "源穎"
                ElseIf UCase(Left(.Fields("產品代號"), 1)) = "O" Then strCompany = "優鮮沛"
                Else
                strCompany = "維他露"
                End If
            
               rs_Access.AddNew
               rs_Access.Fields("出貨單號碼").Value = .Fields("出貨單號碼").Value
               rs_Access.Fields("TMS單號").Value = .Fields("TMS單號").Value
               rs_Access.Fields("排出日期").Value = .Fields("排出日期").Value
               rs_Access.Fields("路線編號").Value = .Fields("路線編號").Value
               rs_Access.Fields("二次排車路編").Value = .Fields("二次排車路編").Value
               rs_Access.Fields("帳款客戶代號").Value = .Fields("帳款客戶代號").Value
               rs_Access.Fields("帳款客戶").Value = .Fields("帳款客戶").Value
               rs_Access.Fields("送貨客戶代號").Value = .Fields("送貨客戶代號").Value
               rs_Access.Fields("送貨客戶").Value = .Fields("送貨客戶").Value
               rs_Access.Fields("棧板使用").Value = .Fields("棧板使用").Value
               rs_Access.Fields("送貨地址").Value = .Fields("送貨地址").Value & ""
               rs_Access.Fields("電話").Value = .Fields("電話").Value
               rs_Access.Fields("承運商代號").Value = .Fields("承運商代號").Value
               rs_Access.Fields("承運商名稱").Value = .Fields("承運商名稱").Value
               rs_Access.Fields("車號").Value = .Fields("車號").Value
               rs_Access.Fields("噸數").Value = .Fields("噸數").Value
               rs_Access.Fields("項次").Value = .Fields("項次").Value
               rs_Access.Fields("原因").Value = .Fields("原因").Value
               rs_Access.Fields("產品代號").Value = .Fields("產品代號").Value
               rs_Access.Fields("產品名稱").Value = .Fields("產品名稱").Value
               rs_Access.Fields("打數").Value = .Fields("打數").Value
               rs_Access.Fields("罐數").Value = .Fields("罐數").Value
'               rs_Access.Fields("箱").Value = .Fields("箱").Value
'               rs_Access.Fields("罐").Value = .Fields("罐").Value
'               rs_Access.Fields("總罐數").Value = .Fields("總罐數").Value
               rs_Access.Fields("備註").Value = .Fields("備註").Value
               rs_Access.Fields("USER").Value = User_Name
               rs_Access.Fields("公司別").Value = strCompany
               rs_Access.Update
               .MoveNext
            Loop
        
        .MoveFirst
        End With
End If

'YFY出貨單
Dim blnPrintYFYorder As Boolean: blnPrintYFYorder = True
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "select * from ev_Report_YFYOrder Where 路線編號 in (" & strSelectedRouteNo & ") or 二次路線編號 in (" & strSelectedRouteNo & ")"

tmp_Rs.Open str_SQL, cn
If tmp_Rs.EOF Then '無資料時無須列印

        blnPrintYFYorder = False

Else
        str_SQL = "Delete From YFY出貨單"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "YFY出貨單", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            Do While Not .EOF
               rs_Access.AddNew
               rs_Access.Fields("TMS單號").Value = .Fields("TMS單號").Value
               rs_Access.Fields("訂單號碼").Value = .Fields("訂單號碼").Value
               'rs_Access.Fields("訂單細項").Value = .Fields("訂單細項").Value
               rs_Access.Fields("貨主名稱").Value = .Fields("貨主名稱").Value
               rs_Access.Fields("客戶名稱").Value = .Fields("客戶名稱").Value
               rs_Access.Fields("備註").Value = .Fields("備註").Value
               rs_Access.Fields("地址").Value = .Fields("地址").Value
               rs_Access.Fields("聯絡人").Value = .Fields("聯絡人").Value
               rs_Access.Fields("列印日期").Value = .Fields("列印日期").Value
               rs_Access.Fields("路線編號").Value = .Fields("路線編號").Value
               'rs_Access.Fields("二次路線編號").Value = .Fields("二次路線編號").Value
               rs_Access.Fields("車號").Value = .Fields("車號").Value
               rs_Access.Fields("品號").Value = .Fields("品號").Value
               rs_Access.Fields("品名").Value = .Fields("品名").Value
               rs_Access.Fields("數量").Value = .Fields("數量").Value
               rs_Access.Fields("材積").Value = .Fields("材積").Value
               rs_Access.Fields("客戶訂單號碼").Value = .Fields("客戶訂單號碼").Value
               rs_Access.Fields("採購單號").Value = .Fields("採購單號").Value
               rs_Access.Fields("出貨日").Value = .Fields("出貨日").Value
               rs_Access.Update
               .MoveNext
            Loop

        .MoveFirst
        End With
End If

cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'5. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab0_PreView.Value = vbChecked Then
   '預覽列印
   MSAccessAP.Visible = True
   If chkVllPallet = vbChecked Then
      MSAccessAP.DoCmd.OpenReport "VLL裝載總表_簡版", acViewPreview
   Else
      MSAccessAP.DoCmd.OpenReport "VLL裝載總表", acViewPreview
   End If
      
   MSAccessAP.DoCmd.OpenReport "車輛放行單", acViewPreview
   
   If chkSUMDetail.Value = vbChecked Then
    MSAccessAP.DoCmd.OpenReport "VLL裝載明細匯總表", acViewPreview
   End If
   
   If chkDetail.Value = vbChecked Then
    MSAccessAP.DoCmd.OpenReport "VLL裝載明細表", acViewPreview
   End If
   
   If blnPrintVLLorder = True And i > 0 Then MSAccessAP.DoCmd.OpenReport "VLL出貨單", acViewPreview
   If blnPrintLCHForder = True Then MSAccessAP.DoCmd.OpenReport "VLL出貨單_LCHF01", acViewPreview
   If blnPrintLNVAorder = True Then MSAccessAP.DoCmd.OpenReport "VLL出貨單_LNVA01", acViewPreview
   If blnPrintVTLorder = True Then MSAccessAP.DoCmd.OpenReport "VTL出貨單", acViewPreview
   If blnPrintYFYorder = True Then MSAccessAP.DoCmd.OpenReport "YFY出貨單", acViewPreview
   MSAccessAP.DoCmd.Maximize
Else
   '直接列印至印表機
   MSAccessAP.Visible = False
   If chkVllPallet = vbChecked Then
        MSAccessAP.DoCmd.OpenReport "VLL裝載總表_簡版", acViewNormal
   Else
        MSAccessAP.DoCmd.OpenReport "VLL裝載總表", acViewNormal
   End If
   MSAccessAP.DoCmd.OpenReport "車輛放行單", acViewPreview
   
   If chkSUMDetail.Value = vbChecked Then
        MSAccessAP.DoCmd.OpenReport "VLL裝載明細匯總表", acViewNormal
   End If
   
   If chkDetail.Value = vbChecked Then
        MSAccessAP.DoCmd.OpenReport "VLL裝載明細表", acViewNormal
   End If
   
   If blnPrintVLLorder = True And i > 0 Then MSAccessAP.DoCmd.OpenReport "VLL出貨單", acViewNormal
   If blnPrintLCHForder = True Then MSAccessAP.DoCmd.OpenReport "VLL出貨單_LCHF01", acViewNormal
   If blnPrintLNVAorder = True Then MSAccessAP.DoCmd.OpenReport "VLL出貨單_LNVA01", acViewNormal
   If blnPrintVTLorder = True Then MSAccessAP.DoCmd.OpenReport "VTL出貨單", acViewNormal
   If blnPrintYFYorder = True Then MSAccessAP.DoCmd.OpenReport "YFY出貨單", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If

cmd_Tab0_PrintReport.Enabled = True     'daniel
Screen.MousePointer = 0
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
   cmd_Tab0_PrintReport.Enabled = True  'daniel
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL上貨表-列印", Me.Caption, "cmd_Tab0_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab0_Query_Click()
On Error GoTo err_Handle

'回傳揀貨量
str_SQL = "select o.route " & _
        ",o.storerkey " & _
        ",o.orderkey " & _
        ",o.updatesource " & _
        ",o.Externorderkey " & _
        ",ExternLineno = case when o.storerkey = 'LLFA01' and o.IncoTerm <> '' then od.orderlinenumber else od.ExternLineno end " & _
        ",od.sku " & _
        ",shippedqty = (od.shippedqty + od.qtyallocated + od.qtypicked) " & _
        ",od.editdate " & _
        ",o.status " & _
        "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey and o.yfystatus = '0' and convert(char(8),o.deliverydate,112) > convert(char(8),getdate()-7,112) " & _
        "where (od.shippedqty + od.qtyallocated + od.qtypicked) > 0 " & _
        "and len(rtrim(isnull(o.updatesource,''))) > 9 and o.updatesource in (select t2.receipt_no from trp02t t2 where t2.receipt_no = o.updatesource and t2.exe_confirm = 2) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'無資料
If Not tmp_Rs.EOF Then

    tmp_Rs.MoveFirst
    Tran_Level = cn.BeginTrans
    Do While Not tmp_Rs.EOF
    
            str_SQL = "UPDATE TRP03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03W set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '寫入紀錄
'            Call WriteLog(Err.Number & Chr(9) & "揀貨數量確認" & Chr(9) & "WMS: " & tmp_Rs("orderkey") & ",TMS: " & tmp_Rs("route") & "," & tmp_Rs("storerkey") & "," & tmp_Rs("updatesource") & "," & RTrim(tmp_Rs("Externorderkey")) & "," & tmp_Rs("Externlineno") & "," & tmp_Rs("sku") & "," & tmp_Rs("shippedqty") & "," & User_id)
            
            '更新YFYstatus回傳狀態
            If Trim(tmp_Rs("status")) = "9" And Trim(tmp_Rs("storerkey")) <> "LTKK01" Then
                str_SQL = "UPDATE " & strWMSDB & "..Orders set YFYstatus = '1' ,TrafficCop = null where orderkey = '" & tmp_Rs("orderkey") & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
        
        tmp_Rs.MoveNext
    Loop
    
    '直接出貨量=訂單量 mark by Gemini @20150805 4 SHIP_QTY 改由TRP02T Trigger 寫入
'            str_SQL = "UPDATE TRP03T set TRP03T.SHIP_QTY=TRP03T.order_qty from trp02t join trp03t on trp02t.receipt_no = trp03t.receipt_no where trp02t.priority = 'C' and ship_qty = 0 "
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
    cn.CommitTrans: Tran_Level = 0
End If

tmp_Rs.Close

'補LF的配置量 add by Eric
Call LLFA01Ship2TMS

'VLL上貨表 >> 查詢
Set dg_Tab0_VLL.DataSource = Nothing
Set rs_Tab0_VLL = Nothing
blVLLReportEventEnable = False  '諮詢

Screen.MousePointer = vbHourglass

'str_SQL = "Select ' ' as '＊',路線編號 ,出車日期,列印次數,列印時間,車牌號碼,車次,駕駛人,運輸公司,排車箱數,揀貨箱數 " & _
'          "From Report_VLL_RouteList "
str_SQL = "Select ' ' as '＊',t01t.Route_No as 路線編號 , Convert(varchar(8),t01t.Delivery_Date,112) as 出車日期 , Isnull(t01t.VLListCount,0) as 列印次數 , " & _
        "Isnull((Convert(varchar,Isnull(t01t.VLListPrintDate,Getdate()),111) + ' ' + Convert(varchar,Isnull(t01t.VLListPrintDate,Getdate()) , 108)),'') as 列印時間 , " & _
        "Rtrim(t05t.Vehicle_ID_No) as 車牌號碼 , t05t.Drive_Times as 車次 , Rtrim(t05t.Driver) as 駕駛人 , " & _
        "Rtrim(Isnull(t08m.C_Name,'')) as 運輸公司, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(order_qty),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(order_qty),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as 排車個數, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(SHIP_QTY),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(SHIP_QTY),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as 揀貨個數, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(WEIGHT),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(WEIGHT),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as 排車重量, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(VOLUMN_WEIGHT),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(VOLUMN_WEIGHT),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as 排車材積 " & _
        "From TRP01T t01t " & _
        "inner join TRP05T t05t on t05t.Route_No = t01t.Route_No and convert(char(8),t01t.delivery_date,112) > convert(char(8),getdate()-7,112) " & _
        "inner join TRP09M t09m on t09m.Vehicle_ID_No = t05t.Vehicle_ID_No " & _
        "Left outer join TRP08M t08m on t08m.Company_Code = t09m.TRP_Company_Code "
        
'揀貨材積重量
'        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(t3.ship_qty * sp.stdgrosswgt),0),0) from trp03t t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where t3.Route_No=t01t.Route_No ) else (select isnull(round(sum(t3.ship_qty * sp.stdgrosswgt),0),0) from trp03t_S t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where Route_No_S=t01t.Route_No ) end as 排車重量, " & _
'        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(t3.ship_qty * sp.stdcube),0),0) from trp03t t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where t3.Route_No=t01t.Route_No ) else (select isnull(round(sum(t3.ship_qty * sp.stdcube),0),0) from trp03t_S t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where Route_No_S=t01t.Route_No ) end as 排車材積 " & _

Dim strWhere As String, strTmp As String
strWhere = "Where t01t.Route_No <> 'D'"

'出車日期
strTmp = ""
If Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),t01t.Delivery_Date,112) between '" & txt_Tab0_DeliveryDate_Start.Text & "' and '" & txt_Tab0_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(varchar(8),t01t.Delivery_Date,112) = '" & txt_Tab0_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) = 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),t01t.Delivery_Date,112) = '" & txt_Tab0_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'路線編號
strTmp = ""
If Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   strTmp = " t01t.Route_No between '" & txt_Tab0_RouteNo_Start.Text & "' and '" & txt_Tab0_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) = 0 Then
   strTmp = " t01t.Route_No = '" & txt_Tab0_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) = 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   strTmp = " t01t.Route_No = '" & txt_Tab0_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'含已列印過的 Wave
strTmp = ""
If chk_Tab0_PrintedRoute.Value = vbUnchecked Then
   strTmp = " Isnull(t01t.VLListCount,0) = 0 "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp & " and t01t.C_ROUTE_NO is null"
   Else
      strWhere = strWhere & strTmp & " and t01t.C_ROUTE_NO is null"
   End If
End If

'只顯示大route
If Len(strWhere) > 0 Then
   strWhere = strWhere & " and t01t.C_ROUTE_NO is null"
Else
   strWhere = "t01t.C_ROUTE_NO is null"
End If
If strWhere <> "" Then
   str_SQL = str_SQL & strWhere
Else
   msg_text = "基於縮小查詢資料量，請適度設定查詢條件"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by 路線編號 "
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
cn.CommandTimeout = 300

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   blVLLReportEventEnable = True
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0_VLL)
tmp_Rs.Close

With dg_Tab0_VLL
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab0_VLL.MoveFirst
Set dg_Tab0_VLL.DataSource = rs_Tab0_VLL

SetDataGridColWidth "VLL裝載", dg_Tab0_VLL

With dg_Tab0_VLL
    .ColumnHeaders = True         '標題行顯示
    .RowHeight = 300

End With
rs_Tab0_VLL.MoveFirst
rs_Tab0_VLL.Filter = adFilterNone
rs_Tab0_VLL.Sort = " 編號 "
rs_Tab0_VLL.MoveFirst
blVLLReportEventEnable = True
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL上貨表-查詢", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Reset_Click()
'VLL上貨表 >> 清除
txt_Tab0_DeliveryDate_Start.Text = "": txt_Tab0_DeliveryDate_End.Text = ""
txt_Tab0_RouteNo_Start.Text = "": txt_Tab0_RouteNo_End.Text = ""
chk_Tab0_PrintedRoute.Value = vbUnchecked
Set dg_Tab0_VLL.DataSource = Nothing
Set rs_Tab0_VLL = Nothing
End Sub

Private Sub cmd_Tab0_SaveToExcel_Click()
'VLL上貨表 >> 轉 EXCEL
blVLLReportEventEnable = False

Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel 檔案名稱
CmnDialog.DialogTitle = "轉存 Excel 檔"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "VLL上貨表_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel檔案(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
   msg_text = "選擇 [取消] 按鈕，必須於 Excel 中自行存檔"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab0_VLL) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab0_VLL.MoveFirst
Exit Sub

err_Handle:
   blVLLReportEventEnable = True
   Dim tmpString As String
   Screen.MousePointer = vbDefault
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL上貨表-轉 EXCEL", Me.Caption, "cmd_Tab0_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab1_PrintReport_Click()
'車輛裝載匯總表 >> 報表列印
If rs_Tab1_VLLSum Is Nothing Then Exit Sub
If rs_Tab1_VLLSum.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. 資料寫出 Access 資料庫 >> 車輛裝載匯總表
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 車輛裝載匯總表"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "車輛裝載匯總表", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab1_VLLSum.MoveFirst
Do While Not rs_Tab1_VLLSum.EOF
   rs_Access.AddNew
   rs_Access.Fields("SerialNo").Value = rs_Tab1_VLLSum.Fields("編號").Value
   rs_Access.Fields("路線編號").Value = rs_Tab1_VLLSum.Fields("路線編號").Value
   rs_Access.Fields("車號").Value = rs_Tab1_VLLSum.Fields("車牌號碼").Value
   rs_Access.Fields("車次").Value = rs_Tab1_VLLSum.Fields("車次").Value
   rs_Access.Fields("出車日期").Value = rs_Tab1_VLLSum.Fields("出車日期").Value
   rs_Access.Fields("運輸公司簡稱").Value = rs_Tab1_VLLSum.Fields("運輸公司簡稱").Value
   rs_Access.Fields("特殊需求1").Value = rs_Tab1_VLLSum.Fields("特殊需求2").Value
   rs_Access.Fields("特殊需求2").Value = rs_Tab1_VLLSum.Fields("特殊需求1").Value
   rs_Access.Fields("貨主").Value = rs_Tab1_VLLSum.Fields("貨主").Value
   rs_Access.Fields("客戶訂單編號").Value = rs_Tab1_VLLSum.Fields("貨主單號").Value
   rs_Access.Fields("客戶編號").Value = rs_Tab1_VLLSum.Fields("客戶編號").Value
   rs_Access.Fields("客戶名稱").Value = rs_Tab1_VLLSum.Fields("客戶名稱").Value
   rs_Access.Fields("郵遞區號").Value = rs_Tab1_VLLSum.Fields("zip").Value
   rs_Access.Fields("送貨地址").Value = rs_Tab1_VLLSum.Fields("送貨地址").Value
   rs_Access.Fields("箱數").Value = rs_Tab1_VLLSum.Fields("箱數").Value
   rs_Access.Fields("個數").Value = rs_Tab1_VLLSum.Fields("個數").Value
   rs_Access.Fields("材積").Value = rs_Tab1_VLLSum.Fields("材積").Value
   rs_Access.Fields("重量").Value = rs_Tab1_VLLSum.Fields("重量").Value
   rs_Access.Fields("板數").Value = rs_Tab1_VLLSum.Fields("板數").Value
   rs_Access.Fields("二次排車路編").Value = rs_Tab1_VLLSum.Fields("二次排車路編").Value
   rs_Access.Fields("碼頭暫存").Value = rs_Tab1_VLLSum.Fields("碼頭暫存").Value
   rs_Access.Fields("預計報到日期時間").Value = rs_Tab1_VLLSum.Fields("預計報到日期時間").Value
   rs_Access.Fields("訂單類型").Value = rs_Tab1_VLLSum.Fields("訂單類型").Value
   rs_Access.Fields("客戶需求").Value = rs_Tab1_VLLSum.Fields("客戶需求").Value
   rs_Access.Fields("訂單備註").Value = rs_Tab1_VLLSum.Fields("訂單備註").Value
   rs_Access.Update
   rs_Tab1_VLLSum.MoveNext
Loop
rs_Tab1_VLLSum.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab1_PreView.Value = vbChecked Then
   '預覽列印
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "車輛裝載匯總表", acViewPreview
Else
   '直接列印至印表機
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "車輛裝載匯總表", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
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
   CreateErrorLog Me.Name & "--車輛裝載匯總表列印", Me.Caption, "cmd_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab1_Query_Click()
'車輛裝載匯總表 >> 查詢
Set dg_Tab1_VLLSum.DataSource = Nothing
Set rs_Tab1_VLLSum = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select 車牌號碼,路線編號,車次,出車日期,運輸公司簡稱,特殊需求1,特殊需求2,訂單編號,貨主,貨主單號," & _
          "   客戶編號,客戶名稱,ZIP,送貨地址,箱數,個數,板數,材積,重量,運送區域,二次排車路編,碼頭暫存,預計報到日期時間,訂單類型,客戶需求,訂單備註  " & _
          "From Report_LoadingSummary "

Dim strWhere As String, strTmp As String
strWhere = ""
'訂單日期
strTmp = ""
If Len(txt_Tab1_DeliveryDate_Start.Text) > 0 And Len(txt_Tab1_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 between '" & txt_Tab1_DeliveryDate_Start.Text & "' and '" & txt_Tab1_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab1_DeliveryDate_Start.Text) > 0 And Len(txt_Tab1_DeliveryDate_End.Text) = 0 Then
   strTmp = " 出車日期 = '" & txt_Tab1_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab1_DeliveryDate_Start.Text) = 0 And Len(txt_Tab1_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 = '" & txt_Tab1_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'路線編號
strTmp = ""
If Len(txt_Tab1_route_Start.Text) > 0 And Len(txt_Tab1_route_End.Text) > 0 Then
   strTmp = " 路線編號 between '" & txt_Tab1_route_Start.Text & "' and '" & txt_Tab1_route_End.Text & "' "
ElseIf Len(txt_Tab1_route_Start.Text) > 0 And Len(txt_Tab1_route_End.Text) = 0 Then
   strTmp = " 路線編號 = '" & txt_Tab1_route_Start.Text & "' "
ElseIf Len(txt_Tab1_route_Start.Text) = 0 And Len(txt_Tab1_route_End.Text) > 0 Then
   strTmp = " 路線編號 = '" & txt_Tab1_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'運送區域
strTmp = ""
If cmb_Tab1_AreaCode.ListIndex <> -1 Then
   strTmp = " 運送區域代碼 = '" & arAreaCode(cmb_Tab1_AreaCode.ListIndex) & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "基於縮小查詢資料量，請適度設定查詢條件"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by 車牌號碼,路線編號,車次,出車日期 "
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab1_VLLSum)
tmp_Rs.Close

With dg_Tab1_VLLSum
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab1_VLLSum.MoveFirst
Set dg_Tab1_VLLSum.DataSource = rs_Tab1_VLLSum

With dg_Tab1_VLLSum
    .ColumnHeaders = True          '標題行顯示
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '車牌號碼
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1100       '路線編號
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 500        '車次
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 900        '出車日期
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1200       '運輸公司簡稱
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1500       '特殊需求1
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1500       '特殊需求2
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1100       '訂單編號
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 500        '貨主
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 900       '貨主單號
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1100      '客戶編號
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 2500      '客戶名稱
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 500       'ZIP
    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 2400      '送貨地址
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 800       '箱數
    .Columns(15).Alignment = dbgRight
    .Columns(16).Width = 800       '板數
    .Columns(16).Alignment = dbgRight
    .Columns(17).Width = 800       '材積
    .Columns(17).Alignment = dbgRight
    .Columns(18).Width = 800       '重量
    .Columns(18).Alignment = dbgRight
    .Columns(19).Width = 3400      '運送區域
    .Columns(19).Alignment = dbgLeft
    .Columns(20).Width = 1300      '二次排車路編
    .Columns(20).Alignment = dbgLeft
    .Columns(21).Width = 1300      '碼頭暫存
    .Columns(21).Alignment = dbgLeft
    .Columns(22).Width = 1300      '預計報到日期時間
    .Columns(22).Alignment = dbgLeft
End With
rs_Tab1_VLLSum.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車輛裝載匯總表-查詢", Me.Caption, "cmd_Tab1_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Reset_Click()
'車輛裝載匯總表 >> 清除
cmb_Tab1_AreaCode.ListIndex = -1
txt_Tab1_DeliveryDate_Start.Text = ""
txt_Tab1_DeliveryDate_End.Text = ""
txt_Tab1_route_Start.Text = ""
txt_Tab1_route_End.Text = ""

Set dg_Tab1_VLLSum.DataSource = Nothing
Set rs_Tab1_VLLSum = Nothing
End Sub

Private Sub cmd_Tab1_SaveToExcel_Click()
'車輛裝載匯總表 >> 轉 EXCEL
Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel 檔案名稱
CmnDialog.DialogTitle = "轉存 Excel 檔"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "車輛裝載匯總表_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel檔案(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
   msg_text = "選擇 [取消] 按鈕，必須於 Excel 中自行存檔"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab1_VLLSum) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab1_VLLSum.MoveFirst
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車輛裝載匯總表-轉 EXCEL", Me.Caption, "cmd_Tab1_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_PrintReport_Click()
'訂單總表 >> 報表列印
If rs_Tab2_OrdersSum Is Nothing Then Exit Sub
If rs_Tab2_OrdersSum.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. 資料寫出 Access 資料庫 >> 訂單總表
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 訂單總表"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords

Call ReDim_Recordset(rs_Access)
rs_Access.Open "訂單總表", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab2_OrdersSum.MoveFirst

Do While Not rs_Tab2_OrdersSum.EOF

'    '檢查欣臨是否件數確認
'    If rs_Tab2_OrdersSum.Fields("貨主編號").Value = "LTHL01" And Len(RTrim(rs_Tab2_OrdersSum.Fields("件數確認時間").Value)) = 0 Then
'        MsgBox "欣臨尚有未確認件數!", 64, "訂單總表列印中止"
'        cnAccess.RollbackTrans
'        rs_Access.Close
'        Set rs_Access = Nothing
'        Exit Sub
'    End If
    
   rs_Access.AddNew
   rs_Access.Fields("序號").Value = rs_Tab2_OrdersSum.Fields("編號").Value
   rs_Access.Fields("路線編號").Value = rs_Tab2_OrdersSum.Fields("路線編號").Value
   rs_Access.Fields("出貨日期").Value = rs_Tab2_OrdersSum.Fields("到貨日").Value
   rs_Access.Fields("車號").Value = rs_Tab2_OrdersSum.Fields("車牌號碼").Value
   rs_Access.Fields("司機").Value = rs_Tab2_OrdersSum.Fields("駕駛人").Value
   rs_Access.Fields("車次").Value = rs_Tab2_OrdersSum.Fields("車次").Value
   rs_Access.Fields("出車日期").Value = rs_Tab2_OrdersSum.Fields("出車日期").Value
   rs_Access.Fields("貨運行").Value = rs_Tab2_OrdersSum.Fields("運輸公司").Value
   rs_Access.Fields("貨主單號").Value = rs_Tab2_OrdersSum.Fields("貨主單號").Value
   rs_Access.Fields("客戶編號").Value = rs_Tab2_OrdersSum.Fields("客戶編號").Value
   rs_Access.Fields("客戶名稱").Value = rs_Tab2_OrdersSum.Fields("客戶名稱").Value
   rs_Access.Fields("郵遞區號").Value = rs_Tab2_OrdersSum.Fields("zip").Value
   rs_Access.Fields("送貨地址").Value = rs_Tab2_OrdersSum.Fields("送貨地址").Value
   rs_Access.Fields("送貨備註").Value = rs_Tab2_OrdersSum.Fields("訂單備註").Value
   rs_Access.Fields("指定到期日註記").Value = rs_Tab2_OrdersSum.Fields("註記").Value
   rs_Access.Fields("貨主PO").Value = rs_Tab2_OrdersSum.Fields("貨主PO").Value
   rs_Access.Fields("訂單類型").Value = rs_Tab2_OrdersSum.Fields("訂單類型").Value
   rs_Access.Fields("冷藏").Value = rs_Tab2_OrdersSum.Fields("冷藏").Value
   rs_Access.Fields("暫存區").Value = rs_Tab2_OrdersSum.Fields("暫存區").Value
   rs_Access.Fields("客戶需求").Value = rs_Tab2_OrdersSum.Fields("客戶需求").Value
   rs_Access.Fields("箱數").Value = rs_Tab2_OrdersSum.Fields("箱數").Value
   rs_Access.Fields("個數").Value = rs_Tab2_OrdersSum.Fields("個數").Value
   rs_Access.Update
   rs_Tab2_OrdersSum.MoveNext
Loop
rs_Tab2_OrdersSum.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab2_PreView.Value = vbChecked Then
   '預覽列印
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "訂單總表", acViewPreview
Else
   '直接列印至印表機
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "訂單總表", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cnAccess.RollbackTrans
      Tran_Level = 0
   End If
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單總表-列印", Me.Caption, "cmd_Tab2_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab2_Query_Click()
'訂單總表 >> 查詢
Set dg_Tab2_OrdersSum.DataSource = Nothing
Set rs_Tab2_OrdersSum = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select 出車日期,到貨日,路線編號,車牌號碼,車次,駕駛人,運輸公司,貨主編號,貨主單號,客戶編號,客戶名稱,送貨地址,訂單備註,訂單日,'          ' as 註記,ZIP,貨主PO,訂單類型, 冷藏, 暫存區, 客戶需求, 箱數, 個數 ,件數確認時間 " & _
          "From Report_OrdersSum "

Dim strWhere As String, strTmp As String
strWhere = ""
'出車日期
strTmp = ""
If Len(txt_Tab2_DeliveryDate_Start.Text) > 0 And Len(txt_Tab2_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 between '" & txt_Tab2_DeliveryDate_Start.Text & "' and '" & txt_Tab2_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab2_DeliveryDate_Start.Text) > 0 And Len(txt_Tab2_DeliveryDate_End.Text) = 0 Then
   strTmp = " 出車日期 = '" & txt_Tab2_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab2_DeliveryDate_Start.Text) = 0 And Len(txt_Tab2_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 = '" & txt_Tab2_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'路線編號
strTmp = ""
If Len(txt_Tab2_RouteNo_Start.Text) > 0 And Len(txt_Tab2_RouteNo_End.Text) > 0 Then
   strTmp = " 路線編號 between '" & txt_Tab2_RouteNo_Start.Text & "' and '" & txt_Tab2_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab2_RouteNo_Start.Text) > 0 And Len(txt_Tab2_RouteNo_End.Text) = 0 Then
   strTmp = " 路線編號 = '" & txt_Tab2_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab2_RouteNo_Start.Text) = 0 And Len(txt_Tab2_RouteNo_End.Text) > 0 Then
   strTmp = " 路線編號 = '" & txt_Tab2_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "基於縮小查詢資料量，請適度設定查詢條件"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by 路線編號,貨主單號 "
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab2_OrdersSum)
tmp_Rs.Close

With dg_Tab0_VLL
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab2_OrdersSum.MoveFirst
Set dg_Tab2_OrdersSum.DataSource = rs_Tab2_OrdersSum

With dg_Tab2_OrdersSum
    .ColumnHeaders = True         '標題行顯示
    .RowHeight = 250
    .Columns(0).Width = 500       '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000      '出車日期
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000      '到貨日
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1100      '路線編號
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 900       '車牌號碼
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 500       '車次
    .Columns(5).Alignment = dbgCenter
    .Columns(6).Width = 1000      '駕駛人
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1500      '運輸公司
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 900       '貨主單號
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 1200      '客戶編號
    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 2000     '客戶名稱
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 2000     '送貨地址
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1200     '客戶備註
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 900      '訂單日
    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 1400      '指定到期日註記
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 500      'ZIP
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1200      '貨主PO
    .Columns(16).Alignment = dbgLeft
End With
'更新 [註記] 欄位  ==> 訂單有 [指定到期日] 出貨，寫入 [***]
rs_Tab2_OrdersSum.MoveFirst
Do While Not rs_Tab2_OrdersSum.EOF
   str_SQL = "Select Count(LotTable05) as 'CNT' From OrderDetail Where ExternOrderKey = '" & rs_Tab2_OrdersSum.Fields("貨主單號").Value & "' and LotTable05 is not null "
   tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
   If tmp_Rs.Fields("CNT").Value = 0 Then
      rs_Tab2_OrdersSum.Fields("註記").Value = ""
   Else
      rs_Tab2_OrdersSum.Fields("註記").Value = "指定到期日"
   End If
   tmp_Rs.Close
   rs_Tab2_OrdersSum.MoveNext
Loop
rs_Tab2_OrdersSum.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單總表-查詢", Me.Caption, "cmd_Tab2_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Reset_Click()
'訂單總表 >> 清除
txt_Tab2_DeliveryDate_Start.Text = "": txt_Tab2_DeliveryDate_End.Text = ""
txt_Tab2_RouteNo_Start.Text = "": txt_Tab2_RouteNo_End.Text = ""
Set dg_Tab2_OrdersSum.DataSource = Nothing
Set rs_Tab2_OrdersSum = Nothing

End Sub

Private Sub cmd_Tab3_Query_Click()
'揀貨裝載稽核表 >> 查詢
Set dg_Tab3_PickLoadCheck.DataSource = Nothing
Set rs_Tab3_PickLoadCheck = Nothing

txt_Tab3_UploadDate_Start.Text = Trim(txt_Tab3_UploadDate_Start.Text)
txt_Tab3_UploadHour_Start.Text = Format(Val(txt_Tab3_UploadHour_Start.Text), "00")
txt_Tab3_UploadMinute_Start.Text = Format(Val(txt_Tab3_UploadMinute_Start.Text), "00")
txt_Tab3_UploadDate_End.Text = Trim(txt_Tab3_UploadDate_End.Text)
txt_Tab3_UploadHour_End.Text = Format(Val(txt_Tab3_UploadHour_End.Text), "00")
txt_Tab3_UploadMinute_End.Text = Format(Val(txt_Tab3_UploadMinute_End.Text), "00")

'If Len(txt_Tab3_UploadDate_Start.Text) = 0 Or Len(txt_Tab3_UploadDate_End.Text) = 0 Then
'   msg_text = "資料驗證：請輸入 [回傳日期] "
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'ElseIf Val(txt_Tab3_UploadHour_Start.Text) = 24 Or Val(txt_Tab3_UploadHour_End.Text) = 24 Then
'   'Wave 建立時間範圍不接受 24：00
'   msg_text = "資料驗證：[回傳日期] 資料錯誤，" & vbCrLf & "" & vbCrLf & _
'              "可接受的資料範圍：yyyymmdd 00：00 ～ yyyymmdd 23：59"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   If Val(txt_Tab3_UploadHour_Start.Text) = 24 Then
'      txt_Tab3_UploadHour_Start.SelStart = 0: txt_Tab3_UploadHour_Start.SelLength = Len(txt_Tab3_UploadHour_Start.Text)
'      txt_Tab3_UploadHour_Start.SetFocus
'   Else
'      txt_Tab3_UploadHour_End.SelStart = 0: txt_Tab3_UploadHour_End.SelLength = Len(txt_Tab3_UploadHour_End.Text)
'      txt_Tab3_UploadHour_End.SetFocus
'   End If
'   Exit Sub
'End If
'txt_Tab3_UploadMinute_Start.Text = Trim(txt_Tab3_UploadMinute_Start.Text)
'txt_Tab3_UploadMinute_End.Text = Trim(txt_Tab3_UploadMinute_End.Text)
'If Len(txt_Tab3_UploadMinute_Start.Text) = 0 Or Len(txt_Tab3_UploadMinute_Start.Text) = 0 Then
'   msg_text = "資料驗證：請輸入 [回傳時間] "
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'End If

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select 運送區域,路線編號,訂單數,送貨點,客戶簡稱,箱數,板數,重量,材積,貨運公司,車號,車次,回傳日期,出車日期,預計報到時間,碼頭 " & _
          "From Report_PickLoadCheck "

Dim tmpString1 As String, tmpString2 As String
Dim strWhere As String, strTmp As String
strWhere = ""
'運送區域
strTmp = ""
If cmb_Tab3_AreaCode.ListIndex <> -1 Then
   strTmp = " 運送區域碼 = '" & arAreaCode(cmb_Tab3_AreaCode.ListIndex) & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'指定 [回傳日期]
If txt_Tab3_UploadDate_Start.Text <> "" And txt_Tab3_UploadDate_End.Text <> "" Then
   strTmp = ""
   tmpString1 = Mid(txt_Tab3_UploadDate_Start.Text, 1, 4) & "-" & Mid(txt_Tab3_UploadDate_Start.Text, 5, 2) & "-" & Mid(txt_Tab3_UploadDate_Start.Text, 7, 2) & " " & _
                txt_Tab3_UploadHour_Start.Text & ":" & txt_Tab3_UploadMinute_Start.Text & ":00"
   tmpString2 = Mid(txt_Tab3_UploadDate_End.Text, 1, 4) & "-" & Mid(txt_Tab3_UploadDate_End.Text, 5, 2) & "-" & Mid(txt_Tab3_UploadDate_End.Text, 7, 2) & " " & _
                txt_Tab3_UploadHour_End.Text & ":" & txt_Tab3_UploadMinute_End.Text & ":00"
   strTmp = "UploadDate between convert(datetime,'" & tmpString1 & "',120) and convert(datetime,'" & tmpString2 & "',120) "
   If Len(strTmp) > 0 Then
      If Len(strWhere) = 0 Then
         strWhere = strTmp
      Else
         strWhere = strWhere & " and " & strTmp
      End If
   End If
End If
'路線編號
strTmp = ""
If Len(txt_Tab3_route_Start.Text) > 0 And Len(txt_Tab3_route_End.Text) > 0 Then
   strTmp = " 路線編號 between '" & txt_Tab3_route_Start.Text & "' and '" & txt_Tab3_route_End.Text & "' "
ElseIf Len(txt_Tab3_route_Start.Text) > 0 And Len(txt_Tab3_route_End.Text) = 0 Then
   strTmp = " 路線編號 = '" & txt_Tab3_route_Start.Text & "' "
ElseIf Len(txt_Tab3_route_Start.Text) = 0 And Len(txt_Tab3_route_End.Text) > 0 Then
   strTmp = " 路線編號 = '" & txt_Tab3_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'出車日期
strTmp = ""
If Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 between '" & txt_Tab3_DeliveryDate_Start.Text & "' and '" & txt_Tab3_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) = 0 Then
   strTmp = " 出車日期 = '" & txt_Tab3_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) = 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 = '" & txt_Tab3_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If



If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "基於縮小查詢資料量，請適度設定查詢條件"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by 運送區域,路線編號 "
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab3_PickLoadCheck)
tmp_Rs.Close

With dg_Tab3_PickLoadCheck
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab3_PickLoadCheck.MoveFirst
Set dg_Tab3_PickLoadCheck.DataSource = rs_Tab3_PickLoadCheck

With dg_Tab3_PickLoadCheck
    .ColumnHeaders = True         '標題行顯示
    .RowHeight = 250
    .Columns(0).Width = 500       '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 2500      '運送區域
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000      '路線編號
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 700      '訂單數
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 700       '送貨點
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1000       '客戶簡稱
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 700      '箱數
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 700      '板數
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700       '重量
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700      '材積
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 1000     '貨運公司
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1000     '車號
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 500     '車次
    .Columns(12).Alignment = dbgCenter
    .Columns(13).Width = 850     '回傳日期
    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 850     '出車日期
    .Columns(14).Alignment = dbgCenter
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-揀貨裝載稽核-查詢", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_SaveToExcel_Click()
'揀貨裝載稽核表 >> 轉 EXCEL
Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel 檔案名稱
CmnDialog.DialogTitle = "轉存 Excel 檔"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "揀貨裝載稽核表_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel檔案(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
   msg_text = "選擇 [取消] 按鈕，必須於 Excel 中自行存檔"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab3_PickLoadCheck) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab3_PickLoadCheck.MoveFirst
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-揀貨裝載稽核表-轉 EXCEL", Me.Caption, "cmd_Tab3_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_PrintReport_Click()
'轉運站路線匯總表 >> 列印資料 >> 報表列印
If rs_Tab4_OrderDetail Is Nothing Then Exit Sub
If rs_Tab4_OrderDetail.RecordCount = 0 Then Exit Sub
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass

'1. 資料寫出 Access 資料庫
Call AccessDB_Connect

' Wave 的訂單資料轉出
str_SQL = "Delete From 轉運站路線匯總表_OrderList"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access1)

Dim i As Double
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
rs_Access1.Open "轉運站路線匯總表_OrderList", cnAccess, adOpenStatic, adLockOptimistic
With dg_Tab4_RouteList


   For i = 1 To .Rows - 2
       .Row = i
       .Col = 1   '列印註記
       If Len(Trim(.Text)) > 0 Then
          rs_Access1.AddNew
          .Col = 3
          rs_Access1.Fields("二次排車路線編號").Value = txt_Tab4_SecondRouteNo.Text
          .Col = 2
          rs_Access1.Fields("ㄧ次排車路線編號").Value = Trim(.Text)
          .Col = 7
          rs_Access1.Fields("貨主單號").Value = Trim(.Text)
          .Col = 8
          rs_Access1.Fields("訂單日期").Value = Trim(.Text)
          .Col = 9
          rs_Access1.Fields("送貨日期").Value = Trim(.Text)
          .Col = 10
          rs_Access1.Fields("客戶編號").Value = Trim(.Text)
          .Col = 11
          rs_Access1.Fields("客戶名稱").Value = Trim(.Text)
          .Col = 5
          rs_Access1.Fields("車牌號碼").Value = Trim(.Text)
          .Col = 6
          rs_Access1.Fields("車次").Value = Trim(.Text)
          rs_Access1.Update
       End If
   Next i
End With
cnAccess.CommitTrans

'二次排車路線的路線匯總資料轉出
str_SQL = "Delete From 轉運站路線匯總表_RouteList"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access2)
cnAccess.BeginTrans
rs_Access2.Open "轉運站路線匯總表_RouteList", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab4_OrderDetail.MoveFirst
Do While Not rs_Tab4_OrderDetail.EOF
   rs_Access2.AddNew
   rs_Access2.Fields("序號").Value = rs_Tab4_OrderDetail.Fields("編號").Value
   rs_Access2.Fields("二次排車路編").Value = rs_Tab4_OrderDetail.Fields("二次路編").Value
   rs_Access2.Fields("貨號").Value = rs_Tab4_OrderDetail.Fields("貨號").Value
   rs_Access2.Fields("中文品名").Value = rs_Tab4_OrderDetail.Fields("品名").Value
   rs_Access2.Fields("指定到期日_標示").Value = rs_Tab4_OrderDetail.Fields("註記").Value
   rs_Access2.Fields("指定到期日_日期").Value = rs_Tab4_OrderDetail.Fields("註記內容").Value
   rs_Access2.Fields("訂單量_CaseQty").Value = rs_Tab4_OrderDetail.Fields("訂單箱數").Value
   rs_Access2.Fields("揀貨量_CaseQty").Value = rs_Tab4_OrderDetail.Fields("揀貨箱數").Value
   rs_Access2.Fields("板數").Value = rs_Tab4_OrderDetail.Fields("揀貨板數").Value
   rs_Access2.Fields("材積").Value = rs_Tab4_OrderDetail.Fields("揀貨材積").Value
   rs_Access2.Fields("重量").Value = rs_Tab4_OrderDetail.Fields("揀貨重量").Value
   rs_Access2.Fields("二次排車車號").Value = rs_Tab4_OrderDetail.Fields("二次排車車號").Value
   rs_Access2.Fields("二次排車車次").Value = rs_Tab4_OrderDetail.Fields("二次排車車次").Value
   rs_Access2.Update
   rs_Tab4_OrderDetail.MoveNext
Loop
rs_Tab4_OrderDetail.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab4_PreView.Value = vbChecked Then
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "轉運站路線匯總表_RouteList", acViewPreview
Else
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "轉運站路線匯總表_RouteList", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
chk_Tab4_PreView.Value = vbUnchecked
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cnAccess.RollbackTrans
      Tran_Level = 0
   End If
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--報表列印", Me.Caption, "cmd_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_QueryBySRouteNo_Click()
'轉運站路線匯總表 >> 資料篩選 >> 路編篩選

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass

'轉運站路線匯總表：二次排車路編與一次排車路編、訂單對應列表
Call SetGridFormat_Tab4_RouteList

str_SQL = "Select ㄧ次排車路編,二次排車路編,出車日期,車牌號碼,車次,貨主單號,訂單日期,送貨日期,客戶編號,客戶名稱 " & _
          "From Report_DCRouteSumSrc "
          

Dim strWhere As String, strTmp As String
strWhere = ""
'路線編號
If txt_Tab4_SecondRouteNo.Text <> "" Then
   strTmp = " (二次排車路編 = '" & txt_Tab4_SecondRouteNo.Text & "' or ㄧ次排車路編 = '" & txt_Tab4_SecondRouteNo.Text & "') "
   If Len(strTmp) > 0 Then
      If Len(strWhere) > 0 Then
         strWhere = strWhere & " and " & strTmp
      Else
         strWhere = strWhere & strTmp
      End If
   End If
End If
'出車日期
strTmp = ""
If Len(txt_Tab4_DeliveryDate_Start.Text) > 0 And Len(txt_Tab4_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 between '" & txt_Tab4_DeliveryDate_Start.Text & "' and '" & txt_Tab4_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab4_DeliveryDate_Start.Text) > 0 And Len(txt_Tab4_DeliveryDate_End.Text) = 0 Then
   strTmp = " 出車日期 = '" & txt_Tab4_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab4_DeliveryDate_Start.Text) = 0 And Len(txt_Tab4_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 = '" & txt_Tab4_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "請輸入適當的查詢條件"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " Order by ㄧ次排車路編,車牌號碼,貨主單號"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "訂單資料查詢結果：無相關的訂單資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   txt_Tab4_SecondRouteNo.SelStart = 0: txt_Tab4_SecondRouteNo.SelLength = Len(txt_Tab4_SecondRouteNo.Text): txt_Tab4_SecondRouteNo.SetFocus
   Exit Sub
End If

Dim iLoop As Double
iLoop = 0
dg_Tab4_RouteList.Visible = False
Do While Not tmp_Rs.EOF
   iLoop = iLoop + 1
   With dg_Tab4_RouteList
        If iLoop + 1 >= .Rows Then .Rows = .Rows + 1
            .Row = iLoop
            .Col = 0: .Text = iLoop
            If chk_Tab4_Selected.Value = vbChecked Then
               .Col = 1: .Text = "Ｖ"
            Else
               .Col = 1: .Text = " " '"Ｖ"
            End If
            .Col = 2: .Text = tmp_Rs.Fields("ㄧ次排車路編").Value
            .Col = 3: .Text = tmp_Rs.Fields("二次排車路編").Value
            .Col = 4: .Text = tmp_Rs.Fields("出車日期").Value
            .Col = 5: .Text = tmp_Rs.Fields("車牌號碼").Value
            .Col = 6: .Text = tmp_Rs.Fields("車次").Value
            .Col = 7: .Text = tmp_Rs.Fields("貨主單號").Value
            .Col = 8: .Text = tmp_Rs.Fields("訂單日期").Value
            .Col = 9: .Text = tmp_Rs.Fields("送貨日期").Value
            .Col = 10: .Text = tmp_Rs.Fields("客戶編號").Value
            .Col = 11: .Text = tmp_Rs.Fields("客戶名稱").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
dg_Tab4_RouteList.Visible = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "轉運站路線匯總表-資料篩選-路編篩選", Me.Caption, "cmd_Tab4_QueryBySRouteNo_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_Query_RouteDetail_Click()
'轉運站路線匯總表 >> 資料篩選 >> 路編篩選
Dim strOrderkey As String, i As Double
Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
strOrderkey = ""
With dg_Tab4_RouteList
     For i = 1 To .Rows - 2   'OrderList Grid 永遠會保留一列的空白
         .Row = i: .Col = 1   '是否選取：要轉出的
         If Len(Trim(.Text)) > 0 Then
            .Col = 7   '貨主單號 欄位
            If Len(strOrderkey) > 0 Then
               strOrderkey = strOrderkey & ",'" & RTrim(.Text) & "'"
            Else
               strOrderkey = "'" & RTrim(.Text) & "'"
            End If
         End If
     Next i
End With
If Len(strOrderkey) = 0 Then
   msg_text = "路線匯總資料篩選作業錯誤訊息：沒有選取要的訂單"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

SSTab2.Tab = 1
'str_SQL = "Select 二次路編,貨號,品名,註記,註記內容,sum(訂單箱數) as 訂單箱數,sum(揀貨箱數) as 揀貨箱數,sum(檢貨板數) as 揀貨板數," & _
'          "      sum(檢貨重量) as 揀貨重量,sum(檢貨材積) as 揀貨材積,二次排車車號,二次排車車次 " & _
'          "From  Report_DCRouteSum " & _
'          "Where 貨主單號 in (" & strOrderKey & ") and 路線編號 = '" & txt_Tab4_SecondRouteNo.Text & "' " & _
'          "Group by 二次路編,貨號,品名,註記,註記內容,二次排車車號,二次排車車次 Order by 二次路編,貨號,品名,註記 "
'daniel_2004100
str_SQL = "Select 二次路編,貨號,品名,註記,註記內容,sum(訂單箱數) as 訂單箱數,sum(揀貨箱數) as 揀貨箱數,sum(檢貨板數) as 揀貨板數," & _
          "      sum(檢貨重量) as 揀貨重量,sum(檢貨材積) as 揀貨材積,二次排車車號,二次排車車次,客戶編號 " & _
          "From  Report_DCRouteSum " & _
          "Where 貨主單號 in (" & strOrderkey & ")  " & _
          "Group by 二次路編,貨號,品名,註記,註記內容,二次排車車號,二次排車車次,客戶編號 Order by 二次路編,貨號,品名,註記 "
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Call Replication_Recordset(tmp_Rs, rs_Tab4_OrderDetail)
With dg_Tab4_OrderDetail
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 230                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab4_OrderDetail.MoveFirst
Set dg_Tab4_OrderDetail.DataSource = rs_Tab4_OrderDetail
With dg_Tab4_OrderDetail
    .Columns(0).Width = 500        '編號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1100       '二次路編
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900        '貨號
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '中文品名
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500        '訂單指定到期日註記＊
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 900        '訂單指定到期日--日期
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '訂單箱數
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800        '揀貨箱數
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 800        '揀貨板數
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 800        '揀貨重量
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800        '揀貨材積
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 1300        '二次排車車牌號碼
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1300        '二次排車車次
    .Columns(12).Alignment = dbgLeft
End With
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "轉運站路線匯總表-資料篩選-路線匯總", Me.Caption, "cmd_Tab4_Query_RouteDetail_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5_PrintReport_Click()
'排車一覽表 >> 報表列印
If rs_Tab5_PlanList Is Nothing Then Exit Sub
If rs_Tab5_PlanList.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. 資料寫出 Access 資料庫 >> 車輛裝載匯總表
Dim iLoop As Double
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 排車一覽表"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "排車一覽表", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab5_PlanList.MoveFirst

Do While Not rs_Tab5_PlanList.EOF
   rs_Access.AddNew
   For iLoop = 0 To rs_Tab5_PlanList.Fields.Count - 1
       rs_Access.Fields(iLoop).Value = rs_Tab5_PlanList.Fields(iLoop).Value
   Next iLoop
   rs_Access.Update
   rs_Tab5_PlanList.MoveNext
Loop

rs_Tab5_PlanList.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
'MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab5_PreView.Value = vbChecked Then
   '預覽列印
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "排車ㄧ覽表", acViewPreview
Else
   '直接列印至印表機
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "排車ㄧ覽表", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If

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
   CreateErrorLog Me.Name & "-排車一覽表-列印", Me.Caption, "cmd_Tab5_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab5_PrintReport1_Click()
'排車一覽表 >> 分貨表列印
Dim strTmp As String, strShortName As String, strRoute_No As String, strOrderkey As String

On Error GoTo err_Handle
str_SQL = "Select * From Report_TRPPlanList1 where 1 = 1 "
    
'訂單日期
If Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = "and 出車日期 between '" & txt_Tab5_DeliveryDate_Start.Text & "' and '" & txt_Tab5_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) = 0 Then
   strTmp = "and 出車日期 = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) = 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = "and 出車日期 = '" & txt_Tab5_DeliveryDate_End.Text & "' "
End If

str_SQL = str_SQL & strTmp

'路線編號
strTmp = ""
If Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = "and Rtrim(路線編號) between '" & txt_Tab5_route_Start.Text & "' and '" & txt_Tab5_route_End.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) = 0 Then
   strTmp = "and Rtrim(路線編號) = '" & txt_Tab5_route_Start.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) = 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = "and Rtrim(路線編號) = '" & txt_Tab5_route_End.Text & "' "
End If

str_SQL = str_SQL & strTmp

'運送區域
strTmp = ""
If cmb_Tab5_AreaCode.ListIndex <> -1 Then
   strTmp = "and 區碼 = '" & mySplit(cmb_Tab5_AreaCode, " ", 0) & "' "
End If

str_SQL = str_SQL & strTmp & " order by 出車日期,路線編號,品號 "

Screen.MousePointer = 11

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim rs_Tab5_PlanList1 As New ADODB.Recordset

Call Replication_Recordset(tmp_Rs, rs_Tab5_PlanList1)
tmp_Rs.Close

'1. 資料寫出 Access 資料庫 >> 車輛裝載匯總表
Dim iLoop As Double
Call AccessDB_Connect
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 分貨表"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "分貨表", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab5_PlanList1.MoveFirst

Do While Not rs_Tab5_PlanList1.EOF
   
   rs_Access.AddNew
   For iLoop = 0 To rs_Tab5_PlanList1.Fields.Count - 1
       rs_Access.Fields(iLoop).Value = rs_Tab5_PlanList1.Fields(iLoop).Value
   Next iLoop
   
   If strRoute_No <> rs_Tab5_PlanList1("路線編號") Then
   strRoute_No = rs_Tab5_PlanList1("路線編號")
   '取客戶名稱
    str_SQL = "select distinct isnull(t1m.Short_Name,'') as Short_Name , m2t.route_no " & _
                "from trp02t m2t join TRP01M t1m on t1m.ConsigneeKey = m2t.ConsigneeKey and t1m.storerkey = m2t.storerkey " & _
                "where m2t.route_no = '" & strRoute_No & "' order by t1m.Short_Name desc "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    strShortName = ""
    tmp_Rs.MoveFirst
    Do While Not tmp_Rs.EOF
        strShortName = strShortName & RTrim(tmp_Rs("Short_Name")) & ";"
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    
    '取訂單號碼
    str_SQL = "select distinct RTRIM(Extern) as Orderkey from trp02t where route_no = '" & strRoute_No & "' GROUP BY Extern order by RTRIM(Extern) "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    strOrderkey = ""
    tmp_Rs.MoveFirst
    
    Do While Not tmp_Rs.EOF
        strOrderkey = strOrderkey & RTrim(tmp_Rs("Orderkey")) & ";"
    tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    
    End If

    rs_Access.Fields("客戶簡稱") = strShortName
    rs_Access.Fields("訂單號碼") = strOrderkey
   
'   rs_Access.Fields("客戶簡稱").Value = .Fields("編號").Value
'   rs_Access.Fields("訂單號碼").Value = .Fields("編號").Value

   rs_Access.Update
   rs_Tab5_PlanList1.MoveNext
Loop

cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)
Set rs_Tab5_PlanList1 = Nothing

'2. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
'MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab5_PreView.Value = vbChecked Then
   '預覽列印
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "分貨表", acViewPreview
Else
   '直接列印至印表機
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "分貨表", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If

Exit Sub

err_Handle:
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
      Tran_Level = 0
   End If
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd_Tab5_Query_Click()
'排車一覽表 >> 查詢
Set dg_Tab5_PlanList.DataSource = Nothing
Set rs_Tab5_PlanList = Nothing

Screen.MousePointer = 11
On Error GoTo err_Handle
'
str_SQL = "Select 出車日期,區域,運送區域,暫存區,縣市別,貨運公司,車牌號碼,車次,一單多車,駕駛人," & _
          "       可載重量,可載材積,路線編號,運送點數,運送箱數,運送個數,運送板數,運送重量,運送材積,貨運公司代碼,備註,預計報到日期時間,貨主名稱,客戶簡稱,本倉,外倉,加工 " & _
          "From Report_TRPPlanList "
          
Dim strWhere As String, strTmp As String
strWhere = ""
'訂單日期
strTmp = ""
If Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 between '" & txt_Tab5_DeliveryDate_Start.Text & "' and '" & txt_Tab5_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) = 0 Then
   strTmp = " 出車日期 = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) = 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = " 出車日期 = '" & txt_Tab5_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'路線編號
strTmp = ""
If Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = " Rtrim(路線編號) between '" & txt_Tab5_route_Start.Text & "' and '" & txt_Tab5_route_End.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) = 0 Then
   strTmp = " Rtrim(路線編號) = '" & txt_Tab5_route_Start.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) = 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = " Rtrim(路線編號) = '" & txt_Tab5_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'運送區域
strTmp = ""
If cmb_Tab5_AreaCode.ListIndex <> -1 Then
   strTmp = " 區域 = '" & arAreaCode(cmb_Tab5_AreaCode.ListIndex) & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "基於縮小查詢資料量，請適度設定查詢條件"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

str_SQL = str_SQL & " order by 出車日期,路線編號 "
'str_SQL = str_SQL & " order by 出車日期,備註,路線編號,車牌號碼 "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab5_PlanList)
tmp_Rs.Close

With dg_Tab5_PlanList
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With

rs_Tab5_PlanList.MoveFirst
Set dg_Tab5_PlanList.DataSource = rs_Tab5_PlanList
dg_Tab5_PlanList.Visible = False

With dg_Tab5_PlanList
    .ColumnHeaders = True          '標題行顯示
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '出車日期
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 500        '區域
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1500       '運送區域
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 800        '暫存區
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1500       '縣市別
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1500       '貨運公司
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1000       '車牌號碼
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 500        '車次
    .Columns(8).Alignment = dbgCenter
    .Columns(9).Width = 800        '一單多車
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 1000       '駕駛人
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 800        '可載重量
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800       '可載材積
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 1100      '路線編號
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 800       '運送點數
    .Columns(14).Alignment = dbgCenter
    .Columns(15).Width = 800       '運送箱數
    .Columns(15).Alignment = dbgRight
    .Columns(16).Width = 800       '運送板數
    .Columns(16).Alignment = dbgRight
    .Columns(17).Width = 800       '運送重量
    .Columns(17).Alignment = dbgRight
    .Columns(18).Width = 800       '運送材積
    .Columns(18).Alignment = dbgRight
    .Columns(19).Width = 1200       '貨運公司代碼
    .Columns(19).Alignment = dbgLeft
    .Columns(20).Width = 1200       '備註(二次排車路線編號)
    .Columns(20).Alignment = dbgLeft
    .Columns(21).Width = 1200       '預計報到日期時間
    .Columns(21).Alignment = dbgCenter
    .Columns(22).Width = 3000       '客戶簡稱
    .Columns(22).Alignment = dbgLeft
End With
rs_Tab5_PlanList.MoveFirst

'取所有客戶名稱
Dim strShort_name As String
Call Confirm_Recordset_Closed(tmp_Rs)

Do While Not rs_Tab5_PlanList.EOF
    '取客戶名稱
    str_SQL = "select distinct isnull(t1m.Short_Name,'') as Short_Name , m2t.route_no " & _
              "from trp02t m2t join TRP01M t1m on t1m.ConsigneeKey = m2t.ConsigneeKey and t1m.storerkey = m2t.storerkey " & _
              "where m2t.route_no = '" & rs_Tab5_PlanList("路線編號") & "' order by t1m.Short_Name desc "

    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    strShort_name = ""
    tmp_Rs.MoveFirst
    
    Do While Not tmp_Rs.EOF

    strShort_name = strShort_name & RTrim(tmp_Rs("Short_Name")) & ";"

    tmp_Rs.MoveNext

    Loop
    tmp_Rs.Close

    rs_Tab5_PlanList("客戶簡稱") = strShort_name
        
    '取配置資料
    str_SQL = "select Facility = sum(case when sectionkey = 'FACILITY' then 1 else 0 end) , Wild = sum(case when sectionkey <> 'FACILITY' then 1 else 0 end) , Repacking = sum(len(rtrim(isnull(od.updatesource,'')))) " & _
                "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..orderdetail od on od.orderkey = o.orderkey " & _
                "join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = od.orderkey and od.orderlinenumber = p.orderlinenumber " & _
                "join " & strWMSDB & "..loc l (nolock) on l.loc = p.loc " & _
                "where o.route = '" & rs_Tab5_PlanList("路線編號") & "' "
    
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn
    
    If Not tmp_Rs.EOF Then
    
        If tmp_Rs("facility") > 0 Then rs_Tab5_PlanList("本倉") = "V"
        If tmp_Rs("wild") > 0 Then rs_Tab5_PlanList("外倉") = "V"
        If tmp_Rs("repacking") > 0 Then rs_Tab5_PlanList("加工") = "V"
    
    End If
    
    tmp_Rs.Close

rs_Tab5_PlanList.MoveNext
Loop
rs_Tab5_PlanList.MoveFirst
dg_Tab5_PlanList.Visible = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車一覽表-查詢", Me.Caption, "cmd_Tab5_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5_ReSet_Click()
'排車一覽表 >> 清除
cmb_Tab5_AreaCode.ListIndex = -1
txt_Tab5_DeliveryDate_Start.Text = ""
txt_Tab5_DeliveryDate_End.Text = ""
txt_Tab5_route_Start.Text = ""
txt_Tab5_route_End.Text = ""
Set dg_Tab5_PlanList.DataSource = Nothing
Set rs_Tab5_PlanList = Nothing

End Sub

Private Sub cmd_Tab5_SaveToExcel_Click()
'排車一覽表 >> 轉 EXCEL

    If rs_Tab5_PlanList Is Nothing Then Exit Sub
    rs_Tab5_PlanList.MoveFirst
    On Error GoTo err_Handle
    
    Screen.MousePointer = 11
    
    '將資料寫入excel檔
    Dim MyXlsApp As Excel.Application   '開啟excel檔
    Dim objFld As Field
    Dim i, j As Integer
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '新增Wookbooks
    MyXlsApp.Workbooks.Add
    '新增Sheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "排車一覽表"
    MyXlsApp.ActiveSheet.Name = "排車一覽表"
    i = 1
    'tr_SQL = "Select 出車日期,區域,運送區域,暫存區,縣市別,貨運公司,車牌號碼,車次,一單多車,駕駛人,可載重量,可載材積,路線編號,運送點數,運送箱數,運送板數,運送重量,運送材積,貨運公司代碼,備註,預計報到日期時間,客戶簡稱 "
    MyXlsApp.Cells(i, 1).Value = "編號"
    MyXlsApp.Cells(i, 2).Value = "出車日期"
    MyXlsApp.Cells(i, 3).Value = "區域"
    MyXlsApp.Cells(i, 4).Value = "暫存區"
    MyXlsApp.Cells(i, 5).Value = "縣市別"
    MyXlsApp.Cells(i, 6).Value = "車牌號碼"
    MyXlsApp.Cells(i, 7).Value = "車次"
    MyXlsApp.Cells(i, 8).Value = "駕駛人"
    MyXlsApp.Cells(i, 9).Value = "路線編號"
    MyXlsApp.Cells(i, 10).Value = "運送箱數"
    MyXlsApp.Cells(i, 11).Value = "運送個數"
    MyXlsApp.Cells(i, 12).Value = "運送重量"
    MyXlsApp.Cells(i, 13).Value = "運送材積"
    MyXlsApp.Cells(i, 14).Value = "備註"
    MyXlsApp.Cells(i, 15).Value = "時間"
    MyXlsApp.Cells(i, 16).Value = "貨主名稱"
    MyXlsApp.Cells(i, 17).Value = "客戶簡稱"
    MyXlsApp.Cells(i, 18).Value = "本倉"
    MyXlsApp.Cells(i, 19).Value = "外倉"
    MyXlsApp.Cells(i, 20).Value = "加工"
    MyXlsApp.Cells(i, 21).Value = "追蹤時間"
    MyXlsApp.Cells(i, 22).Value = "確認"
    MyXlsApp.Cells(i, 23).Value = "借出"
    MyXlsApp.Cells(i, 24).Value = "回收"
    MyXlsApp.Cells(i, 25).Value = "隔板"
    i = i + 1
    j = i
    rs_Tab5_PlanList.MoveFirst
    '日期,車號,單號,班別,借出,還入
    Do While Not rs_Tab5_PlanList.EOF
        If i > 2 Then
            If MyXlsApp.Cells(i - 1, 6).Value <> rs_Tab5_PlanList.Fields(7) Then
                '車號不同,隔一行在寫入excel
                MyXlsApp.Cells(i, 10).Value = "=SUM(J" & CStr(j) & ":J" & CStr(i - 1) & ")"  '運送箱數
                MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")"  '運送個數
                MyXlsApp.Cells(i, 13).Value = "=SUM(M" & CStr(j) & ":M" & CStr(i - 1) & ")" '運送重量
                MyXlsApp.Cells(i, 12).Value = "=SUM(L" & CStr(j) & ":L" & CStr(i - 1) & ")" '運送材積
                i = i + 2
                j = i
            End If
        End If
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab5_PlanList.Fields(1)) '出車日期
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rs_Tab5_PlanList.Fields(2) '區域
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rs_Tab5_PlanList.Fields(4) '暫存區
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rs_Tab5_PlanList.Fields(5) '縣市別
        MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 6).Value = rs_Tab5_PlanList.Fields(7) '車號
        MyXlsApp.Cells(i, 7).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 7).Value = rs_Tab5_PlanList.Fields(8) '車次
        'MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 8).Value = rs_Tab5_PlanList.Fields(10) '駕駛人
        MyXlsApp.Cells(i, 9).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 9).Value = rs_Tab5_PlanList.Fields(13)    '路線編號
        MyXlsApp.Cells(i, 10).Value = rs_Tab5_PlanList.Fields(15)    '運送箱數
        MyXlsApp.Cells(i, 11).Value = rs_Tab5_PlanList.Fields(16)    '運送個數
        MyXlsApp.Cells(i, 13).Value = rs_Tab5_PlanList.Fields(19)   '運送重量
        MyXlsApp.Cells(i, 12).Value = rs_Tab5_PlanList.Fields(18)   '運送材積
        MyXlsApp.Cells(i, 14).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 14).Value = rs_Tab5_PlanList.Fields(21)   '備註
        MyXlsApp.Cells(i, 15).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 15).Value = Mid(rs_Tab5_PlanList.Fields(22), 10, 4)   '時間
        MyXlsApp.Cells(i, 16).Value = rs_Tab5_PlanList.Fields(23)   '貨主名稱
        MyXlsApp.Cells(i, 17).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 17).Value = rs_Tab5_PlanList.Fields(24)   '客戶簡稱
        MyXlsApp.Cells(i, 18).Value = rs_Tab5_PlanList.Fields(25)   '本倉
        MyXlsApp.Cells(i, 19).Value = rs_Tab5_PlanList.Fields(26)   '外倉
        MyXlsApp.Cells(i, 20).Value = rs_Tab5_PlanList.Fields(27)   '加工
        rs_Tab5_PlanList.MoveNext
        i = i + 1
    Loop
    '計算箱數個數重量材積
    MyXlsApp.Cells(i, 10).Value = "=SUM(J" & CStr(j) & ":J" & CStr(i - 1) & ")"  '運送箱數
    MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")"  '運送個數
    MyXlsApp.Cells(i, 13).Value = "=SUM(M" & CStr(j) & ":M" & CStr(i - 1) & ")" '運送重量
    MyXlsApp.Cells(i, 12).Value = "=SUM(L" & CStr(j) & ":L" & CStr(i - 1) & ")" '運送材積
    i = i + 1
    '最適欄寬
    MyXlsApp.Columns("A:X").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '儲存格格式設定,重量和材積
    MyXlsApp.Columns("L:M").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A1:X" & i - 1).Select
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
    
    '車籍資料
    str_SQL = "select VEHICLE_ID_NO,DRIVER,DRIVER_PHONE from TRP09M"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.MoveFirst
        MyXlsApp.Sheets.Add
'        MyXlsApp.Sheets("Sheet2").Select
'        MyXlsApp.Sheets("Sheet2").Name = "車籍資料"
        MyXlsApp.ActiveSheet.Name = "車籍資料"
        i = 1
        MyXlsApp.Cells(i, 1).Value = "車號"
        MyXlsApp.Cells(i, 2).Value = "司機"
        MyXlsApp.Cells(i, 3).Value = "電話"
        i = i + 1
        Do While Not tmp_Rs.EOF
            MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
            MyXlsApp.Cells(i, 1).Value = Trim(tmp_Rs.Fields(0))
            MyXlsApp.Cells(i, 2).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 2).Value = Trim(tmp_Rs.Fields(1))
            MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 3).Value = Trim(tmp_Rs.Fields(2))
            tmp_Rs.MoveNext
            i = i + 1
        Loop
        '司機對應
        MyXlsApp.Sheets("排車一覽表").Select
        MyXlsApp.Range("H2").Select
        MyXlsApp.ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-2])=0,"""",VLOOKUP(RC[-2],車籍資料!C[-7]:C[-5],2,FALSE))"
    End If
    tmp_Rs.Close
    
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車一覽表-轉 EXCEL", Me.Caption, "cmd_Tab5_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5_SaveToExcel_NEW_Click()
'排車一覽表 >> 轉 EXCEL NEW

    If rs_Tab5_PlanList Is Nothing Then Exit Sub
    rs_Tab5_PlanList.MoveFirst
    On Error GoTo err_Handle
    Screen.MousePointer = 11
    
    Dim strWhere As String, strTmp As String, lngAR As Long, lngAP As Long, lngSorting As Long
    
    str_SQL = "select * from gv_TRPPlanLst where 1 = 1 and "
    
    strWhere = ""
    '訂單日期
    strTmp = ""
    If Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
       strTmp = " 出車日期 between '" & txt_Tab5_DeliveryDate_Start.Text & "' and '" & txt_Tab5_DeliveryDate_End.Text & "' "
    ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) = 0 Then
       strTmp = " 出車日期 = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
    ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) = 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
       strTmp = " 出車日期 = '" & txt_Tab5_DeliveryDate_End.Text & "' "
    End If
    
    If Len(strTmp) > 0 Then
       If Len(strWhere) > 0 Then
          strWhere = strWhere & " and " & strTmp
       Else
          strWhere = strWhere & strTmp
       End If
    End If
    
    '路線編號
    strTmp = ""
    If Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) > 0 Then
       strTmp = " Rtrim(路線編號) between '" & txt_Tab5_route_Start.Text & "' and '" & txt_Tab5_route_End.Text & "' "
    ElseIf Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) = 0 Then
       strTmp = " Rtrim(路線編號) = '" & txt_Tab5_route_Start.Text & "' "
    ElseIf Len(txt_Tab5_route_Start.Text) = 0 And Len(txt_Tab5_route_End.Text) > 0 Then
       strTmp = " Rtrim(路線編號) = '" & txt_Tab5_route_End.Text & "' "
    End If
    
    If Len(strTmp) > 0 Then
       If Len(strWhere) > 0 Then
          strWhere = strWhere & " and " & strTmp
       Else
          strWhere = strWhere & strTmp
       End If
    End If
        
    str_SQL = str_SQL & strWhere & " order by 出車日期,left(區域,1),備註,貨主"
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無符合搜尋條件之排車資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    
    Call Replication_Recordset(tmp_Rs, rs_Tab5_TRPPlanList)
    tmp_Rs.Close

    '將資料寫入excel檔
    Dim MyXlsApp As Excel.Application   '開啟excel檔
    Dim objFld As Field
    Dim i, j As Integer
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '新增Wookbooks
    MyXlsApp.Workbooks.Add
    '新增Sheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "排車一覽表"
    MyXlsApp.ActiveSheet.Name = "排車一覽表"
    i = 1
    'tr_SQL = "Select 出車日期,區域,運送區域,貨運公司,車牌號碼,車次,一單多車,駕駛人,可載重量,可載材積,路線編號,運送點數,運送箱數,運送板數,運送重量,運送材積,貨運公司代碼,備註,預計報到日期時間,客戶簡稱 "
    MyXlsApp.Cells(i, 1).Value = "編號"
    MyXlsApp.Cells(i, 2).Value = "出車日期"
    MyXlsApp.Cells(i, 3).Value = "區域"
    MyXlsApp.Cells(i, 4).Value = "車牌號碼"
    MyXlsApp.Cells(i, 5).Value = "車次"
    MyXlsApp.Cells(i, 6).Value = "駕駛人"
    MyXlsApp.Cells(i, 7).Value = "路線編號"
    MyXlsApp.Cells(i, 8).Value = "貨主"
    MyXlsApp.Cells(i, 9).Value = "應收"
    MyXlsApp.Cells(i, 10).Value = "應付"
    MyXlsApp.Cells(i, 11).Value = "翻板理貨"
    MyXlsApp.Cells(i, 12).Value = "運送件數"
    MyXlsApp.Cells(i, 13).Value = "運送箱數"
    MyXlsApp.Cells(i, 14).Value = "運送重量"
    MyXlsApp.Cells(i, 15).Value = "運送材積"
    MyXlsApp.Cells(i, 16).Value = "備註"
    MyXlsApp.Cells(i, 17).Value = "時間"
    MyXlsApp.Cells(i, 18).Value = "客戶簡稱"
    MyXlsApp.Cells(i, 19).Value = "追蹤時間"
    MyXlsApp.Cells(i, 20).Value = "確認"
    MyXlsApp.Cells(i, 21).Value = "借出"
    MyXlsApp.Cells(i, 22).Value = "回收"
    MyXlsApp.Cells(i, 23).Value = "隔板"
    i = i + 1
    j = i
    
    rs_Tab5_TRPPlanList.MoveFirst
    '日期,車號,單號,班別,借出,還入
    Do While Not rs_Tab5_TRPPlanList.EOF
        If i > 2 Then
            If RTrim(MyXlsApp.Cells(i - 1, 4).Value) & RTrim(MyXlsApp.Cells(i - 1, 15).Value) <> RTrim(rs_Tab5_TRPPlanList.Fields(3)) & RTrim(rs_Tab5_TRPPlanList.Fields(12)) Then
                '車號不同,隔一行再寫入excel
                MyXlsApp.Cells(i, 9).Value = "=SUM(i" & CStr(j) & ":i" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 10).Value = "=SUM(j" & CStr(j) & ":j" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 11).Value = "=SUM(k" & CStr(j) & ":k" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 12).Value = "=SUM(l" & CStr(j) & ":l" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 13).Value = "=SUM(m" & CStr(j) & ":m" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 14).Value = "=SUM(n" & CStr(j) & ":n" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 15).Value = "=SUM(o" & CStr(j) & ":o" & CStr(i - 1) & ")"
                i = i + 2
                j = i
            End If
        End If
        
        '取所有客戶名稱
        Dim strShort_name As String
        Call Confirm_Recordset_Closed(tmp_Rs)
               
        str_SQL = "select distinct isnull(t1m.Short_Name,'') as Short_Name , m2t.route_no " & _
                    "from trp02t m2t join TRP01M t1m on t1m.ConsigneeKey = m2t.ConsigneeKey and m2t.storerkey = t1m.storerkey " & _
                    "join trp16m t16m on t16m.storerkey = t1m.storerkey and t16m.short_name = '" & rs_Tab5_TRPPlanList("貨主") & "' " & _
                    "where m2t.route_no = '" & rs_Tab5_TRPPlanList("路線編號") & "' order by t1m.Short_Name desc "
    
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        strShort_name = ""
        tmp_Rs.MoveFirst
        
        Do While Not tmp_Rs.EOF
    
            strShort_name = strShort_name & RTrim(tmp_Rs("Short_Name")) & ";"
    
        tmp_Rs.MoveNext
    
        Loop
        tmp_Rs.Close
        
        '運費預先計算
        If rs_Tab5_TRPPlanList("出車日期") > Format(Now - 7, "YYYYMMDD") Then cn.Execute "exec gs_precost '" & IIf(Len(Trim(rs_Tab5_TRPPlanList("備註"))) = 0, rs_Tab5_TRPPlanList("路線編號"), rs_Tab5_TRPPlanList("備註")) & "','" & rs_Tab5_TRPPlanList("貨主") & "' ", RowsAffect, adExecuteNoRecords
        
        '取路線編號應收付資料
        lngAR = 0: lngAP = 0: lngSorting = 0
        
        Call Confirm_Recordset_Closed(tmp_Rs)
               
        str_SQL = "select ar=sum(t2.receivable),ap = sum(t2.payable) ,sorting = isnull((select sum((palletqty * 50 )+ ((case when storer = 'LTHL01' then 45 else 40 end) * sortingqty/1000)) from gt_LoadSorting where route_no = t2.route_no and storer = t2.storerkey ),0) from trp02t t2 join trp16m t16m on t16m.storerkey = t2.storerkey and t16m.short_name = '" & rs_Tab5_TRPPlanList("貨主") & "' where t2.route_no = '" & rs_Tab5_TRPPlanList("路線編號") & "' group by t2.route_no ,t2.storerkey"
    
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
'        lngAR = tmp_rs("ar")不顯示應收金額
        lngAP = tmp_Rs("ap")
        lngSorting = tmp_Rs("sorting")
        
        tmp_Rs.Close
        
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab5_TRPPlanList.Fields(1))
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rs_Tab5_TRPPlanList.Fields(2)
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rs_Tab5_TRPPlanList.Fields(3)
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rs_Tab5_TRPPlanList.Fields(4)
        'MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 6).Value = rs_Tab5_TRPPlanList.Fields(5)
        MyXlsApp.Cells(i, 7).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 7).Value = rs_Tab5_TRPPlanList.Fields(6)
        MyXlsApp.Cells(i, 8).Value = rs_Tab5_TRPPlanList.Fields(7)
        MyXlsApp.Cells(i, 9).Value = lngAR
        MyXlsApp.Cells(i, 10).Value = lngAP
        MyXlsApp.Cells(i, 11).Value = lngSorting
        MyXlsApp.Cells(i, 12).Value = rs_Tab5_TRPPlanList.Fields(8)
        MyXlsApp.Cells(i, 13).Value = rs_Tab5_TRPPlanList.Fields(9)
        MyXlsApp.Cells(i, 14).Value = rs_Tab5_TRPPlanList.Fields(10)
        MyXlsApp.Cells(i, 15).Value = rs_Tab5_TRPPlanList.Fields(11)
        MyXlsApp.Cells(i, 16).Value = rs_Tab5_TRPPlanList.Fields(12)
        MyXlsApp.Cells(i, 17).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 17).Value = Mid(rs_Tab5_TRPPlanList.Fields(13), 10, 4)
        MyXlsApp.Cells(i, 18).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 18).Value = strShort_name
        rs_Tab5_TRPPlanList.MoveNext
        i = i + 1
    Loop
    
    MyXlsApp.Cells(i, 9).Value = "=SUM(I" & CStr(j) & ":I" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 10).Value = "=SUM(j" & CStr(j) & ":j" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 11).Value = "=SUM(k" & CStr(j) & ":k" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 12).Value = "=SUM(l" & CStr(j) & ":l" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 13).Value = "=SUM(m" & CStr(j) & ":m" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 14).Value = "=SUM(n" & CStr(j) & ":n" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 15).Value = "=SUM(o" & CStr(j) & ":o" & CStr(i - 1) & ")"
    i = i + 1
    '最適欄寬
    MyXlsApp.Columns("A:W").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '儲存格格式設定
    MyXlsApp.Columns("m:n").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A1:W" & i - 1).Select
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
    
    '車籍資料
    str_SQL = "select VEHICLE_ID_NO,DRIVER,DRIVER_PHONE from TRP09M"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.MoveFirst
        MyXlsApp.Sheets.Add
'        MyXlsApp.Sheets("Sheet2").Select
'        MyXlsApp.Sheets("Sheet2").Name = "車籍資料"
        MyXlsApp.ActiveSheet.Name = "車籍資料"
        i = 1
        MyXlsApp.Cells(i, 1).Value = "車號"
        MyXlsApp.Cells(i, 2).Value = "司機"
        MyXlsApp.Cells(i, 3).Value = "電話"
        i = i + 1
        Do While Not tmp_Rs.EOF
            MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
            MyXlsApp.Cells(i, 1).Value = Trim(tmp_Rs.Fields(0))
            MyXlsApp.Cells(i, 2).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 2).Value = Trim(tmp_Rs.Fields(1))
            MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 3).Value = Trim(tmp_Rs.Fields(2))
            tmp_Rs.MoveNext
            i = i + 1
        Loop
        '司機對應
        MyXlsApp.Sheets("排車一覽表").Select
        MyXlsApp.Range("F2").Select
        MyXlsApp.ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-2])=0,"""",VLOOKUP(RC[-2],車籍資料!C[-5]:C[-3],2,FALSE))"
    End If
    tmp_Rs.Close
    
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-排車一覽表-轉 EXCEL", Me.Caption, "cmd_Tab5_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6_Query_Click()
'棧板維護by路編 >> 查詢
' add by Terry 20180518
Set dg_Tab6_PlanList.DataSource = Nothing
Set rs_Tab6_PlanList = Nothing


Screen.MousePointer = 11
On Error GoTo err_Handle

str_SQL = "select 出車日期 = Convert(VarChar, s01t.Delivery_Date, 112),貨運公司 = Rtrim(Isnull(t08m.C_Name,'')),貨運公司代碼 = t08m.COMPANY_CODE,客戶簡稱 = ISNULL(RTRIM(t01m.short_name),'') " & _
          ",車牌號碼 = Rtrim(Isnull(s01t.C_VEHICLE_ID_NO,'')),駕駛人 = Isnull(rtrim(s01t.Driver),''),路線編號 = Rtrim(Isnull(s02t.ROUTE_NO,'')),二次路編 = Rtrim(Isnull(s02t.C_ROUTE_NO,'')) " & _
          ",運送箱數 = sum(case when p.casecnt = 0 then 0 else floor((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end) /p.casecnt) end) " & _
          ",運送個數 = sum(case when p.casecnt = 0 then (case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end) else cast((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end) as int)%cast(p.casecnt as int) end) " & _
          ",運送板數 = sum(case when p.pallet = 0 then 0 else round((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end)/p.pallet,2) end) " & _
          ",運送重量 = sum(round((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end)*s.STDGROSSWGT,2)) " & _
          ",運送材積 = sum(round((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end)*s.stdcube,2)) " & _
          ",棧板維護 = s01t.PalletDefend " & _
          "From SDN01T s01t join SDN02T s02t on s01t.C_Route_No = s02t.C_ROUTE_NO " & _
          "join SDN03T s03t on s03t.receipt_no = s02t.receipt_no " & _
          "join " & strWMSDB & "..sku s on s03t.PRODUCT_NO = s.sku " & _
          "join " & strWMSDB & "..pack p on s.packkey = p.packkey " & _
          "join TRP01M t01m on s02t.CONSIGNEEKEY = t01m.CONSIGNEEKEY " & _
          "Left join TRP05T t05t on t05t.Route_No = s02t.Route_No " & _
          "Left join TRP08M t08m on t08m.Company_Code = t05t.TRP_Company_Code " & _
          "Where s01t.Delivery_Date > getdate() - 30 "


Dim tmpString1 As String, tmpString2 As String
Dim strWhere As String, strTmp As String
strWhere = ""



'訂單日期
strTmp = ""
If Len(txt_Tab6_DeliveryDate_Start.Text) > 0 And Len(txt_Tab6_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(VarChar, s01t.Delivery_Date, 112) between '" & txt_Tab6_DeliveryDate_Start.Text & "' and '" & txt_Tab6_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab6_DeliveryDate_Start.Text) > 0 And Len(txt_Tab6_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(VarChar, s01t.Delivery_Date, 112) = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab6_DeliveryDate_Start.Text) = 0 And Len(txt_Tab6_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(VarChar, s01t.Delivery_Date, 112) = '" & txt_Tab6_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'路線編號
strTmp = ""
If Len(txt_Tab6_route_Start.Text) > 0 And Len(txt_Tab6_route_End.Text) > 0 Then
   strTmp = " Rtrim(s02t.route_no) between '" & txt_Tab6_route_Start.Text & "' and '" & txt_Tab6_route_End.Text & "' "
ElseIf Len(txt_Tab6_route_Start.Text) > 0 And Len(txt_Tab6_route_End.Text) = 0 Then
   strTmp = " Rtrim(s02t.route_no) = '" & txt_Tab5_route_Start.Text & "' "
ElseIf Len(txt_Tab6_route_Start.Text) = 0 And Len(txt_Tab6_route_End.Text) > 0 Then
   strTmp = " Rtrim(s02t.route_no) = '" & txt_Tab6_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If


If strWhere <> "" Then
   str_SQL = str_SQL & " and " & strWhere
End If

str_SQL = str_SQL & " group by s01t.Delivery_Date,t08m.C_Name,s01t.C_VEHICLE_ID_NO,s01t.Driver,s02t.ROUTE_NO,t08m.COMPANY_CODE,s02t.C_ROUTE_NO,t01m.SHORT_NAME,s01t.PalletDefend order by 出車日期,路線編號 "
str_SQL_Excel = str_SQL
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab6_PlanList)
tmp_Rs.Close

With dg_Tab6_PlanList
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With

rs_Tab6_PlanList.MoveFirst
Set dg_Tab6_PlanList.DataSource = rs_Tab6_PlanList
dg_Tab6_PlanList.Visible = False

With dg_Tab6_PlanList
    .ColumnHeaders = True          '標題行顯示
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '出車日期
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1500       '貨運公司
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1200       '貨運公司代碼
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 3000       '客戶簡稱
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1000       '車牌號碼
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1000       '駕駛人
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1100      '路線編號
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1200       '備註(二次排車路線編號)
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 800       '運送箱數
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800       '運送個數
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 800       '運送板數
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800       '運送重量
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 800       '運送材積
    .Columns(13).Alignment = dbgRight
    .Columns(14).Width = 800       '棧板維護
    .Columns(14).Alignment = dbgCenter
End With
rs_Tab6_PlanList.MoveFirst

dg_Tab6_PlanList.Visible = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-棧板維護by路編-查詢", Me.Caption, "cmd_Tab6_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmd_Tab6_SaveToExcel_Click()
On Error GoTo err_Handle
If rs_Tab6_PlanList Is Nothing Then Exit Sub
If rs_Tab6_PlanList.RecordCount = 0 Then Exit Sub

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL_Excel, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then MsgBox "查無資料!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'轉Excel
Call Recordset2Excel("PalletDefend", rsTmp)

Set MyXlsApp = Nothing
rsTmp.Close: Set rsTmp = Nothing
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdLTHL01ShipDate_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = adUseClient
str_SQL = "select * from gv_LTHL01ShipData where 1 = 1 "

'出車日期
Dim strTmp As String
strTmp = ""
If Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = "and 出車日期 between '" & txt_Tab0_DeliveryDate_Start.Text & "' and '" & txt_Tab0_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) = 0 Then
   strTmp = "and 出車日期 = '" & txt_Tab0_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) = 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = "and 出車日期 = '" & txt_Tab0_DeliveryDate_End.Text & "' "
End If

str_SQL = str_SQL & strTmp

rsTmp.Open str_SQL, cn
If rsTmp.EOF Then MsgBox "", 64, cmdLTHL01ShipDate.Caption: Screen.MousePointer = 0: Exit Sub

'轉文字檔
'If Dir("C:\LTHL01\出貨回檔", vbDirectory) = "" Then MkDirs "C:\LTHL01\出貨回檔"
Open "C:\ShipDate.txt" For Output As #1

rsTmp.Sort = "訂單號碼"

rsTmp.MoveFirst
Do While Not rsTmp.EOF
    Print #1, rsTmp("訂單號碼"); rsTmp("路線編號"); rsTmp("TMS單號")
    rsTmp.MoveNext
Loop

'關閉檔案
Close

MsgBox "共轉出 " & rsTmp.RecordCount & "筆訂單，文字檔存放C:\ShipDate.txt", 64, "出貨資料轉出"
Screen.MousePointer = 0

Exit Sub
err_Handle:
Close
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub dg_Tab0_VLL_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dg_Tab0_VLL
If Len(dg.Columns(ColIndex).DataField) = 0 Then Exit Sub
SaveSetting App.title, "VLL裝載" & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dg_Tab0_VLL_HeadClick(ByVal ColIndex As Integer)
'VLL 裝載報表
'以滑鼠點選 dg_Tab0_VLL 欄位標題區
Dim OrderFieldName As String
If TypeName(rs_Tab0_VLL) <> "Nothing" Then
   OrderFieldName = "[" & dg_Tab0_VLL.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_Tab0_VLL.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_Tab0_VLL.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

Private Sub dg_Tab0_VLL_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If blVLLReportEventEnable Then
   
   With dg_Tab0_VLL
        '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
        If Trim(rs_Tab0_VLL.Fields(1).Value) = "" Then
           'If rs_Tab0_VLL("揀貨個數") = 0 Then MsgBox "揀貨量為0無法選取!", 64, "注意": Exit Sub
           rs_Tab0_VLL.Fields(1).Value = "V"
           dg_Tab0_VLL.SelBookmarks.Add rs_Tab0_VLL.Bookmark
           If rs_Tab0_VLL("排車個數") <> rs_Tab0_VLL("揀貨個數") Then MsgBox "排車個數不等於揀貨個數或揀貨量為0，請確認配置揀貨量是否不足！", 16, "注意"
        Else
           rs_Tab0_VLL.Fields(1).Value = " "
           If dg_Tab0_VLL.SelBookmarks.Count <> 0 Then dg_Tab0_VLL.SelBookmarks.Remove 0
           
        End If
   End With
End If
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "排車系統作業報表"
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

Dim tmp_cnt As Double
'取出所有運送區域代碼 TRP03M
cmb_Tab1_AreaCode.Clear: cmb_Tab3_AreaCode.Clear: cmb_Tab5_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_Tab1_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab3_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab5_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If
cmb_Tab1_AreaCode.ListIndex = -1
cmb_Tab3_AreaCode.ListIndex = -1
cmb_Tab5_AreaCode.ListIndex = -1
tmp_Rs.Close

cmd_Exit(0).Picture = BaseObject.cmdExit.Picture

SSTab1.Tab = 0

'轉運站路線匯總表：二次排車路編與一次排車路編、訂單對應列表
Call SetGridFormat_Tab4_RouteList
End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '變小
   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   fam_Tab0_Header.Left = fam_Tab0_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab0_VLL.Width = dg_Tab0_VLL.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab0_VLL.Height = dg_Tab0_VLL.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab1_Header.Left = fam_Tab1_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab1_VLLSum.Width = dg_Tab1_VLLSum.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab1_VLLSum.Height = dg_Tab1_VLLSum.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab2.Left = fam_Tab2.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab2_OrdersSum.Width = dg_Tab2_OrdersSum.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab2_OrdersSum.Height = dg_Tab2_OrdersSum.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab3.Left = fam_Tab3.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab3_PickLoadCheck.Width = dg_Tab3_PickLoadCheck.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab3_PickLoadCheck.Height = dg_Tab3_PickLoadCheck.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   SSTab2.Left = SSTab2.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   SSTab2.Top = SSTab2.Top - ((dbsrcFormHeight - Me.ScaleHeight) / 2)
   
   fam_Tab5_Header.Left = fam_Tab5_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab5_PlanList.Width = dg_Tab5_PlanList.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab5_PlanList.Height = dg_Tab5_PlanList.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab6_Header.Left = fam_Tab6_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab6_PlanList.Width = dg_Tab6_PlanList.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab6_PlanList.Height = dg_Tab6_PlanList.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   fam_Tab0_Header.Left = fam_Tab0_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab0_VLL.Width = dg_Tab0_VLL.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab0_VLL.Height = dg_Tab0_VLL.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab1_Header.Left = fam_Tab1_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab1_VLLSum.Width = dg_Tab1_VLLSum.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab1_VLLSum.Height = dg_Tab1_VLLSum.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab2.Left = fam_Tab2.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab2_OrdersSum.Width = dg_Tab2_OrdersSum.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab2_OrdersSum.Height = dg_Tab2_OrdersSum.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab3.Left = fam_Tab3.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab3_PickLoadCheck.Width = dg_Tab3_PickLoadCheck.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab3_PickLoadCheck.Height = dg_Tab3_PickLoadCheck.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   SSTab2.Left = SSTab2.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   SSTab2.Top = SSTab2.Top + ((Me.ScaleHeight - dbsrcFormHeight) / 2)
   
   fam_Tab5_Header.Left = fam_Tab5_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab5_PlanList.Width = dg_Tab5_PlanList.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab5_PlanList.Height = dg_Tab5_PlanList.Height + (Me.ScaleHeight - dbsrcFormHeight)

   fam_Tab6_Header.Left = fam_Tab6_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab6_PlanList.Width = dg_Tab6_PlanList.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab6_PlanList.Height = dg_Tab6_PlanList.Height - (dbsrcFormHeight - Me.ScaleHeight)

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
Set frm_Report_TRPPlan = Nothing

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'日期選取
Select Case mvDate.Tag
    Case "VLL裝載表.出車日期.起"
         txt_Tab0_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "VLL裝載表.出車日期.迄"
         txt_Tab0_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "車輛裝載匯總表.出車日期.起"
         txt_Tab1_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "車輛裝載匯總表.出車日期.迄"
         txt_Tab1_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "訂單總表.出車日期.起"
         txt_Tab2_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "訂單總表.出車日期.迄"
         txt_Tab2_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "揀貨裝載稽核表.回傳日期.起"
         txt_Tab3_UploadDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "揀貨裝載稽核表.回傳日期.迄"
         txt_Tab3_UploadDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "揀貨裝載稽核表.出車日期.起"
         txt_Tab3_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "揀貨裝載稽核表.出車日期.迄"
         txt_Tab3_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "轉運站路線匯總表.出車日期.起"
         txt_Tab4_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "轉運站路線匯總表.出車日期.迄"
         txt_Tab4_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "排車一覽表.出車日期.起"
         txt_Tab5_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "排車一覽表.出車日期.迄"
         txt_Tab5_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "棧板維護by路編.回傳日期.起"
         txt_Tab6_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "棧板維護by路編.回傳日期.迄"
         txt_Tab6_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case Else
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    mvDate.Visible = False
End Sub

Private Sub txt_Tab0_DeliveryDate_End_Click()
'VLL裝載 >> 出車日期 >> 迄
If Trim(txt_Tab0_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab0_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab0_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "VLL裝載表.出車日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab0_Header.Top + txt_Tab0_DeliveryDate_End.Top + txt_Tab0_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Header.Left + txt_Tab0_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab0_DeliveryDate_Start_Click()
'VLL裝載表 >> 出車日期 >> 起
If Trim(txt_Tab0_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab0_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab0_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "VLL裝載表.出車日期.起"
mvDate.Top = SSTab1.Top + fam_Tab0_Header.Top + txt_Tab0_DeliveryDate_Start.Top + txt_Tab0_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Header.Left + txt_Tab0_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab0_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'VLL裝載表 >> 出車日期 >> 起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab0_DeliveryDate_Start.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_Start.Text) = 1 Then
                msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_DeliveryDate_Start.SelStart = 0: txt_Tab0_DeliveryDate_Start.SelLength = Len(txt_Tab0_DeliveryDate_Start.Text): txt_Tab0_DeliveryDate_Start.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab0_DeliveryDate_End.SelStart = 0
          txt_Tab0_DeliveryDate_End.SelLength = Len(txt_Tab0_DeliveryDate_End.Text)
          txt_Tab0_DeliveryDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab0_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'VLL裝載表 >> 出車日期 >> 迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab0_DeliveryDate_End.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_End.Text) = 1 Then
                msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_DeliveryDate_End.SelStart = 0: txt_Tab0_DeliveryDate_End.SelLength = Len(txt_Tab0_DeliveryDate_End.Text): txt_Tab0_DeliveryDate_End.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab0_RouteNo_Start.SelStart = 0: txt_Tab0_RouteNo_Start.SelLength = Len(txt_Tab0_RouteNo_Start.Text)
          txt_Tab0_RouteNo_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab0_RouteNo_End_KeyPress(KeyAscii As Integer)
'VLL上貨表 >> 路線編號 >> 迄
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab0_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab0_RouteNo_Start_KeyPress(KeyAscii As Integer)
'VLL上貨表 >> 路線編號 >> 起
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab0_RouteNo_End.SelStart = 0: txt_Tab0_RouteNo_End.SelLength = Len(txt_Tab0_RouteNo_End.Text)
          txt_Tab0_RouteNo_End.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_DeliveryDate_End_Click()
'車輛裝載匯總表 >> 出車日期 >> 迄
If Trim(txt_Tab1_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "車輛裝載匯總表.出車日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_End.Top + txt_Tab1_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DeliveryDate_Start_Click()
'車輛裝載匯總表 >> 出車日期 >> 起
If Trim(txt_Tab1_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "車輛裝載匯總表.出車日期.起"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_Start.Top + txt_Tab1_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab2_DeliveryDate_End_Click()
'訂單總表 >> 出車日期 >> 迄
If Trim(txt_Tab2_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab2_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab2_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab2_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "訂單總表.出車日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab2.Top + txt_Tab2_DeliveryDate_End.Top + txt_Tab2_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab2.Left + txt_Tab2_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab2_DeliveryDate_Start_Click()
'訂單總表 >> 出車日期 >> 起
If Trim(txt_Tab2_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab2_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab2_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab2_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "訂單總表.出車日期.起"
mvDate.Top = SSTab1.Top + fam_Tab2.Top + txt_Tab2_DeliveryDate_Start.Top + txt_Tab2_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab2.Left + txt_Tab2_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab2_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'訂單總表 >> 送貨日期 >> 起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab2_DeliveryDate_Start.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_Start.Text) = 1 Then
                msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab2_DeliveryDate_Start.SelStart = 0: txt_Tab2_DeliveryDate_Start.SelLength = Len(txt_Tab2_DeliveryDate_Start.Text): txt_Tab2_DeliveryDate_Start.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab2_DeliveryDate_End.SelStart = 0
          txt_Tab2_DeliveryDate_End.SelLength = Len(txt_Tab2_DeliveryDate_End.Text)
          txt_Tab2_DeliveryDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab2_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'訂單總表 >> 出車日期 >> 迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab2_DeliveryDate_End.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_End.Text) = 1 Then
                msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab2_DeliveryDate_End.SelStart = 0: txt_Tab2_DeliveryDate_End.SelLength = Len(txt_Tab2_DeliveryDate_End.Text): txt_Tab2_DeliveryDate_End.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab2_RouteNo_Start.SelStart = 0: txt_Tab2_RouteNo_Start.SelLength = Len(txt_Tab2_RouteNo_Start.Text)
          txt_Tab2_RouteNo_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab2_RouteNo_End_KeyPress(KeyAscii As Integer)
'訂單總表 >> 路線編號 >> 迄
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab2_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab2_RouteNo_Start_KeyPress(KeyAscii As Integer)
'訂單總表 >> 路線編號 >> 起
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab2_RouteNo_End.SelStart = 0: txt_Tab2_RouteNo_End.SelLength = Len(txt_Tab2_RouteNo_End.Text)
          txt_Tab2_RouteNo_End.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'裝載匯總表 >> 出車日期 >> 起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab1_DeliveryDate_Start.SelStart = 0: txt_Tab1_DeliveryDate_Start.SelLength = Len(txt_Tab1_DeliveryDate_Start.Text): txt_Tab1_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab1_DeliveryDate_End.SelStart = 0
             txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text)
             txt_Tab1_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab1_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'裝載匯總表 >> 出車日期 >> 迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab1_DeliveryDate_End.SelStart = 0: txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text): txt_Tab1_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab1_Query.SetFocus
          End If
   End Select
End Sub

Private Sub txt_Tab3_UploadDate_End_Click()
'揀貨裝載稽核表 >> 回傳日期 >> 日期：迄
If Trim(txt_Tab3_UploadDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_UploadDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_UploadDate_End.Text, 4) & "/" & Mid(txt_Tab3_UploadDate_End.Text, 5, 2) & "/" & Right(txt_Tab3_UploadDate_End.Text, 2))
   End If
End If
mvDate.Tag = "揀貨裝載稽核表.回傳日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_UploadDate_End.Top + txt_Tab3_UploadDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_UploadDate_End.Left
mvDate.Visible = True
End Sub


Private Sub txt_Tab3_UploadDate_Start_Click()
'揀貨裝載稽核表 >> 回傳日期 >> 日期：起
If Trim(txt_Tab3_UploadDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_UploadDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_UploadDate_Start.Text, 4) & "/" & Mid(txt_Tab3_UploadDate_Start.Text, 5, 2) & "/" & Right(txt_Tab3_UploadDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "揀貨裝載稽核表.回傳日期.起"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_UploadDate_Start.Top + txt_Tab3_UploadDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_UploadDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab3_uploadDate_Start_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 回傳日期 >> 日期：起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_Start.SelStart = 0: txt_Tab3_UploadHour_Start.SelLength = Len(txt_Tab3_UploadHour_Start.Text)
          txt_Tab3_UploadHour_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab3_uploadhour_Start_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 回傳日期 >> 分鐘：起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_Start.Text = Format(Val(txt_Tab3_UploadHour_Start.Text), "00")
          txt_Tab3_UploadMinute_Start.SelStart = 0: txt_Tab3_UploadMinute_Start.SelLength = Len(txt_Tab3_UploadMinute_Start.Text)
          txt_Tab3_UploadMinute_Start.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_uploadminute_Start_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 回傳日期 >> 秒數：起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadMinute_Start.Text = Format(Val(txt_Tab3_UploadMinute_Start.Text), "00")
          txt_Tab3_UploadDate_End.SelStart = 0: txt_Tab3_UploadDate_End.SelLength = Len(txt_Tab3_UploadDate_End.Text)
          txt_Tab3_UploadDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_uploadDate_End_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 回傳日期 >> 日期：迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_End.SelStart = 0: txt_Tab3_UploadHour_End.SelLength = Len(txt_Tab3_UploadHour_End.Text)
          txt_Tab3_UploadHour_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_uploadhour_End_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 回傳日期 >> 分鐘：迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_End.Text = Format(Val(txt_Tab3_UploadHour_End.Text), "00")
          txt_Tab3_UploadMinute_End.SelStart = 0: txt_Tab3_UploadMinute_End.SelLength = Len(txt_Tab3_UploadMinute_End.Text)
          txt_Tab3_UploadMinute_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_Uploadminute_End_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 回傳日期 >> 秒數：迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadMinute_End.Text = Format(Val(txt_Tab3_UploadMinute_End.Text), "00")
          cmd_Tab3_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab3_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 出車日期 >> 起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_Start.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab3_DeliveryDate_Start.SelStart = 0: txt_Tab3_DeliveryDate_Start.SelLength = Len(txt_Tab3_DeliveryDate_Start.Text): txt_Tab3_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab3_DeliveryDate_End.SelStart = 0
             txt_Tab3_DeliveryDate_End.SelLength = Len(txt_Tab3_DeliveryDate_End.Text)
             txt_Tab3_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab3_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'揀貨裝載稽核表 >> 出車日期 >> 迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_End.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab3_DeliveryDate_End.SelStart = 0: txt_Tab3_DeliveryDate_End.SelLength = Len(txt_Tab3_DeliveryDate_End.Text): txt_Tab3_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab3_Query.SetFocus
          End If
   End Select
End Sub

Private Sub txt_Tab3_DeliveryDate_Start_Click()
'揀貨裝載稽核表 >> 出車日期 >> 起
If Trim(txt_Tab3_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "揀貨裝載稽核表.出車日期.起"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_DeliveryDate_Start.Top + txt_Tab3_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab3_DeliveryDate_End_Click()
'揀貨裝載稽核表 >> 出車日期 >> 迄
If Trim(txt_Tab3_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "揀貨裝載稽核表.出車日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_DeliveryDate_End.Top + txt_Tab3_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub cmd_Tab3_Reset_Click()
'揀貨裝載稽核表 >> 重設
Set dg_Tab3_PickLoadCheck.DataSource = Nothing
Set rs_Tab3_PickLoadCheck = Nothing
cmb_Tab3_AreaCode.ListIndex = -1
txt_Tab3_UploadDate_Start.Text = ""
txt_Tab3_UploadHour_Start.Text = ""
txt_Tab3_UploadMinute_Start.Text = ""
txt_Tab3_UploadDate_End.Text = ""
txt_Tab3_UploadHour_End.Text = ""
txt_Tab3_UploadMinute_End.Text = ""
End Sub


Private Sub txt_Tab4_SecondRouteNo_KeyPress(KeyAscii As Integer)
'轉運站路線匯總表 >> 資料篩選 >> 二次排車路線編號 >> 起
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab4_QueryBySRouteNo.SetFocus
   End Select
End Sub

Private Sub txt_Tab4_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'轉運站路線彙總表 >> 出車日期 >> 起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_Start.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab4_DeliveryDate_Start.SelStart = 0: txt_Tab4_DeliveryDate_Start.SelLength = Len(txt_Tab4_DeliveryDate_Start.Text): txt_Tab4_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab4_DeliveryDate_End.SelStart = 0
             txt_Tab4_DeliveryDate_End.SelLength = Len(txt_Tab4_DeliveryDate_End.Text)
             txt_Tab4_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab4_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'轉運站路線彙總表 >> 出車日期 >> 迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_End.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab4_DeliveryDate_End.SelStart = 0: txt_Tab4_DeliveryDate_End.SelLength = Len(txt_Tab4_DeliveryDate_End.Text): txt_Tab4_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab4_QueryBySRouteNo.SetFocus
          End If
   End Select
End Sub

Private Sub txt_Tab4_DeliveryDate_Start_Click()
'轉運站路線彙總表 >> 出車日期 >> 起
If Trim(txt_Tab4_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab4_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab4_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab4_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "轉運站路線匯總表.出車日期.起"
mvDate.Top = SSTab1.Top + SSTab2.Top + txt_Tab4_DeliveryDate_Start.Top + txt_Tab4_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + SSTab2.Left + txt_Tab4_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab4_DeliveryDate_End_Click()
'轉運站路線彙總表 >> 出車日期 >> 迄
If Trim(txt_Tab4_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab4_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab4_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab4_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "轉運站路線匯總表.出車日期.迄"
mvDate.Top = SSTab1.Top + SSTab2.Top + txt_Tab4_DeliveryDate_End.Top + txt_Tab4_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + SSTab2.Left + txt_Tab4_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub SetGridFormat_Tab4_RouteList()
'設定 轉運站路線匯總表 的 [ㄧ次排車路編資料] Grid 格式
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Tab4_RouteList.Visible = False
With dg_Tab4_RouteList
     .Rows = 2: .Cols = 12
     .FixedRows = 1
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 500    '序號
     .ColWidth(1) = 500    '列印與否
     .ColWidth(2) = 1300   'ㄧ次排車路編
     .ColWidth(3) = 1300   '二次排車路編
     .ColWidth(4) = 1000   '出車日期
     .ColWidth(5) = 900    '車牌號碼
     .ColWidth(6) = 500    '車次
     .ColWidth(7) = 1000   '貨主單號
     .ColWidth(8) = 1000   '訂單日期
     .ColWidth(9) = 1000   '送貨日期
     .ColWidth(10) = 1200   '客戶編號
     .ColWidth(11) = 2600   '客戶名稱
     
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "序號"
     .Col = 1: .Text = "列印"
     .Col = 2: .Text = "ㄧ次排車路編"
     .Col = 3: .Text = "二次排車路編"
     .Col = 4: .Text = "出車日期"
     .Col = 5: .Text = "車牌號碼"
     .Col = 6: .Text = "車次"
     .Col = 7: .Text = "貨主單號"
     .Col = 8: .Text = "訂單日期"
     .Col = 9: .Text = "送貨日期"
     .Col = 10: .Text = "客戶編號"
     .Col = 11: .Text = "客戶名稱"
     
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignCenterCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignCenterCenter
     .ColAlignment(7) = flexAlignCenterCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .ColAlignment(10) = flexAlignCenterCenter
     .ColAlignment(11) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Tab4_RouteList.Visible = True
End Sub

Private Sub DG_TAB4_ROUTELIST_Click()
'Wave 所屬訂單資料
'點一次：選取，點第二次：取消選取
Dim i As Double
With dg_Tab4_RouteList
     .Col = 5   'Exceed貨主單號
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 1
     If Len(Trim(.Text)) = 0 Then
        .Text = "Ｖ"
     Else
        .Text = ""
     End If
     .Col = 0
'     For i = 0 To .Cols - 1
'         .ColSel = i
'     Next i
End With
End Sub

Private Sub txt_Tab5_DeliveryDate_End_Click()
'排車一覽表 >> 出車日期 >> 迄
If Trim(txt_Tab5_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab5_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab5_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab5_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "排車一覽表.出車日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab5_Header.Top + txt_Tab1_DeliveryDate_End.Top + txt_Tab5_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab5_Header.Left + txt_Tab5_DeliveryDate_End.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab5_DeliveryDate_Start_Click()
'排車一覽表 >> 出車日期 >> 起
If Trim(txt_Tab5_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab5_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab5_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab5_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "排車一覽表.出車日期.起"
mvDate.Top = SSTab1.Top + fam_Tab5_Header.Top + txt_Tab1_DeliveryDate_Start.Top + txt_Tab5_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab5_Header.Left + txt_Tab5_DeliveryDate_Start.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab5_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'排車一覽表 >> 出車日期 >> 起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_Start.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab5_DeliveryDate_Start.SelStart = 0: txt_Tab5_DeliveryDate_Start.SelLength = Len(txt_Tab5_DeliveryDate_Start.Text): txt_Tab5_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab5_DeliveryDate_End.SelStart = 0
             txt_Tab5_DeliveryDate_End.SelLength = Len(txt_Tab5_DeliveryDate_End.Text)
             txt_Tab5_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab5_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'排車一覽表 >> 出車日期 >> 迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_End.Text) = 1 Then
             msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab5_DeliveryDate_End.SelStart = 0: txt_Tab5_DeliveryDate_End.SelLength = Len(txt_Tab5_DeliveryDate_End.Text): txt_Tab5_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab5_Query.SetFocus
          End If
   End Select
End Sub


Private Sub LLFA01Ship2TMS()
On Error GoTo err_Handle
'回傳揀貨量

str_SQL = "select o.route " & _
        ",o.storerkey " & _
        ",o.orderkey " & _
        ",o.updatesource " & _
        ",o.Externorderkey " & _
        ",ExternLineno = case when o.storerkey = 'LLFA01' then od.orderlinenumber else od.ExternLineno end " & _
        ",od.sku " & _
        ",shippedqty = (od.shippedqty + od.qtyallocated + od.qtypicked) " & _
        ",od.editdate " & _
        "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey  and convert(char(8),o.deliverydate,112) > convert(char(8),getdate()-7,112) " & _
        "where (od.shippedqty + od.qtyallocated + od.qtypicked) > 0 " & _
        "and len(rtrim(isnull(o.updatesource,''))) > 9 and o.updatesource in (select distinct receipt_no from trp03t where storerkey = 'LLFA01' and ship_qty = 0) "


Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'無資料
If Not tmp_Rs.EOF Then

    tmp_Rs.MoveFirst
    Tran_Level = cn.BeginTrans
    Do While Not tmp_Rs.EOF
    
            str_SQL = "UPDATE TRP03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03W set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '寫入紀錄
'            Call WriteLog(Err.Number & Chr(9) & "揀貨數量確認" & Chr(9) & "WMS: " & tmp_Rs("orderkey") & ",TMS: " & tmp_Rs("route") & "," & tmp_Rs("storerkey") & "," & tmp_Rs("updatesource") & "," & RTrim(tmp_Rs("Externorderkey")) & "," & tmp_Rs("Externlineno") & "," & tmp_Rs("sku") & "," & tmp_Rs("shippedqty") & "," & User_id)
            
'            '更新YFYstatus回傳狀態
'            str_SQL = "UPDATE " & strWMSDB & "..Orders set YFYstatus = '1' ,TrafficCop = null where orderkey = '" & tmp_Rs("orderkey") & "'"
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        tmp_Rs.MoveNext
    Loop
    
    '直接出貨量=訂單量
            str_SQL = "UPDATE TRP03T set TRP03T.SHIP_QTY=TRP03T.order_qty from trp02t join trp03t on trp02t.receipt_no = trp03t.receipt_no where trp02t.priority = 'C' and ship_qty = 0 "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
    cn.CommitTrans: Tran_Level = 0
End If

tmp_Rs.Close

Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL上貨表-查詢", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub
Private Sub txt_Tab6_DeliveryDate_End_Click()
'棧板維護by路編 >> 回傳日期 >> 日期：迄
If Trim(txt_Tab6_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab6_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab6_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab6_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab6_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "棧板維護by路編.回傳日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab6_Header.Top + txt_Tab6_DeliveryDate_End.Top + txt_Tab6_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + txt_Tab6_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab6_DeliveryDate_Start_Click()
'棧板維護by路編 >> 回傳日期 >> 日期：起
If Trim(txt_Tab6_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab6_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab6_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab6_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab6_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "棧板維護by路編.回傳日期.起"
mvDate.Top = SSTab1.Top + fam_Tab6_Header.Top + txt_Tab6_DeliveryDate_Start.Top + txt_Tab6_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + txt_Tab6_DeliveryDate_Start.Left
mvDate.Visible = True
End Sub
