VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_OP_DCRouteMerge 
   Caption         =   "二次排車作業"
   ClientHeight    =   7530
   ClientLeft      =   195
   ClientTop       =   855
   ClientWidth     =   13290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15630
   ScaleWidth      =   28560
   Begin TabDlg.SSTab SSTab1 
      Height          =   7440
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   13123
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "二次排車"
      TabPicture(0)   =   "frm_OP_DCRouteMerge.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fam_SrcRoute"
      Tab(0).Control(1)=   "fam_SelectedOrders"
      Tab(0).Control(2)=   "fam_RouteData"
      Tab(0).Control(3)=   "fra_ExtraQuery"
      Tab(0).Control(4)=   "mvDate"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "二次排車路編列表"
      TabPicture(1)   =   "frm_OP_DCRouteMerge.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fam_Tab1_Delete"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dg_Tab1_Route"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dg_Tab1_RouteDC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fam_Tab1_Query"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin MSComCtl2.MonthView mvDate 
         Height          =   2220
         Left            =   -73350
         TabIndex        =   70
         Top             =   4635
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
         StartOfWeek     =   114163713
         TitleBackColor  =   -2147483646
         TitleForeColor  =   16777215
         TrailingForeColor=   -2147483643
         CurrentDate     =   38232
         MaxDate         =   2958455
      End
      Begin VB.Frame fra_ExtraQuery 
         Appearance      =   0  '平面
         BackColor       =   &H00E0E0E0&
         Caption         =   "查詢條件設定"
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   3600
         Begin VB.TextBox txt_FPlanDate_Start 
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
            Left            =   1005
            TabIndex        =   78
            Top             =   225
            Width           =   1125
         End
         Begin VB.TextBox txt_FPlanDate_End 
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
            Left            =   2385
            TabIndex        =   77
            Top             =   225
            Width           =   1125
         End
         Begin VB.TextBox txt_FDeliveryDate_Start 
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
            Left            =   990
            TabIndex        =   76
            Top             =   540
            Width           =   1125
         End
         Begin VB.TextBox txt_FDeliveryDate_End 
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
            Left            =   2385
            TabIndex        =   75
            Top             =   540
            Width           =   1125
         End
         Begin VB.CheckBox chk_AddWho 
            Caption         =   "排車人員篩選"
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
            Left            =   990
            TabIndex        =   73
            Top             =   885
            Value           =   1  '核取
            Width           =   1875
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "排車日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   90
            TabIndex        =   82
            Top             =   270
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
            Index           =   13
            Left            =   2145
            TabIndex        =   81
            Top             =   255
            Width           =   240
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   80
            Top             =   585
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
            Index           =   15
            Left            =   2145
            TabIndex        =   79
            Top             =   570
            Width           =   240
         End
      End
      Begin VB.Frame fam_Tab1_Query 
         BackColor       =   &H00404000&
         Height          =   2160
         Left            =   9420
         TabIndex        =   58
         Top             =   285
         Width           =   1995
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
            Left            =   120
            TabIndex        =   60
            Top             =   630
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Tab1_RouteNoQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "二次排車查詢"
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
            Picture         =   "frm_OP_DCRouteMerge.frx":0038
            Style           =   1  '圖片外觀
            TabIndex        =   59
            Top             =   1230
            Width           =   1785
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
            ForeColor       =   &H0080FF80&
            Height          =   240
            Left            =   465
            TabIndex        =   61
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.Frame fam_RouteData 
         Height          =   540
         Left            =   -71295
         TabIndex        =   12
         Top             =   315
         Width           =   9915
         Begin VB.CommandButton cmd_ShowQuery 
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
            Left            =   1575
            Style           =   1  '圖片外觀
            TabIndex        =   74
            Top             =   135
            Width           =   375
         End
         Begin VB.CommandButton cmd_Tab0_CreateRoute 
            Appearance      =   0  '平面
            BackColor       =   &H00FF8080&
            Caption         =   "建立二次路編"
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
            Left            =   7260
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   69
            Top             =   75
            Width           =   1410
         End
         Begin VB.TextBox txt_Tab0_CarCheckInDate 
            Alignment       =   2  '置中對齊
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
            Left            =   4455
            TabIndex        =   66
            Top             =   105
            Width           =   1140
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
            Left            =   2490
            TabIndex        =   53
            Top             =   105
            Width           =   1215
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
            Height          =   435
            Index           =   0
            Left            =   8745
            Style           =   1  '圖片外觀
            TabIndex        =   15
            Top             =   90
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Tab0_ImportRoute 
            BackColor       =   &H00C0C0FF&
            Caption         =   "匯入一次路編"
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
            Left            =   105
            Style           =   1  '圖片外觀
            TabIndex        =   14
            Top             =   90
            Width           =   1440
         End
         Begin VB.TextBox txt_Tab0_CarCheckInTime 
            Alignment       =   2  '置中對齊
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
            Left            =   6375
            TabIndex        =   13
            Top             =   105
            Width           =   660
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   2
            Left            =   3630
            Top             =   75
            Width           =   1875
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
            ForeColor       =   &H00400000&
            Height          =   435
            Index           =   20
            Left            =   3795
            TabIndex        =   67
            Top             =   120
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
            ForeColor       =   &H00400000&
            Height          =   435
            Index           =   18
            Left            =   2040
            TabIndex        =   54
            Top             =   120
            Width           =   435
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   0
            Left            =   1995
            Top             =   75
            Width           =   1740
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
            ForeColor       =   &H00400000&
            Height          =   435
            Index           =   19
            Left            =   5700
            TabIndex        =   16
            Top             =   120
            Width           =   675
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   1
            Left            =   5640
            Top             =   75
            Width           =   1425
         End
      End
      Begin VB.Frame fam_SelectedOrders 
         Height          =   2790
         Left            =   -74895
         TabIndex        =   17
         Top             =   765
         Width           =   12300
         Begin VB.TextBox txt_Tab0_DeliveryCarTypeCode 
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
            Left            =   5955
            TabIndex        =   83
            Top             =   675
            Visible         =   0   'False
            Width           =   570
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
            Height          =   420
            Left            =   7875
            TabIndex        =   68
            Top             =   180
            Width           =   780
         End
         Begin VB.CommandButton cmd_Tab0_srcRouteReset 
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
            Height          =   435
            Left            =   9990
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   57
            Top             =   2190
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_srcRouteQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ㄧ次排車路編搜尋"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   9990
            Style           =   1  '圖片外觀
            TabIndex        =   56
            Top             =   1455
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_SelectedRemove_All 
            BackColor       =   &H000080FF&
            Caption         =   "已選路編移除(全)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   9975
            Style           =   1  '圖片外觀
            TabIndex        =   55
            Top             =   630
            Width           =   1140
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
            Left            =   4320
            TabIndex        =   24
            Top             =   150
            Width           =   1320
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
            Left            =   6165
            TabIndex        =   23
            Top             =   150
            Width           =   1230
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
            Left            =   7380
            Style           =   1  '圖片外觀
            TabIndex        =   22
            Top             =   150
            Width           =   330
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarType 
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
            Left            =   11970
            TabIndex        =   21
            Top             =   315
            Width           =   555
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
            Left            =   9600
            TabIndex        =   20
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab0_DeliveryCompany 
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
            Left            =   8760
            TabIndex        =   19
            Top             =   315
            Width           =   825
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
            Left            =   10785
            TabIndex        =   18
            Top             =   315
            Width           =   1170
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_SelectedRoute 
            Height          =   1695
            Left            =   45
            TabIndex        =   25
            Top             =   1035
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   2990
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
            Left            =   15
            TabIndex        =   44
            Top             =   525
            Width           =   5895
            Begin VB.TextBox txt_Tab0_Selected_Weight 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   4845
               TabIndex        =   48
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Volumn 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   3555
               TabIndex        =   47
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Pallet 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   2265
               TabIndex        =   46
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Case 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   1005
               TabIndex        =   45
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "累計：箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   7
               Left            =   75
               TabIndex        =   52
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "板數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   6
               Left            =   1890
               TabIndex        =   51
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
               Left            =   3165
               TabIndex        =   50
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
               Left            =   4470
               TabIndex        =   49
               Top             =   210
               Width           =   360
            End
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
            Left            =   5670
            TabIndex        =   30
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
            Left            =   3840
            TabIndex        =   31
            Top             =   150
            Width           =   435
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00400000&
            FillStyle       =   0  '實心
            Height          =   1350
            Left            =   9930
            Top             =   1380
            Width           =   1245
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
            Left            =   11940
            TabIndex        =   29
            Top             =   120
            Width           =   600
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
            Left            =   9885
            TabIndex        =   28
            Top             =   120
            Width           =   630
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
            Left            =   8760
            TabIndex        =   27
            Top             =   120
            Width           =   840
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
            Left            =   11145
            TabIndex        =   26
            Top             =   120
            Width           =   540
         End
      End
      Begin VB.Frame fam_SrcRoute 
         Height          =   3795
         Left            =   -74895
         TabIndex        =   1
         Top             =   3585
         Width           =   12300
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
            Left            =   5985
            Style           =   1  '圖片外觀
            TabIndex        =   43
            Top             =   135
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
            Left            =   6405
            Style           =   1  '圖片外觀
            TabIndex        =   42
            Top             =   135
            Width           =   345
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
            Left            =   7245
            Style           =   1  '圖片外觀
            TabIndex        =   41
            Top             =   135
            Width           =   1260
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel_All 
            BackColor       =   &H00FF80FF&
            Caption         =   "待選取消(全)"
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
            Left            =   8520
            Style           =   1  '圖片外觀
            TabIndex        =   40
            Top             =   135
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Case 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   38
            Top             =   795
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Pallet 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   36
            Top             =   1365
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Volumn 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   34
            Top             =   1950
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Weight 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   32
            Top             =   2505
            Width           =   1080
         End
         Begin MSDataGridLib.DataGrid dg_TRP01T 
            Height          =   2475
            Left            =   60
            TabIndex        =   11
            Top             =   525
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   4366
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
         Begin VB.Frame fam_SelectedSum 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   5895
            Begin VB.TextBox txt_Tab0_srcSelected_Weight 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   4830
               TabIndex        =   6
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Volumn 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   3555
               TabIndex        =   5
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Pallet 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   2265
               TabIndex        =   4
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Case 
               Alignment       =   1  '靠右對齊
               Appearance      =   0  '平面
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   1005
               TabIndex        =   3
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "選取：箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   10
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "板數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   1890
               TabIndex        =   9
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
               Left            =   3165
               TabIndex        =   8
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
               Left            =   4440
               TabIndex        =   7
               Top             =   195
               Width           =   360
            End
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_Orders 
            Height          =   690
            Left            =   60
            TabIndex        =   71
            Top             =   3045
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   1217
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "不含已出車確認"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   21
            Left            =   10320
            TabIndex        =   84
            Top             =   240
            Width           =   1260
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00400040&
            FillStyle       =   0  '實心
            Height          =   450
            Left            =   7200
            Top             =   90
            Width           =   2925
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '不透明
            Height          =   435
            Left            =   5940
            Top             =   105
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "箱    數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   11
            Left            =   10020
            TabIndex        =   39
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "板    數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   10
            Left            =   10020
            TabIndex        =   37
            Top             =   1170
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材    積"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   9
            Left            =   10020
            TabIndex        =   35
            Top             =   1755
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重    量"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   8
            Left            =   10020
            TabIndex        =   33
            Top             =   2310
            Width           =   540
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_RouteDC 
         Height          =   3480
         Left            =   105
         TabIndex        =   62
         Top             =   3510
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   6138
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
         Height          =   3105
         Left            =   105
         TabIndex        =   63
         Top             =   360
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   5477
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
      Begin VB.Frame fam_Tab1_Delete 
         Appearance      =   0  '平面
         BackColor       =   &H00000040&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   9450
         TabIndex        =   64
         Top             =   2385
         Width           =   1965
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
            Left            =   90
            Picture         =   "frm_OP_DCRouteMerge.frx":0342
            Style           =   1  '圖片外觀
            TabIndex        =   65
            ToolTipText     =   "刪除"
            Top             =   180
            Width           =   1785
         End
      End
   End
End
Attribute VB_Name = "frm_OP_DCRouteMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬

Private blTRP01TEventEnable As Boolean              '待選取ㄧ次排車路編 Event 觸發有效控制
Private blTab0SelectedRouteEventEnable As Boolean   '已選取ㄧ次排車路編之 Event 觸發有效控制
Private blTab1RouteEventEnable As Boolean           '二次排車路編之 Event 觸發有效控制

Private rs_TRP01T As ADODB.Recordset                '二次排車作業：匯入之ㄧ次排車{DC}路線編號
Private rs_Tab0_SelectedRoute As ADODB.Recordset    '已選取欲進行二次排車{併車}之ㄧ次排車{DC}路線編號
Private rs_Tab0_Orders As ADODB.Recordset           '路編對應之訂單
Private rs_Tab1_Route As ADODB.Recordset            'ㄧ次排車{DC}路編進行二次排車產生之路線編號
Private rs_Tab1_RouteDC As ADODB.Recordset          'ㄧ次排車路線編號

Private strSourceFilter As String        '待排車之ㄧ次排車路編篩選
Private strSourceOrderBy As String       '待排車之一次排車路編排序方式
Private dbsrcSelected_Case As Double     'ㄧ次排車之 [DC] 路編：選取箱數
Private dbsrcSelected_Pallet As Double   'ㄧ次排車之 [DC] 路編：選取板數
Private dbsrcSelected_Volumn As Double   'ㄧ次排車之 [DC] 路編：選取材積
Private dbsrcSelected_Weight As Double   'ㄧ次排車之 [DC] 路編：選取重量
Private dbSelectedCount As Double        '選取ㄧ次排車路編筆數

Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub

Private Sub cmd_ShowQuery_Click()
'二次排車 >> 匯入一次排車路編 >> 查詢條件
fra_ExtraQuery.Visible = Not fra_ExtraQuery.Visible
End Sub

Private Sub cmd_Tab0_CreateRoute_Click()
'二次排車  >> 建立二次排車路編
Dim Str_RouteNo As String
Dim str_FirstRouteNo As String
str_FirstRouteNo = ""
Str_RouteNo = ""
If rs_Tab0_SelectedRoute.RecordCount = 0 Then Exit Sub

'add by Terry 20190614 檢查路編是否已組成二次路編
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
    str_FirstRouteNo = str_FirstRouteNo & "'" & rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value & "',"
    rs_Tab0_SelectedRoute.MoveNext
Loop

str_FirstRouteNo = str_FirstRouteNo & "''"
rs_Tab0_SelectedRoute.MoveFirst
str_SQL = "select c_route_no from trp01t where route_no in (" & str_FirstRouteNo & ") and c_route_no is not null"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
    MsgBox ("有一次路編已組成二次路編，請重新載入一次路編並清空[已選取的一次路編]"), vbOKOnly + vbCritical
    tmp_Rs.Close
    Exit Sub
End If
tmp_Rs.Close


'add by Eric 20141211 檢查是否已經被出車確認
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
    Str_RouteNo = Str_RouteNo & "'" & rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value & "',"
    rs_Tab0_SelectedRoute.MoveNext
Loop
rs_Tab0_SelectedRoute.MoveFirst
str_SQL = "select route_no from trp05t t5 where t5.sdnstatus = 1 and t5.route_no in (" & Mid(Str_RouteNo, 1, Len(Str_RouteNo) - 1) & ")"

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Str_RouteNo = ""
If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    Do While Not tmp_Rs.EOF
        '有已出車的路線編號
        Str_RouteNo = Str_RouteNo & tmp_Rs.Fields("route_no") & " , "
        tmp_Rs.MoveNext
    Loop
    msg_text = "發現有路編已經出車確認，請確認一次路編是否已被出車。" & Chr(13) + Chr(10) & "並重新載入一次路編進行二次路編作業"
    MsgBox msg_text, vbOKOnly + vbCritical, msg_title

    msg_text = "已出車的路線編號:" & Chr(13) + Chr(10) & Str_RouteNo
    MsgBox msg_text, vbOKOnly + vbCritical, msg_title
    tmp_Rs.Close
    Exit Sub
Else
    tmp_Rs.Close
End If

If Len(Trim(txt_Tab0_TRPDate.Text)) = 0 Then
   msg_text = "資料錯誤：未輸入出車日期"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_TRPDate.SetFocus
   Exit Sub
End If
If Len(Trim(txt_Tab0_DeliveryCarNo.Text)) = 0 Then
   msg_text = "資料錯誤：未輸入車牌號碼"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_DeliveryCarNo.SetFocus
   Exit Sub
End If

'資料檢核

'a.出車日期：格式 yyyymmdd
txt_Tab0_TRPDate.Text = Trim(txt_Tab0_TRPDate.Text)
If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
   msg_text = "出車日期：" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
   Exit Sub
End If
'a2.出車日期 >= 今天
If txt_Tab0_TRPDate.Text < Format(Now, "yyyymmdd") Then
   msg_text = "出車日期不得小於今天"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
   Exit Sub
End If

'b.檢核 [車牌號碼] 是否有效
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

txt_Tab0_DeliveryCarNo.Text = Trim(txt_Tab0_DeliveryCarNo.Text)
str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "資料錯誤：車牌號碼 " & txt_Tab0_DeliveryCarNo.Text & " 未建檔"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
   txt_Tab0_DeliveryCarNo.SetFocus
   Exit Sub
End If
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
    txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
    txt_Tab0_DeliveryCarNo.SetFocus
    Exit Sub
End If
If intableCBM < Val(txt_Tab0_Selected_Volumn.Text) Then
    msg_text = "排車重量超過車輛可載材積,車輛可載材積:" & intableCBM
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
    txt_Tab0_DeliveryCarNo.SetFocus
    Exit Sub
End If

'指定碼頭暫存：必須輸入
txt_Tab0_DockNo.Text = Trim(txt_Tab0_DockNo.Text)
If Len(Trim(txt_Tab0_DockNo.Text)) = 0 Then
   msg_text = "資料錯誤：[碼頭暫存] 必須輸入"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_DockNo.SetFocus
End If

'預計報到日期
txt_Tab0_CarCheckInDate.Text = Trim(txt_Tab0_CarCheckInDate.Text)
If Len(Trim(txt_Tab0_CarCheckInDate.Text)) <> 8 Then
   msg_text = "預計報到日期：資料格式 yyyymmdd "
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
   txt_Tab0_CarCheckInDate.SetFocus
   Exit Sub
End If

If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
   msg_text = "預計報到日期：資料錯誤 yyyymmdd，" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
   txt_Tab0_CarCheckInDate.SetFocus
   Exit Sub
End If

'預計報到日期 >= 今天
If txt_Tab0_CarCheckInDate.Text < Format(Now, "yyyymmdd") Then
   msg_text = "預計報到日期不得小於今天"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text): txt_Tab0_CarCheckInDate.SetFocus
   Exit Sub
End If

'預計報到時間
txt_Tab0_CarCheckInTime.Text = Trim(txt_Tab0_CarCheckInTime.Text)
If Len(Trim(txt_Tab0_CarCheckInTime.Text)) <> 4 Then
   msg_text = "預計報到時間：資料格式 hhss "
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
   txt_Tab0_CarCheckInTime.SetFocus
   Exit Sub
End If
Select Case Left(txt_Tab0_CarCheckInTime.Text, 2)
       Case "00" To "24"
       Case Else
            msg_text = "預計報到時間：資料格式 hhss "
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
            txt_Tab0_CarCheckInTime.SetFocus
            Exit Sub
End Select
Select Case Right(txt_Tab0_CarCheckInTime.Text, 2)
       Case "00" To "59"
       Case Else
            msg_text = "預計報到時間：資料格式 hhss "
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
            txt_Tab0_CarCheckInTime.SetFocus
            Exit Sub
End Select

On Error GoTo err_Handle
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
          "From TRP01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'S'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
strRouteNo = "S" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
tmp_Rs.Close

'3.Insert into TRP01T 路線編號主檔
'  TRP01T.EXE_CONFIRM = '0' 新建立路線編號，尚未回傳過 exe
str_SQL = "Insert into TRP01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
          strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
          txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'3-1.更新總計欄位值
txt_Tab0_srcTotal_Case.Text = Val(txt_Tab0_srcTotal_Case.Text) - Val(txt_Tab0_Selected_Case.Text)
txt_Tab0_srcTotal_Pallet.Text = Val(txt_Tab0_srcTotal_Pallet.Text) - Val(txt_Tab0_Selected_Pallet.Text)
txt_Tab0_srcTotal_Volumn.Text = Val(txt_Tab0_srcTotal_Volumn.Text) - Val(txt_Tab0_Selected_Volumn.Text)
txt_Tab0_srcTotal_Weight.Text = Val(txt_Tab0_srcTotal_Weight.Text) - Val(txt_Tab0_Selected_Weight.Text)

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
rs_Tab1_Route.Fields("二次排車路編").Value = strRouteNo
rs_Tab1_Route.Fields("出車日期").Value = txt_Tab0_TRPDate.Text
rs_Tab1_Route.Fields("車牌號碼").Value = txt_Tab0_DeliveryCarNo.Text
rs_Tab1_Route.Fields("車次").Value = intDriveTimes
rs_Tab1_Route.Fields("駕駛人").Value = txt_Tab0_DeliveryDriver.Text
rs_Tab1_Route.Fields("箱數").Value = txt_Tab0_Selected_Case.Text
rs_Tab1_Route.Fields("板數").Value = txt_Tab0_Selected_Pallet.Text
rs_Tab1_Route.Fields("材積").Value = txt_Tab0_Selected_Volumn.Text
rs_Tab1_Route.Fields("重量").Value = txt_Tab0_Selected_Weight.Text
rs_Tab1_Route.Fields("碼頭暫存").Value = txt_Tab0_DockNo.Text
rs_Tab1_Route.Fields("預計報到日期").Value = txt_Tab0_CarCheckInDate.Text
rs_Tab1_Route.Fields("預計報到時間").Value = txt_Tab0_CarCheckInTime.Text
rs_Tab1_Route.Fields("車種").Value = txt_Tab0_DeliveryCarTypeCode.Text
rs_Tab1_Route.Fields("排車者").Value = User_id
rs_Tab1_Route.Update
blTab1RouteEventEnable = True

'5.update TRP01T & TRP05T [ㄧ次排車路編 & ㄧ次排車車輛管制]
'  寫至 SSTab1.Tab 1 [二次排車路線編號所屬之ㄧ次排車路線編號]
blTab0SelectedRouteEventEnable = False
rs_Tab1_RouteDC.Filter = adFilterNone
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
   'UPDATE TRP01T
   str_SQL = "Update TRP01T Set C_ROUTE_NO = '" & strRouteNo & "',C_VEHICLE_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "',C_DRIVE_TIMES = " & intDriveTimes & " " & _
             "Where Route_No = '" & rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Update TRP05T Set C_ROUTE_NO = '" & strRouteNo & "',C_VEHICLE_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "',C_DRIVE_TIMES = " & intDriveTimes & " " & _
             "Where Route_No = '" & rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value & "' "
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '寫至 SSTab1.Tab 1 [二次排車路編所屬之ㄧ次排車路線編號列表]
   rs_Tab1_RouteDC.AddNew
   rs_Tab1_RouteDC.Fields("編號").Value = rs_Tab1_RouteDC.RecordCount
   rs_Tab1_RouteDC.Fields("二次排車路編").Value = strRouteNo
   rs_Tab1_RouteDC.Fields("ㄧ次排車路編").Value = rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value
   rs_Tab1_RouteDC.Fields("出車日期").Value = rs_Tab0_SelectedRoute.Fields("出車日期").Value
   rs_Tab1_RouteDC.Fields("車牌號碼").Value = rs_Tab0_SelectedRoute.Fields("車牌號碼").Value
   rs_Tab1_RouteDC.Fields("車次").Value = rs_Tab0_SelectedRoute.Fields("車次").Value
   rs_Tab1_RouteDC.Fields("駕駛人").Value = rs_Tab0_SelectedRoute.Fields("駕駛人").Value
   rs_Tab1_RouteDC.Fields("箱數").Value = rs_Tab0_SelectedRoute.Fields("箱數").Value
   rs_Tab1_RouteDC.Fields("板數").Value = rs_Tab0_SelectedRoute.Fields("板數").Value
   rs_Tab1_RouteDC.Fields("材積").Value = rs_Tab0_SelectedRoute.Fields("材積").Value
   rs_Tab1_RouteDC.Fields("重量").Value = rs_Tab0_SelectedRoute.Fields("重量").Value
   rs_Tab1_RouteDC.Fields("車種").Value = rs_Tab0_SelectedRoute.Fields("車種").Value
   rs_Tab1_RouteDC.Update
   rs_Tab0_SelectedRoute.MoveNext
Loop


'到貨追蹤APP刪除與寫入
'cn.Execute "update apporderdate set status = 'C',editdate = getdate() where c_route_no = '" & strRouteNo & "' ", RowsAffect, adExecuteNoRecords
cn.Execute "delete apporderdate where receipt_no in (select t2.receipt_no from trp02t t2 join trp01t t1 on t1.route_no = t2.route_no and t1.c_route_no = '" & strRouteNo & "') ", RowsAffect, adExecuteNoRecords

str_SQL = "insert into apporderdate(wh,C_Route_no,C_VEHICLE_ID_NO,Priority,Receipt_no,OrderGroup,Storerkey,Arrive_date,Company,Status) " & _
    "select WH = 'GYDC' ,C_Route_no = t1.C_route_no " & _
    ",C_VEHICLE_ID_NO = isnull(t1.C_VEHICLE_ID_NO,t2.VEHICLE_ID_NO) " & _
    ",Priority = t2.Priority " & _
    ",Receipt_no = t2.receipt_no " & _
    ",OrderGroup = t1m.address " & _
    ",Storerkey = rtrim(t16m.storerkey) + '_' + t16m.short_name " & _
    ",Arrive_date = convert(char(8),t2.arrive_date,112) " & _
    ",Company = t1m.short_name " & _
    ",Status = '0' " & _
    "from trp02t t2 join trp01t t1 on t1.route_no = t2.route_no and c_route_no is not null and t1.C_ROUTE_NO = '" & strRouteNo & "' " & _
    "join trp01m t1m on t1m.storerkey = t2.storerkey and t2.consigneekey = t1m.consigneekey " & _
    "join trp16m t16m on t16m.storerkey = t2.storerkey "
    
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

If dg_Tab1_Route.SelBookmarks.Count > 0 Then
   dg_Tab1_Route.SelBookmarks.Remove 0
End If
dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
rs_Tab1_RouteDC.Filter = " 二次排車路編 = '" & rs_Tab1_Route.Fields("二次排車路編").Value & "'"

blTab0SelectedRouteEventEnable = True

'5.清除 [已選取之ㄧ次排車路線編號列表]
blTab0SelectedRouteEventEnable = False
'排車作業：已選取之ㄧ次排車路線編號列表 DBGrid 格式設定-ReSet
Call CreateRS_Tab0_SelectedRoute
'重新計算已選取ㄧ次排車路線編號：箱數，板數，材積，重量 + 編號重新產生
Call Calculate_SelectedRoute
blTab0SelectedRouteEventEnable = True

'6.清除排車作業欄位值
txt_Tab0_DockNo.Text = ""               '碼頭暫存
txt_Tab0_CarCheckInDate.Text = ""       '車輛預計報到時間
txt_Tab0_CarCheckInTime.Text = ""       '車輛預計報到時間
txt_Tab0_TRPDate.Text = ""              '出車日期
txt_Tab0_DeliveryCarNo.Text = ""        '車牌號碼
txt_Tab0_DeliveryCompany.Text = ""      '運輸公司
txt_Tab0_DeliveryDriver.Text = ""       '駕駛人
txt_Tab0_DeliveryPhone.Text = ""        '電話
txt_Tab0_DeliveryCarType.Text = ""      '車種
txt_Tab0_DeliveryCarTypeCode.Text = ""  '車種代碼

SSTab1.Tab = 1
DoEvents: DoEvents

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
   rs_Tab1_Route.Filter = "二次排車路編='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   rs_Tab1_RouteDC.Filter = "二次排車路編='" & strRouteNo & "'"
   If Not rs_Tab1_RouteDC.EOF Then
      Do While Not rs_Tab1_RouteDC.EOF
         rs_Tab1_RouteDC.Delete
         rs_Tab1_RouteDC.MoveFirst
      Loop
   End If
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
      
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-二次排車-建立二次排車路編", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_ImportRoute_Click()
'排車作業 >> 匯入ㄧ次排車路線編號
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_TRP01T.DataSource = Nothing
Set rs_TRP01T = Nothing
fra_ExtraQuery.Visible = False
strSourceFilter = adFilterNone
DoEvents

'有已選取ㄧ次排車路編者：詢問 user 是否要清除
If rs_Tab0_SelectedRoute.RecordCount <> 0 Then
   msg_text = "[已選取之ㄧ次排車路線編號] 是否進行清除"
   If MsgBox(msg_text, vbYesNo + vbInformation + vbDefaultButton2, msg_title) = vbYes Then
      '併車作業：已選取之ㄧ次排車路線編號列表 DBGrid 格式設定
      Call CreateRS_Tab0_SelectedRoute
      '清除欄位：累計選取之ㄧ次排車路編：小計歸 0
      txt_Tab0_Selected_Case.Text = ""
      txt_Tab0_Selected_Pallet.Text = ""
      txt_Tab0_Selected_Volumn.Text = ""
      txt_Tab0_Selected_Weight.Text = ""
   End If
End If

'ㄧ次排車路編：選取小計：歸零
dbSelectedCount = 0
dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""

Dim str_SQL2 As String
'取回ㄧ次排車路線編號
str_SQL = "Select " & _
        "' ' as '＊' " & _
        ",T1.ROUTE_NO as ㄧ次排車路編 " & _
        ",Isnull(Rtrim(T5.Dock_No),'') as 碼頭 " & _
        ",Round(T1.Case_cnt,2) as 箱數 " & _
        ",Round(T1.Pallet_Qty,2) as 板數 " & _
        ",Round(T1.Volumn_Weight,2) as 材積 " & _
        ",Round(T1.Weight,2) as 重量 " & _
        ",Rtrim(T5.VEHICLE_ID_NO) as 車牌號碼 " & _
        ",T5.Drive_Times as 車次 " & _
        ",Rtrim(Isnull(T5.Driver,'')) as 駕駛人 " & _
        ",Convert(varchar , T1.Delivery_Date,112) as 出車日期 " & _
        ",Rtrim(Isnull(a1.Vehicle_Type,'')) as 車種 " & _
        ",Case T1.EXE_Confirm When '0' Then '新建路編' When '1' Then '設定回傳' When '2' Then '已回傳' When '9' Then '預先揀貨' else '未知狀態' End  AS EXE回傳 " & _
        ",cast(' ' as char(300)) as 客戶 " & _
        ",Rtrim(Isnull(T1.AddWho,'')) as 排車者 " & _
        "From TRP01T T1 " & _
        "inner join TRP05T T5  on T1.ROUTE_NO=T5.ROUTE_NO and sdnstatus = 0 " & _
        "inner join TRP09M A1 on A1.Vehicle_ID_No = T5.Vehicle_ID_No " & _
        "Where Left(T1.ROUTE_NO,1) <> 'S'  and T1.ROUTE_NO <> 'D' and rtrim(isnull(T1.C_ROUTE_NO,''))='' and T5.Valid_Vehicle = '1'  and 1=1 "
        
Dim str_Where As String, intloop As Integer
str_Where = ""

'排車日期
If Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   str_Where = "and Convert(varchar,T1.AddDate,112) Between '" & txt_FPlanDate_Start.Text & "' and '" & txt_FPlanDate_End.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) = 0 Then
   str_Where = "and Convert(varchar,T1.AddDate,112) = '" & txt_FPlanDate_Start.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) = 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   str_Where = "and  Convert(varchar,T1.AddDate,112) = '" & txt_FPlanDate_End.Text & "' "
End If

'出車日期
If Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   str_Where = str_Where & "and Convert(varchar , T1.Delivery_Date,112) Between '" & txt_FDeliveryDate_Start.Text & "' and '" & txt_FDeliveryDate_End.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) = 0 Then
   str_Where = "and Convert(varchar , T1.Delivery_Date,112) = '" & txt_FDeliveryDate_Start.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) = 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   str_Where = "and Convert(varchar , T1.Delivery_Date,112) = '" & txt_FDeliveryDate_End.Text & "' "
End If

'排車人員篩選
If chk_AddWho.Value = vbChecked Then str_Where = str_Where & "and Rtrim(Isnull(T1.AddWho,'')) = '" & User_id & "' "

'str_Where = str_Where & "and EXE回傳 <> '已回傳' "

str_SQL = str_SQL & str_Where & "Order by T1.ROUTE_NO "
'str_SQL2 = str_SQL2 & str_Where & " "
          
strSourceOrderBy = " ㄧ次排車路編 "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之ㄧ次排車路線編號資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_TRP01T)
tmp_Rs.Close

With dg_TRP01T
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_TRP01T.MoveFirst
blTRP01TEventEnable = False
Set dg_TRP01T.DataSource = rs_TRP01T
With dg_TRP01T
    .RowHeight = 250
    .Columns(0).Width = 500         '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 350         '選取識別欄位
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1200        'ㄧ次排車路線編號
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 500        'ㄧ次排車路線暫存碼頭
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 800         '箱數
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 800         '板數
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 800         '材積
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800         '重量
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 900         '車牌號碼
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 500         '車次
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 700         '駕駛人
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1000       '出車日期
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 500       '車種
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 900       'EXE回傳
    .Columns(13).Alignment = dbgLeft
End With

'顯示客戶名稱
rs_TRP01T.MoveFirst
Do While Not rs_TRP01T.EOF
    
    str_SQL = "select distinct t1m.short_name as short_name from trp02t t2 join trp01m t1m on t2.consigneekey = t1m.consigneekey and t2.storerkey = t1m.storerkey and t2.route_no = '" & rs_TRP01T("ㄧ次排車路編") & "' order by short_name"
    
    tmp_Rs.Open str_SQL, cn
    If Not tmp_Rs.EOF Then
        tmp_Rs.MoveFirst
        Do While Not tmp_Rs.EOF
            rs_TRP01T("客戶") = Trim(rs_TRP01T("客戶")) & tmp_Rs("short_name") & ","
        tmp_Rs.MoveNext
        Loop
    End If
    tmp_Rs.Close
    
    '取的待進行二次排車之一次排車路線編號：總計資料值
    txt_Tab0_srcTotal_Case.Text = Val(txt_Tab0_srcTotal_Case.Text) + Val(rs_TRP01T.Fields("箱數").Value)
    txt_Tab0_srcTotal_Pallet.Text = Val(txt_Tab0_srcTotal_Pallet.Text) + Val(rs_TRP01T.Fields("板數").Value)
    txt_Tab0_srcTotal_Volumn.Text = Val(txt_Tab0_srcTotal_Volumn.Text) + Val(rs_TRP01T.Fields("材積").Value)
    txt_Tab0_srcTotal_Weight.Text = Val(txt_Tab0_srcTotal_Weight.Text) + Val(rs_TRP01T.Fields("重量").Value)
    
rs_TRP01T.MoveNext
Loop

rs_TRP01T.MoveFirst

blTRP01TEventEnable = True


'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'cn.CommandTimeout = 0   '無限期等待
'tmp_Rs.Open str_SQL2, cn, adOpenForwardOnly, adLockReadOnly
'If Not tmp_Rs.EOF Then
'   txt_Tab0_srcTotal_Case.Text = tmp_Rs.Fields("總箱數").Value
'   txt_Tab0_srcTotal_Pallet.Text = tmp_Rs.Fields("總板數").Value
'   txt_Tab0_srcTotal_Volumn.Text = tmp_Rs.Fields("總材積").Value
'   txt_Tab0_srcTotal_Weight.Text = tmp_Rs.Fields("總重量").Value
'End If
'tmp_Rs.Close
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-二次排車-ㄧ次排車路編匯入", Me.Caption, "cmd_Tab0_ImportRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Remove_Click()
'二次排車作業 >> ↓ 已選取ㄧ次排車路編取消
If rs_TRP01T Is Nothing Then Exit Sub
If rs_Tab0_SelectedRoute Is Nothing Then Exit Sub
'已選取ㄧ次排車路編若無反白選取：Disable 已選取消的動作，防止誤刪
If dg_Tab0_SelectedRoute.SelBookmarks.Count = 0 Then Exit Sub

blTab0SelectedRouteEventEnable = False

'欲移除之ㄧ次排車路線編號
Dim strRouteNo As String
strRouteNo = rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value
   
'將欲刪除之 [已選取ㄧ次排車路線編號] 加入 [待二次排車之ㄧ次排車路線編號]
Call SelectedRoute_Removeto_TRP01T(strRouteNo)
'重新產生 [待二次排車路線編號] 之 [編號] 欄位值
Call ReSet_TRP01T_SeqNo

'刪除反白選取之ㄧ次排車路編：已選取ㄧ次排車路編部分
rs_Tab0_SelectedRoute.Delete
If Not rs_Tab0_SelectedRoute.EOF Then rs_Tab0_SelectedRoute.MoveFirst
If dg_Tab0_SelectedRoute.SelBookmarks.Count > 0 Then dg_Tab0_SelectedRoute.SelBookmarks.Remove 0
'重新計算已選取ㄧ次排車路編：箱數，板數，材積，重量 + 編號重新產生
Call Calculate_SelectedRoute
blTab0SelectedRouteEventEnable = True

'還原 [篩選] 與 [排序] 之設定值
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = strSourceOrderBy
blTRP01TEventEnable = True

'重新計算 [待排車一次排車路編] 總計
Call ReCaculate_FirstRouteSum

End Sub

Private Sub cmd_Tab0_SelectCar_Click()
'DC路編併車 >> 司機選取
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
'ㄧ次排車之路編：選取

'ㄧ次排車路編：選取小計：歸零
dbSelectedCount = 0
dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""

'還原所有篩選設定，並以預設 [編號] 排列
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大

'篩選已選取者
rs_TRP01T.Filter = "＊='V'"
If Not rs_TRP01T.EOF Then
   dg_Tab0_SelectedRoute.Visible = False
   blTab0SelectedRouteEventEnable = False
   Do While Not rs_TRP01T.EOF
      '判斷是否已經選取過
      rs_Tab0_SelectedRoute.Filter = adFilterNone
      rs_Tab0_SelectedRoute.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
      rs_Tab0_SelectedRoute.Filter = "ㄧ次排車路編 = '" & rs_TRP01T.Fields("ㄧ次排車路編").Value & "'"
      If rs_Tab0_SelectedRoute.EOF Then
         '新增選取之ㄧ次排車路線編號
         rs_Tab0_SelectedRoute.AddNew
         rs_Tab0_SelectedRoute.Fields("編號").Value = 999
         rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value = rs_TRP01T.Fields("ㄧ次排車路編").Value
         rs_Tab0_SelectedRoute.Fields("箱數").Value = rs_TRP01T.Fields("箱數").Value
         rs_Tab0_SelectedRoute.Fields("板數").Value = rs_TRP01T.Fields("板數").Value
         rs_Tab0_SelectedRoute.Fields("材積").Value = rs_TRP01T.Fields("材積").Value
         rs_Tab0_SelectedRoute.Fields("重量").Value = rs_TRP01T.Fields("重量").Value
         rs_Tab0_SelectedRoute.Fields("車牌號碼").Value = rs_TRP01T.Fields("車牌號碼").Value
         rs_Tab0_SelectedRoute.Fields("車次").Value = rs_TRP01T.Fields("車次").Value
         rs_Tab0_SelectedRoute.Fields("駕駛人").Value = rs_TRP01T.Fields("駕駛人").Value
         rs_Tab0_SelectedRoute.Fields("車種").Value = rs_TRP01T.Fields("車種").Value
         rs_Tab0_SelectedRoute.Fields("出車日期").Value = rs_TRP01T.Fields("出車日期").Value
         rs_Tab0_SelectedRoute.Update
      Else
         '更新選選之ㄧ次排車資料
         rs_Tab0_SelectedRoute.Fields("編號").Value = 999
         rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value = rs_TRP01T.Fields("ㄧ次排車路編").Value
         rs_Tab0_SelectedRoute.Fields("箱數").Value = rs_TRP01T.Fields("箱數").Value
         rs_Tab0_SelectedRoute.Fields("板數").Value = rs_TRP01T.Fields("板數").Value
         rs_Tab0_SelectedRoute.Fields("材積").Value = rs_TRP01T.Fields("材積").Value
         rs_Tab0_SelectedRoute.Fields("重量").Value = rs_TRP01T.Fields("重量").Value
         rs_Tab0_SelectedRoute.Fields("車牌號碼").Value = rs_TRP01T.Fields("車牌號碼").Value
         rs_Tab0_SelectedRoute.Fields("車次").Value = rs_TRP01T.Fields("車次").Value
         rs_Tab0_SelectedRoute.Fields("駕駛人").Value = rs_TRP01T.Fields("駕駛人").Value
         rs_Tab0_SelectedRoute.Fields("車種").Value = rs_TRP01T.Fields("車種").Value
         rs_Tab0_SelectedRoute.Fields("出車日期").Value = rs_TRP01T.Fields("出車日期").Value
      End If
      rs_TRP01T.MoveNext
   Loop
   '重新對 [已選取ㄧ次排車路編] 產生 [編號] 與相關資料統計：箱數，板數，材積，重量
   Call Calculate_SelectedRoute
   dg_Tab0_SelectedRoute.Visible = True
   blTab0SelectedRouteEventEnable = True
   
   '[待選取ㄧ次排車路編] 中，刪除已選取之ㄧ次排車路線編號
   rs_TRP01T.MoveFirst
   Do While Not rs_TRP01T.EOF
      rs_TRP01T.Delete
      rs_TRP01T.MoveFirst
   Loop
   
End If
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone

rs_TRP01T.Sort = strSourceOrderBy '套用排序
'取消反白選取狀態
If dg_TRP01T.SelBookmarks.Count > 0 Then
   dg_TRP01T.SelBookmarks.Remove 0
End If

'清除路編之訂單明細
Set dg_Tab0_Orders.DataSource = Nothing
Set rs_Tab0_Orders = Nothing

blTRP01TEventEnable = True

'重新計算 [待排車一次排車路編] 總計
Call ReCaculate_FirstRouteSum


End Sub

Private Sub cmd_Tab0_SelectedCancel_All_Click()
'二次排車 >> X待選全部取消

'ㄧ次排車路編：選取小計：歸零
dbSelectedCount = 0
dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""

'還原所有篩選設定，並以預設 [編號] 排列
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大

'篩選已選取者
rs_TRP01T.Filter = "＊='V'"
If Not rs_TRP01T.EOF Then
   Do While Not rs_TRP01T.EOF
      rs_TRP01T.Fields("＊").Value = " "
      rs_TRP01T.MoveNext
   Loop
End If
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then
   rs_TRP01T.Filter = adFilterNone
End If
rs_TRP01T.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
'取消反白選取狀態
If dg_TRP01T.SelBookmarks.Count > 0 Then
   dg_TRP01T.SelBookmarks.Remove 0
End If
blTRP01TEventEnable = True

End Sub

Private Sub cmd_Tab0_SelectedCancel_Click()
'二次排車 >> X待選取消
If rs_TRP01T Is Nothing Then Exit Sub
'待選取ㄧ次排車路編若無反白選取：Disable 待選取消，防止誤刪
If dg_TRP01T.SelBookmarks.Count = 0 Then Exit Sub

If Trim(rs_TRP01T.Fields(1).Value) = "V" Then
   dbSelectedCount = dbSelectedCount - 1
   rs_TRP01T.Fields(1).Value = " "
   '待選定ㄧ次排車路編：選取小計更新
   If dbSelectedCount <> 0 Then
      dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP01T.Fields("箱數").Value
      dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP01T.Fields("板數").Value
      dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP01T.Fields("材積").Value
      dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP01T.Fields("重量").Value
   Else
      dbsrcSelected_Case = 0
      dbsrcSelected_Pallet = 0
      dbsrcSelected_Volumn = 0
      dbsrcSelected_Weight = 0
   End If
   txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
   txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
   '取消反白選取狀態
   If dg_TRP01T.SelBookmarks.Count > 0 Then
      dg_TRP01T.SelBookmarks.Remove 0
   End If
End If
'套用 [篩選] 與 [排序] 設定值
blTRP01TEventEnable = False
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = strSourceOrderBy
blTRP01TEventEnable = True
End Sub

Private Sub cmd_Tab0_SelectedRemove_All_Click()
'二次排車 >> ↓ 已選取ㄧ次排車路編取消-全部
If rs_TRP01T Is Nothing Then Exit Sub
If rs_Tab0_SelectedRoute Is Nothing Then Exit Sub
If rs_Tab0_SelectedRoute.RecordCount = 0 Then Exit Sub

blTab0SelectedRouteEventEnable = False

'欲移除之ㄧ次排車路線編號
Dim strRouteNo As String
'逐筆寫回 [ㄧ次排車路編 TRP01T]
rs_Tab0_SelectedRoute.Filter = adFilterNone
rs_Tab0_SelectedRoute.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
   strRouteNo = rs_Tab0_SelectedRoute.Fields("ㄧ次排車路編").Value
   '將欲刪除之 [已選取ㄧ次排車路編] 加入 [ㄧ次排車路編]
   Call SelectedRoute_Removeto_TRP01T(strRouteNo)
   rs_Tab0_SelectedRoute.MoveNext
Loop
   
'重新產生 [ㄧ次排車路編] 之 [編號] 欄位值
Call ReSet_TRP01T_SeqNo

'排車作業：已選取之ㄧ次排車路編列表 DBGrid 格式設定-ReSet
Call CreateRS_Tab0_SelectedRoute

'重新計算已選取ㄧ次排車路編：箱數，板數，材積，重量 + 編號重新產生
Call Calculate_SelectedRoute
blTab0SelectedRouteEventEnable = True

'套用 [篩選] 與 [排序] 設定值
blTRP01TEventEnable = False
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = strSourceOrderBy
blTRP01TEventEnable = True

'重新計算 [待排車一次排車路編] 總計
Call ReCaculate_FirstRouteSum

End Sub

Private Sub cmd_Tab0_srcRouteQuery_Click()
'二次排車作業 >> ㄧ次排車路編搜尋
If rs_TRP01T Is Nothing Then Exit Sub
If rs_TRP01T.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_TRP01T"

If ShowForm_RS_FilterAndSort(rs_TRP01T, "ㄧ次排車路線編號", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = 2
End Sub

Private Sub cmd_Tab0_srcRouteReset_Click()
'取消篩選排序
'移除篩選條件，重設排序依據
 blTRP01TEventEnable = False
 rs_TRP01T.Filter = adFilterNone
 rs_TRP01T.Sort = strSourceOrderBy  '套用排序，一般資料序號由小至大
 blTRP01TEventEnable = True

'重新計算 [待排車一次排車路編] 總計
Call ReCaculate_FirstRouteSum

End Sub

Private Sub cmd_Tab1_RouteNoDelete_Click()
'二次排車路編列表 >> 二次排車路線編號刪除
If rs_Tab1_Route.RecordCount = 0 Then Exit Sub
If dg_Tab1_Route.SelBookmarks.Count = 0 Then Exit Sub

Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
strDeleteRouteNo = Trim(rs_Tab1_Route.Fields("二次排車路編").Value)
strCarno = Trim(rs_Tab1_Route.Fields("車牌號碼").Value)
dbDriveTimes = Trim(rs_Tab1_Route.Fields("車次").Value)

'欲刪除之路編：是否已出車確認
Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "Select c_Route_No From SDN01T Where c_Route_No = '" & strDeleteRouteNo & "'"
'Terry 20191127 改為檢查出車狀態
str_SQL = "Select Route_No From TRP05T Where Route_No = '" & strDeleteRouteNo & "' and sdnstatus = '1' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
    tmp_Rs.Close
    msg_text = "注意：此路線編號已出車確認，無法刪除! "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If

msg_text = "確認刪除二次排車路線編號：" & strDeleteRouteNo
If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)

'驗證欲刪除之路編，排車者是否為此時登入之使用者
str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "資料異常：找不到欲刪除之二次排車路線編號"
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

'欲刪除之路編：車輛報到、離倉時間是否已登錄
str_SQL = "Select Convert(varchar(8),Vehicle_Check_in,112) as Checkin,Convert(varchar(8),Vehicle_Check_out,112) as Checkout From TRP05T Where Route_No = '" & strDeleteRouteNo & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("Checkin").Value <> "" Or tmp_Rs.Fields("CheckOut").Value <> "" Then
   tmp_Rs.Close
   msg_text = "資料異常：此路線編號已執行 [車輛報到] 或 [車輛離倉]，欲刪除此路編，請清除車輛進出紀錄後再進行刪除"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
tmp_Rs.Close

Screen.MousePointer = vbHourglass
blTab1RouteEventEnable = False
Tran_Level = 0
Tran_Level = cn.BeginTrans

''APP刪訂單
'cn.Execute "delete apporderdate where receipt_no in (select t2.receipt_no from trp02t t2 join trp01t t1 on t1.route_no = t2.route_no and t1.c_route_no = '" & strDeleteRouteNo & "') ", RowsAffect, adExecuteNoRecords

'刪除併車路線編號

'(1).將 TRP05T 之ㄧ次排車路編之 [二次排車路線編號] 清除
str_SQL = "Update TRP05T Set C_Route_No = null,C_Vehicle_ID_No = null,C_Drive_Times = null Where C_Route_No = '" & strDeleteRouteNo & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(2).刪除 TRP05T 二次排車路編
str_SQL = "Delete From TRP05T Where Route_No = '" & strDeleteRouteNo & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(3).將 TRP01T 之 ㄧ次排車路編 之 [二次排車路線編號] 清除
str_SQL = "Update TRP01T Set C_Route_No = null,C_Vehicle_ID_No = null,C_Drive_Times = null Where C_Route_No = '" & strDeleteRouteNo & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(4).刪除 TRP01T 之二次排車路線編號
str_SQL = "Delete From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(5).刪除查詢結果 [二次排車所屬之ㄧ次排車路編] 中該筆路線編號--rs_Tab1_RouteDC
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   rs_Tab1_RouteDC.Filter = "二次排車路編='" & strDeleteRouteNo & "'"
   If Not rs_Tab1_RouteDC.EOF Then
      Do While Not rs_Tab1_RouteDC.EOF
         rs_Tab1_RouteDC.Delete
         rs_Tab1_RouteDC.MoveFirst
      Loop
   End If
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大

'(6).刪除查詢結果 [二次排車路編] 中該筆路線編號--rs_Tab1_Route
rs_Tab1_Route.Delete
If Not rs_Tab1_Route.EOF Then rs_Tab1_Route.MoveFirst

blTab1RouteEventEnable = True

cn.CommitTrans
Tran_Level = 0
Screen.MousePointer = vbDefault


On Error GoTo err_Handle2
    
    Dim HttpClient As Object
    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/DeleteRouteNoByWareHouse?Route_NO=" & strDeleteRouteNo & "&WareHouse=DYDC_BEST", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    
    Exit Sub

err_Handle2:
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If

   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-二次排車路編列表-二次排車路編刪除", Me.Caption, "cmd_Tab1_RouteNoDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_RouteNoQuery_Click()
'二次排車路編列表 >> 二次排車路線編號查詢
If Len(Trim(txt_Tab1_RouteNo.Text)) = 0 Then Exit Sub

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle

'設定二次排車路編列表
blTab1RouteEventEnable = False
Call CreateRS_Tab1_Route
blTab1RouteEventEnable = True
'設定二次排車所屬依次排車路編列表
Call CreateRS_Tab1_RouteDC

str_SQL = "Select 二次排車路編,出車日期,車牌號碼,車次,駕駛人,箱數,板數,重量,材積,碼頭暫存,預計報到日期,預計報到時間,車種,排車者 " & _
          "From DCRouteMerge_RouteData Where 二次排車路編 like '%" & txt_Tab1_RouteNo.Text & "%' and left(二次排車路編,1) = 'S' order by 二次排車路編"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之二次排車路編資料(TRP01T)"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
blTab1RouteEventEnable = False
Do While Not tmp_Rs.EOF
   rs_Tab1_Route.AddNew
   rs_Tab1_Route.Fields("編號").Value = rs_Tab1_Route.RecordCount
   rs_Tab1_Route.Fields("二次排車路編").Value = tmp_Rs.Fields("二次排車路編").Value
   rs_Tab1_Route.Fields("出車日期").Value = tmp_Rs.Fields("出車日期").Value
   rs_Tab1_Route.Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
   rs_Tab1_Route.Fields("車次").Value = tmp_Rs.Fields("車次").Value
   rs_Tab1_Route.Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
   rs_Tab1_Route.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
   rs_Tab1_Route.Fields("板數").Value = tmp_Rs.Fields("板數").Value
   rs_Tab1_Route.Fields("材積").Value = tmp_Rs.Fields("材積").Value
   rs_Tab1_Route.Fields("重量").Value = tmp_Rs.Fields("重量").Value
   rs_Tab1_Route.Fields("碼頭暫存").Value = tmp_Rs.Fields("碼頭暫存").Value
   rs_Tab1_Route.Fields("預計報到日期").Value = tmp_Rs.Fields("預計報到日期").Value
   rs_Tab1_Route.Fields("預計報到時間").Value = tmp_Rs.Fields("預計報到時間").Value
   rs_Tab1_Route.Fields("車種").Value = tmp_Rs.Fields("車種").Value
   rs_Tab1_Route.Fields("排車者").Value = tmp_Rs.Fields("排車者").Value
   rs_Tab1_Route.Update
   tmp_Rs.MoveNext
Loop
rs_Tab1_Route.MoveFirst
blTab1RouteEventEnable = True
tmp_Rs.Close
'TRP05T
str_SQL = "Select 二次排車路編,ㄧ次排車路編,出車日期,車牌號碼,車次,駕駛人,箱數,板數,重量,材積,車種 " & _
          "From DCRouteMerge_RouteDCData " & _
           "Where 二次排車路編 like '%" & txt_Tab1_RouteNo.Text & "%' Order by 二次排車路編,ㄧ次排車路編"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定二次排車路線編號資料(TRP05T)"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Do While Not tmp_Rs.EOF
   rs_Tab1_RouteDC.AddNew
   rs_Tab1_RouteDC.Fields("編號").Value = rs_Tab1_RouteDC.RecordCount
   rs_Tab1_RouteDC.Fields("二次排車路編").Value = tmp_Rs.Fields("二次排車路編").Value
   rs_Tab1_RouteDC.Fields("ㄧ次排車路編").Value = tmp_Rs.Fields("ㄧ次排車路編").Value
   rs_Tab1_RouteDC.Fields("出車日期").Value = tmp_Rs.Fields("出車日期").Value
   rs_Tab1_RouteDC.Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
   rs_Tab1_RouteDC.Fields("車次").Value = tmp_Rs.Fields("車次").Value
   rs_Tab1_RouteDC.Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
   rs_Tab1_RouteDC.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
   rs_Tab1_RouteDC.Fields("板數").Value = tmp_Rs.Fields("板數").Value
   rs_Tab1_RouteDC.Fields("重量").Value = tmp_Rs.Fields("重量").Value
   rs_Tab1_RouteDC.Fields("材積").Value = tmp_Rs.Fields("材積").Value
   rs_Tab1_RouteDC.Fields("車種").Value = tmp_Rs.Fields("車種").Value
   rs_Tab1_RouteDC.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab1_RouteDC.MoveFirst

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-二次排車路編列表-二次排車路編查詢", Me.Caption, "cmd_Tab1_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub dg_Tab0_SelectedRoute_HeadClick(ByVal ColIndex As Integer)
'以滑鼠點選 [已選取ㄧ次排車路線編號] dg_Tab0_SelectedRoute 欄位標題區：排序欄位選取
Dim OrderFieldName As String
If TypeName(rs_Tab0_SelectedRoute) <> "Nothing" Then
   OrderFieldName = "[" & dg_Tab0_SelectedRoute.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_Tab0_SelectedRoute.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_Tab0_SelectedRoute.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

Private Sub dg_Tab0_SelectedRoute_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'排車作業 >> 已選取ㄧ次排車路線編號 DBGrid
If blTab0SelectedRouteEventEnable Then
   With dg_Tab0_SelectedRoute
        '反白顯示選取之資料列
        If Not rs_Tab0_SelectedRoute.EOF Then
           dg_Tab0_SelectedRoute.SelBookmarks.Add rs_Tab0_SelectedRoute.Bookmark
        End If
   End With
End If
End Sub

Private Sub dg_Tab1_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'二次排車路線編號列表：整行選取
If blTab1RouteEventEnable Then
   If Not rs_Tab1_Route.EOF Then
      dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
      rs_Tab1_RouteDC.Filter = " 二次排車路編 = '" & rs_Tab1_Route.Fields("二次排車路編").Value & "'"
   End If
End If
End Sub

Private Sub dg_TRP01T_HeadClick(ByVal ColIndex As Integer)
'以滑鼠點選 [ㄧ次排車路編] dg_TRP01T 欄位標題區：排序欄位選取
Dim OrderFieldName As String
If TypeName(rs_TRP01T) <> "Nothing" Then
   '避免產生 [選取] 的動作
   blTRP01TEventEnable = False
   OrderFieldName = "[" & dg_TRP01T.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_TRP01T.Sort = OrderFieldName & " DESC "
      strSourceOrderBy = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_TRP01T.Sort = OrderFieldName & " ASC "
      strSourceOrderBy = OrderFieldName & " ASC "
   End If
   blTRP01TEventEnable = True
End If
End Sub

Private Sub dg_TRP01T_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'二次排車 >> ㄧ次排車路線編號列表 DBGrid
If blTRP01TEventEnable Then
   With dg_TRP01T
        '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
        If Trim(rs_TRP01T.Fields(1).Value) = "" Then
           dbSelectedCount = dbSelectedCount + 1
           rs_TRP01T.Fields(1).Value = "V"
           '選取小計更新
           dbsrcSelected_Case = dbsrcSelected_Case + rs_TRP01T.Fields("箱數").Value
           dbsrcSelected_Pallet = dbsrcSelected_Pallet + rs_TRP01T.Fields("板數").Value
           dbsrcSelected_Volumn = dbsrcSelected_Volumn + rs_TRP01T.Fields("材積").Value
           dbsrcSelected_Weight = dbsrcSelected_Weight + rs_TRP01T.Fields("重量").Value
           txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
           txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
        Else
           dbSelectedCount = dbSelectedCount - 1
           rs_TRP01T.Fields(1).Value = " "
           '選取小計更新
           If dbSelectedCount <> 0 Then
              dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP01T.Fields("箱數").Value
              dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP01T.Fields("板數").Value
              dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP01T.Fields("材積").Value
              dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP01T.Fields("重量").Value
           Else
              dbsrcSelected_Case = 0
              dbsrcSelected_Pallet = 0
              dbsrcSelected_Volumn = 0
              dbsrcSelected_Weight = 0
           End If
           txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
           txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
        End If
        '顯示選取之路編明細
        
        '反白顯示選取之資料列
        If Not rs_TRP01T.EOF Then
           dg_TRP01T.SelBookmarks.Add rs_TRP01T.Bookmark
        End If
        '顯示路編的訂單資料
        Call Display_SelectOrdersData(rs_TRP01T.Fields("ㄧ次排車路編").Value)
   End With
End If
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "二次排車作業"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'排車作業：已選取之路線編號列表 DBGrid 格式設定
Call CreateRS_Tab0_SelectedRoute

'二次排車產生之新路編列表：DBGrid 格式設定
Call CreateRS_Tab1_Route
'被二次排車之ㄧ次排車路編列表：DBGrid 格式設定
Call CreateRS_Tab1_RouteDC
'路線編號對應之訂單資料
Call CreateRS_Tab0_Orders
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'攔截整個表單鍵盤按鍵事件
'用途：使用者按下 Esc 則不傳回任何資料，且關閉日期選取視窗
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
End If
End Sub

Private Sub Form_Resize()
On Error GoTo err_Handle

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '變小
'   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Height = Me.ScaleHeight - 120
   SSTab1.Width = Me.ScaleWidth - 120
   
   fam_SelectedOrders.Width = SSTab1.Width - 120
   fam_SrcRoute.Width = fam_SelectedOrders.Width
   dg_Tab0_SelectedRoute.Width = fam_SelectedOrders.Width - 1600
   dg_TRP01T.Width = fam_SrcRoute.Width - 1600
   dg_Tab0_Orders.Width = dg_TRP01T.Width
   
   fam_SrcRoute.Height = SSTab1.Height - fam_RouteData.Top - fam_RouteData.Height - fam_SelectedOrders.Height
   dg_Tab0_Orders.Height = fam_SrcRoute.Height - fam_SelectedSum.Height - dg_TRP01T.Height + 900

   cmd_Tab0_SelectedRemove_All.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   Shape3.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   cmd_Tab0_srcRouteQuery.Left = Shape3.Left + 60
   cmd_Tab0_srcRouteReset.Left = Shape3.Left + 60

   Label1(11).Left = dg_TRP01T.Left + dg_TRP01T.Width + 120
   Label1(10).Left = Label1(11).Left
   Label1(9).Left = Label1(11).Left
   Label1(8).Left = Label1(11).Left
   txt_Tab0_srcTotal_Case.Left = Label1(11).Left
   txt_Tab0_srcTotal_Pallet.Left = Label1(11).Left
   txt_Tab0_srcTotal_Volumn.Left = Label1(11).Left
   txt_Tab0_srcTotal_Weight.Left = Label1(11).Left

   fam_Tab1_Query.Left = fam_Tab1_Query.Left - (dbsrcFormWidth - Me.ScaleWidth)
   fam_Tab1_Delete.Left = fam_Tab1_Delete.Left - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab1_Route.Width = dg_Tab1_Route.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab1_RouteDC.Height = dg_Tab1_RouteDC.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Tab1_RouteDC.Width = dg_Tab1_RouteDC.Width - (dbsrcFormWidth - Me.ScaleWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = Me.ScaleHeight - 120
   SSTab1.Width = Me.ScaleWidth - 120
   
   fam_SelectedOrders.Width = SSTab1.Width - 120
   fam_SrcRoute.Width = fam_SelectedOrders.Width
   dg_Tab0_SelectedRoute.Width = fam_SelectedOrders.Width - 1600
   dg_TRP01T.Width = fam_SrcRoute.Width - 1600
   dg_Tab0_Orders.Width = dg_TRP01T.Width
   
   fam_SrcRoute.Height = SSTab1.Height - fam_RouteData.Top - fam_RouteData.Height - fam_SelectedOrders.Height
   dg_Tab0_Orders.Height = fam_SrcRoute.Height - fam_SelectedSum.Height - dg_TRP01T.Height - 120
   
   cmd_Tab0_SelectedRemove_All.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   Shape3.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   cmd_Tab0_srcRouteQuery.Left = Shape3.Left + 60
   cmd_Tab0_srcRouteReset.Left = Shape3.Left + 60

   Label1(11).Left = dg_TRP01T.Left + dg_TRP01T.Width + 120
   Label1(10).Left = Label1(11).Left
   Label1(9).Left = Label1(11).Left
   Label1(8).Left = Label1(11).Left
   txt_Tab0_srcTotal_Case.Left = Label1(11).Left
   txt_Tab0_srcTotal_Pallet.Left = Label1(11).Left
   txt_Tab0_srcTotal_Volumn.Left = Label1(11).Left
   txt_Tab0_srcTotal_Weight.Left = Label1(11).Left
   
   fam_Tab1_Query.Left = fam_Tab1_Query.Left + (Me.ScaleWidth - dbsrcFormWidth)
   fam_Tab1_Delete.Left = fam_Tab1_Delete.Left + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab1_Route.Width = dg_Tab1_Route.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab1_RouteDC.Height = dg_Tab1_RouteDC.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Tab1_RouteDC.Width = dg_Tab1_RouteDC.Width + (Me.ScaleWidth - dbsrcFormWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
End If

Exit Sub
err_Handle:
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_OP_DCRouteMerge = Nothing
End Sub

Private Sub CreateRS_Tab0_SelectedRoute()
'排車作業：已選取之ㄧ次排車路編列表
Call ReDim_Recordset(rs_Tab0_SelectedRoute)
With rs_Tab0_SelectedRoute
     .Fields.Append "編號", adDouble
     .Fields.Append "ㄧ次排車路編", adVarChar, 10
     .Fields.Append "箱數", adDouble
     .Fields.Append "板數", adDouble
     .Fields.Append "材積", adDouble
     .Fields.Append "重量", adDouble
     .Fields.Append "車牌號碼", adVarChar, 20
     .Fields.Append "車次", adDouble
     .Fields.Append "駕駛人", adVarChar, 30
     .Fields.Append "出車日期", adVarChar, 12
     .Fields.Append "車種", adVarChar, 10
     .Fields.Append "EXE回傳", adVarChar, 20
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
Set dg_Tab0_SelectedRoute.DataSource = rs_Tab0_SelectedRoute
'設定顯示欄位
With dg_Tab0_SelectedRoute
    .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
    .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
    .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
    .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
    .Columns(0).Width = 500         '編號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1200        'ㄧ次排車路線編號
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 800         '箱數
    .Columns(2).Alignment = dbgRight
    .Columns(3).Width = 800         '板數
    .Columns(3).Alignment = dbgRight
    .Columns(4).Width = 800         '材積
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 800         '重量
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 900         '車牌號碼
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 500         '車次
    .Columns(7).Alignment = dbgCenter
    .Columns(8).Width = 800         '駕駛人
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 1000        '出車日期
    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 500       '車種
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 800        'EXE回傳
    .Columns(11).Alignment = dbgLeft
End With
End Sub

Private Sub Calculate_SelectedRoute()
'作業內容：
'1.針對已選取ㄧ次排出路編列表，依ㄧ次排車路編重新產生 [編號] 欄位值
'2.計算已選取ㄧ次排車路編之累計資料
Dim dbSeqNo As Double
dbSeqNo = 0
txt_Tab0_Selected_Case.Text = ""
txt_Tab0_Selected_Pallet.Text = ""
txt_Tab0_Selected_Volumn.Text = ""
txt_Tab0_Selected_Weight.Text = ""

rs_Tab0_SelectedRoute.Filter = adFilterNone
rs_Tab0_SelectedRoute.Sort = "ㄧ次排車路編 asc"  '原始排序，一般資料序號由小至大
If Not rs_Tab0_SelectedRoute.EOF Then
   rs_Tab0_SelectedRoute.MoveFirst
Else
   '清出篩選條件，仍無資料者，結束 SubProgram 執行
   Exit Sub
End If
Do While Not rs_Tab0_SelectedRoute.EOF
   dbSeqNo = dbSeqNo + 1
   rs_Tab0_SelectedRoute.Fields("編號").Value = dbSeqNo
   txt_Tab0_Selected_Case.Text = Val(txt_Tab0_Selected_Case.Text) + rs_Tab0_SelectedRoute.Fields("箱數").Value
   txt_Tab0_Selected_Pallet.Text = Val(txt_Tab0_Selected_Pallet.Text) + rs_Tab0_SelectedRoute.Fields("板數").Value
   txt_Tab0_Selected_Volumn.Text = Val(txt_Tab0_Selected_Volumn.Text) + rs_Tab0_SelectedRoute.Fields("材積").Value
   txt_Tab0_Selected_Weight.Text = Val(txt_Tab0_Selected_Weight.Text) + rs_Tab0_SelectedRoute.Fields("重量").Value
   rs_Tab0_SelectedRoute.MoveNext
Loop
rs_Tab0_SelectedRoute.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
If Not rs_Tab0_SelectedRoute.EOF Then rs_Tab0_SelectedRoute.MoveFirst
End Sub

Private Sub SelectedRoute_Removeto_TRP01T(ByVal strRouteNo As String)
'將指定之 [ㄧ次排車路編] 加入 [ㄧ次排車路編列表]
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "編號 asc"  '原始排序，一般資料序號由小至大

rs_TRP01T.Filter = "ㄧ次排車路編 = '" & strRouteNo & "'"
If Not rs_TRP01T.EOF Then
   'ㄧ次排車路線編號已存在的話，不進行新增，也不更新
   rs_TRP01T.Filter = adFilterNone
   rs_TRP01T.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
   blTRP01TEventEnable = True
   Exit Sub
End If

'取回ㄧ次排車路線編號
str_SQL = "Select ' ' as '＊',ㄧ次排車路編,碼頭,箱數,板數,材積,重量,車牌號碼,車次,駕駛人,出車日期,車種,EXE回傳 " & _
          "From DCRouteMerge_DCRouteData Where ㄧ次排車路編 = '" & strRouteNo & "' Order by ㄧ次排車路編 "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合選定之ㄧ次排車路線編號資料可以取消"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   blTRP01TEventEnable = True
   Exit Sub
End If

rs_TRP01T.AddNew
rs_TRP01T.Fields("編號").Value = rs_TRP01T.RecordCount
rs_TRP01T.Fields("ㄧ次排車路編").Value = tmp_Rs.Fields("ㄧ次排車路編").Value
rs_TRP01T.Fields("碼頭").Value = tmp_Rs.Fields("碼頭").Value
rs_TRP01T.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
rs_TRP01T.Fields("板數").Value = tmp_Rs.Fields("板數").Value
rs_TRP01T.Fields("材積").Value = tmp_Rs.Fields("材積").Value
rs_TRP01T.Fields("重量").Value = tmp_Rs.Fields("重量").Value
rs_TRP01T.Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
rs_TRP01T.Fields("車次").Value = tmp_Rs.Fields("車次").Value
rs_TRP01T.Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
rs_TRP01T.Fields("出車日期").Value = tmp_Rs.Fields("出車日期").Value
rs_TRP01T.Fields("車種").Value = tmp_Rs.Fields("車種").Value
rs_TRP01T.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
rs_TRP01T.Update
tmp_Rs.Close

rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "ㄧ次排車路編 asc"  '原始排序，一般資料序號由小至大
If Not rs_TRP01T.EOF Then rs_TRP01T.MoveFirst
blTRP01TEventEnable = True
End Sub

Private Sub ReSet_TRP01T_SeqNo()
'重新產生 [ㄧ次排車路線編號] 之 [編號] 欄位值
blTRP01TEventEnable = False
dg_TRP01T.Visible = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "ㄧ次排車路編 asc"  '原始排序，一般資料序號由小至大
If Not rs_TRP01T.EOF Then rs_TRP01T.MoveFirst

Dim dbSeqNo As Double
dbSeqNo = 0
Do While Not rs_TRP01T.EOF
   dbSeqNo = dbSeqNo + 1
   rs_TRP01T.Fields("編號").Value = dbSeqNo
   rs_TRP01T.MoveNext
Loop
If rs_TRP01T.RecordCount > 0 Then rs_TRP01T.MoveFirst
dg_TRP01T.Visible = True
blTRP01TEventEnable = True
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'日期選取
Select Case mvDate.Tag
       Case "出車日期"
            txt_Tab0_TRPDate.Text = Format(mvDate.Value, "yyyymmdd")
       Case "預計報到日期"
            txt_Tab0_CarCheckInDate.Text = Format(mvDate.Value, "yyyymmdd")
       Case "排車日期.起"
            txt_FPlanDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "排車日期.迄"
            txt_FPlanDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "出車日期.起"
            txt_FDeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "出車日期.迄"
            txt_FDeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_FPlanDate_Start_Click()
'二次排車 >> 匯入一次排車路編 >> 排車日期：起
If Trim(txt_FPlanDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_Start.Text, 4) & "/" & Mid(txt_FPlanDate_Start.Text, 5, 2) & "/" & Right(txt_FPlanDate_Start.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FPlanDate_Start.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FPlanDate_Start.Top + txt_FPlanDate_Start.Height
mvDate.Tag = "排車日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_Start_KeyPress(KeyAscii As Integer)
'二次排車 >> 匯入一次排車路編 >> 排車日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanDate_End.SelStart = 0: txt_FPlanDate_End.SelLength = Len(txt_FPlanDate_End.Text)
         txt_FPlanDate_End.SetFocus
End Select
End Sub

Private Sub txt_FPlanDate_End_Click()
'二次排車 >> 匯入一次排車路編 >> 排車日期：迄
If Trim(txt_FPlanDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_End.Text, 4) & "/" & Mid(txt_FPlanDate_End.Text, 5, 2) & "/" & Right(txt_FPlanDate_End.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FPlanDate_End.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FPlanDate_End.Top + txt_FPlanDate_End.Height
mvDate.Tag = "排車日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_End_KeyPress(KeyAscii As Integer)
'二次排車 >> 匯入一次排車路編 >> 排車日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_Start.SelStart = 0: txt_FDeliveryDate_Start.SelLength = Len(txt_FDeliveryDate_Start.Text)
         txt_FDeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_Start_Click()
'二次排車 >> 匯入一次排車路編 >> 出車日期：起
If Trim(txt_FDeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_Start.Text, 4) & "/" & Mid(txt_FDeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FDeliveryDate_Start.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FDeliveryDate_Start.Top + txt_FDeliveryDate_Start.Height
mvDate.Tag = "出車日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_Start_KeyPress(KeyAscii As Integer)
'二次排車 >> 匯入一次排車路編 >> 出車日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_End.SelStart = 0: txt_FDeliveryDate_End.SelLength = Len(txt_FDeliveryDate_End.Text)
         txt_FDeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_End_Click()
'二次排車 >> 匯入一次排車路編 >> 出車日期：迄
If Trim(txt_FDeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_End.Text, 4) & "/" & Mid(txt_FDeliveryDate_End.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FDeliveryDate_End.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FDeliveryDate_End.Top + txt_FDeliveryDate_End.Height
mvDate.Tag = "出車日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_End_KeyPress(KeyAscii As Integer)
'二次排車 >> 匯入一次排車路編 >> 出車日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         cmd_Tab0_ImportRoute.SetFocus
End Select
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
'二次排車 >> 預計報到時間
Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_KeyPress(KeyAscii As Integer)
'二次排車 >> 車牌號碼
Select Case KeyAscii
       Case 97 To 122   '轉換為大寫字元
            KeyAscii = KeyAscii - 32
End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_LostFocus()
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
'二次排車 >> 碼頭暫存
Select Case KeyAscii
       Case 97 To 122   '轉換為大寫字元
            KeyAscii = KeyAscii - 32
       Case vbKeyReturn
            KeyAscii = 0
            txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
            txt_Tab0_CarCheckInDate.SetFocus
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
'二次排車 > [出車日期] 資料格式：yyyymmdd
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
Public Sub frm_OP_DCRouteMerge_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
'表單公用副程式，由 frm_RS_FilterAndSort 表單呼叫
'傳入值：strCode      動作識別碼
'                     [FILTER] 自訂篩選    [SORT] 排序
'        strReturn    篩選 or 排序 之設定字串

Select Case strCode
       Case "FILTER"  '自訂篩選
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP01T"   'ㄧ次排車路線編號資料
                        blTRP01TEventEnable = False
                        rs_TRP01T.Filter = adFilterNone
                        rs_TRP01T.Filter = strReturn
                        strSourceFilter = strReturn
                        If rs_TRP01T.RecordCount = 0 Then
                           msg_text = "抱歉ㄟ，找不到符合條件的ㄧ次排車路線編號喔"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_TRP01T.Filter = adFilterNone
                           strSourceFilter = adFilterNone
                           rs_TRP01T.Sort = strSourceOrderBy  '套用排序，一般資料序號由小至大
                           blTRP01TEventEnable = True
                           Exit Sub
                        End If
                        blTRP01TEventEnable = True
                        '重新計算 [待排車一次排車路編] 總計
                        Call ReCaculate_FirstRouteSum

            End Select
       Case "SORT"    '排序
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP01T"   'ㄧ次排車路線編號資料
                        blTRP01TEventEnable = False
                        rs_TRP01T.Sort = strReturn
                        strSourceOrderBy = strReturn
                        blTRP01TEventEnable = True
            End Select
End Select
End Sub

Private Sub CreateRS_Tab1_Route()
'排車作業：二次排車產生之路線編號列表
Call ReDim_Recordset(rs_Tab1_Route)
With rs_Tab1_Route
     .Fields.Append "編號", adVarChar, 10
     .Fields.Append "二次排車路編", adVarChar, 10
     .Fields.Append "出車日期", adVarChar, 8
     .Fields.Append "車牌號碼", adVarChar, 10
     .Fields.Append "車次", adDouble
     .Fields.Append "駕駛人", adVarChar, 20
     .Fields.Append "箱數", adDouble
     .Fields.Append "板數", adDouble
     .Fields.Append "材積", adDouble
     .Fields.Append "重量", adDouble
     .Fields.Append "碼頭暫存", adVarChar, 10
     .Fields.Append "預計報到日期", adVarChar, 8
     .Fields.Append "預計報到時間", adVarChar, 4
     .Fields.Append "車種", adVarChar, 10
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
    .Columns(1).Width = 1200        '二次排車路線編號
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900         '出車日期
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 850         '車牌號碼
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 500         '車次
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 900         '駕駛人
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 700         '箱數
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 700         '板數
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700         '材積
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700         '重量
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 900        '碼頭暫存
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1200       '預計車輛報到日期
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1200       '預計車輛報到時間
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 500       '車種
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1300       '排車者
    .Columns(14).Alignment = dbgLeft
End With
End Sub

Private Sub CreateRS_Tab1_RouteDC()
'排車作業：已編妥二次排車路編之 ㄧ次排車 路編列表
Call ReDim_Recordset(rs_Tab1_RouteDC)
With rs_Tab1_RouteDC
     .Fields.Append "編號", adVarChar, 10
     .Fields.Append "二次排車路編", adVarChar, 10
     .Fields.Append "ㄧ次排車路編", adVarChar, 10
     .Fields.Append "出車日期", adVarChar, 8
     .Fields.Append "車牌號碼", adVarChar, 10
     .Fields.Append "車次", adDouble
     .Fields.Append "駕駛人", adVarChar, 20
     .Fields.Append "箱數", adDouble
     .Fields.Append "板數", adDouble
     .Fields.Append "材積", adDouble
     .Fields.Append "重量", adDouble
     .Fields.Append "車種", adVarChar, 10
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
Set dg_Tab1_RouteDC.DataSource = rs_Tab1_RouteDC
'設定顯示欄位
With dg_Tab1_RouteDC
    .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
    .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
    .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
    .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
    .Columns(0).Width = 500         '編號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1200        '二次排車路編
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1200        'ㄧ次排車路編
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 900         '出車日期
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 850         '車牌號碼
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 500         '車次
    .Columns(5).Alignment = dbgCenter
    .Columns(6).Width = 900         '駕駛人
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 700         '箱數
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700         '板數
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700         '材積
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 700        '重量
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 500       '車種
    .Columns(11).Alignment = dbgLeft
End With
End Sub

Private Sub txt_Tab1_RouteNo_KeyPress(KeyAscii As Integer)
'併車路線編號列表 >> 併車路線編號
Select Case KeyAscii
     Case 97 To 122   '轉換大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          cmd_Tab1_RouteNoQuery.SetFocus
End Select
End Sub

Private Sub Display_SelectOrdersData(ByVal strRouteNo As String)
'顯示傳入之路編對應之訂單
Set dg_Tab0_Orders.DataSource = Nothing
Set rs_Tab0_Orders = Nothing
Call CreateRS_Tab0_Orders

'str_SQL = "Select 路線編號,送貨日,訂單編號,ZIP,Area,客戶名稱,箱數,板數,材積,重量,訂單備註,Receipt_No,EXE回傳 " & _
'          "From DCRouteMerge_RouteOrders " & _
'           "Where 路線編號 like '" & strRouteNo & "%' Order by Receipt_No"
           
str_SQL = "Select  Rtrim(a1.Route_No) as 路線編號 " & _
        ", Convert(varchar,a1.Arrive_Date,112) as 送貨日 " & _
        ", Rtrim(a1.Receipt_No) + '(' + Rtrim(a1.StorerKey)+'-'+Rtrim(a1.Extern)+')' as 訂單編號 " & _
        ", Rtrim(a2.ZIP) as ZIP , Rtrim(Isnull(a2.Area_Code,'')) as Area , Rtrim(a2.Full_Name) as 客戶名稱 " & _
        ", Round(a1.Case_cnt,2) as 箱數 ,  Round(a1.Pallet_Qty,2) as 板數 " & _
        ", Round(a1.Volumn_Weight,2) as 材積 " & _
        ", Round(a1.Weight,2) as 重量 " & _
        ",訂單備註 = rtrim(a1.description) " & _
        ", Rtrim(a1.Receipt_No) as Receipt_No " & _
        ", Case a1.EXE_Confirm When '0' Then '新建路編' When '1' Then '設定回傳' When '2' Then '已回傳' When '9' Then '預先揀貨' else '未知狀態' End  AS EXE回傳 " & _
        "From TRP02T a1(nolock) inner join TRP01M a2(nolock) on a2.ConsigneeKey = a1.ConsigneeKey and a2.storerkey = a1.storerkey " & _
        "where Rtrim(a1.Route_No) = '" & strRouteNo & "' order by Rtrim(a1.Receipt_No) "

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
   rs_Tab0_Orders.AddNew
   rs_Tab0_Orders.Fields("編號").Value = rs_Tab0_Orders.RecordCount
   rs_Tab0_Orders.Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
   rs_Tab0_Orders.Fields("送貨日").Value = tmp_Rs.Fields("送貨日").Value
   rs_Tab0_Orders.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
   rs_Tab0_Orders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
   rs_Tab0_Orders.Fields("Area").Value = tmp_Rs.Fields("Area").Value
   rs_Tab0_Orders.Fields("客戶名稱").Value = tmp_Rs.Fields("客戶名稱").Value
   rs_Tab0_Orders.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
   rs_Tab0_Orders.Fields("板數").Value = tmp_Rs.Fields("板數").Value
   rs_Tab0_Orders.Fields("材積").Value = tmp_Rs.Fields("材積").Value
   rs_Tab0_Orders.Fields("重量").Value = tmp_Rs.Fields("重量").Value
   rs_Tab0_Orders.Fields("訂單備註").Value = tmp_Rs.Fields("訂單備註").Value & ""
   rs_Tab0_Orders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
   rs_Tab0_Orders.Fields("EXE回傳").Value = tmp_Rs.Fields("EXE回傳").Value
   rs_Tab0_Orders.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab0_Orders.MoveFirst

End Sub

Private Sub CreateRS_Tab0_Orders()
'排車作業：已編妥路編之訂單列表
Call ReDim_Recordset(rs_Tab0_Orders)
With rs_Tab0_Orders
     .Fields.Append "編號", adVarChar, 10
     .Fields.Append "路線編號", adVarChar, 10
     .Fields.Append "送貨日", adVarChar, 20
     .Fields.Append "訂單編號", adVarChar, 60
     .Fields.Append "ZIP", adVarChar, 60
     .Fields.Append "Area", adVarChar, 60
     .Fields.Append "客戶名稱", adVarChar, 120
     .Fields.Append "箱數", adDouble
     .Fields.Append "板數", adDouble
     .Fields.Append "材積", adDouble
     .Fields.Append "重量", adDouble
     .Fields.Append "訂單備註", adVarChar, 300
     .Fields.Append "Receipt_No", adVarChar, 60
     .Fields.Append "EXE回傳", adVarChar, 20
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
Set dg_Tab0_Orders.DataSource = rs_Tab0_Orders
'設定顯示欄位
With dg_Tab0_Orders
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
    .Columns(5).Width = 400         'Area 運送區域代碼
    .Columns(5).Alignment = dbgCenter
    .Columns(6).Width = 1500        '客戶名稱
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 700         '箱數
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700         '板數
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700         '材積
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 700         '重量
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 2500       '訂單備註
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1100       'Receipt_No
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 1100       'EXE回傳
    .Columns(13).Alignment = dbgLeft
End With
End Sub

Private Sub ReCaculate_FirstRouteSum()
'計算 [待排車一次排車路編] 總計
'不採用再次擷取統計，因為會有套用 [篩選條件] 的問題
txt_Tab0_srcTotal_Case.Text = ""
txt_Tab0_srcTotal_Pallet.Text = ""
txt_Tab0_srcTotal_Volumn.Text = ""
txt_Tab0_srcTotal_Weight.Text = ""

If rs_TRP01T Is Nothing Then Exit Sub
If rs_TRP01T.RecordCount = 0 Then Exit Sub

Dim dbTotalCase As Double
Dim dbTotalPallet As Double
Dim dbTotalWeight As Double
Dim dbTotalVolumn As Double
dbTotalCase = 0: dbTotalPallet = 0: dbTotalVolumn = 0: dbTotalWeight = 0
blTRP01TEventEnable = False
dg_TRP01T.Visible = False
rs_TRP01T.MoveFirst
Do While Not rs_TRP01T.EOF
   dbTotalCase = dbTotalCase + rs_TRP01T.Fields("箱數").Value
   dbTotalPallet = dbTotalPallet + rs_TRP01T.Fields("板數").Value
   dbTotalVolumn = dbTotalVolumn + rs_TRP01T.Fields("材積").Value
   dbTotalWeight = dbTotalWeight + rs_TRP01T.Fields("重量").Value
   rs_TRP01T.MoveNext
Loop
rs_TRP01T.MoveFirst

txt_Tab0_srcTotal_Case.Text = dbTotalCase
txt_Tab0_srcTotal_Pallet.Text = dbTotalPallet
txt_Tab0_srcTotal_Volumn.Text = dbTotalVolumn
txt_Tab0_srcTotal_Weight.Text = dbTotalWeight
dg_TRP01T.Visible = True
blTRP01TEventEnable = True
End Sub
