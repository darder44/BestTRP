VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_RouteConfirm 
   Caption         =   "出車確認"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11400
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   240
      TabIndex        =   116
      Top             =   4680
      Visible         =   0   'False
      Width           =   3465
      _ExtentX        =   6112
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
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   97320961
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14215660
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "出車確認"
      TabPicture(0)   =   "frm_OP_RouteConfirm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dg_Tab0_C_RouteList"
      Tab(0).Control(1)=   "fam_Tab0_Consignee"
      Tab(0).Control(2)=   "cmd_Tab0_ConsigneeShow"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "cmd_Tab0_Confirm"
      Tab(0).Control(5)=   "cmd_Tab0_Delete"
      Tab(0).Control(6)=   "chkPalletDefend1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "待重組排車"
      TabPicture(1)   =   "frm_OP_RouteConfirm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_Tab1_Reset"
      Tab(1).Control(1)=   "cmd_Tab1_Add"
      Tab(1).Control(2)=   "cmd_Tab1_Query"
      Tab(1).Control(3)=   "fam_SelectedOrders"
      Tab(1).Control(4)=   "fam_SrcOrders"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "已確認路編 "
      TabPicture(2)   =   "frm_OP_RouteConfirm.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "dg_Tab2_Route"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dg_Tab2_RouteOrders"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   " 訂單切割"
      TabPicture(3)   =   "frm_OP_RouteConfirm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fam_Tab3_Orders"
      Tab(3).Control(1)=   "fam_Tab3_OrderDetail"
      Tab(3).Control(2)=   "cmd_Tab3_DisplaySelectedOrder"
      Tab(3).Control(3)=   "cmd_Tab3_DisplayOrders"
      Tab(3).Control(4)=   "dg_Tab3_SDN02W"
      Tab(3).ControlCount=   5
      Begin VB.CheckBox chkPalletDefend1 
         BackColor       =   &H8000000A&
         Caption         =   "是否維護棧板"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   -71280
         TabIndex        =   134
         Top             =   360
         Value           =   1  '核取
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Frame Frame6 
         Height          =   645
         Left            =   120
         TabIndex        =   119
         Top             =   360
         Width           =   9075
         Begin VB.TextBox txt_Tab2_Route 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
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
            Left            =   600
            TabIndex        =   128
            Top             =   150
            Width           =   1365
         End
         Begin VB.CommandButton cmd_Tab2_SelectCar 
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
            Left            =   3630
            Style           =   1  '圖片外觀
            TabIndex        =   124
            Top             =   150
            Width           =   330
         End
         Begin VB.TextBox txt_Tab2_VehicleNo 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   123
            Top             =   150
            Width           =   1125
         End
         Begin VB.TextBox txt_Tab2_DELIVERY_DATE 
            BackColor       =   &H00E0E0E0&
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
            Left            =   6645
            TabIndex        =   122
            Top             =   150
            Width           =   1110
         End
         Begin VB.TextBox txt_Tab2_Driver 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Left            =   4560
            TabIndex        =   121
            Top             =   150
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab2_CreateRoute 
            Appearance      =   0  '平面
            BackColor       =   &H00FF8080&
            Caption         =   "確定存檔"
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
            Left            =   7920
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   120
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "路編"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   31
            Left            =   120
            TabIndex        =   129
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "車號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   30
            Left            =   2040
            TabIndex        =   127
            Top             =   240
            Width           =   540
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
            Height          =   195
            Index           =   29
            Left            =   5760
            TabIndex        =   126
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "司機"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   14
            Left            =   4080
            TabIndex        =   125
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame fam_Tab3_Orders 
         BackColor       =   &H8000000A&
         Caption         =   "新增訂單資料"
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   100
         Top             =   4920
         Width           =   10905
         Begin VB.CommandButton cmd_Tab3_Query 
            BackColor       =   &H00FFC0C0&
            Caption         =   "？"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            Style           =   1  '圖片外觀
            TabIndex        =   115
            Top             =   158
            Width           =   330
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   5325
            TabIndex        =   114
            Top             =   150
            Width           =   1230
         End
         Begin VB.CommandButton cmd_Tab3_DelOrders 
            BackColor       =   &H00FF8080&
            Caption         =   "刪除訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9600
            Style           =   1  '圖片外觀
            TabIndex        =   113
            Top             =   465
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab3_CaseQty 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   9780
            TabIndex        =   112
            Top             =   120
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_Volumn 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   8835
            TabIndex        =   111
            Top             =   120
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_Weight 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   7890
            TabIndex        =   110
            Top             =   120
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_FullName 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   945
            TabIndex        =   103
            Top             =   465
            Width           =   5610
         End
         Begin VB.TextBox txt_Tab3_Extern 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   3390
            TabIndex        =   102
            Top             =   165
            Width           =   1050
         End
         Begin VB.TextBox txt_Tab3_OrderKey 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   945
            TabIndex        =   101
            Top             =   165
            Width           =   1170
         End
         Begin MSDataGridLib.DataGrid dg_Tab3_SDN03W 
            Height          =   1155
            Left            =   120
            TabIndex        =   109
            Top             =   840
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   2037
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重量/材積/箱數"
            Height          =   180
            Index           =   38
            Left            =   6645
            TabIndex        =   108
            Top             =   195
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "出貨日期"
            Height          =   180
            Index           =   37
            Left            =   4560
            TabIndex        =   107
            Top             =   210
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶名稱"
            Height          =   180
            Index           =   28
            Left            =   165
            TabIndex        =   106
            Top             =   525
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主單號"
            Height          =   180
            Index           =   27
            Left            =   2610
            TabIndex        =   105
            Top             =   225
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單編號"
            Height          =   180
            Index           =   24
            Left            =   165
            TabIndex        =   104
            Top             =   210
            Width           =   720
         End
      End
      Begin VB.Frame fam_Tab3_OrderDetail 
         BackColor       =   &H80000000&
         Caption         =   "訂單明細"
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   -74760
         TabIndex        =   86
         Top             =   2520
         Width           =   10875
         Begin VB.CommandButton cmd_Tab3_ClearQty 
            BackColor       =   &H00FF80FF&
            Caption         =   "Ｘ"
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
            Left            =   9105
            Style           =   1  '圖片外觀
            TabIndex        =   93
            Top             =   150
            Width           =   420
         End
         Begin VB.CommandButton cmd_Tab3_CutOrders 
            BackColor       =   &H00FF8080&
            Caption         =   "訂單切割"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9600
            Style           =   1  '圖片外觀
            TabIndex        =   92
            Top             =   165
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab3_CutQty 
            BackColor       =   &H00C0C0FF&
            Caption         =   "數量切割"
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
            Left            =   7920
            Style           =   1  '圖片外觀
            TabIndex        =   91
            Top             =   150
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab3_CutCaseQty 
            Alignment       =   2  '置中對齊
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
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   7200
            TabIndex        =   90
            Top             =   210
            Width           =   700
         End
         Begin VB.TextBox txt_Tab3_SelectedCaseQty 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1995
            TabIndex        =   89
            Top             =   225
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_SelectedWeight 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   3555
            TabIndex        =   88
            Top             =   225
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_SelectedVolumn 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   5100
            TabIndex        =   87
            Top             =   225
            Width           =   945
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab3_SelectedOrderDetail 
            Height          =   1845
            Left            =   45
            TabIndex        =   94
            Top             =   525
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   3254
            _Version        =   393216
            Cols            =   9
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "選取項次小計"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   195
            TabIndex        =   99
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "箱數切割"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   19
            Left            =   6315
            TabIndex        =   98
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重量"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   3045
            TabIndex        =   97
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材積"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   4575
            TabIndex        =   96
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "箱數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   1500
            TabIndex        =   95
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.CommandButton cmd_Tab3_DisplaySelectedOrder 
         BackColor       =   &H00C0E0FF&
         Caption         =   "訂單切割明細"
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
         Left            =   -72120
         Style           =   1  '圖片外觀
         TabIndex        =   84
         Top             =   360
         Width           =   2250
      End
      Begin VB.CommandButton cmd_Tab3_DisplayOrders 
         BackColor       =   &H00FF8080&
         Caption         =   "匯入待排車訂單"
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
         Left            =   -74760
         Style           =   1  '圖片外觀
         TabIndex        =   83
         Top             =   360
         Width           =   2250
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '平面
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   3345
         Left            =   9300
         TabIndex        =   77
         Top             =   360
         Width           =   1995
         Begin VB.CommandButton cmd_Tab2_Excel 
            BackColor       =   &H00FFFF80&
            Caption         =   "轉 Excel"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   133
            ToolTipText     =   "刪除"
            Top             =   2760
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab2_DeliveryDate_Start 
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
            Left            =   173
            TabIndex        =   131
            Top             =   1290
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab2_RouteNoDelete 
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
            Height          =   465
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   130
            ToolTipText     =   "刪除"
            Top             =   2280
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab2_Route_Start 
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
            Left            =   173
            MaxLength       =   10
            TabIndex        =   79
            Top             =   525
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab2_RouteNoQuery 
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
            Height          =   465
            Left            =   120
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   78
            Top             =   1800
            Width           =   1785
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "出車日期"
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
            TabIndex        =   132
            Top             =   960
            Width           =   1020
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
            TabIndex        =   80
            Top             =   195
            Width           =   1020
         End
      End
      Begin VB.Frame fam_SrcOrders 
         Height          =   2955
         Left            =   -74865
         TabIndex        =   45
         Top             =   4020
         Width           =   11220
         Begin VB.Frame fam_SelectedSum 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   5595
            Begin VB.TextBox txt_Tab1_srcSelected_Weight 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   4695
               TabIndex        =   56
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcSelected_Volumn 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   2865
               TabIndex        =   55
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcSelected_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   990
               TabIndex        =   54
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "選取：箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   25
               Left            =   75
               TabIndex        =   59
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "材積"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   23
               Left            =   2475
               TabIndex        =   58
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "重量"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   22
               Left            =   4320
               TabIndex        =   57
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Enabled         =   0   'False
            Height          =   525
            Left            =   5610
            TabIndex        =   46
            Top             =   0
            Width           =   5595
            Begin VB.TextBox txt_Tab1_srcTotal_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   975
               TabIndex        =   49
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcTotal_Volumn 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   2865
               TabIndex        =   48
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcTotal_Weight 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   4680
               TabIndex        =   47
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "重量"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   21
               Left            =   4305
               TabIndex        =   52
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "材積"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   20
               Left            =   2475
               TabIndex        =   51
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "總計：箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   18
               Left            =   75
               TabIndex        =   50
               Top             =   210
               Width           =   900
            End
         End
         Begin MSDataGridLib.DataGrid dg_SDN02W 
            Height          =   2190
            Left            =   45
            TabIndex        =   60
            Top             =   600
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   3863
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
         Height          =   3495
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   11220
         Begin VB.CommandButton cmd_Tab1_CreateRoute 
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
            Left            =   8520
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   76
            Top             =   120
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab1_Driver0 
            BackColor       =   &H00E0E0E0&
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
            Left            =   4920
            TabIndex        =   74
            Top             =   150
            Width           =   1080
         End
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   67
            Top             =   2835
            Width           =   5595
            Begin VB.TextBox txt_Tab1_Selected_Weight 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   4695
               TabIndex        =   70
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_Selected_Volumn 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   2865
               TabIndex        =   69
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_Selected_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   990
               TabIndex        =   68
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "累計：箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   9
               Left            =   75
               TabIndex        =   73
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "材積"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   7
               Left            =   2475
               TabIndex        =   72
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "重量"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   6
               Left            =   4320
               TabIndex        =   71
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.CommandButton cmd_Tab1_Selected 
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
            Left            =   5655
            Style           =   1  '圖片外觀
            TabIndex        =   66
            Top             =   2955
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab1_SelectedCancel_All 
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
            Left            =   6630
            Style           =   1  '圖片外觀
            TabIndex        =   65
            Top             =   2955
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton cmd_Tab1_Remove 
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
            Left            =   6015
            Style           =   1  '圖片外觀
            TabIndex        =   64
            Top             =   2955
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab1_srcOrderQuery 
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
            Left            =   9525
            Style           =   1  '圖片外觀
            TabIndex        =   63
            Top             =   2955
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab1_srcOrderReset 
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
            Left            =   10650
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   62
            Top             =   2955
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab1_SelectedCancel 
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
            Left            =   8085
            Style           =   1  '圖片外觀
            TabIndex        =   61
            Top             =   2955
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txt_Tab1_DELIVERY_DATE 
            BackColor       =   &H00E0E0E0&
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
            Left            =   7005
            TabIndex        =   41
            Top             =   150
            Width           =   1110
         End
         Begin VB.TextBox txt_Tab1_VehicleNo 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2880
            TabIndex        =   40
            Top             =   150
            Width           =   1125
         End
         Begin VB.CommandButton cmd_Tab1_SelectCar 
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
            Left            =   3990
            Style           =   1  '圖片外觀
            TabIndex        =   39
            Top             =   150
            Width           =   330
         End
         Begin VB.CommandButton cmd_Tab1_ImportOrders 
            BackColor       =   &H00C0C0FF&
            Caption         =   "載入待重組訂單"
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
            TabIndex        =   38
            Top             =   105
            Width           =   1815
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
            Index           =   1
            Left            =   9960
            Style           =   1  '圖片外觀
            TabIndex        =   37
            Top             =   120
            Width           =   1110
         End
         Begin MSDataGridLib.DataGrid dg_Tab1_SelectedOrders 
            Height          =   2235
            Left            =   0
            TabIndex        =   42
            Top             =   600
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   3942
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
            BackStyle       =   0  '透明
            Caption         =   "司機"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   10
            Left            =   4440
            TabIndex        =   75
            Top             =   240
            Width           =   900
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '不透明
            Height          =   435
            Left            =   5610
            Top             =   2925
            Width           =   795
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00400040&
            FillStyle       =   0  '實心
            Height          =   435
            Index           =   0
            Left            =   6600
            Top             =   2925
            Visible         =   0   'False
            Width           =   2790
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
            Height          =   195
            Index           =   12
            Left            =   6120
            TabIndex        =   44
            Top             =   240
            Width           =   915
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
            Height          =   270
            Index           =   11
            Left            =   2040
            TabIndex        =   43
            Top             =   240
            Width           =   1020
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '不透明
            BorderColor     =   &H00FF0000&
            FillStyle       =   0  '實心
            Height          =   435
            Left            =   9495
            Top             =   2925
            Visible         =   0   'False
            Width           =   1680
         End
      End
      Begin VB.CommandButton cmd_Tab1_Query 
         BackColor       =   &H0080FF80&
         Caption         =   "查  詢"
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
         Height          =   360
         Left            =   -73800
         Style           =   1  '圖片外觀
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Tab1_Add 
         BackColor       =   &H00FFFF80&
         Caption         =   "新  增"
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
         Height          =   360
         Left            =   -72600
         Style           =   1  '圖片外觀
         TabIndex        =   34
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Tab1_Reset 
         BackColor       =   &H000080FF&
         Caption         =   "重組"
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
         Height          =   360
         Left            =   -71400
         Style           =   1  '圖片外觀
         TabIndex        =   32
         Top             =   840
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Tab0_Delete 
         BackColor       =   &H0080FFFF&
         Caption         =   "非經輪"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72120
         Style           =   1  '圖片外觀
         TabIndex        =   26
         Top             =   390
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_Tab0_Confirm 
         BackColor       =   &H000080FF&
         Caption         =   "確 認"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         Style           =   1  '圖片外觀
         TabIndex        =   25
         Top             =   390
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '平面
         BackColor       =   &H8000000C&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   -69600
         TabIndex        =   11
         Top             =   3960
         Width           =   8640
         Begin VB.CommandButton cmd_Tab0_Clear1 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   118
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt_Driver1 
            BackColor       =   &H00E0E0E0&
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
            Left            =   3585
            TabIndex        =   28
            Top             =   600
            Width           =   1080
         End
         Begin VB.TextBox txt_DELIVERY_DATE1 
            BackColor       =   &H00E0E0E0&
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
            Left            =   5505
            TabIndex        =   27
            Top             =   600
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab0_RouteSelect1 
            BackColor       =   &H00FF8080&
            Caption         =   "→"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   24
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt_Tab0_C_Route_No1 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1665
            TabIndex        =   19
            Top             =   195
            Width           =   1980
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
            Height          =   360
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   18
            Top             =   195
            Width           =   495
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
            Height          =   360
            Left            =   4200
            Style           =   1  '圖片外觀
            TabIndex        =   17
            Top             =   195
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab0_SelectCar1 
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
            Left            =   2760
            Style           =   1  '圖片外觀
            TabIndex        =   13
            Top             =   615
            Width           =   330
         End
         Begin VB.TextBox txt_VehicleNo1 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   615
            Width           =   1080
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_RouteList1 
            Height          =   1920
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3387
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
            BackStyle       =   0  '透明
            Caption         =   "司機"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   3120
            TabIndex        =   30
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "出車日"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   4800
            TabIndex        =   29
            Top             =   720
            Width           =   660
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
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   16
            Top             =   735
            Width           =   900
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   15
            Top             =   315
            Width           =   840
         End
      End
      Begin VB.CommandButton cmd_Tab0_ConsigneeShow 
         BackColor       =   &H00FFC0C0&
         Caption         =   "顯示未確認路編"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74805
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   360
         Width           =   1830
      End
      Begin VB.Frame fam_Tab0_Consignee 
         Appearance      =   0  '平面
         BackColor       =   &H8000000C&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   -69600
         TabIndex        =   1
         Top             =   840
         Width           =   8640
         Begin VB.CheckBox chkPalletDefend2 
            BackColor       =   &H8000000A&
            Caption         =   "是否維護棧板"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   6600
            MaskColor       =   &H00808080&
            TabIndex        =   136
            Top             =   600
            Value           =   1  '核取
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CommandButton cmd_Tab0_RouteSelect0 
            BackColor       =   &H00FF8080&
            Caption         =   "→"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   135
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab0_Clear0 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   117
            Top             =   600
            Width           =   495
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
            Height          =   360
            Index           =   0
            Left            =   6120
            Style           =   1  '圖片外觀
            TabIndex        =   33
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Del 
            BackColor       =   &H0080FFFF&
            Caption         =   "打散重組"
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
            Height          =   360
            Left            =   4920
            Style           =   1  '圖片外觀
            TabIndex        =   31
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox txt_DELIVERY_DATE0 
            BackColor       =   &H00E0E0E0&
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
            Left            =   5505
            TabIndex        =   22
            Top             =   600
            Width           =   1080
         End
         Begin VB.TextBox txt_Driver0 
            BackColor       =   &H00E0E0E0&
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
            Left            =   3585
            TabIndex        =   20
            Top             =   600
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab0_SelectCar0 
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
            Left            =   2760
            Style           =   1  '圖片外觀
            TabIndex        =   9
            Top             =   615
            Width           =   330
         End
         Begin VB.TextBox txt_VehicleNo0 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab0_OK 
            BackColor       =   &H000080FF&
            Caption         =   "確  認"
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
            Height          =   360
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   6
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab0_C_Route_No0 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1665
            TabIndex        =   2
            Top             =   195
            Width           =   1980
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_RouteList0 
            Height          =   1800
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3175
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
            BackStyle       =   0  '透明
            Caption         =   "出車日"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   4800
            TabIndex        =   23
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '透明
            Caption         =   "司機"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   3120
            TabIndex        =   21
            Top             =   720
            Width           =   900
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
            Height          =   270
            Index           =   13
            Left            =   720
            TabIndex        =   10
            Top             =   735
            Width           =   900
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   16
            Left            =   720
            TabIndex        =   3
            Top             =   315
            Width           =   840
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_C_RouteList 
         Height          =   6120
         Left            =   -74820
         TabIndex        =   5
         Top             =   840
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   10795
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
      Begin MSDataGridLib.DataGrid dg_Tab2_RouteOrders 
         Height          =   3240
         Left            =   120
         TabIndex        =   81
         Top             =   3735
         Width           =   11180
         _ExtentX        =   19711
         _ExtentY        =   5715
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
      Begin MSDataGridLib.DataGrid dg_Tab2_Route 
         Height          =   2625
         Left            =   120
         TabIndex        =   82
         Top             =   1080
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   4630
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
      Begin MSDataGridLib.DataGrid dg_Tab3_SDN02W 
         Height          =   1755
         Left            =   -74760
         TabIndex        =   85
         Top             =   720
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   3096
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
   End
End
Attribute VB_Name = "frm_OP_RouteConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_Tab0_C_RouteList As ADODB.Recordset    '未出車確認之路編
Private rs_Tab0_RouteList0 As ADODB.Recordset     'Tab0
Private rs_Tab0_RouteList1 As ADODB.Recordset
Private rs_Tab1_SelectedOrders As ADODB.Recordset
Private rs_Tab2_Route As ADODB.Recordset
Private rs_Tab2_RouteOrders As ADODB.Recordset
Private rs_SDN02W As ADODB.Recordset
Private rs_Tab3_SDN02W As ADODB.Recordset
Private rs_Tab3_SDN03W As ADODB.Recordset
Private str_route As String                      '新增路編
Private dbsrcFormHeight As Double                'Form 設計時期的高
Private dbsrcFormWidth As Double                 'Form 設計時期的寬
Private Tab0_RouteListEventEnable As Boolean     'Tab0是否啟動選取事件
Private Tab1_RouteListEventEnable As Boolean     'Tab1是否啟動選取事件
Private Tab2_RouteListEventEnable As Boolean     'Tab1是否啟動選取事件
Private CutOrderkey As String                    '新切割出來之訂單編號
Private dbCut_TotalCaseQty As Double
Private dbCut_TotalWeight As Double
Private dbCut_TotalVolumn As Double
Private rs_Tab2_RouteEvent As Boolean
Private intColumnIndex As Integer

Private Sub cmd_Exit_Click(Index As Integer)
    '離開
    Unload Me
End Sub

Private Sub cmd_Tab0_Clear0_Click()
    Call clear_Tab0_RouteList0
End Sub

Private Sub cmd_Tab0_Clear1_Click()
    Call clear_Tab0_RouteList1
End Sub

Private Sub cmd_Tab0_Confirm_Click()
    Dim Str_Receiver As String
    '出車確認 -確認路編
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    
    'Terry 20180515 新增是否維護棧板按鈕
    Dim str_PalletDefend As String
    str_PalletDefend = ""
    If chkPalletDefend1.Value = vbChecked Then
        str_PalletDefend = "Y"
    Else
        str_PalletDefend = "N"
    End If
    
    
    
    Tab0_RouteListEventEnable = False
    Call WriteOut_RunLog("1.開始 >> 存入SDN01T;SDN02T;SDN03T")
    
    rs_Tab0_C_RouteList.Filter = "＊='V'"
    If Not rs_Tab0_C_RouteList.EOF Then

        Do While Not rs_Tab0_C_RouteList.EOF
        Tran_Level = cn.BeginTrans
        '此路編是否已確認
        Str_Receiver = ""
        Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)
        str_SQL = "Select t05t.Route_No as 路線編號 From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & rs_Tab0_C_RouteList("路線編號") & "' Union All Select t05t.Route_No as 路線編號 From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & rs_Tab0_C_RouteList("路線編號") & "' "
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        If Not tmp_Rs.EOF Then MsgBox "路線編號 " & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & " 已作過出車確認!", 16, "注意": tmp_Rs.Close: GoTo nextROUTE
        tmp_Rs.Close
        
        '抓取請款人
        Call ReDim_Recordset(tmp_Rs)
        str_SQL = "select 請款人=isnull(receiver,driver) from trp09m(nolock) where vehicle_id_no = '" & rs_Tab0_C_RouteList.Fields("車牌號碼").Value & "'"
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        If tmp_Rs.EOF Then Str_Receiver = "" Else Str_Receiver = RTrim(tmp_Rs.Fields("請款人"))
        tmp_Rs.Close
        
            '聯合利華過來的時候，就是退貨訂單，現在已經有A2B
            If Trim(rs_Tab0_C_RouteList.Fields("類別").Value) = "退貨訂單" Then
'                DoEvents: DoEvents
                Call WriteOut_RunLog("確認路編: " & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "")
                
                '存表頭,SDN01T
                str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,receiver,PalletDefend) " & _
                        "Values ( '" & Trim(rs_Tab0_C_RouteList.Fields("出車日期").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("車牌號碼").Value) & "', " & _
                        "'" & Trim(rs_Tab0_C_RouteList.Fields("駕駛人").Value) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("車次").Value) & "','" & Str_Receiver & "','" & str_PalletDefend & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                                              
                '存訂單,SDN02T
                str_SQL = "INSERT dbo.SDN02T(C_ROUTE_NO,ROUTE_NO,STORERKEY,EXTERN,RECEIPT_DATE,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,OnTimeDelivery,PODOnTime,RejectOrder,DESCRIPTION,CONFIRM_DATE,CONSIGNEEKEY,CONFIRM_USERID,CUSTSIGNDATE,RBCCode,RSCCode,CONFIRM_Notes,PRIORITY,SCHEDULEDATE,CustomerOrderkey1,Scan,SDNSendDate,CUST_Handle,TRP_Handle,Advance,INV_Handle,TRP_Cost,Sorting_Cost,Total_Cost,VEHICLE_ID_NO,ExpectReceiptOK,SdnFeedBack,InvBack,C_RECEIPT_NO) " & _
                        "SELECT ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO ,t2.STORERKEY, t2.EXTERN , CONVERT(varchar(8),t2.RECEIPT_DATE, 112) AS RECEIPT_DATE,CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                        "SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.ORDER_QTY / sp.Casecnt end, 3), 0)) AS SHIP_CS, " & _
                        "SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDCUBE, 3), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDGrossWGT, 3), 0)) AS SHIP_WT, " & _
                        "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO,0,0,0,isnull(t2.description,''),null,t2.consigneekey,'',null,'','','',t2.priority,null,'','N',null,'','','','',0,0,0,t2.vehicle_id_no ,'N',0,'N',t2.receipt_no " & _
                        "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                        "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                        "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                        "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                        "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' " & _
                        "GROUP BY t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO,t2.CONSIGNEEKEY,t2.PRIORITY,t2.VEHICLE_ID_NO,t2.STORERKEY,t2.description,CONVERT(varchar(8),t2.RECEIPT_DATE, 112),t2.receipt_no"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '存明細,SDN03T
                str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                        "select  '" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ORDER_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                        "from ORT03t where route_no in( " & _
                        "select  route_no from ORT01t where  isnull(c_route_no,route_no)='" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' and left(route_no,1) <>'S')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                'ORT05T
                str_SQL = "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Else
'                DoEvents: DoEvents
                Call WriteOut_RunLog("確認路編: " & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "")
                
                '存表頭,SDN01T
                str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,receiver,PalletDefend) " & _
                        "Values ( '" & Trim(rs_Tab0_C_RouteList.Fields("出車日期").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("車牌號碼").Value) & "', " & _
                        "'" & Trim(rs_Tab0_C_RouteList.Fields("駕駛人").Value) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("車次").Value) & "','" & Str_Receiver & "', '" & str_PalletDefend & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

                '存訂單,SDN02T
                str_SQL = "INSERT dbo.SDN02T(C_ROUTE_NO,ROUTE_NO,STORERKEY,EXTERN,RECEIPT_DATE,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,OnTimeDelivery,PODOnTime,RejectOrder,DESCRIPTION,CONFIRM_DATE,CONSIGNEEKEY,CONFIRM_USERID,CUSTSIGNDATE,RBCCode,RSCCode,CONFIRM_Notes,PRIORITY,SCHEDULEDATE,CustomerOrderkey1,Scan,SDNSendDate,CUST_Handle,TRP_Handle,Advance,INV_Handle,TRP_Cost,Sorting_Cost,Total_Cost,VEHICLE_ID_NO,ExpectReceiptOK,SdnFeedBack,InvBack,C_RECEIPT_NO) " & _
                        "SELECT  ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO ,t2.STORERKEY , t2.EXTERN ,CONVERT(varchar(8),t2.RECEIPT_DATE, 112) AS RECEIPT_DATE, CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                        "SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.SHIP_QTY / sp.Casecnt end, 2), 0)) AS SHIP_CS, " & _
                        "SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.stdcube, 2), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDGrossWGT, 2), 0)) AS SHIP_WT, " & _
                        "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO,0,0,0,t2.description " & _
                        ",null,t2.consigneekey,'',null,'','','',t2.priority,t2.scheduledate,t2.customerorderkey1,'N',null,'','','','',0,0,0,t2.vehicle_id_no,'N',0,'N',t2.c_receipt_no " & _
                        "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                        "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                        "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                        "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                        "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' " & _
                        "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO,t2.CONSIGNEEKEY,t2.PRIORITY,t2.SCHEDULEDATE,t2.VEHICLE_ID_NO,t2.CustomerOrderkey1,t2.STORERKEY,t2.description,CONVERT(varchar(8),t2.RECEIPT_DATE, 112),t2.c_receipt_no "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '存明細,SDN03T
                str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                        "select  '" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ship_qty,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                        "from trp03t where  route_no in( " & _
                        "select  route_no from trp01t where  isnull(c_route_no,route_no)='" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' and left(route_no,1) <>'S')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '更新TRP05T狀態
                str_SQL = "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '更新SDN01T狀態
                str_SQL = "Update SDN01T set SDN01T.APPStatus = TRP05T.APPStatus ,SDN01T.VLListCount = TRP05T.VLListCount ,SDN01T.VLListPrintDate = TRP05T.VLListPrintDate from trp05t join sdn01t on isnull(trp05t.c_route_no,trp05t.route_no) = sdn01t.c_route_no and sdn01t.c_Route_No= '" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '更新APPOrderDate狀態
                str_SQL = "update AppOrderDate set status = '5' where status < '6' and c_Route_No= '" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "' "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                               
            End If
            
                '更新OrderType Add by Gemini @20190604
                str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & Trim(rs_Tab0_C_RouteList.Fields("路線編號")) & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & Trim(rs_Tab0_C_RouteList.Fields("路線編號")) & "' "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '更新請款人 edit by Eric 先抓出請款人，在insert就補進去，避免更新兩次
                'cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & rs_Tab0_C_RouteList.Fields("車牌號碼").Value & "') where c_route_no = '" & rs_Tab0_C_RouteList.Fields("路線編號").Value & "'", RowsAffect, adExecuteNoRecords
            
nextROUTE:
        cn.CommitTrans: Tran_Level = 0
            rs_Tab0_C_RouteList.MoveNext
            
        Loop

        
        '[刪除已選取之訂單
        Call WriteOut_RunLog("3.刪除已選取之路線編號")
'        DoEvents: DoEvents
        rs_Tab0_C_RouteList.MoveFirst
        Do While Not rs_Tab0_C_RouteList.EOF
            rs_Tab0_C_RouteList.Delete
            rs_Tab0_C_RouteList.MoveFirst
        Loop
    End If
  
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    Call WriteOut_RunLog("4.確認完成")
'    DoEvents: DoEvents
    Call Unload_RunLogForm
    Call ReSet_Tab0_C_RouteList_SeqNo
    Call clear_Tab0_RouteList0
    Call clear_Tab0_RouteList1
    Tab0_RouteListEventEnable = True

    Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
    End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "出車確認-確認路編", Me.Caption, "cmd_Tab0_Confirm_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_ConsigneeShow_Click()
    '出車確認-匯入未確認路編
    Set dg_Tab0_C_RouteList.DataSource = Nothing
    Set rs_Tab0_C_RouteList = Nothing
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    str_SQL = "Select ' ' as '＊',t01t.Route_No as 路線編號 , Convert(varchar(8),t01t.Delivery_Date,112) as 出車日期 ,  " & _
            "Rtrim(t05t.Vehicle_ID_No) as 車牌號碼 , t05t.Drive_Times as 車次 , Rtrim(t05t.Driver) as 駕駛人 , Rtrim(Isnull(t08m.SHORT_NAME,'')) as 運輸公司,'正常訂單' as 類別 " & _
            "From TRP01T t01t " & _
            "inner join TRP05T t05t on t05t.Route_No = t01t.Route_No " & _
            "left join TRP09M t09m on t09m.Vehicle_ID_No = t05t.Vehicle_ID_No " & _
            "Left outer join TRP08M t08m on t08m.Company_Code = t09m.TRP_Company_Code " & _
            "Where t01t.Route_No <> 'D' and  t05t.SDNStatus='0'  and t01t.C_ROUTE_NO is null " & _
            "Union All " & _
            "Select ' ' as '＊',t01t.Route_No as 路線編號 , Convert(varchar(8),t01t.Delivery_Date,112) as 出車日期 , " & _
            "Rtrim(t05t.Vehicle_ID_No) as 車牌號碼 , t05t.Drive_Times as 車次 , Rtrim(t05t.Driver) as 駕駛人 , Rtrim(Isnull(t08m.SHORT_NAME,'')) as 運輸公司,'退貨訂單' as 類別 " & _
            "From ORT01T t01t " & _
            "inner join ORT05T t05t on t05t.Route_No = t01t.Route_No " & _
            "left join TRP09M t09m on t09m.Vehicle_ID_No = t05t.Vehicle_ID_No " & _
            "Left outer join TRP08M t08m on t08m.Company_Code = t09m.TRP_Company_Code " & _
            "Where t01t.Route_No <> 'D' and  t05t.SDNStatus='0'  and t01t.C_ROUTE_NO is null order by t01t.Route_No "

    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無符合搜尋條件之排車資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab0_C_RouteList)
    tmp_Rs.Close
    
    With dg_Tab0_C_RouteList
         .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
         .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
         .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
         .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
    End With
    rs_Tab0_C_RouteList.MoveFirst
    Set dg_Tab0_C_RouteList.DataSource = rs_Tab0_C_RouteList
    With dg_Tab0_C_RouteList
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 400       '序號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 300       '序號
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000      '路線編號
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 900      '出車日期
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 900      '車牌號碼
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 400       '車次
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 1000      '駕駛人
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1200       '運輸簡稱
        .Columns(7).Alignment = dbgLeft
        .Columns(8).Width = 1200       '報到時間
        .Columns(8).Alignment = dbgLeft
    End With
    rs_Tab0_C_RouteList.MoveFirst
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = " 編號 "
    rs_Tab0_C_RouteList.MoveFirst
    blVLLReportEventEnable = True
    Tab0_RouteListEventEnable = True
    Screen.MousePointer = vbDefault
    Call cmd_Tab0_Clear0_Click
    Call cmd_Tab0_Clear1_Click
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-匯入未確認路編", Me.Caption, "cmd_Tab0_ConsigneeShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Del_Click()
    '出車確認-打散重組
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If Len(Trim(txt_Tab0_C_Route_No0.Text)) = 0 Then Exit Sub
    
    '此路編是否已確認
    Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)

    str_SQL = "Select t05t.Route_No as 路線編號 From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0 & "' Union All Select t05t.Route_No as 路線編號 From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0 & "' "
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then MsgBox "路線編號 " & txt_Tab0_C_Route_No0 & " 已作過出車確認!", 16, "注意": tmp_Rs.Close: cmd_Tab0_Del.Enabled = False: Exit Sub
     
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    Call WriteOut_RunLog("1.開始 >> 存入SDN02W")
    
    rs_Tab0_C_RouteList.Filter = "路線編號='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
    If Not rs_Tab0_C_RouteList.EOF Then
        Tran_Level = cn.BeginTrans
        If Trim(rs_Tab0_C_RouteList.Fields("類別").Value) = "退貨訂單" Then
            '存訂單,SDN02W
            str_SQL = "INSERT dbo.SDN02W (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,C_RECEIPT_NO)" & _
                    "SELECT ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO , t2.EXTERN , CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                    "SUM(ISNULL(ROUND(case when s1.Casecnt = 0 then 0 else t3.SHIP_QTY / s1.Casecnt end , 2), 0)) AS SHIP_CS, " & _
                    "SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDCUBE, 2), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDGrossWGT, 2), 0)) AS SHIP_WT, " & _
                    "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO ,t2.receipt_no " & _
                    "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                    "INNER JOIN gv_SKUxpack s1 ON s1.Sku = t3.PRODUCT_NO and s1.storerkey = t3.storerkey " & _
                    "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                    "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                    "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and t2.RECEIPT_NO not in (select receipt_no from sdn02w ) " & _
                    "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '存明細,SDN03W
            str_SQL = "Insert into SDN03W (ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME) " & _
                    "select  ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME " & _
                    "from ORT03t where route_no in( " & _
                    "select  route_no from ORT01t where  isnull(c_route_no,route_no)='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and left(route_no,1) <>'S')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新ORT05T狀態
            str_SQL = "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            rs_Tab0_C_RouteList.MoveNext
        Else
            '存訂單,SDN02W
            str_SQL = "INSERT dbo.SDN02W (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,C_RECEIPT_NO) " & _
                    "SELECT  ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO , t2.EXTERN , CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                    "SUM(ISNULL(ROUND(case when s1.Casecnt = 0 then 0 else t3.SHIP_QTY / s1.Casecnt end , 2), 0)) AS SHIP_CS, " & _
                    "SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDCUBE, 2), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDGrossWGT, 2), 0)) AS SHIP_WT, " & _
                    "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO ,t2.c_receipt_no " & _
                    "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                    "INNER JOIN gv_SKUxpack s1 ON s1.Sku = t3.PRODUCT_NO and s1.storerkey = t3.storerkey " & _
                    "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                    "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                    "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and t2.RECEIPT_NO not in (select receipt_no from sdn02w ) " & _
                    "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO,t2.c_receipt_no"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '存明細,SDN03W
            str_SQL = "Insert into SDN03W (ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME) " & _
                    "select  ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME " & _
                    "from trp03t where  route_no in( " & _
                    "select  route_no from trp01t where  isnull(c_route_no,route_no)='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and left(route_no,1) <>'S')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新TRP05T狀態
            str_SQL = "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            rs_Tab0_C_RouteList.MoveNext
        End If
        
        cn.CommitTrans: Tran_Level = 0
        
        '[刪除已選取之訂單
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.MoveFirst
        Do While Not rs_Tab0_C_RouteList.EOF
            rs_Tab0_C_RouteList.Delete
            rs_Tab0_C_RouteList.MoveFirst
        Loop
    End If
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    Call WriteOut_RunLog("4.匯入完成")
    DoEvents: DoEvents
    Call Unload_RunLogForm
    '畫面處理
    Call ReSet_Tab0_C_RouteList_SeqNo
    Set dg_Tab0_RouteList0.DataSource = Nothing
    txt_Tab0_C_Route_No0.Text = ""
    txt_DELIVERY_DATE0.Text = ""
    txt_VehicleNo0.Text = ""
    txt_Driver0.Text = ""
    cmd_Tab0_OK.Enabled = False
    cmd_Tab0_Del.Enabled = False
    Tab0_RouteListEventEnable = True
    Call cmd_Tab0_Clear1_Click
    Exit Sub
    
err_Handle:
   If Tran_Level <> 0 Then cn.RollbackTrans
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-打散重組", Me.Caption, "cmd_Tab0_Del_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Delete_Click()
    '非經綸
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    rs_Tab0_C_RouteList.Filter = "＊='V'"
    If Not rs_Tab0_C_RouteList.EOF Then
        Do While Not rs_Tab0_C_RouteList.EOF
        
            '更新TRP05T狀態
            str_SQL = "Update TRP05T set SDNStatus = '1' ,Receiver ='非經綸' where Route_No='" & Trim(rs_Tab0_C_RouteList.Fields("路線編號").Value) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            rs_Tab0_C_RouteList.MoveNext
        Loop
        '[刪除已選取之訂單
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.MoveFirst
        Do While Not rs_Tab0_C_RouteList.EOF
            rs_Tab0_C_RouteList.Delete
            rs_Tab0_C_RouteList.MoveFirst
        Loop
    End If
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    DoEvents: DoEvents
    Call Unload_RunLogForm
    Call ReSet_Tab0_C_RouteList_SeqNo
    Call clear_Tab0_RouteList0
    Call clear_Tab0_RouteList1
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-非經綸", Me.Caption, "cmd_Tab0_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmd_Tab0_OK_Click()
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
        
    '此路編是否已確認
    Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)

    str_SQL = "Select t05t.Route_No as 路線編號 From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0.Text & "' Union All Select t05t.Route_No as 路線編號 From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0.Text & "' "
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then MsgBox "路線編號 " & txt_Tab0_C_Route_No0.Text & " 已作過出車確認!", 16, "注意": tmp_Rs.Close: cmd_Tab0_OK.Enabled = False: Exit Sub
    
    '此路編是否已確認txt_Tab0_C_Route_No1.Text
    Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)

    str_SQL = "Select t05t.Route_No as 路線編號 From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No1 & "' Union All Select t05t.Route_No as 路線編號 From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No1 & "' "
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then MsgBox "路線編號 " & txt_Tab0_C_Route_No1 & " 已作過出車確認!", 16, "注意": tmp_Rs.Close: cmd_Tab0_OK.Enabled = False: Exit Sub
    
    On Error GoTo err_Handle
    
    'Terry 20180515 新增是否維護棧板按鈕
    Dim str_PalletDefend As String
    str_PalletDefend = ""
    If chkPalletDefend2.Value = vbChecked Then
        str_PalletDefend = "Y"
    Else
        str_PalletDefend = "N"
    End If
    
    
    
    Tab0_RouteListEventEnable = False
    If Len(Trim(txt_Tab0_C_Route_No0.Text)) > 0 Then
        
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.Filter = "路線編號='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
        rs_Tab0_C_RouteList.MoveFirst
                
        If Not rs_Tab0_C_RouteList.EOF Then
        
        Tran_Level = cn.BeginTrans
        
        Call WriteOut_RunLog("1.開始 >> 存入SDN01T;SDN02T")
            If Trim(rs_Tab0_C_RouteList.Fields("類別").Value) = "退貨訂單" Then     '處理退貨訂單
                If rs_Tab0_RouteList0.RecordCount > 0 Then
                    '存表頭,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE0.Text) & "','" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(txt_VehicleNo0.Text) & "', " & _
                            "'" & Trim(txt_Driver0.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("車次").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    rs_Tab0_RouteList0.MoveFirst
                    Do While Not rs_Tab0_RouteList0.EOF
                        '存訂單,SDN02T
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(rs_Tab0_RouteList0.Fields("路線編號").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("客戶單號").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("日期").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("指送客戶").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("箱數").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("材積").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("重量").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("多車").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '補所需資料 by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = ort02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(char(8),ort02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = ort02t.consigneekey " & _
                                    ",sdn02t.description = ort02t.description " & _
                                    ",sdn02t.priority = ort02t.priority " & _
                                    ",sdn02t.scheduledate = ort02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = ort02t.vehicle_id_no " & _
                                    ",sdn02t.c_receipt_no = ort02t.c_receipt_no " & _
                                    "from ort02t join sdn02t on ort02t.receipt_no = sdn02t.receipt_no " & _
                                    "where ort02t.receipt_no = '" & Trim(rs_Tab0_RouteList0.Fields("訂單號碼").Value) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '存明細,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No0.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ORDER_QTY,SIGN_QTY, SHIP_TIME ,weight,volumn_weight " & _
                            "from ORT03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList0.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        rs_Tab0_RouteList0.MoveNext
                    Loop
                End If
                              
            Else '一般訂單
                If rs_Tab0_RouteList0.RecordCount > 0 Then
                                  
                    '存表頭,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE0.Text) & "','" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(txt_VehicleNo0.Text) & "', " & _
                            "'" & Trim(txt_Driver0.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("車次").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    rs_Tab0_RouteList0.MoveFirst
                    Do While Not rs_Tab0_RouteList0.EOF
                        '存訂單,SDN02T
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(rs_Tab0_RouteList0.Fields("路線編號").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("客戶單號").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("日期").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("指送客戶").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("箱數").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("材積").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("重量").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("多車").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '補所需資料 by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = trp02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(varchar(8),trp02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = trp02t.consigneekey " & _
                                    ",sdn02t.description = trp02t.description " & _
                                    ",sdn02t.priority = trp02t.priority " & _
                                    ",sdn02t.scheduledate = trp02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = trp02t.vehicle_id_no " & _
                                    ",sdn02t.c_receipt_no = trp02t.c_receipt_no " & _
                                    "from trp02t join sdn02t on sdn02t.receipt_no = trp02t.receipt_no " & _
                                    "where sdn02t.receipt_no = '" & Trim(rs_Tab0_RouteList0.Fields("訂單號碼").Value) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '存明細,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No0.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY, ship_qty,SIGN_QTY, SHIP_TIME ,weight,volumn_weight " & _
                            "from trp03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList0.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        rs_Tab0_RouteList0.MoveNext
                    Loop
                    
                End If
                
            End If
            
            '更新TRP05T狀態
            cn.Execute "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'", RowsAffect, adExecuteNoRecords
                    
            '更新ORT05T狀態
            cn.Execute "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'", RowsAffect, adExecuteNoRecords
                    
            '更新SDN01T狀態
            str_SQL = "Update SDN01T set SDN01T.APPStatus = TRP05T.APPStatus ,SDN01T.VLListCount = TRP05T.VLListCount ,SDN01T.VLListPrintDate = TRP05T.VLListPrintDate from trp05t join sdn01t on isnull(trp05t.c_route_no,trp05t.route_no) = sdn01t.c_route_no and sdn01t.c_Route_No= '" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新APPOrderDate狀態
            str_SQL = "update AppOrderDate set status = '5' where status < '6' and c_Route_No= '" & Trim(txt_Tab0_C_Route_No0.Text) & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新請款人
            cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & Trim(txt_VehicleNo0.Text) & "') where c_route_no = '" & txt_Tab0_C_Route_No0.Text & "'", RowsAffect, adExecuteNoRecords
                        
            '更新OrderType Add by Gemini @20190604
            str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & txt_Tab0_C_Route_No0.Text & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & txt_Tab0_C_Route_No0.Text & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
            cn.CommitTrans: Tran_Level = 0
            
            '[刪除已選取之訂單
            DoEvents: DoEvents
            rs_Tab0_C_RouteList.MoveFirst
            Do While Not rs_Tab0_C_RouteList.EOF
                rs_Tab0_C_RouteList.Delete
                rs_Tab0_C_RouteList.MoveFirst
            Loop
        End If
        
        '畫面處理
        Call clear_Tab0_RouteList0
        rs_Tab0_C_RouteList.Filter = adFilterNone
        rs_Tab0_C_RouteList.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    End If
    
    If Len(Trim(txt_Tab0_C_Route_No1.Text)) > 0 Then
        Call WriteOut_RunLog("1.1.開始 >> 存入SDN01T")
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.Filter = "路線編號='" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
        rs_Tab0_C_RouteList.MoveFirst
        If Not rs_Tab0_C_RouteList.EOF Then
            cn.BeginTrans
            If Trim(rs_Tab0_C_RouteList.Fields("類別").Value) = "退貨訂單" Then     '處理退貨訂單
                If rs_Tab0_RouteList1.RecordCount > 0 Then
                    '存表頭,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE1.Text) & "','" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(txt_VehicleNo1.Text) & "', " & _
                            "'" & Trim(txt_Driver1.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("車次").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    '存訂單,SDN02T
                    rs_Tab0_RouteList1.MoveFirst
                    Do While Not rs_Tab0_RouteList1.EOF
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(rs_Tab0_RouteList1.Fields("路線編號").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("客戶單號").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("日期").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("指送客戶").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("箱數").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("材積").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("重量").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("多車").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '補所需資料 by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = ort02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(varchar(8),ort02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = ort02t.consigneekey " & _
                                    ",sdn02t.description = ort02t.description " & _
                                    ",sdn02t.priority = ort02t.priority " & _
                                    ",sdn02t.scheduledate = ort02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = '" & Trim(txt_VehicleNo1.Text) & "' " & _
                                    ",sdn02t.c_receipt_no = ort02t.c_receipt_no " & _
                                    "from ort02t join sdn02t on sdn02t.receipt_no = ort02t.receipt_no " & _
                                    "where ort02t.receipt_no = '" & Trim(rs_Tab0_RouteList1.Fields("訂單號碼").Value) & "'"
                         cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                         
                        '存明細,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No1.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY, ship_qty,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                            "from ORT03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList1.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                                              
                       '補所需資料 by gemini
                       str_SQL = "Update sdn03t Set sdn03t.Weight = ort03t.Weight ,sdn03t.volumn_weight = ort03t.volumn_weight " & _
                            "from ort03t join sdn03t on sdn03t.receipt_no = ort03t.receipt_no and sdn03t.seq_no = ort03t.seq_no and isnull(sdn03t.subseq_no,'') = isnull(ort03t.subseq_no,'') " & _
                            "where sdn03t.receipt_no = '" & Trim(rs_Tab0_RouteList1.Fields("訂單號碼").Value) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                       rs_Tab0_RouteList1.MoveNext
                    Loop
                End If
            
            Else
                If rs_Tab0_RouteList1.RecordCount > 0 Then
                    '存表頭,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE1.Text) & "','" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(txt_VehicleNo1.Text) & "', " & _
                            "'" & Trim(txt_Driver1.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("車次").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    '存訂單,SDN02T
                    rs_Tab0_RouteList1.MoveFirst
                    Do While Not rs_Tab0_RouteList1.EOF
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(rs_Tab0_RouteList1.Fields("路線編號").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("客戶單號").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("日期").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("指送客戶").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("箱數").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("材積").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("重量").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("多車").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '補所需資料 by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = trp02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(varchar(8),trp02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = trp02t.consigneekey " & _
                                    ",sdn02t.description = trp02t.description " & _
                                    ",sdn02t.priority = trp02t.priority " & _
                                    ",sdn02t.scheduledate = trp02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = '" & Trim(txt_VehicleNo1.Text) & "' " & _
                                    ",sdn02t.c_receipt_no = trp02t.c_receipt_no " & _
                                    "from trp02t join sdn02t on sdn02t.receipt_no = trp02t.receipt_no " & _
                                    "where sdn02t.receipt_no = '" & Trim(rs_Tab0_RouteList1.Fields("訂單號碼").Value) & "'"
                         cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '存明細,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No1.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                            "from trp03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList1.Fields("訂單號碼").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        rs_Tab0_RouteList1.MoveNext
                    Loop
                End If
                
            End If
            
            '更新TRP05T狀態
            str_SQL = "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新ORT05T狀態
            str_SQL = "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新SDN01T狀態
            str_SQL = "Update SDN01T set SDN01T.APPStatus = TRP05T.APPStatus ,SDN01T.VLListCount = TRP05T.VLListCount ,SDN01T.VLListPrintDate = TRP05T.VLListPrintDate from trp05t join sdn01t on isnull(trp05t.c_route_no,trp05t.route_no) = sdn01t.c_route_no and sdn01t.c_Route_No= '" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新APPOrderDate狀態
            str_SQL = "update AppOrderDate set status = '5' where status < '6' and c_Route_No= '" & Trim(txt_Tab0_C_Route_No1.Text) & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新請款人
            cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & Trim(txt_VehicleNo1.Text) & "') where c_route_no = '" & txt_Tab0_C_Route_No1.Text & "'", RowsAffect, adExecuteNoRecords
            
            '更新OrderType Add by Gemini @20190604
            str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & txt_Tab0_C_Route_No1.Text & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & txt_Tab0_C_Route_No1.Text & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            cn.CommitTrans: Tran_Level = 0
            
            '[刪除已選取之訂單
            DoEvents: DoEvents
            rs_Tab0_C_RouteList.MoveFirst
            Do While Not rs_Tab0_C_RouteList.EOF
                rs_Tab0_C_RouteList.Delete
                rs_Tab0_C_RouteList.MoveFirst
            Loop
        End If
        
        '畫面處理
        Call clear_Tab0_RouteList1
        rs_Tab0_C_RouteList.Filter = adFilterNone
        rs_Tab0_C_RouteList.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    End If
    
    Call WriteOut_RunLog("4.匯入完成")
    DoEvents: DoEvents
    Call Unload_RunLogForm
    
    '畫面處理
    Call ReSet_Tab0_C_RouteList_SeqNo
    cmd_Tab0_OK.Enabled = False
    cmd_Tab0_Del.Enabled = False
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   If Tran_Level <> 0 Then cn.RollbackTrans
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-確認", Me.Caption, "cmd_Tab0_Del_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Remove_Click()
    '篩選已選取者
    
    '排除退貨訂單 Gemini @ 20060728
    If Left(txt_Tab0_C_Route_No0.Text, 1) = "R" Or Left(txt_Tab0_C_Route_No1.Text, 1) = "R" Then If Left(txt_Tab0_C_Route_No0.Text, 1) <> Left(txt_Tab0_C_Route_No1.Text, 1) Then MsgBox "退貨路編無法併入其他路編！", vbOKOnly, Me.Caption: Exit Sub
    
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList1.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    rs_Tab0_RouteList0.Filter = "＊='V'"
    If Not rs_Tab0_RouteList0.EOF Then
       Do While Not rs_Tab0_RouteList0.EOF
          '判斷是否已經選取過
          rs_Tab0_RouteList1.Filter = adFilterNone
          rs_Tab0_RouteList1.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
          rs_Tab0_RouteList1.Filter = "訂單號碼 = '" & rs_Tab0_RouteList0.Fields("訂單號碼").Value & "'"
          '如果是查詢所顯示之有效路編，設定路編異動識別旗標
          If blRouteModify Then blRouteChange = True
          If rs_Tab0_RouteList1.EOF Then
             '新增選取之訂單
             rs_Tab0_RouteList1.AddNew
             rs_Tab0_RouteList1.Fields("編號").Value = 999
             rs_Tab0_RouteList1.Fields("二次排車").Value = rs_Tab0_RouteList0.Fields("二次排車").Value
             rs_Tab0_RouteList1.Fields("客戶單號").Value = rs_Tab0_RouteList0.Fields("客戶單號").Value
             rs_Tab0_RouteList1.Fields("路線編號").Value = rs_Tab0_RouteList0.Fields("路線編號").Value
             rs_Tab0_RouteList1.Fields("日期").Value = rs_Tab0_RouteList0.Fields("日期").Value
             rs_Tab0_RouteList1.Fields("指送客戶").Value = rs_Tab0_RouteList0.Fields("指送客戶").Value
             rs_Tab0_RouteList1.Fields("箱數").Value = rs_Tab0_RouteList0.Fields("箱數").Value
             rs_Tab0_RouteList1.Fields("重量").Value = rs_Tab0_RouteList0.Fields("重量").Value
             rs_Tab0_RouteList1.Fields("材積").Value = rs_Tab0_RouteList0.Fields("材積").Value
             rs_Tab0_RouteList1.Fields("多車").Value = rs_Tab0_RouteList0.Fields("多車").Value
             rs_Tab0_RouteList1.Fields("訂單號碼").Value = rs_Tab0_RouteList0.Fields("訂單號碼").Value
             rs_Tab0_RouteList1.Update
          Else
             '更新選取之訂單資料
             rs_Tab0_RouteList1.Fields("二次排車").Value = rs_Tab0_RouteList0.Fields("二次排車").Value
             rs_Tab0_RouteList1.Fields("客戶單號").Value = rs_Tab0_RouteList0.Fields("客戶單號").Value
             rs_Tab0_RouteList1.Fields("路線編號").Value = rs_Tab0_RouteList0.Fields("路線編號").Value
             rs_Tab0_RouteList1.Fields("日期").Value = rs_Tab0_RouteList0.Fields("日期").Value
             rs_Tab0_RouteList1.Fields("指送客戶").Value = rs_Tab0_RouteList0.Fields("指送客戶").Value
             rs_Tab0_RouteList1.Fields("箱數").Value = rs_Tab0_RouteList0.Fields("箱數").Value
             rs_Tab0_RouteList1.Fields("重量").Value = rs_Tab0_RouteList0.Fields("重量").Value
             rs_Tab0_RouteList1.Fields("材積").Value = rs_Tab0_RouteList0.Fields("材積").Value
             rs_Tab0_RouteList1.Fields("多車").Value = rs_Tab0_RouteList0.Fields("多車").Value
             rs_Tab0_RouteList1.Fields("訂單號碼").Value = rs_Tab0_RouteList0.Fields("訂單號碼").Value
          End If
          rs_Tab0_RouteList0.MoveNext
       Loop
       
       '[刪除已選取之訂單
       rs_Tab0_RouteList0.MoveFirst
       Do While Not rs_Tab0_RouteList0.EOF
          rs_Tab0_RouteList0.Delete
          rs_Tab0_RouteList0.MoveFirst
       Loop
        
    End If
    rs_Tab0_RouteList1.Filter = adFilterNone
    rs_Tab0_RouteList1.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    rs_Tab0_RouteList0.Filter = adFilterNone
    rs_Tab0_RouteList0.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    Call ReSet_Tab0_RouteList1_SeqNo
    Call ReSet_Tab0_RouteList0_SeqNo
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-選取向下", Me.Caption, "cmd_Tab0_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_RouteSelect0_Click()
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If Trim(txt_Tab0_C_Route_No1.Text) = Trim(rs_Tab0_C_RouteList.Fields(2).Value) Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    str_route = Trim(rs_Tab0_C_RouteList.Fields(2).Value)
    txt_Tab0_C_Route_No0.Text = str_route
    txt_DELIVERY_DATE0.Text = Trim(rs_Tab0_C_RouteList.Fields(3).Value)
    txt_VehicleNo0.Text = Trim(rs_Tab0_C_RouteList.Fields(4).Value)
    txt_Driver0.Text = Trim(rs_Tab0_C_RouteList.Fields(6).Value)
    If Trim(rs_Tab0_C_RouteList.Fields("類別").Value) = "退貨訂單" Then
                str_SQL = "SELECT  ' ' as '＊',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS 二次排車, t2.ROUTE_NO AS 路線編號, t2.EXTERN AS 客戶單號, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS 日期, m1.FULL_NAME as 指送客戶,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.ORDER_QTY / sp.Casecnt end , 2), 0)) AS 箱數, " & _
                  "SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDCUBE, 2), 0)) AS 材積,SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDGrossWGT, 2), 0)) AS 重量, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS 多車,t2.RECEIPT_NO as 訂單號碼  " & _
                  "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    Else
        str_SQL = "SELECT  ' ' as '＊',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS 二次排車, t2.ROUTE_NO AS 路線編號, t2.EXTERN AS 客戶單號, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS 日期, m1.FULL_NAME as 指送客戶,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.SHIP_QTY / sp.Casecnt end , 2), 0)) AS 箱數, " & _
                  "SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDCUBE, 2), 0)) AS 材積,SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDGrossWGT, 2), 0)) AS 重量, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS 多車,t2.RECEIPT_NO as 訂單號碼  " & _
                  "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    End If
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
    
    Call Replication_Recordset(tmp_Rs, rs_Tab0_RouteList0)
    Set dg_Tab0_RouteList0.DataSource = rs_Tab0_RouteList0
    tmp_Rs.Close
    With dg_Tab0_RouteList0
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500       '選取
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000      '二次排車
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000      '客戶單號
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800      '路線編號
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 800       '日期
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1200      '指送客戶
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 800       '箱數
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '材積
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 800       '重量
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 800       '多車
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 1000       '訂單號碼
        .Columns(11).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    cmd_Tab0_Delete.Enabled = True
    cmd_Tab0_Del.Enabled = True
    cmd_Tab0_OK.Enabled = True
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-路編選取上", Me.Caption, "cmd_Tab0_RouteSelect0_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_RouteSelect1_Click()
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If Trim(txt_Tab0_C_Route_No0.Text) = Trim(rs_Tab0_C_RouteList.Fields(2).Value) Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    str_route = Trim(rs_Tab0_C_RouteList.Fields(2).Value)
    txt_Tab0_C_Route_No1.Text = str_route
    txt_DELIVERY_DATE1.Text = Trim(rs_Tab0_C_RouteList.Fields(3).Value)
    txt_VehicleNo1.Text = Trim(rs_Tab0_C_RouteList.Fields(4).Value)
    txt_Driver1.Text = Trim(rs_Tab0_C_RouteList.Fields(6).Value)
    If Trim(rs_Tab0_C_RouteList.Fields("類別").Value) = "退貨訂單" Then
                str_SQL = "SELECT  ' ' as '＊',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS 二次排車, t2.ROUTE_NO AS 路線編號, t2.EXTERN AS 客戶單號, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS 日期, m1.FULL_NAME as 指送客戶,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.ORDER_QTY / sp.Casecnt end , 2), 0)) AS 箱數, " & _
                  "SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDCUBE, 2), 0)) AS 材積,SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDGrossWGT, 2), 0)) AS 重量, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS 多車,t2.RECEIPT_NO as 訂單號碼 " & _
                  "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    Else
        str_SQL = "SELECT  ' ' as '＊',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS 二次排車, t2.ROUTE_NO AS 路線編號, t2.EXTERN AS 客戶單號, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS 日期, m1.FULL_NAME as 指送客戶,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.SHIP_QTY / sp.Casecnt end , 2), 0)) AS 箱數, " & _
                  "SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDCUBE, 2), 0)) AS 材積,SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDGrossWGT, 2), 0)) AS 重量, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS 多車,t2.RECEIPT_NO as 訂單號碼  " & _
                  "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    End If
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
    cmd_Tab0_Delete.Enabled = True
    Call Replication_Recordset(tmp_Rs, rs_Tab0_RouteList1)
    Set dg_Tab0_RouteList1.DataSource = rs_Tab0_RouteList1
    tmp_Rs.Close
    With dg_Tab0_RouteList1
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500       '選取
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000      '二次排車
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000      '客戶單號
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800      '路線編號
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 800       '日期
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1200      '指送客戶
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 800       '箱數
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '材積
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 800       '重量
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 800       '多車
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 1000       '訂單號碼
        .Columns(11).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-路編選取下", Me.Caption, "cmd_Tab0_RouteSelect1_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_SelectCar0_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar0.Name & "2")
End Sub

Private Sub cmd_Tab0_SelectCar1_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar1.Name & "2")
End Sub

Private Sub cmd_Tab0_Selected_Click()
    '篩選已選取者
    
    '排除退貨訂單 Gemini @ 20060728
    If Left(txt_Tab0_C_Route_No0.Text, 1) = "R" Or Left(txt_Tab0_C_Route_No1.Text, 1) = "R" Then If Left(txt_Tab0_C_Route_No0.Text, 1) <> Left(txt_Tab0_C_Route_No1.Text, 1) Then MsgBox "退貨路編無法併入其他路編！", vbOKOnly, Me.Caption: Exit Sub
   
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList1.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    rs_Tab0_RouteList1.Filter = "＊='V'"
    If Not rs_Tab0_RouteList1.EOF Then
       Do While Not rs_Tab0_RouteList1.EOF
          '判斷是否已經選取過
          rs_Tab0_RouteList0.Filter = adFilterNone
          rs_Tab0_RouteList0.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
          rs_Tab0_RouteList0.Filter = "訂單號碼 = '" & rs_Tab0_RouteList1.Fields("訂單號碼").Value & "'"
          '如果是查詢所顯示之有效路編，設定路編異動識別旗標
          If blRouteModify Then blRouteChange = True
          If rs_Tab0_RouteList0.EOF Then
             '新增選取之訂單
             rs_Tab0_RouteList0.AddNew
             rs_Tab0_RouteList0.Fields("編號").Value = 999
             rs_Tab0_RouteList0.Fields("二次排車").Value = rs_Tab0_RouteList1.Fields("二次排車").Value
             rs_Tab0_RouteList0.Fields("客戶單號").Value = rs_Tab0_RouteList1.Fields("客戶單號").Value
             rs_Tab0_RouteList0.Fields("路線編號").Value = rs_Tab0_RouteList1.Fields("路線編號").Value
             rs_Tab0_RouteList0.Fields("日期").Value = rs_Tab0_RouteList1.Fields("日期").Value
             rs_Tab0_RouteList0.Fields("指送客戶").Value = rs_Tab0_RouteList1.Fields("指送客戶").Value
             rs_Tab0_RouteList0.Fields("箱數").Value = rs_Tab0_RouteList1.Fields("箱數").Value
             rs_Tab0_RouteList0.Fields("重量").Value = rs_Tab0_RouteList1.Fields("重量").Value
             rs_Tab0_RouteList0.Fields("材積").Value = rs_Tab0_RouteList1.Fields("材積").Value
             rs_Tab0_RouteList0.Fields("多車").Value = rs_Tab0_RouteList1.Fields("多車").Value
             rs_Tab0_RouteList0.Fields("訂單號碼").Value = rs_Tab0_RouteList1.Fields("訂單號碼").Value
             rs_Tab0_RouteList0.Update
          Else
             '更新選取之訂單資料
             rs_Tab0_RouteList0.Fields("二次排車").Value = rs_Tab0_RouteList1.Fields("二次排車").Value
             rs_Tab0_RouteList0.Fields("客戶單號").Value = rs_Tab0_RouteList1.Fields("客戶單號").Value
             rs_Tab0_RouteList0.Fields("路線編號").Value = rs_Tab0_RouteList1.Fields("路線編號").Value
             rs_Tab0_RouteList0.Fields("日期").Value = rs_Tab0_RouteList1.Fields("日期").Value
             rs_Tab0_RouteList0.Fields("指送客戶").Value = rs_Tab0_RouteList1.Fields("指送客戶").Value
             rs_Tab0_RouteList0.Fields("箱數").Value = rs_Tab0_RouteList1.Fields("箱數").Value
             rs_Tab0_RouteList0.Fields("重量").Value = rs_Tab0_RouteList1.Fields("重量").Value
             rs_Tab0_RouteList0.Fields("材積").Value = rs_Tab0_RouteList1.Fields("材積").Value
             rs_Tab0_RouteList0.Fields("多車").Value = rs_Tab0_RouteList1.Fields("多車").Value
             rs_Tab0_RouteList0.Fields("訂單號碼").Value = rs_Tab0_RouteList1.Fields("訂單號碼").Value
          End If
          rs_Tab0_RouteList1.MoveNext
       Loop
       
       '[刪除已選取之訂單
       rs_Tab0_RouteList1.MoveFirst
       Do While Not rs_Tab0_RouteList1.EOF
          rs_Tab0_RouteList1.Delete
          rs_Tab0_RouteList1.MoveFirst
       Loop
       Call ReSet_Tab0_RouteList1_SeqNo
       Call ReSet_Tab0_RouteList0_SeqNo
    End If
    rs_Tab0_RouteList1.Filter = adFilterNone
    rs_Tab0_RouteList1.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    rs_Tab0_RouteList0.Filter = adFilterNone
    rs_Tab0_RouteList0.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-選取向上", Me.Caption, "cmd_Tab0_Selected_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmd_Tab1_CreateRoute_Click()
    If rs_Tab1_SelectedOrders.RecordCount = 0 Then
        msg_text = "資料錯誤：無裝載資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(txt_Tab1_DELIVERY_DATE.Text)) = 0 Then
        msg_text = "資料錯誤：未輸入出車日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    If Len(Trim(txt_Tab1_VehicleNo.Text)) = 0 Then
        msg_text = "資料錯誤：未輸入車牌號碼"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_VehicleNo.SetFocus
        Exit Sub
    End If
    '出車日期：格式 yyyymmdd
    txt_Tab1_DELIVERY_DATE.Text = Trim(txt_Tab1_DELIVERY_DATE.Text)
    If Fun_ChkDateFormat(txt_Tab1_DELIVERY_DATE.Text) = 1 Then
        msg_text = "出車日期：" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_DELIVERY_DATE.SelStart = 0: txt_Tab1_DELIVERY_DATE.SelLength = Len(txt_Tab1_DELIVERY_DATE.Text): txt_Tab1_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    On Error GoTo err_Handle
    Tab1_RouteListEventEnable = False
    
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    '檢核 [車牌號碼] 是否有效
    txt_Tab1_VehicleNo.Text = Trim(txt_Tab1_VehicleNo.Text)
    str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab1_VehicleNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       msg_text = "資料錯誤：車牌號碼 " & txt_Tab1_VehicleNo.Text & " 未建檔"
       MsgBox msg_text, vbOKOnly + vbCritical, msg_title
       txt_Tab1_VehicleNo.SelStart = 0: txt_Tab1_VehicleNo.SelLength = Len(txt_Tab1_VehicleNo.Text)
       txt_Tab1_VehicleNo.SetFocus
       Exit Sub
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    Dim intDriveTimes As Integer    '車次
    Dim strRouteNo As String        '路線編號
    
    '產生車次
    str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
              "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab1_DELIVERY_DATE.Text & "' and Vehicle_ID_No = '" & txt_Tab1_VehicleNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    tmp_Rs.Close
    
    '產生路線編號
    str_SQL = "Select Isnull(Max(Cast(Right(C_ROUTE_NO,3) as integer))+1,1) as RouteSN " & _
              "From SDN01T Where Substring(C_ROUTE_NO,2,6)='" & Mid(txt_Tab1_DELIVERY_DATE.Text, 3, 6) & "' and Left(C_ROUTE_NO,1) = 'N'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strRouteNo = "N" & Mid(txt_Tab1_DELIVERY_DATE, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
    tmp_Rs.Close
    DoEvents: DoEvents
    Tran_Level = cn.BeginTrans
    
        '存表頭,SDN01T
        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times) " & _
            "Values ( '" & Trim(txt_Tab1_DELIVERY_DATE.Text) & "','" & Trim(strRouteNo) & "','" & Trim(txt_Tab1_VehicleNo.Text) & "', " & _
            "'" & Trim(txt_Tab1_Driver0.Text) & "','0','" & User_id & "','" & intDriveTimes & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
        '更新請款人
        cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & txt_Tab1_VehicleNo.Text & "') where c_route_no = '" & strRouteNo & "'", RowsAffect, adExecuteNoRecords
                    
        rs_Tab1_SelectedOrders.MoveFirst
        Do While Not rs_Tab1_SelectedOrders.EOF
            '存訂單,SDN02T
            str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO,C_RECEIPT_NO) " & _
                    "Values ( '" & Trim(strRouteNo) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("路線編號").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("客戶單號").Value) & "', " & _
                    "'" & Trim(rs_Tab1_SelectedOrders.Fields("日期").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("指送客戶").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("箱數").Value) & "', " & _
                    "'" & Trim(rs_Tab1_SelectedOrders.Fields("材積").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("重量").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("多車").Value) & "', " & _
                    "'" & Trim(rs_Tab1_SelectedOrders.Fields("訂單號碼").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("C_RECEIPT_NO").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

            '補一般訂單所需資料 by gemini
            str_SQL = "Update sdn02t " & _
                        "Set sdn02t.storerkey = orders.storerkey " & _
                        ",sdn02t.receipt_date = convert(varchar(8),orders.orderdate,112) " & _
                        ",sdn02t.consigneekey = orders.consigneekey " & _
                        ",sdn02t.description = orders.notes " & _
                        ",sdn02t.priority = orders.priority " & _
                        ",sdn02t.scheduledate = (select top 1 trp02t.scheduledate from trp02t where trp02t.c_receipt_no = orders.orderkey order by trp02t.scheduledate ) " & _
                        ",sdn02t.scan = 'N' " & _
                        ",sdn02t.vehicle_id_no = '" & Trim(txt_Tab1_VehicleNo.Text) & "' " & _
                        "from orders join sdn02t on sdn02t.c_receipt_no = orders.orderkey " & _
                        "where sdn02t.receipt_no = '" & Trim(rs_Tab1_SelectedOrders.Fields("訂單號碼").Value) & "' and sdn02t.priority not in ('R','A2B','RC') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '補退貨訂單所需資料 by gemini
            str_SQL = "Update sdn02t " & _
                        "Set sdn02t.storerkey = orders.storerkey " & _
                        ",sdn02t.receipt_date = convert(varchar(8),orders.orderdate,112) " & _
                        ",sdn02t.consigneekey = orders.consigneekey " & _
                        ",sdn02t.description = orders.notes " & _
                        ",sdn02t.priority = orders.priority " & _
                        ",sdn02t.scheduledate = (select top 1 ort02t.scheduledate from ort02t where ort02t.c_receipt_no = orders.orderkey order by ort02t.scheduledate ) " & _
                        ",sdn02t.scan = 'N' " & _
                        ",sdn02t.vehicle_id_no = '" & Trim(txt_Tab1_VehicleNo.Text) & "' " & _
                        "from orders join sdn02t on sdn02t.c_receipt_no = orders.orderkey " & _
                        "where sdn02t.receipt_no = '" & Trim(rs_Tab1_SelectedOrders.Fields("訂單號碼").Value) & "' and sdn02t.priority in ('R','A2B','RC') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '存明細,SDN03T
            str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME) " & _
                "select  '" & Trim(strRouteNo) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ship_qty,SIGN_QTY, SHIP_TIME " & _
                "from SDN03W where  RECEIPT_NO in('" & Trim(rs_Tab1_SelectedOrders.Fields("訂單號碼").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '補訂單資料 by gemini
             str_SQL = "Update sdn03t Set sdn03t.Weight = sdn03t.order_qty * sp.stdgrosswgt ,sdn03t.volumn_weight = sdn03t.order_qty * sp.stdcube " & _
                  "from gv_skuxpack sp join sdn03t on sdn03t.product_no = sp.sku and sp.storerkey = sdn03t.storerkey " & _
                  "where sdn03t.receipt_no = '" & Trim(rs_Tab1_SelectedOrders.Fields("訂單號碼").Value) & "'"
              cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                      
            '刪除訂單,SDN02W
            str_SQL = "delete SDN02W where RECEIPT_NO in('" & Trim(rs_Tab1_SelectedOrders.Fields("訂單號碼").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '刪除明細,SDN03W
            str_SQL = "delete SDN03W where RECEIPT_NO in('" & Trim(rs_Tab1_SelectedOrders.Fields("訂單號碼").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            rs_Tab1_SelectedOrders.MoveNext
        Loop
        
        '刪除已選取之訂單
        DoEvents: DoEvents
        rs_Tab1_SelectedOrders.MoveFirst
        Do While Not rs_Tab1_SelectedOrders.EOF
            rs_Tab1_SelectedOrders.Delete
            rs_Tab1_SelectedOrders.MoveFirst
        Loop
        
'        '更新APPOrderDate狀態
'        str_SQL = "update AppOrderDate set status = '5' where and status < '6' and c_Route_No= '" & Trim(strRouteNo) & "' "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

            '更新OrderType Add by Gemini @20190604
            str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & strRouteNo & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & strRouteNo & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
    cn.CommitTrans
    
    '畫面處理
    txt_Tab1_DELIVERY_DATE.Text = ""
    txt_Tab1_VehicleNo.Text = ""
    txt_Tab1_Driver0.Text = ""
    txt_Tab1_Selected_Case.Text = "0"
    txt_Tab1_Selected_Volumn.Text = "0"
    txt_Tab1_Selected_Weight.Text = "0"
    '顯示新建之路編
    txt_Tab2_Route_Start.Text = strRouteNo
    Call cmd_Tab2_RouteNoQuery_Click
    Tab1_RouteListEventEnable = True
    SSTab1.Tab = 2
    Exit Sub
    
err_Handle:
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, Me.Name & "出車確認-建立路線編號")
End Sub

Private Sub cmd_Tab1_ImportOrders_Click()

 '更新箱材重資料
 cn.Execute "exec gs_UpdateSDNW", RowsAffect, adExecuteNoRecords
 
 '出車確認>>匯入待重整訂單
 Screen.MousePointer = vbHourglass
 DoEvents: DoEvents
 Tab1_RouteListEventEnable = False
 Set dg_SDN02W.DataSource = Nothing
 strSourceFilter = adFilterNone
 DoEvents
 
 Call CreateRS_Tab1_SelectedOrders
 '待排車訂單載入：選取小計：歸零
 txt_Tab1_srcSelected_Case.Text = ""
 txt_Tab1_srcSelected_Volumn.Text = ""
 txt_Tab1_srcSelected_Weight.Text = ""
 
 '取回待排車訂單
 str_SQL = "SELECT  ' ' as '＊',C_ROUTE_NO AS 二次排車, ROUTE_NO AS 路線編號,EXTERN AS 客戶單號,ARRIVE_DATE AS 日期,CUST_NAME as 指送客戶, " & _
         "SHIP_CS As 箱數, SHIP_CBM As 材積, SHIP_WT As 重量, RECEIPT_NO As 訂單號碼, CAR_NOTES As 多車 , C_Receipt_no " & _
         "FROM dbo.SDN02W Order by 二次排車,路線編號,客戶單號,箱數 "
 strSourceOrderBy = " 二次排車,路線編號,客戶單號,箱 "
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
 Call Replication_Recordset(tmp_Rs, rs_SDN02W)
 Set dg_SDN02W.DataSource = rs_SDN02W
 tmp_Rs.Close
 With dg_SDN02W
     .ColumnHeaders = True         '標題行顯示
     .RowHeight = 250
     .Columns(0).Width = 500       '序號
     .Columns(0).Alignment = dbgLeft
     .Columns(1).Width = 500       '選取
     .Columns(1).Alignment = dbgCenter
     .Columns(2).Width = 1000      '二次排車
     .Columns(2).Alignment = dbgLeft
     .Columns(3).Width = 1000      '客戶單號
     .Columns(3).Alignment = dbgLeft
     .Columns(4).Width = 800      '路線編號
     .Columns(4).Alignment = dbgLeft
     .Columns(5).Width = 800       '日期
     .Columns(5).Alignment = dbgLeft
     .Columns(6).Width = 1500      '指送客戶
     .Columns(6).Alignment = dbgLeft
     .Columns(7).Width = 800       '箱數
     .Columns(7).Alignment = dbgRight
     .Columns(8).Width = 800       '材積
     .Columns(8).Alignment = dbgRight
     .Columns(9).Width = 800       '重量
     .Columns(9).Alignment = dbgRight
     .Columns(10).Width = 1000       '訂單號碼
     .Columns(10).Alignment = dbgLeft
     .Columns(11).Width = 800       '多車
     .Columns(11).Alignment = dbgLeft
 End With
 DoEvents: DoEvents

 rs_SDN02W.MoveFirst
 '待排車訂單總計資訊
 Call Retrive_OrderSum
 Screen.MousePointer = vbDefault
 Tab1_RouteListEventEnable = True
 Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認>>匯入待重整訂單", Me.Caption, "cmd_Tab1_ImportOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Remove_Click()
    '篩選已選取者
    If dg_Tab1_SelectedOrders.DataSource Is Nothing Then Exit Sub
    If dg_SDN02W.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab1_RouteListEventEnable = False
    rs_Tab1_SelectedOrders.Filter = "＊='V'"
    If Not rs_Tab1_SelectedOrders.EOF Then
        Do While Not rs_Tab1_SelectedOrders.EOF
            '判斷是否已經選取過
            rs_SDN02W.Filter = adFilterNone
            rs_SDN02W.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
            rs_SDN02W.Filter = "訂單號碼 = '" & rs_Tab1_SelectedOrders.Fields("訂單號碼").Value & "'"
            '如果是查詢所顯示之有效路編，設定路編異動識別旗標
            If blRouteModify Then blRouteChange = True
            If rs_SDN02W.EOF Then
                '新增選取之訂單
                rs_SDN02W.AddNew
                rs_SDN02W.Fields("編號").Value = 999
                rs_SDN02W.Fields("二次排車").Value = rs_Tab1_SelectedOrders.Fields("二次排車").Value
                rs_SDN02W.Fields("客戶單號").Value = rs_Tab1_SelectedOrders.Fields("客戶單號").Value
                rs_SDN02W.Fields("路線編號").Value = rs_Tab1_SelectedOrders.Fields("路線編號").Value
                rs_SDN02W.Fields("日期").Value = rs_Tab1_SelectedOrders.Fields("日期").Value
                rs_SDN02W.Fields("指送客戶").Value = rs_Tab1_SelectedOrders.Fields("指送客戶").Value
                rs_SDN02W.Fields("箱數").Value = rs_Tab1_SelectedOrders.Fields("箱數").Value
                rs_SDN02W.Fields("重量").Value = rs_Tab1_SelectedOrders.Fields("重量").Value
                rs_SDN02W.Fields("材積").Value = rs_Tab1_SelectedOrders.Fields("材積").Value
                rs_SDN02W.Fields("多車").Value = rs_Tab1_SelectedOrders.Fields("多車").Value
                rs_SDN02W.Fields("訂單號碼").Value = rs_Tab1_SelectedOrders.Fields("訂單號碼").Value
                rs_SDN02W.Fields("C_Receipt_No").Value = rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value
                rs_SDN02W.Update
            Else
                '更新選取之訂單資料
                rs_SDN02W.Fields("二次排車").Value = rs_Tab1_SelectedOrders.Fields("二次排車").Value
                rs_SDN02W.Fields("客戶單號").Value = rs_Tab1_SelectedOrders.Fields("客戶單號").Value
                rs_SDN02W.Fields("路線編號").Value = rs_Tab1_SelectedOrders.Fields("路線編號").Value
                rs_SDN02W.Fields("日期").Value = rs_Tab1_SelectedOrders.Fields("日期").Value
                rs_SDN02W.Fields("指送客戶").Value = rs_Tab1_SelectedOrders.Fields("指送客戶").Value
                rs_SDN02W.Fields("箱數").Value = rs_Tab1_SelectedOrders.Fields("箱數").Value
                rs_SDN02W.Fields("重量").Value = rs_Tab1_SelectedOrders.Fields("重量").Value
                rs_SDN02W.Fields("材積").Value = rs_Tab1_SelectedOrders.Fields("材積").Value
                rs_SDN02W.Fields("多車").Value = rs_Tab1_SelectedOrders.Fields("多車").Value
                rs_SDN02W.Fields("訂單號碼").Value = rs_Tab1_SelectedOrders.Fields("訂單號碼").Value
                rs_SDN02W.Fields("C_Receipt_No").Value = rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value
            End If
            rs_Tab1_SelectedOrders.MoveNext
        Loop
        
        '[刪除已選取之訂單
        rs_Tab1_SelectedOrders.MoveFirst
        Do While Not rs_Tab1_SelectedOrders.EOF
            rs_Tab1_SelectedOrders.Delete
            rs_Tab1_SelectedOrders.MoveFirst
        Loop
    End If
    rs_SDN02W.Filter = adFilterNone
    rs_SDN02W.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    rs_Tab1_SelectedOrders.Filter = adFilterNone
    rs_Tab1_SelectedOrders.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    Call ReSet_Tab1_SelectedOrders_SeqNo
    Call ReSet_Tab1_SDN02W_SeqNo
    Tab1_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-Tab1選取向下", Me.Caption, "cmd_Tab1_Selected_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_SelectCar_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab1_SelectCar.Name & "2")
End Sub

Private Sub cmd_Tab1_Selected_Click()
    '篩選已選取者
    If dg_Tab1_SelectedOrders.DataSource Is Nothing Then Exit Sub
    If dg_SDN02W.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab1_RouteListEventEnable = False
    rs_SDN02W.Filter = "＊='V'"
    If Not rs_SDN02W.EOF Then
        Do While Not rs_SDN02W.EOF
            '判斷是否已經選取過
            rs_Tab1_SelectedOrders.Filter = adFilterNone
            rs_Tab1_SelectedOrders.Sort = "編號 asc"  '原始排序，一般資料序號由小至大
            rs_Tab1_SelectedOrders.Filter = "訂單號碼 = '" & rs_SDN02W.Fields("訂單號碼").Value & "'"
            '如果是查詢所顯示之有效路編，設定路編異動識別旗標
            If blRouteModify Then blRouteChange = True
            If rs_Tab1_SelectedOrders.EOF Then
                '新增選取之訂單
                rs_Tab1_SelectedOrders.AddNew
                rs_Tab1_SelectedOrders.Fields("編號").Value = 999
                rs_Tab1_SelectedOrders.Fields("二次排車").Value = rs_SDN02W.Fields("二次排車").Value
                rs_Tab1_SelectedOrders.Fields("客戶單號").Value = rs_SDN02W.Fields("客戶單號").Value
                rs_Tab1_SelectedOrders.Fields("路線編號").Value = rs_SDN02W.Fields("路線編號").Value
                rs_Tab1_SelectedOrders.Fields("日期").Value = rs_SDN02W.Fields("日期").Value
                rs_Tab1_SelectedOrders.Fields("指送客戶").Value = rs_SDN02W.Fields("指送客戶").Value
                rs_Tab1_SelectedOrders.Fields("箱數").Value = rs_SDN02W.Fields("箱數").Value
                rs_Tab1_SelectedOrders.Fields("重量").Value = rs_SDN02W.Fields("重量").Value
                rs_Tab1_SelectedOrders.Fields("材積").Value = rs_SDN02W.Fields("材積").Value
                rs_Tab1_SelectedOrders.Fields("多車").Value = rs_SDN02W.Fields("多車").Value
                rs_Tab1_SelectedOrders.Fields("訂單號碼").Value = rs_SDN02W.Fields("訂單號碼").Value
                rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value = rs_SDN02W.Fields("C_Receipt_No").Value
                rs_Tab1_SelectedOrders.Update
            Else
                '更新選取之訂單資料
                rs_Tab1_SelectedOrders.Fields("二次排車").Value = rs_SDN02W.Fields("二次排車").Value
                rs_Tab1_SelectedOrders.Fields("客戶單號").Value = rs_SDN02W.Fields("客戶單號").Value
                rs_Tab1_SelectedOrders.Fields("路線編號").Value = rs_SDN02W.Fields("路線編號").Value
                rs_Tab1_SelectedOrders.Fields("日期").Value = rs_SDN02W.Fields("日期").Value
                rs_Tab1_SelectedOrders.Fields("指送客戶").Value = rs_SDN02W.Fields("指送客戶").Value
                rs_Tab1_SelectedOrders.Fields("箱數").Value = rs_SDN02W.Fields("箱數").Value
                rs_Tab1_SelectedOrders.Fields("重量").Value = rs_SDN02W.Fields("重量").Value
                rs_Tab1_SelectedOrders.Fields("材積").Value = rs_SDN02W.Fields("材積").Value
                rs_Tab1_SelectedOrders.Fields("多車").Value = rs_SDN02W.Fields("多車").Value
                rs_Tab1_SelectedOrders.Fields("訂單號碼").Value = rs_SDN02W.Fields("訂單號碼").Value
                rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value = rs_SDN02W.Fields("C_Receipt_No").Value
            End If
            txt_Tab1_Selected_Case.Text = Val(txt_Tab1_Selected_Case.Text) + Val(rs_SDN02W.Fields("箱數").Value)
            txt_Tab1_Selected_Weight.Text = Val(txt_Tab1_Selected_Weight.Text) + Val(rs_SDN02W.Fields("重量").Value)
            txt_Tab1_Selected_Volumn.Text = Val(txt_Tab1_Selected_Volumn.Text) + Val(rs_SDN02W.Fields("材積").Value)
            rs_SDN02W.MoveNext
        Loop
       
       '[刪除已選取之訂單
        rs_SDN02W.MoveFirst
        Do While Not rs_SDN02W.EOF
            rs_SDN02W.Delete
            rs_SDN02W.MoveFirst
        Loop
        txt_Tab1_srcSelected_Case.Text = 0
        txt_Tab1_srcSelected_Volumn.Text = 0
        txt_Tab1_srcSelected_Weight.Text = 0
    End If
    rs_SDN02W.Filter = adFilterNone
    rs_SDN02W.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    rs_Tab1_SelectedOrders.Filter = adFilterNone
    rs_Tab1_SelectedOrders.Sort = "編號 asc"  '原始排序，一定要有這行資料才會重新顯示
    Call ReSet_Tab1_SelectedOrders_SeqNo
    Call ReSet_Tab1_SDN02W_SeqNo
    Tab1_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-Tab1選取向上", Me.Caption, "cmd_Tab1_Selected_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_CreateRoute_Click()
    If Len(Trim(txt_Tab2_Route.Text)) = 0 Then Exit Sub
    
    If Len(Trim(txt_Tab2_DELIVERY_DATE.Text)) = 0 Then
        msg_text = "資料錯誤：未輸入出車日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    If Len(Trim(txt_Tab2_VehicleNo.Text)) = 0 Then
        msg_text = "資料錯誤：未輸入車牌號碼"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_VehicleNo.SetFocus
        Exit Sub
    End If
    '出車日期：格式 yyyymmdd
    txt_Tab2_DELIVERY_DATE.Text = Trim(txt_Tab2_DELIVERY_DATE.Text)
    If Fun_ChkDateFormat(txt_Tab2_DELIVERY_DATE.Text) = 1 Then
        msg_text = "出車日期：" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab2_DELIVERY_DATE.SelStart = 0: txt_Tab2_DELIVERY_DATE.SelLength = Len(txt_Tab2_DELIVERY_DATE.Text): txt_Tab2_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    On Error GoTo err_Handle
    Call DB_CheckConnectStatus
    
'    '檢核是否簽單確認
'    Call ReDim_Recordset(tmp_Rs)
'    str_SQL = "Select Count(*) as Receiver From SDN02T Where C_Route_No = '" & Trim(rs_Tab2_Route.Fields("二次排車").Value) & "'  and confirm_notes <> '' "
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.Fields("Receiver").Value > 0 Then
'        tmp_Rs.Close
'        msg_text = "已有部份簽單完成維護,無法修改："
'        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    tmp_Rs.Close
'
'    'Terry 20190320 新增防呆 已維護棧板之路編不可修改車號
'    Call ReDim_Recordset(tmp_Rs)
'    str_SQL = "select count(*) from pallet_cds where checkno = '" & Trim(rs_Tab2_Route.Fields("二次排車").Value) & "'"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.Fields(0).Value > 0 Then
'        tmp_Rs.Close
'        MsgBox ("此路編已維護棧板，無法變更車號!")
'        Exit Sub
'    End If
'    tmp_Rs.Close

    'Terry 20190327 新增防呆 已有計費資料之路編不可修改車號
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from sdn05t where c_route_no = '" & Trim(rs_Tab2_Route.Fields("二次排車").Value) & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("此路編已有維護資料，無法變更車號!")
        Exit Sub
    End If
    tmp_Rs.Close
    
    
    
    '檢核 [車牌號碼] 是否有效
    Call ReDim_Recordset(tmp_Rs)
    txt_Tab1_VehicleNo.Text = Trim(txt_Tab1_VehicleNo.Text)
    str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab2_VehicleNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "資料錯誤：車牌號碼 " & txt_Tab1_VehicleNo.Text & " 未建檔"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        txt_Tab1_VehicleNo.SelStart = 0: txt_Tab1_VehicleNo.SelLength = Len(txt_Tab1_VehicleNo.Text)
        txt_Tab1_VehicleNo.SetFocus
        Exit Sub
    End If
    tmp_Rs.Close
    
    DoEvents: DoEvents
    cn.BeginTrans
        '存表頭,SDN01T
        str_SQL = "Update SDN01T set DELIVERY_DATE='" & txt_Tab2_DELIVERY_DATE.Text & "',C_VEHICLE_ID_NO='" & Trim(txt_Tab2_VehicleNo.Text) & "',Driver='" & Trim(txt_Tab2_Driver.Text) & "',edituser = '" & User_id & _
                "',editdate = getdate() where C_Route_No='" & Trim(txt_Tab2_Route.Text) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '更新請款人
        cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & Trim(txt_Tab2_VehicleNo.Text) & "') where c_route_no = '" & Trim(txt_Tab2_Route.Text) & "'", RowsAffect, adExecuteNoRecords
        
    cn.CommitTrans

    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-Tab2確認存檔", Me.Caption, "cmd_Tab1_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Excel_Click()
    '排車一覽表 >> 轉 EXCEL

    If rs_Tab2_Route Is Nothing Then Exit Sub
    rs_Tab2_Route.MoveFirst
    On Error GoTo err_Handle
    Tab2_RouteListEventEnable = False
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
'    MyXlsApp.Sheets("Sheet1").Name = "出車確認一覽表"
    MyXlsApp.ActiveSheet.Name = "出車確認一覽表"
    
    i = 1
    'select convert(char,s1.DELIVERY_DATE,112) as 日期,s1.C_VEHICLE_ID_NO as 車號,s1.Driver as 司機,s1.C_Route_No as 二次排車, " & _
            "sum(s2.SHIP_CBM) as 材積,sum(s2.SHIP_WT) as 重量,Max(Distinct s2.CUST_NAME) as 客戶簡稱
    MyXlsApp.Cells(i, 1).Value = "編號"
    MyXlsApp.Cells(i, 2).Value = "出車日期"
    MyXlsApp.Cells(i, 3).Value = "車牌號碼"
    MyXlsApp.Cells(i, 4).Value = "駕駛人"
    MyXlsApp.Cells(i, 5).Value = "路線編號"
    MyXlsApp.Cells(i, 6).Value = "運送重量"
    MyXlsApp.Cells(i, 7).Value = "運送材積"
    MyXlsApp.Cells(i, 8).Value = "客戶簡稱"
    i = i + 1
    rs_Tab2_Route.MoveFirst
    '日期,車號,單號,班別,借出,還入
    Do While Not rs_Tab2_Route.EOF
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab2_Route.Fields(0))
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab2_Route.Fields(1))
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rs_Tab2_Route.Fields(2)
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rs_Tab2_Route.Fields(3)
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rs_Tab2_Route.Fields(4)
        MyXlsApp.Cells(i, 6).Value = rs_Tab2_Route.Fields(5)
        MyXlsApp.Cells(i, 7).Value = rs_Tab2_Route.Fields(6)
        MyXlsApp.Cells(i, 8).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 8).Value = rs_Tab2_Route.Fields(7)
        rs_Tab2_Route.MoveNext
        i = i + 1
    Loop
    i = i + 1
    '最適欄寬
    MyXlsApp.Columns("A:H").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '儲存格格式設定
    MyXlsApp.Columns("F:G").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A1:H" & i - 1).Select
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
    Tab2_RouteListEventEnable = True
    Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-出車確認-轉 EXCEL", Me.Caption, "cmd_Tab2_Excel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_RouteNoDelete_Click()

    'If blAdmin = False Then MsgBox "系統管理員才有權限執行此作業!", 64, "權限不足": Exit Sub
    '路線編號列表 >> 路線編號刪除
    If rs_Tab2_Route.RecordCount = 0 Then Exit Sub
    If dg_Tab2_Route.SelBookmarks.Count = 0 Then MsgBox "未選取路線編號!", 64, "路線編號刪除": Exit Sub
    On Error GoTo err_Handle
    Tab2_RouteListEventEnable = False
    Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
    strDeleteRouteNo = Trim(rs_Tab2_Route.Fields("二次排車").Value)
    strCarno = Trim(rs_Tab2_Route.Fields("車號").Value)
    'dbDriveTimes = Trim(rs_Tab2_Route.Fields("車次").Value)
        
'    '檢查刪除的路編是否有已回傳的訂單
    rs_Tab2_RouteOrders.MoveFirst
    Do While Not rs_Tab2_RouteOrders.EOF
        Call ReDim_Recordset(tmp_Rs)
        str_SQL = "Select returnstatus From SDN02t Where receipt_no = '" & Trim(rs_Tab2_RouteOrders.Fields("訂單號碼").Value) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        tmp_Rs.MoveFirst
        Do While Not tmp_Rs.EOF
            If tmp_Rs.Fields("returnstatus").Value > 0 Then
                msg_text = "訂單號碼:" & Trim(rs_Tab2_RouteOrders.Fields("訂單號碼").Value) & " 資料已回傳，無法進行刪除!"
                MsgBox msg_text, vbOKOnly + vbCritical, msg_title
                Screen.MousePointer = vbDefault
                blTab1RouteEventEnable = True
                Tab2_RouteListEventEnable = True
                tmp_Rs.Close
                Exit Sub
            End If
            tmp_Rs.MoveNext
        Loop
        rs_Tab2_RouteOrders.MoveNext
    Loop
    rs_Tab2_RouteOrders.MoveFirst
    
'    msg_text = "確認刪除路線編號：" & strDeleteRouteNo
'    If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub

    'Terry 20190311 新增防呆 已維護棧板之路編不可打散重組
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from pallet_cds where checkno = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("此路編已維護棧板，無法刪除!")
        Exit Sub
    End If
    tmp_Rs.Close
    
    'Terry 20190327 新增防呆 已有計費資料之路編不可打散重組
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from sdn05t where c_route_no = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("此路編已有計費資料，無法刪除!")
        Exit Sub
    End If
    tmp_Rs.Close
    

    If MsgBox("此路編所有訂單的運費與簽單確認將一併刪除，訂單將轉入打散重組，是否繼續?", vbOKCancel, Trim(strDeleteRouteNo) & "==>打散重組") <> vbOK Then blTab1RouteEventEnable = True: Tab2RouteListEventEnable = True: Tab2_RouteListEventEnable = True: Exit Sub
    
    '刪除路編
    Call Delete_RouteNo(strDeleteRouteNo)
    
    '刪除查詢結果中該筆路線編號--rs_Tab1_RouteOrders
    rs_Tab2_RouteOrders.Filter = adFilterNone
    rs_Tab2_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    rs_Tab2_RouteOrders.Filter = "二次排車='" & strDeleteRouteNo & "'"
    If Not rs_Tab2_RouteOrders.EOF Then
        Do While Not rs_Tab2_RouteOrders.EOF
            rs_Tab2_RouteOrders.Delete
            rs_Tab2_RouteOrders.MoveFirst
        Loop
    End If
    rs_Tab2_RouteOrders.Filter = adFilterNone
    rs_Tab2_RouteOrders.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
    
    '(7).刪除查詢結果中該筆路線編號--rs_Tab1_Route
    rs_Tab2_Route.Delete
    If Not rs_Tab2_Route.EOF Then rs_Tab2_Route.MoveFirst
    
    blTab1RouteEventEnable = True
    Tab2_RouteListEventEnable = True
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-出車確認-路線編號刪除", Me.Caption, "cmd_Tab2_RouteNoDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_RouteNoQuery_Click()
    '出車確認 >> Tab2路線編號查詢
    'If Len(Trim(txt_Tab2_Route_Start.Text)) = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    '路線編號
    txt_Tab2_Route_Start.Text = Trim(txt_Tab2_Route_Start.Text)
    strSubwhere = ""
    If Len(txt_Tab2_Route_Start.Text) > 0 Then
        strSubwhere = " s1.C_Route_No = '" & txt_Tab2_Route_Start.Text & "' "
    End If
    If Len(strSubwhere) > 0 Then
        If Len(str_Where) = 0 Then
            str_Where = strSubwhere
        Else
            str_Where = str_Where & " and " & strSubwhere
        End If
    End If
    '出車日期
    txt_Tab2_DeliveryDate_Start.Text = Trim(txt_Tab2_DeliveryDate_Start.Text)
    strSubwhere = ""
    If Len(txt_Tab2_DeliveryDate_Start.Text) > 0 Then
        strSubwhere = " s1.DELIVERY_DATE = '" & txt_Tab2_DeliveryDate_Start.Text & "' "
    End If
    If Len(strSubwhere) > 0 Then
        If Len(str_Where) = 0 Then
            str_Where = strSubwhere
        Else
            str_Where = str_Where & " and " & strSubwhere
        End If
    End If
    '組str_SQL
    'str_SQL = "Select C_Route_No as 二次排車, Convert(varchar,DELIVERY_DATE,112) as 日期, C_VEHICLE_ID_NO as 車號, Driver as 司機 From SDN01T"
    str_SQL = "select convert(char(8),s1.DELIVERY_DATE,112) as 日期,s1.C_VEHICLE_ID_NO as 車號,s1.Driver as 司機,s1.C_Route_No as 二次排車, " & _
            "sum(s2.SHIP_CBM) as 材積,sum(s2.SHIP_WT) as 重量,Max(Distinct s2.CUST_NAME) as 客戶簡稱 from SDN01T s1 " & _
            "inner join SDN02T s2 on s1.C_Route_No=s2.C_Route_No "
    If Len(str_Where) = 0 Then
        str_SQL = str_SQL & "group by s1.DELIVERY_DATE order by s1.C_Route_No"
    Else
        str_SQL = str_SQL & " where" & str_Where & " group by s1.DELIVERY_DATE,s1.C_Route_No,s1.C_VEHICLE_ID_NO,s1.Driver order by s1.C_Route_No"
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       msg_text = "查詢結果：無符合設定條件之路線編號資料(SDN01T)"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab2_Route)
    Set dg_Tab2_Route.DataSource = rs_Tab2_Route
    'tmp_rs.Close
    With dg_Tab2_Route
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000      '二次排車
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '日期
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000      '車號
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1500       '司機
        .Columns(4).Alignment = dbgLeft
    End With
    rs_Tab2_Route.MoveFirst
    txt_Tab2_Route.Text = rs_Tab2_Route.Fields("二次排車").Value
    txt_Tab2_VehicleNo.Text = rs_Tab2_Route.Fields("車號").Value
    txt_Tab2_Driver.Text = rs_Tab2_Route.Fields("司機").Value
    txt_Tab2_DELIVERY_DATE.Text = rs_Tab2_Route.Fields("日期").Value
    'SDN03T
    Call Display_Tab2_RouteOrders
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-出車確認 >> Tab2路線編號查詢", Me.Caption, "cmd_Tab2_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_SelectCar_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab2_SelectCar.Name & "2")
End Sub

Private Sub cmd_Tab3_ClearQty_Click()
    '待切割訂單 >> 清除 [板數切割][箱數切割] 欄位值
    txt_Tab3_CutCaseQty.Text = ""
    txt_Tab3_CutCaseQty.SetFocus
    'RUN Button [數量切割] Click
    Call cmd_Tab3_CutQty_Click
End Sub

Private Sub cmd_Tab3_CutOrders_Click()
    '出車確認 >> 切割訂單
    If rs_Tab3_SDN02W Is Nothing Then Exit Sub
    If rs_Tab3_SDN02W.RecordCount = 0 Then Exit Sub
    
    Dim intTRP02WBookMark As String     '正在進行 [訂單切割作業] 之訂單資料列
    Dim strCutOrder_SrcKey As String    '正在進行 [訂單切割作業] 之訂單編號
    Dim dbMaxKey As Double              '新訂單編號：尾碼 key
    Dim strCutOrder_NewKey As String    '新切割出來之訂單其 [訂單編號]
    Dim i As Double
    
    On Error GoTo err_Handle
    If Len(Trim(CutOrderkey)) = 0 Then Exit Sub
    
    '檢查是有點選欲切割之訂單細項
    dg_Tab3_SelectedOrderDetail.Visible = False
    Dim dbCount As Double
    dbCount = 0
    With dg_Tab3_SelectedOrderDetail
        For i = 1 To .Rows - 2
            .Row = i: .Col = 1
            If Len(Trim(.Text)) <> 0 Then
                dbCount = dbCount + 1
            End If
        Next i
    End With
    dg_Tab3_SelectedOrderDetail.Visible = True
    If dbCount = 0 Then
        msg_text = "資料錯誤：未選取欲切割之訂單喔"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    '資料庫異動交易--起點
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '為新切割出來之訂單決定其 [訂單編號]
    strCutOrder_SrcKey = CutOrderkey
    str_SQL = "Select Cast(Code as integer) as AvailNo From CodeLKUP Where ListName = 'CUTORDERSNO'  "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        strCutOrder_NewKey = "CT" & Format(1, "00000000")
        str_SQL = "Insert into CodeLKUP (ListName,Code,Description,AddWho,EditWho) Values ('CUTORDERSNO',2,'ㄧ單多車重新產生訂單號碼','" & User_id & "','" & User_id & "')"
    Else
        strCutOrder_NewKey = "CT" & Format(tmp_Rs.Fields("AvailNo").Value, "00000000")
        str_SQL = "Update CodeLKUP Set Code = " & (tmp_Rs.Fields("AvailNo").Value + 1) & " Where ListName = 'CUTORDERSNO'"
    End If
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    tmp_Rs.Close
    
    '阻斷原始訂單列表 DBGrid 的 Event 執行
    'blTRP02WEventEnable = False
    rs_Tab3_SDN02W.Filter = adFilterNone
    rs_Tab3_SDN02W.Filter = "訂單號碼 = '" & CutOrderkey & "'"
    If rs_Tab3_SDN02W.RecordCount = 0 Then
        msg_text = "抱歉ㄟ，找不到符合條件的原訂單資料喔"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        rs_Tab3_SDN02W.Filter = adFilterNone
        rs_Tab3_SDN02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
        Exit Sub
    Else
       '產生切割訂單-Header
        intTRP02WBookMark = rs_Tab3_SDN02W.Bookmark
        txt_Tab3_OrderKey.Text = strCutOrder_NewKey     '訂單編號
        txt_Tab3_DeliveryDate.Text = rs_Tab3_SDN02W.Fields("日期").Value    '送貨日
        txt_Tab3_Extern.Text = rs_Tab3_SDN02W.Fields("客戶單號").Value  '客戶編號
        txt_Tab3_CaseQty.Text = txt_Tab3_SelectedCaseQty.Text '箱數
        txt_Tab3_Weight.Text = txt_Tab3_SelectedWeight.Text    '重量
        txt_Tab3_Volumn.Text = txt_Tab3_SelectedVolumn.Text    '材積
        txt_Tab3_FullName.Text = rs_Tab3_SDN02W.Fields("指送客戶").Value    '客戶名稱
               
        '產生新的訂單資料--SDN02W
        str_SQL = "Insert into SDN02W (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,c_RECEIPT_NO) " & _
                  "Select C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME, " & _
                  "" & txt_Tab3_SelectedCaseQty.Text & "," & txt_Tab3_SelectedVolumn.Text & "," & txt_Tab3_SelectedWeight.Text & ",CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,'" & strCutOrder_NewKey & "',c_RECEIPT_NO " & _
                  "From SDN02W Where  Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '更新原訂單之統計數字--SDN02W
        str_SQL = "Update SDN02W Set SHIP_CS=SHIP_CS-" & txt_Tab3_SelectedCaseQty.Text & "," & _
                  "SHIP_WT=SHIP_WT-" & txt_Tab3_SelectedWeight.Text & ",SHIP_CBM=SHIP_CBM-" & txt_Tab3_SelectedVolumn.Text & " " & _
                  "Where Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    End If
    rs_Tab3_SDN02W.Filter = adFilterNone
    rs_Tab3_SDN02W.Sort = "訂單號碼 ASC"
    blTRP02WEventEnable = True

    '切割訂單之 OrderDetail
    Dim dbsrcQty As Double, dbCutQty As Double, dbSeqNo As String
    dbSeqNo = 0
    dg_Tab3_SelectedOrderDetail.Visible = False
    With dg_Tab3_SelectedOrderDetail
         For i = 1 To .Rows - 2
             .Row = i: .Col = 1
             If .Text <> "" Then   '細項被選取進行切割
                .Col = 0: dbSeqNo = .Text          '保留原定單項次編號已為對應
                .Col = 4: dbsrcQty = Val(.Text)    '原訂單箱數
                .Col = 7: dbCutQty = Val(.Text)    '切割箱數
                If dbsrcQty = dbCutQty Then        '若全項次箱數進行切割，註記準備後續刪除此細項
                    .Col = 1: .Text = "X"
                    '直接更新SDN03W的訂單號碼
                    str_SQL = "Update SDN03W Set Receipt_No = '" & strCutOrder_NewKey & "' " & _
                              "Where Receipt_No = '" & CutOrderkey & "' and SEQ_NO = '" & dbSeqNo & "' "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                Else
                   '更新 待切割訂單明細
                   .Col = 1: .Text = ""        '清除註記，更新欄位值
                   .Col = 4: dbsrcQty = Val(.Text)    '原訂單箱數
                   .Col = 7: dbCutQty = Val(.Text)    '切割箱數
                   .Col = 4: .Text = dbsrcQty - dbCutQty
                   .Col = 5: dbsrcQty = Val(.Text)    '原訂單重量
                   .Col = 8: dbCutQty = Val(.Text)    '切割重量
                   .Col = 5: .Text = dbsrcQty - dbCutQty
                   .Col = 6: dbsrcQty = Val(.Text)    '原訂單材積
                   .Col = 9: dbCutQty = Val(.Text)   '切割材積
                   .Col = 6: .Text = dbsrcQty - dbCutQty
                   
                   '更新SDN03W原數量
                   .Col = 7
                   str_SQL = "Update SDN03W Set SDN03W.SHIP_QTY = " & _
                   "SDN03W.SHIP_QTY - (" & .Text & " * sp.casecnt) ,SDN03W.ORDER_QTY= SDN03W.ORDER_QTY - ( " & .Text & " * sp.casecnt) " & _
                   "from sdn03w SDN03W join gv_skuxpack sp on sp.storerkey = SDN03W.storerkey and sp.sku = SDN03W.product_no Where SDN03W.Receipt_No = '" & CutOrderkey & "' and SDN03W.SEQ_NO = '" & dbSeqNo & "' "
                   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   
'                '將箱數換算回個數 by gemini 20071212
'               str_SQL = "Update SDN03W Set SDN03W.Order_Qty = SDN03W.Order_Qty * s1.casecnt,SDN03W.SHIP_QTY = SDN03W.SHIP_QTY * s1.casecnt " & _
'                        "from SDN03W SDN03W join sku s on SDN03W.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = SDN03W.storerkey " & _
'                        "Where Receipt_No = '" & CutOrderkey & "' and SEQ_NO = '" & dbSeqNo & "' "
'                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   
                   '新增新訂單之訂單細項
                   str_SQL = "Insert into SDN03W (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME) " & _
                             "Select s3.C_ROUTE_NO,s3.ROUTE_NO,s3.StorerKey,'" & strCutOrder_NewKey & "',s3.Seq_No,s3.SubSeq_No,s3.EXTERN,s3.Product_No,s3.Ship_Unit,"
                   .Col = 7: str_SQL = str_SQL & .Text & " * sp.casecnt,"
                   str_SQL = str_SQL & "s3.SIGN_QTY,s3.RSC_CODE,s3.RBC_CODE,s3.CONFIRM_DATE,s3.DESCRIPTION,"
                   .Col = 7: str_SQL = str_SQL & .Text & " * sp.casecnt,"
                   str_SQL = str_SQL & "s3.SHIP_TIME From SDN03W s3 join gv_skuxpack sp on sp.sku = s3.product_no and s3.storerkey = sp.storerkey Where s3.Receipt_No = '" & CutOrderkey & "' and s3.SEQ_NO = '" & dbSeqNo & "' "
                   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   
'                '將箱數換算回個數 by gemini 20071212
'               str_SQL = "Update SDN03W Set SDN03W.Order_Qty = SDN03W.Order_Qty * s1.casecnt ,SDN03W.SHIP_QTY = SDN03W.SHIP_QTY * s1.casecnt " & _
'                        "from SDN03W SDN03W join sku s on SDN03W.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = SDN03W.storerkey " & _
'                        "Where Receipt_No = '" & strCutOrder_NewKey & "' and SEQ_NO = '" & dbSeqNo & "' "
'                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   .Col = 7: .Text = ""   '切割箱數
                   .Col = 8: .Text = ""   '切割重量
                   .Col = 9: .Text = ""  '切割材積
                End If
             End If
         Next i
    End With

    '刪除已全數量切割之訂單細項
    Dim j As Double
    With dg_Tab3_SelectedOrderDetail
        For i = 1 To .Rows - 2
            For j = 1 To .Rows - 2
                .Row = j: .Col = 1
                If .Text = "X" Then
                    Call Delete_GridRow(j)
                    Exit For
                End If
            Next j
        Next i
        '重新產生訂單加總統計資料
        txt_Tab3_Weight.Text = 0
        txt_Tab3_Volumn.Text = 0
        For i = 1 To .Rows - 2
            .Row = i
            .Col = 5: txt_Tab3_Weight.Text = Val(txt_Tab3_Weight.Text) + Val(.Text)
            .Col = 6: txt_Tab3_Volumn.Text = Val(txt_Tab3_Volumn.Text) + Val(.Text)
        Next i
    End With
    dg_Tab3_SelectedOrderDetail.Visible = True
    Call Dispaly_dg_Tab3_SDN03W
    
    '清除欄位值-選取項次統計
    txt_Tab3_SelectedCaseQty.Text = ""
    dbCut_TotalCaseQty = 0
    txt_Tab3_SelectedWeight.Text = ""
    dbCut_TotalWeight = 0
    txt_Tab3_SelectedVolumn.Text = ""
    dbCut_TotalVolumn = 0
    
    '細項切割數量欄位：板數，箱數
    txt_Tab3_CutCaseQty.Text = ""
    If dg_Tab3_SelectedOrderDetail.Rows = 2 And txt_Tab3_Weight.Text = 0 And txt_Tab3_Volumn.Text = 0 Then
        '已全部切割之訂單：刪除 TRP02W & TRP03W
        str_SQL = "Delete From SDN02W Where Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Delete From SDN03W Where Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        DoEvents
        SSTab1.Tab = 0
        DoEvents
    End If
    cn.CommitTrans
    Tran_Level = 0
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then Tran_Level = 0: cn.RollbackTrans
   
   dg_Tab3_SelectedOrderDetail.Visible = True
   blTRP02WEventEnable = True
   
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-出車確認 >> 切割訂單", Me.Caption, "cmd_Tab3_CutOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_CutQty_Click()
    '出車確認 >> 數量切割
    If rs_Tab3_SDN02W Is Nothing Then Exit Sub
    If rs_Tab3_SDN02W.RecordCount = 0 Then Exit Sub
    
    cmd_Tab3_CutOrders.Enabled = False
    cmd_Tab3_ClearQty.Enabled = False
    If Val(txt_Tab3_CutCaseQty.Text) = 0 Then
        '把數量清除表示：不選取此項
        dg_Tab3_SelectedOrderDetail.Col = 1: dg_Tab3_SelectedOrderDetail.Text = ""   '取消選取
    End If
    Dim tmpQty As Double
    If Val(txt_Tab3_CutCaseQty.Text) > 0 Then
      
       dg_Tab3_SelectedOrderDetail.Col = 4: tmpQty = Val(dg_Tab3_SelectedOrderDetail.Text)
       If Val(txt_Tab3_CutCaseQty.Text) > tmpQty Then
            msg_text = "資料錯誤：切割箱數 大於 品項總箱數"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            cmd_Tab3_ClearQty.Enabled = True
            cmd_Tab3_CutOrders.Enabled = True
            Exit Sub
       End If
    
       '輸入切割箱數：箱數
       dg_Tab3_SelectedOrderDetail.Col = 9
       dg_Tab3_SelectedOrderDetail.Text = ""
       dg_Tab3_SelectedOrderDetail.Col = 7
       dg_Tab3_SelectedOrderDetail.Text = txt_Tab3_CutCaseQty.Text
    End If
    '計算選取之訂單細項之加總 [箱數] [重量] [才積] [板數]
    Call Calculate_Tab3_SelectedPrderDetail
    
    '清除切割量欄位值
    txt_Tab3_CutCaseQty.Text = ""
    
    cmd_Tab3_ClearQty.Enabled = True
    cmd_Tab3_CutOrders.Enabled = True
End Sub

Private Sub cmd_Tab3_DelOrders_Click()
    '訂單切割明細 >> 刪除
    Dim dbDeleteRow As Double, strOrderkey As String, strStorerkey As String, strExtern As String
    strOrderkey = Trim(txt_Tab3_OrderKey.Text)      '訂單編號 Receipt_No
    strExtern = Trim(txt_Tab3_Extern.Text)        '貨主單號 Extern
    
    msg_text = "刪除作業：確認刪除選取之子訂單：" & strOrderkey
    If MsgBox(msg_text, vbOKCancel + vbInformation, msg_title) = vbCancel Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    
    '檢核欲刪除之訂單：以貨主單號為查詢條件
    str_SQL = "Select Count(*) as RecCount From SDN02W Where Extern = '" & strExtern & "' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields("RecCount").Value = 1 Then
        tmp_Rs.Close
        msg_text = "訂單編號：" & strOrderkey & " 不允許刪除，因其貨主單號只對應ㄧ筆訂單資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf tmp_Rs.Fields("RecCount").Value = 0 Then
        tmp_Rs.Close
        msg_text = "訂單編號：" & strOrderkey & " 已不存在，請重新執行查詢"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    tmp_Rs.Close
    
    '取的最小訂單編號：接收被刪除訂單所有之項目、數量
    Dim strToOrderKey As String
    str_SQL = "Select Min(Receipt_No) as 接收訂單編號 From SDN02W Where Extern = '" & strExtern & "'  and Receipt_No <> '" & strOrderkey & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        strToOrderKey = tmp_Rs.Fields("接收訂單編號").Value
    Else
        tmp_Rs.Close
        msg_text = "貨主單號找不到對應之訂單編號可以接收欲刪除之訂單項次"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    tmp_Rs.Close
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    '更新接收訂單之相關資料 TRP02W
    str_SQL = "Update SDN02W Set SHIP_CS=SHIP_CS+" & Val(txt_Tab3_CaseQty.Text) & ",SHIP_CBM=SHIP_CBM+ " & Val(txt_Tab3_Volumn.Text) & ", " & _
           "SHIP_WT=SHIP_WT+" & Val(txt_Tab3_Weight.Text) & "Where  Receipt_No = '" & strToOrderKey & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    
    '更新接收訂單之相關資料 TRP03W
    
    Do While Not rs_Tab3_SDN03W.EOF
        '找找看接收訂單編號有無相同項次、貨號的訂單細項 SDN03W
        str_SQL = "Select Count(*) AS RecCount From SDN03W " & _
                  "Where  Receipt_No = '" & strToOrderKey & "' and " & _
                  "      Seq_No = " & Trim(rs_Tab3_SDN03W.Fields("項次").Value) & " and Product_No = '" & Trim(rs_Tab3_SDN03W.Fields("貨號").Value) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.Fields("RecCount").Value = 0 Then
           '新增細項 SDN03W
           str_SQL = "Insert into SDN03W (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME) " & _
                     "Select C_ROUTE_NO,ROUTE_NO,StorerKey,'" & strToOrderKey & "',Seq_No,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME " & _
                     "From SDN03W Where  Receipt_No = '" & strOrderkey & "' and " & _
                     "      Seq_No = " & Trim(rs_Tab3_SDN03W.Fields("項次").Value) & " and Product_No = '" & Trim(rs_Tab3_SDN03W.Fields("貨號").Value) & "'"
        Else
           '更新細項 TRP03W
           str_SQL = "Update SDN03W Set Order_Qty = Order_Qty + " & Trim(rs_Tab3_SDN03W.Fields("訂單箱數").Value) & ",SHIP_QTY=SHIP_QTY+" & Trim(rs_Tab3_SDN03W.Fields("揀貨箱數").Value) & " " & _
                     "Where  Receipt_No = '" & strToOrderKey & "' and " & _
                     "      Seq_No = " & Trim(rs_Tab3_SDN03W.Fields("項次").Value) & " and Product_No = '" & Trim(rs_Tab3_SDN03W.Fields("貨號").Value) & "'"
        End If
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        tmp_Rs.Close
        rs_Tab3_SDN03W.MoveNext
    Loop
    '刪除細項
    rs_Tab3_SDN03W.MoveFirst
    Do While Not rs_Tab3_SDN03W.EOF
        rs_Tab3_SDN03W.Delete
        rs_Tab3_SDN03W.MoveFirst
    Loop
    str_SQL = "Delete From SDN03W Where  Receipt_No = '" & strOrderkey & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rs_Tab3_SDN03W.Filter = adFilterNone
    rs_Tab3_SDN03W.Sort = "項次 ASC"  '原始排序，一般資料序號由小至大

    '刪除訂單主檔 TRP02W
    str_SQL = "Delete From SDN02W Where  Receipt_No = '" & strOrderkey & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    cn.CommitTrans
    Tran_Level = 0
    txt_Tab3_OrderKey.Text = ""
    txt_Tab3_DeliveryDate.Text = ""
    txt_Tab3_Extern.Text = ""  '客戶編號
    txt_Tab3_CaseQty.Text = "" '箱數
    txt_Tab3_Weight.Text = ""    '重量
    txt_Tab3_Volumn.Text = ""  '材積
    txt_Tab3_FullName.Text = ""   '客戶名稱
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-出車確認-Tab3訂單刪除", Me.Caption, "cmd_Tab3_DelOrders", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_DisplayOrders_Click()
    '出車確認 >> 顯示待排車訂單
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab3_SDN02W.DataSource = Nothing
    Set rs_Tab3_SDN02W = Nothing
    On Error GoTo err_Handle
    str_SQL = "SELECT  C_ROUTE_NO AS 二次排車, ROUTE_NO AS 路線編號,EXTERN AS 客戶單號,ARRIVE_DATE AS 日期,CUST_NAME as 指送客戶, " & _
            "SHIP_CS As 箱數, SHIP_CBM As 材積, SHIP_WT As 重量, RECEIPT_NO As 訂單號碼, CAR_NOTES As 多車 " & _
            "FROM dbo.SDN02W Order by 二次排車,路線編號,客戶單號,箱數 "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        Screen.MousePointer = vbDefault
        msg_text = "查詢結果：無符合設定條件之待排車訂單資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_SDN02W)
    tmp_Rs.Close
    rs_Tab3_SDN02W.MoveFirst
    Set dg_Tab3_SDN02W.DataSource = rs_Tab3_SDN02W
    With dg_Tab3_SDN02W
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000      '二次排車
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '客戶單號
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800      '路線編號
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800       '日期
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1200      '指送客戶
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 800       '箱數
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 800       '材積
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '重量
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000       '訂單號碼
        .Columns(9).Alignment = dbgLeft
        .Columns(10).Width = 800       '多車
        .Columns(10).Alignment = dbgLeft
    End With
    
'    '清欄位值
    Call SetGrid_Format_Tab3_SelectedOrderDetail
    Call Clear_CutOrderDetail
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單列表-顯示待排車訂單", Me.Caption, "cmd_Tab0_DisplayOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_DisplaySelectedOrder_Click()
    '訂單列表 >> 顯示訂單明細
    If rs_Tab3_SDN02W Is Nothing Then Exit Sub
    If rs_Tab3_SDN02W.RecordCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    '設定已完成切割訂單明細表
    Call Clear_CutOrderDetail
    On Error GoTo err_Handle
    DoEvents: DoEvents
    '設定欲切割訂單之訂單名細
    CutOrderkey = Trim(rs_Tab3_SDN02W.Fields("訂單號碼").Value)
    Call SetGrid_Format_Tab3_SelectedOrderDetail
    
    str_SQL = "Select rtrim(SEQ_NO) as 項次,rtrim(PRODUCT_NO) as 貨號,rtrim(sp.Descr) as 品名,case when sp.casecnt = 0 then 0 else isnull (SHIP_QTY/sp.casecnt,0) end as 箱數,(isnull(SHIP_QTY,0)*sp.Stdgrosswgt) as 重量,(isnull(SHIP_QTY,0)*sp.STDCUBE) as 材積 " & _
            "from SDN03W inner join gv_skuxpack sp on sp.sku=PRODUCT_NO and sp.storerkey = sdn03w.storerkey " & _
            "where RECEIPT_NO='" & CutOrderkey & "' order by SEQ_NO"
            
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無符合設定條件之待排車訂單明細資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Exit Sub
    End If
    Do While Not tmp_Rs.EOF
        With dg_Tab3_SelectedOrderDetail
             .Rows = .Rows + 1
             .Row = .Rows - 2
             .Col = 0    '訂單明細項次
             .Text = tmp_Rs.Fields("項次").Value
             .Col = 2    '貨號
             .Text = tmp_Rs.Fields("貨號").Value
             .Col = 3    '品名
             .Text = tmp_Rs.Fields("品名").Value
             .Col = 4    '箱數
             .Text = tmp_Rs.Fields("箱數").Value
             .Col = 5    '重量
             .Text = tmp_Rs.Fields("重量").Value
             .Col = 6    '材積
             .Text = tmp_Rs.Fields("材積").Value
        End With
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    Set tmp_Rs = Nothing
    Screen.MousePointer = vbDefault
    dbCut_TotalCaseQty = 0
    dbCut_TotalWeight = 0
    dbCut_TotalVolumn = 0
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單列表-顯示訂單名細", Me.Caption, "cmd_Tab0_DisplaySelectedOrder_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   If Not (tmp_Rs Is Nothing) Then
      Set tmp_Rs = Nothing
   End If
End Sub

Private Sub cmd_Tab3_Query_Click()
    If Len(Trim(txt_Tab3_OrderKey.Text)) = 0 Then Exit Sub
    On Error GoTo err_Handle
    str_SQL = "SELECT  EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,RECEIPT_NO " & _
            "from SDN02W where  RECEIPT_NO= '" & Trim(txt_Tab3_OrderKey.Text) & "' "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        Screen.MousePointer = vbDefault
        msg_text = "查詢結果：無符合設定條件之訂單資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    txt_Tab3_DeliveryDate.Text = tmp_Rs.Fields("ARRIVE_DATE").Value    '送貨日
    txt_Tab3_Extern.Text = tmp_Rs.Fields("EXTERN").Value  '客戶編號
    txt_Tab3_CaseQty.Text = tmp_Rs.Fields("SHIP_CS").Value '箱數
    txt_Tab3_Weight.Text = tmp_Rs.Fields("SHIP_WT").Value   '重量
    txt_Tab3_Volumn.Text = tmp_Rs.Fields("SHIP_CBM").Value    '材積
    txt_Tab3_FullName.Text = tmp_Rs.Fields("CUST_NAME").Value    '客戶名稱
    tmp_Rs.Close
    Call Dispaly_dg_Tab3_SDN03W
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認>>Tab3查詢訂單", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab2_SelectCar.Name & "2")
End Sub

Private Sub dg_SDN02W_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_SDN02W.DataSource Is Nothing Then Exit Sub
    If Tab1_RouteListEventEnable = False Then Exit Sub
    '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
    If Trim(rs_SDN02W.Fields(1).Value) = "" Then
        rs_SDN02W.Fields(1).Value = "V"
        txt_Tab1_srcSelected_Case.Text = Val(txt_Tab1_srcSelected_Case.Text) + Val(rs_SDN02W.Fields("箱數").Value)
        txt_Tab1_srcSelected_Volumn.Text = Val(txt_Tab1_srcSelected_Volumn.Text) + Val(rs_SDN02W.Fields("材積").Value)
        txt_Tab1_srcSelected_Weight.Text = Val(txt_Tab1_srcSelected_Weight.Text) + Val(rs_SDN02W.Fields("重量").Value)
    Else
        rs_SDN02W.Fields(1).Value = " "
        txt_Tab1_srcSelected_Case.Text = Val(txt_Tab1_srcSelected_Case.Text) - Val(rs_SDN02W.Fields("箱數").Value)
        txt_Tab1_srcSelected_Volumn.Text = Val(txt_Tab1_srcSelected_Volumn.Text) - Val(rs_SDN02W.Fields("材積").Value)
        txt_Tab1_srcSelected_Weight.Text = Val(txt_Tab1_srcSelected_Weight.Text) - Val(rs_SDN02W.Fields("重量").Value)
    End If
End Sub

Private Sub dg_Tab0_C_RouteList_HeadClick(ByVal ColIndex As Integer)

If dg_Tab0_C_RouteList.Row = -1 Then Exit Sub

Tab0_RouteListEventEnable = False

If intColumnIndex = ColIndex Then
    rs_Tab0_C_RouteList.Sort = dg_Tab0_C_RouteList.Columns(ColIndex).Caption & " DESC"
    dg_Tab0_C_RouteList.ClearSelCols
    intColumnIndex = 255

Else
    rs_Tab0_C_RouteList.Sort = dg_Tab0_C_RouteList.Columns(ColIndex).Caption
    dg_Tab0_C_RouteList.ClearSelCols
    intColumnIndex = ColIndex

End If

Tab0_RouteListEventEnable = True

End Sub

Private Sub dg_Tab0_C_RouteList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If Tab0_RouteListEventEnable = False Then Exit Sub
    
    '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
    
    If Trim(rs_Tab0_C_RouteList.Fields(1).Value) = "" Then
        rs_Tab0_C_RouteList.Fields(1).Value = "V"
    Else
        rs_Tab0_C_RouteList.Fields(1).Value = " "
    End If
End Sub

Private Sub dg_Tab0_RouteList0_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If Tab0_RouteListEventEnable = False Then Exit Sub
    '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
    If Trim(rs_Tab0_RouteList0.Fields(1).Value) = "" Then
        rs_Tab0_RouteList0.Fields(1).Value = "V"
    Else
        rs_Tab0_RouteList0.Fields(1).Value = " "
    End If
End Sub

Private Sub dg_Tab0_RouteList1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab0_RouteList1.DataSource Is Nothing Then Exit Sub
    If Tab0_RouteListEventEnable = False Then Exit Sub
    '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
    If Trim(rs_Tab0_RouteList1.Fields(1).Value) = "" Then
        rs_Tab0_RouteList1.Fields(1).Value = "V"
    Else
        rs_Tab0_RouteList1.Fields(1).Value = " "
    End If
End Sub

Private Sub dg_Tab1_SelectedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab1_SelectedOrders.DataSource Is Nothing Then Exit Sub
    If rs_Tab1_SelectedOrders.RecordCount = 0 Then Exit Sub
    If Tab1_RouteListEventEnable = False Then Exit Sub
    '點選即表示選取，取消選取以其他 Button 專門處理：因為取消選取不方便
    If Trim(rs_Tab1_SelectedOrders.Fields(1).Value) = "" Then
        rs_Tab1_SelectedOrders.Fields(1).Value = "V"
    Else
        rs_Tab1_SelectedOrders.Fields(1).Value = " "
       
    End If
End Sub


Private Sub dg_Tab2_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Tab2_RouteListEventEnable = False Then Exit Sub
    txt_Tab2_Route.Text = rs_Tab2_Route.Fields("二次排車").Value
    txt_Tab2_VehicleNo.Text = rs_Tab2_Route.Fields("車號").Value
    txt_Tab2_Driver.Text = rs_Tab2_Route.Fields("司機").Value
    txt_Tab2_DELIVERY_DATE.Text = rs_Tab2_Route.Fields("日期").Value
    Call Display_Tab2_RouteOrders
End Sub

Private Sub dg_Tab3_SelectedOrderDetail_Click()
    '待切割之訂單：訂單明細項次
    '點一次：選取，除非清除 [切割數量] 否則ㄧ直保持 [選取] 狀態
    txt_Tab3_CutCaseQty.Text = ""
    Dim i As Integer
    Dim tmpQty As Double
    With dg_Tab3_SelectedOrderDetail
        .Col = 2   '貨號
        If Len(Trim(.Text)) = 0 Then Exit Sub
        .Col = 1
        If Len(.Text) = 0 Then
            .Text = "V"
            .Col = 4   '顯示所選取之箱數
            tmpQty = .Text
            dbCut_TotalCaseQty = dbCut_TotalCaseQty + .Text
            txt_Tab3_SelectedCaseQty.Text = dbCut_TotalCaseQty
            .Col = 7: .Text = tmpQty
            txt_Tab3_CutCaseQty.Text = tmpQty
            
            .Col = 5   '顯示所選取之重量
            tmpQty = .Text
            dbCut_TotalWeight = dbCut_TotalWeight + .Text
            txt_Tab3_SelectedWeight.Text = dbCut_TotalWeight
            .Col = 8: .Text = tmpQty
            
            .Col = 6   '顯示所選取之材積
            tmpQty = .Text
            dbCut_TotalVolumn = dbCut_TotalVolumn + .Text
            txt_Tab3_SelectedVolumn.Text = dbCut_TotalVolumn
            .Col = 9: .Text = tmpQty
        Else
            .Col = 7   '切割之箱數
            If Val(.Text) <> 0 Then
                txt_Tab3_CutCaseQty.Text = .Text
            End If
        End If
        '反白選取之資料行
        .Col = 0
        For i = 0 To .Cols - 1
            .ColSel = i
        Next i
    End With
End Sub

Private Sub Form_Load()
    '設定 Form 大小、位置
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 11475
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    SSTab1.Tab = 0
    Tab2_RouteListEventEnable = True
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    If Me.ScaleHeight < dbsrcFormHeight Then
        '變小
        SSTab1.Top = (SSTab1.Top - ((dbsrcFormHeight - Me.ScaleHeight) / 2))
        SSTab1.Left = (SSTab1.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2))
          
        dbsrcFormHeight = Me.ScaleHeight
        dbsrcFormWidth = Me.ScaleWidth
    Else
        SSTab1.Top = (SSTab1.Top + ((Me.ScaleHeight - dbsrcFormHeight) / 2))
        SSTab1.Left = (SSTab1.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2))
        
        dbsrcFormHeight = Me.ScaleHeight
        dbsrcFormWidth = Me.ScaleWidth
    End If
End Sub

Private Sub ReSet_Tab0_RouteList1_SeqNo()
    '重新產生 [dg_Tab0_RouteList0] 之 [編號] 欄位值
    dg_Tab0_RouteList1.Visible = False
    rs_Tab0_RouteList1.Filter = adFilterNone
    rs_Tab0_RouteList1.Sort = "客戶單號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_Tab0_RouteList1.EOF Then rs_Tab0_RouteList1.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab0_RouteList1.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_RouteList1.Fields("編號").Value = dbSeqNo
        rs_Tab0_RouteList1.MoveNext
    Loop
    If rs_Tab0_RouteList1.RecordCount > 0 Then rs_Tab0_RouteList1.MoveFirst
    dg_Tab0_RouteList1.Visible = True
End Sub

Private Sub ReSet_Tab0_RouteList0_SeqNo()
    '重新產生 [dg_Tab0_RouteList1] 之 [編號] 欄位值
    dg_Tab0_RouteList0.Visible = False
    rs_Tab0_RouteList0.Filter = adFilterNone
    rs_Tab0_RouteList0.Sort = "客戶單號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_Tab0_RouteList0.EOF Then rs_Tab0_RouteList0.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab0_RouteList0.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_RouteList0.Fields("編號").Value = dbSeqNo
        rs_Tab0_RouteList0.MoveNext
    Loop
    If rs_Tab0_RouteList0.RecordCount > 0 Then rs_Tab0_RouteList0.MoveFirst
    dg_Tab0_RouteList0.Visible = True
End Sub

Private Sub ReSet_Tab0_C_RouteList_SeqNo()
    '重新產生 [dg_Tab0_RouteList1] 之 [編號] 欄位值
    dg_Tab0_C_RouteList.Visible = False
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "路線編號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_Tab0_C_RouteList.EOF Then rs_Tab0_C_RouteList.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab0_C_RouteList.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_C_RouteList.Fields("編號").Value = dbSeqNo
        rs_Tab0_C_RouteList.MoveNext
    Loop
    If rs_Tab0_C_RouteList.RecordCount > 0 Then rs_Tab0_C_RouteList.MoveFirst
    dg_Tab0_C_RouteList.Visible = True
End Sub

Private Sub ReSet_Tab1_SelectedOrders_SeqNo()
    '重新產生 [dg_Tab0_RouteList1] 之 [編號] 欄位值
    dg_Tab1_SelectedOrders.Visible = False
    rs_Tab1_SelectedOrders.Filter = adFilterNone
    rs_Tab1_SelectedOrders.Sort = "路線編號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_Tab1_SelectedOrders.EOF Then rs_Tab1_SelectedOrders.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab1_SelectedOrders.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab1_SelectedOrders.Fields("編號").Value = dbSeqNo
        rs_Tab1_SelectedOrders.MoveNext
    Loop
    If rs_Tab1_SelectedOrders.RecordCount > 0 Then rs_Tab1_SelectedOrders.MoveFirst
    dg_Tab1_SelectedOrders.Visible = True
    
End Sub

Private Sub ReSet_Tab1_SDN02W_SeqNo()
    '重新產生 [dg_Tab0_RouteList1] 之 [編號] 欄位值
    dg_SDN02W.Visible = False
    rs_SDN02W.Filter = adFilterNone
    rs_SDN02W.Sort = "路線編號 asc"  '原始排序，一般資料序號由小至大
    If Not rs_SDN02W.EOF Then rs_SDN02W.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_SDN02W.EOF
        dbSeqNo = dbSeqNo + 1
        rs_SDN02W.Fields("編號").Value = dbSeqNo
        rs_SDN02W.MoveNext
    Loop
    If rs_SDN02W.RecordCount > 0 Then rs_SDN02W.MoveFirst
    dg_SDN02W.Visible = True
    
End Sub

Private Sub Retrive_OrderSum()
    '取的待排車訂單：總計資料值
    txt_Tab1_srcTotal_Case.Text = ""
    txt_Tab1_srcTotal_Volumn.Text = ""
    txt_Tab1_srcTotal_Weight.Text = ""
    'SHIP_CS,SHIP_CBM,SHIP_WT
    str_SQL = "Select Isnull(Round(sum(SHIP_CS),0),0) as 總箱數,Isnull(Round(sum(SHIP_WT),0),0) as 總重量," & _
              "       Isnull(Round(sum(SHIP_CBM),0),0) as 總材積 " & _
              "From SDN02W  "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
       txt_Tab1_srcTotal_Case.Text = tmp_Rs.Fields("總箱數").Value
       txt_Tab1_srcTotal_Volumn.Text = tmp_Rs.Fields("總材積").Value
       txt_Tab1_srcTotal_Weight.Text = tmp_Rs.Fields("總重量").Value
    End If
    tmp_Rs.Close
End Sub


Private Sub CreateRS_Tab1_SelectedOrders()
    '排車作業：已選取之待排車訂單列表
    Set dg_Tab1_SelectedOrders.DataSource = Nothing
    Call ReDim_Recordset(rs_Tab1_SelectedOrders)
    With rs_Tab1_SelectedOrders
         .Fields.Append "編號", adDouble
         .Fields.Append "＊", adVarChar, 5
         .Fields.Append "二次排車", adVarChar, 10
         .Fields.Append "路線編號", adVarChar, 10
         .Fields.Append "客戶單號", adVarChar, 30
         .Fields.Append "日期", adVarChar, 20
         .Fields.Append "指送客戶", adVarChar, 60
         .Fields.Append "箱數", adDouble
         .Fields.Append "材積", adDouble
         .Fields.Append "重量", adDouble
         .Fields.Append "訂單號碼", adVarChar, 20
         .Fields.Append "多車", adVarChar, 50
         .Fields.Append "C_Receipt_No", adVarChar, 20
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '不需連接物件
    End With
    Set dg_Tab1_SelectedOrders.DataSource = rs_Tab1_SelectedOrders
    '設定顯示欄位
    With dg_Tab1_SelectedOrders
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 500       '序號
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Width = 500       '選取
            .Columns(1).Alignment = dbgCenter
            .Columns(2).Width = 1000      '二次排車
            .Columns(2).Alignment = dbgLeft
            .Columns(3).Width = 1000      '客戶單號
            .Columns(3).Alignment = dbgLeft
            .Columns(4).Width = 800      '路線編號
            .Columns(4).Alignment = dbgLeft
            .Columns(5).Width = 800       '日期
            .Columns(5).Alignment = dbgLeft
            .Columns(6).Width = 1500      '指送客戶
            .Columns(6).Alignment = dbgLeft
            .Columns(7).Width = 800       '箱數
            .Columns(7).Alignment = dbgRight
            .Columns(8).Width = 800       '材積
            .Columns(8).Alignment = dbgRight
            .Columns(9).Width = 800       '重量
            .Columns(9).Alignment = dbgRight
            .Columns(10).Width = 1000       '訂單號碼
            .Columns(10).Alignment = dbgLeft
            .Columns(11).Width = 800       '多車
            .Columns(11).Alignment = dbgLeft
    End With
End Sub


Private Sub SetGrid_Format_Tab3_SelectedOrderDetail()
    '選取作為待切割訂單之項目明細
    Dim sub_var1 As Integer, sub_var2 As Integer
    dg_Tab3_SelectedOrderDetail.Visible = False
    With dg_Tab3_SelectedOrderDetail
         .Rows = 2: .Cols = 10
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
         .ColWidth(0) = 1000
         .ColWidth(1) = 300
         .ColWidth(2) = 1500
         .ColWidth(3) = 2200
         .ColWidth(4) = 600
         .ColWidth(5) = 1000
         .ColWidth(6) = 1000
         .ColWidth(7) = 850
         .ColWidth(8) = 1000
         .ColWidth(9) = 1000

         '設定列表之標題
         .Row = 0
         .Col = 0: .Text = "項次"
         .Col = 1: .Text = "※"
         .Col = 2: .Text = "貨號"
         .Col = 3: .Text = "品名"
         .Col = 4: .Text = "箱數"
         .Col = 5: .Text = "重量"
         .Col = 6: .Text = "材積"
         .Col = 7: .Text = "切割箱數"
         .Col = 8: .Text = "切割重量"
         .Col = 9: .Text = "切割材積"
         '設定列表之文字對齊
         .ColAlignment(0) = flexAlignLeftCenter
         .ColAlignment(1) = flexAlignCenterCenter
         .ColAlignment(2) = flexAlignLeftCenter
         .ColAlignment(3) = flexAlignLeftCenter
         .ColAlignment(4) = flexAlignRightCenter
         .ColAlignment(5) = flexAlignLeftCenter
         .ColAlignment(6) = flexAlignLeftCenter
         .ColAlignment(7) = flexAlignRightCenter
         .ColAlignment(8) = flexAlignLeftCenter
         .ColAlignment(9) = flexAlignLeftCenter
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1
             .CellAlignment = flexAlignCenterCenter
         Next sub_var1
         .Rows = 2: .Row = 1
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1: .Text = ""
         Next sub_var1
    End With
    dg_Tab3_SelectedOrderDetail.Visible = True
End Sub

Private Sub Calculate_Tab3_SelectedPrderDetail()
    '計算選取之訂單細項：箱數，重量，才積，板數
    dbCut_TotalCaseQty = 0
    txt_Tab3_SelectedCaseQty.Text = ""
    dbCut_TotalWeight = 0
    txt_Tab3_SelectedWeight.Text = ""
    dbCut_TotalVolumn = 0
    txt_Tab3_SelectedVolumn.Text = ""
    
    Dim dbCaseQty As Double, dbWeight As Double, dbVolumn As Double, dbPalletQty As Double
    Dim dbCutPLQty As Double, dbCutCSQty As Double
    Dim i As Double
    With dg_Tab3_SelectedOrderDetail

        For i = 1 To .Rows - 2
            .Row = i
            .Col = 1
            If .Text <> "" Then   '被選取
                .Col = 4: dbCaseQty = Val(.Text)     '箱數
                .Col = 5: dbWeight = Val(.Text)      '重量
                .Col = 6: dbVolumn = Val(.Text)      '材積
                .Col = 7   '切割箱數
                If Val(.Text) <> 0 Then
                     dbCutCSQty = Val(.Text)
                     dbCut_TotalCaseQty = dbCut_TotalCaseQty + dbCutCSQty
                    .Col = 8   '切割箱數換算之重量
                    .Text = ((dbCutCSQty / dbCaseQty) * dbWeight)
                     dbCut_TotalWeight = dbCut_TotalWeight + ((dbCutCSQty / dbCaseQty) * dbWeight)
                    .Col = 9   '切割箱數換算之材積
                    .Text = ((dbCutCSQty / dbCaseQty) * dbVolumn)
                     dbCut_TotalVolumn = dbCut_TotalVolumn + ((dbCutCSQty / dbCaseQty) * dbVolumn)
                End If
            Else
                .Col = 7: .Text = ""
                .Col = 8: .Text = ""
                .Col = 9: .Text = ""
            End If
        Next i
    End With
    '顯示選取之細項各欄位之加總值
    txt_Tab3_SelectedCaseQty.Text = dbCut_TotalCaseQty
    txt_Tab3_SelectedWeight.Text = dbCut_TotalWeight
    txt_Tab3_SelectedVolumn.Text = dbCut_TotalVolumn
End Sub

Private Sub Delete_GridRow(ByVal intRow As Double)
    '待切割訂單項次(Detail) 資料刪除
    If intRow = 0 Then Exit Sub
    Dim i As Double, j As Integer
    '1. 將刪除列資料由下一列資料取代
    '   而後的資料列往上移一列
    With dg_Tab3_SelectedOrderDetail
        For i = intRow To .Rows - 2   '會有多一行空白列
            .Row = i
            For j = 0 To .Cols - 1
                .Col = j
                .Text = .TextArray((.Row + 1) * .Cols + .Col)
            Next j
            '防止最後第一列往上移給最後第二列時，會是弄白資料列，[序號] 欄位不能有值
            '有資料的列，[序號] 必須重新編號
            .Col = 0
            If Val(.Text) = 0 Then .Text = ""   'Else .Text = .Row
        Next i
        '2. Grid 總列數 - 1
        .Rows = .Rows - 1
        .Row = 1
        For i = 0 To .Cols - 1
            .ColSel = i
        Next i
    End With
End Sub

Private Sub Dispaly_dg_Tab3_SDN03W()
    '出車確認 >> Tab3顯示新增訂單明細訂單
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab3_SDN03W.DataSource = Nothing
    Set rs_Tab3_SDN03W = Nothing
    On Error GoTo err_Handle
    
    str_SQL = "SELECT  SEQ_NO as 項次 " & _
        ",PRODUCT_NO as 貨號 " & _
        ",sp.Descr as 品名 " & _
        ",isnull(case when sp.casecnt = 0 then 0 else SHIP_QTY/sp.casecnt end ,0) as 揀貨箱數 " & _
        ",isnull(case when sp.casecnt = 0 then 0 else ORDER_QTY/sp.casecnt end ,0) as 訂單箱數 " & _
        ",(isnull(SHIP_QTY,0)*sp.Stdgrosswgt) as 揀貨重量 " & _
        ",(isnull(SHIP_QTY,0)*sp.STDCUBE) as 揀貨材積 " & _
        "from SDN03W inner join gv_skuxpack sp on sp.sku=PRODUCT_NO and sp.storerkey = sdn03w.storerkey " & _
        "where RECEIPT_NO= '" & Trim(txt_Tab3_OrderKey.Text) & "' order by SEQ_NO"
        
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '無限期等待
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無符合設定條件之待排車訂單資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_SDN03W)
    tmp_Rs.Close
    rs_Tab3_SDN03W.MoveFirst
    Set dg_Tab3_SDN03W.DataSource = rs_Tab3_SDN03W
    With dg_Tab3_SDN03W
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
        .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
        .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
        .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500      '項次
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '貨號
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 2500      '品名
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1000      '揀貨箱數
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 1000       '訂單箱數
        .Columns(5).Alignment = dbgRight
        .Columns(6).Width = 1000      '重量
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 1000       '材積
        .Columns(7).Alignment = dbgRight
    End With
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "出車確認-Tab3顯示訂單明細", Me.Caption, "Dispaly_dg_Tab3_SDN03W", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub Clear_CutOrderDetail()
    Set dg_Tab3_SDN03W.DataSource = Nothing
    txt_Tab3_OrderKey.Text = ""
    txt_Tab3_DeliveryDate.Text = ""
    txt_Tab3_Extern.Text = ""  '客戶編號
    txt_Tab3_CaseQty.Text = "" '箱數
    txt_Tab3_Weight.Text = ""    '重量
    txt_Tab3_Volumn.Text = ""  '材積
    txt_Tab3_FullName.Text = ""   '客戶名稱
End Sub

Private Sub Delete_RouteNo(strRouteNo As String)
    Screen.MousePointer = vbHourglass
    blTab1RouteEventEnable = False
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
        On Error GoTo err_Handle
        '刪除 TRP01T 路線編號主檔
        Call DB_CheckConnectStatus
        
        '(1).將 SDN03T 寫回 SDN03W >> 刪除 SDN03T
        str_SQL = "Insert into SDN03W( " & _
                  "C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME) " & _
                  "Select C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME " & _
                  "From SDN03T  Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '(2).將 SDN02T 寫回 SDN02W >> 刪除 SDN02T
        str_SQL = "Insert into SDN02W( " & _
                  "C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,c_receipt_no) " & _
                  "Select C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,c_receipt_no " & _
                  "From SDN02T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '(3).刪除 SDN03T & SDN02T & SDN01T
        str_SQL = "Delete From SDN03T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Delete From SDN02T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Delete From SDN01T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '刪除運費
        str_SQL = "Delete From SDN05T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
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
   CreateErrorLog Me.Name & "出車確認-路線編號刪除", Me.Caption, "Form 內部 SubProgram Delete_RouteNo", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub



Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    '日期選取
    Select Case mvDate.Tag
        Case "Tab0_DELIVERY_DATE0"
             txt_DELIVERY_DATE0.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab0_DELIVERY_DATE1"
             txt_DELIVERY_DATE1.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab1_DELIVERY_DATE"
             txt_Tab1_DELIVERY_DATE.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab2_DELIVERY_DATE"
             txt_Tab2_DELIVERY_DATE.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab2_DELIVERYDATE_START"
             txt_Tab2_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
        Case Else
    End Select
    mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    mvDate.Visible = False
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_DELIVERY_DATE0_Click()
    'Tab0 >> 出車日期
    If Trim(txt_DELIVERY_DATE0.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_DELIVERY_DATE0.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_DELIVERY_DATE0.Text, 4) & "/" & Mid(txt_DELIVERY_DATE0.Text, 5, 2) & "/" & Right(txt_DELIVERY_DATE0.Text, 2))
       End If
    End If
    mvDate.Tag = "Tab0_DELIVERY_DATE0"
    mvDate.Top = SSTab1.Top + fam_Tab0_Consignee.Top + txt_DELIVERY_DATE0.Top + txt_DELIVERY_DATE0.Height
    mvDate.Left = SSTab1.Left + fam_Tab0_Consignee.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub



Private Sub txt_DELIVERY_DATE1_Click()
    'Tab0 >> 出車日期
    If Trim(txt_DELIVERY_DATE1.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_DELIVERY_DATE1.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_DELIVERY_DATE1.Text, 4) & "/" & Mid(txt_DELIVERY_DATE1.Text, 5, 2) & "/" & Right(txt_DELIVERY_DATE1.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab0_DELIVERY_DATE1"
    mvDate.Top = SSTab1.Top + Frame1.Top + txt_DELIVERY_DATE1.Top + txt_DELIVERY_DATE1.Height
    mvDate.Left = SSTab1.Left + fam_Tab0_Consignee.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DELIVERY_DATE_Click()
    'Tab0 >> 出車日期
    If Trim(txt_Tab1_DELIVERY_DATE.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab1_DELIVERY_DATE.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab1_DELIVERY_DATE.Text, 4) & "/" & Mid(txt_Tab1_DELIVERY_DATE.Text, 5, 2) & "/" & Right(txt_Tab1_DELIVERY_DATE.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab1_DELIVERY_DATE"
    mvDate.Top = SSTab1.Top + fam_SelectedOrders.Top + txt_Tab1_DELIVERY_DATE.Top + txt_Tab1_DELIVERY_DATE.Height
    mvDate.Left = SSTab1.Left + fam_SelectedOrders.Left + txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub

Private Sub clear_Tab0_RouteList0()
    '畫面處理
    Set dg_Tab0_RouteList0.DataSource = Nothing
    txt_Tab0_C_Route_No0.Text = ""
    txt_DELIVERY_DATE0.Text = ""
    txt_VehicleNo0.Text = ""
    txt_Driver0.Text = ""
End Sub
Private Sub clear_Tab0_RouteList1()
    '畫面處理
    Set dg_Tab0_RouteList1.DataSource = Nothing
    txt_Tab0_C_Route_No1.Text = ""
    txt_DELIVERY_DATE1.Text = ""
    txt_VehicleNo1.Text = ""
    txt_Driver1.Text = ""
End Sub

Private Sub Display_Tab2_RouteOrders()
    'SDN03T
    If Tab2_RouteListEventEnable = False Then Exit Sub
    str_SQL = "SELECT  C_ROUTE_NO AS 二次排車, ROUTE_NO AS 路線編號,EXTERN AS 客戶單號,ARRIVE_DATE AS 日期,CUST_NAME as 指送客戶, " & _
            "SHIP_CS As 箱數, SHIP_CBM As 材積, SHIP_WT As 重量, RECEIPT_NO As 訂單號碼, CAR_NOTES As 多車 " & _
            "FROM   SDN02T Where C_ROUTE_NO = '" & Trim(rs_Tab2_Route.Fields("二次排車").Value) & "' Order by C_ROUTE_NO,Receipt_No"
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
    Call Replication_Recordset(tmp_Rs, rs_Tab2_RouteOrders)
    Set dg_Tab2_RouteOrders.DataSource = rs_Tab2_RouteOrders
    tmp_Rs.Close
    With dg_Tab2_RouteOrders
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000      '二次排車
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '客戶單號
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800      '路線編號
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800       '日期
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1600      '指送客戶
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 800       '箱數
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 800       '材積
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '重量
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000       '多車
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 800       '訂單號碼
        .Columns(10).Alignment = dbgRight
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-出車確認 >> Tab2路線編號查詢", Me.Caption, "cmd_Tab2_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub txt_Tab1_VehicleNo_LostFocus()

If Len(Trim(txt_Tab1_VehicleNo)) = 0 Then Exit Sub

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

tmp_Rs.Open "select driver=isnull(driver,'') from trp09m where Vehicle_id_No = '" & Trim(txt_Tab1_VehicleNo) & "' ", cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then MsgBox "無此車號!", 16, "注意": txt_Tab1_VehicleNo.SetFocus: Exit Sub

txt_Tab1_Driver0 = RTrim(tmp_Rs("driver")) & ""
tmp_Rs.Close

End Sub

Private Sub txt_Tab2_DELIVERY_DATE_Click()
    'Tab2 >> 出車日期
    If Trim(txt_Tab2_DELIVERY_DATE.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab2_DELIVERY_DATE.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab2_DELIVERY_DATE.Text, 4) & "/" & Mid(txt_Tab2_DELIVERY_DATE.Text, 5, 2) & "/" & Right(txt_Tab2_DELIVERY_DATE.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab2_DELIVERY_DATE"
    mvDate.Top = SSTab1.Top + fam_Tab0_Consignee.Top + txt_Tab2_DELIVERY_DATE.Top + txt_Tab2_DELIVERY_DATE.Height
    mvDate.Left = SSTab1.Left + Frame6.Left + txt_Tab2_DELIVERY_DATE.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub

Private Sub txt_Tab2_DeliveryDate_Start_Click()
    'Tab2 >> 出車日期
    If Trim(txt_Tab2_DeliveryDate_Start.Text) = "" Then
        mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_Start.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab2_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab2_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab2_DeliveryDate_Start.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab2_DELIVERYDATE_START"
    mvDate.Top = SSTab1.Top + fam_Tab0_Consignee.Top + txt_Tab2_DeliveryDate_Start.Top + txt_Tab2_DeliveryDate_Start.Height
    mvDate.Left = SSTab1.Left + Frame6.Left + txt_Tab2_DELIVERY_DATE.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub
