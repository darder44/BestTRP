VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_SDNAbnormal 
   Caption         =   "配送異常維護"
   ClientHeight    =   8925
   ClientLeft      =   135
   ClientTop       =   975
   ClientWidth     =   14160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   14160
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   7200
      TabIndex        =   125
      Top             =   6480
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
      StartOfWeek     =   97255425
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm_OP_SDNAbnormal.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "txt_Tab0_Route_No"
      Tab(0).Control(2)=   "Frame12"
      Tab(0).Control(3)=   "Frame13"
      Tab(0).Control(4)=   "Frame14"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frm_OP_SDNAbnormal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame9"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frm_OP_SDNAbnormal.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape1(4)"
      Tab(2).Control(1)=   "Label3(23)"
      Tab(2).Control(2)=   "Label3(24)"
      Tab(2).Control(3)=   "Label3(25)"
      Tab(2).Control(4)=   "Label3(26)"
      Tab(2).Control(5)=   "Label3(35)"
      Tab(2).Control(6)=   "Frame5"
      Tab(2).Control(7)=   "Frame1"
      Tab(2).Control(8)=   "Frame6"
      Tab(2).Control(9)=   "txt_Tab02_C_Route_No"
      Tab(2).Control(10)=   "txt_Tab02_Receiver"
      Tab(2).Control(11)=   "txt_Tab02_Driver"
      Tab(2).Control(12)=   "txt_Tab02_Delivery_Date"
      Tab(2).Control(13)=   "txt_Tab02_C_VEHICLE_ID_NO"
      Tab(2).Control(14)=   "cmd_Tab2_SelectCar"
      Tab(2).Control(15)=   "Frame7"
      Tab(2).Control(16)=   "Frame8"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "簽單明細確認"
      TabPicture(3)   =   "frm_OP_SDNAbnormal.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fra_MultiOrder_Header"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fra_MultiOrder_Detail"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "fra_OneOrder_Header"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "fra_Function"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fra_OneOrder_Detail"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame14 
         BackColor       =   &H80000004&
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   248
         Top             =   360
         Width           =   10095
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "離開"
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
            Index           =   2
            Left            =   8880
            Picture         =   "frm_OP_SDNAbnormal.frx":0070
            Style           =   1  '圖片外觀
            TabIndex        =   265
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT0 
            BackColor       =   &H00FFFFC0&
            Caption         =   "查詢"
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
            Left            =   7680
            Picture         =   "frm_OP_SDNAbnormal.frx":29C82
            Style           =   1  '圖片外觀
            TabIndex        =   264
            Top             =   120
            Width           =   1065
         End
         Begin VB.ComboBox cboCarT0 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5280
            TabIndex        =   262
            Text            =   "cboCarT0"
            Top             =   600
            Width           =   2325
         End
         Begin VB.ComboBox cboStorerT0 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5280
            TabIndex        =   260
            Text            =   "cboStorerT0"
            Top             =   240
            Width           =   2325
         End
         Begin VB.TextBox txtDeliveryDateST0 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   253
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtDeliveryDateET0 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   252
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtRouteET0 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   251
            Top             =   600
            Width           =   1365
         End
         Begin VB.TextBox txtRouteST0 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   250
            Top             =   600
            Width           =   1365
         End
         Begin VB.CommandButton cmdPickSave 
            BackColor       =   &H00FFFF80&
            Caption         =   "存檔"
            Enabled         =   0   'False
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
            Left            =   4680
            Picture         =   "frm_OP_SDNAbnormal.frx":29F8C
            Style           =   1  '圖片外觀
            TabIndex        =   249
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車號"
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
            Index           =   35
            Left            =   4560
            TabIndex        =   263
            Top             =   660
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   34
            Left            =   4560
            TabIndex        =   261
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   33
            Left            =   2535
            TabIndex        =   257
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   32
            Left            =   120
            TabIndex        =   256
            Top             =   285
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "二次路編"
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
            Left            =   120
            TabIndex        =   255
            Top             =   645
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Left            =   2535
            TabIndex        =   254
            Top             =   660
            Width           =   360
         End
      End
      Begin VB.Frame Frame13 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   246
         Top             =   4800
         Width           =   8295
         Begin VB.TextBox txtCustomerOrderkey 
            BackColor       =   &H0000FFFF&
            Height          =   270
            Left            =   2640
            TabIndex        =   266
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdOpenOrderT0 
            BackColor       =   &H00FFFFC0&
            Caption         =   "檢視明細"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1200
            Style           =   1  '圖片外觀
            TabIndex        =   259
            Top             =   180
            Width           =   960
         End
         Begin VB.CommandButton cmdDeliveryokT0 
            BackColor       =   &H00C0FFC0&
            Caption         =   "正常訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   258
            Top             =   180
            Width           =   960
         End
         Begin MSDataGridLib.DataGrid dgOrderT0 
            Height          =   2295
            Left            =   120
            TabIndex        =   247
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   20
            TabAction       =   1
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
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
      Begin VB.Frame Frame12 
         Caption         =   "Route"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   244
         Top             =   1440
         Width           =   8295
         Begin VB.CommandButton cmdTKPremiamAR 
            BackColor       =   &H00FFFF80&
            Caption         =   "TK議價應收分攤"
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
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   268
            ToolTipText     =   "僅針對台灣麒麟"
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdPremiamAP 
            BackColor       =   &H0080FFFF&
            Caption         =   "議價應付分攤"
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
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   267
            Top             =   240
            Width           =   1065
         End
         Begin MSDataGridLib.DataGrid dgRouteT0 
            Height          =   2295
            Left            =   1320
            TabIndex        =   245
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   20
            TabAction       =   1
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
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
      Begin VB.Frame Frame10 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   165
         Top             =   3000
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMain1 
            Height          =   2295
            Left            =   120
            TabIndex        =   166
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   20
            TabAction       =   1
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
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
      Begin VB.Frame Frame9 
         BackColor       =   &H80000004&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   -74880
         TabIndex        =   151
         Top             =   360
         Width           =   11295
         Begin VB.ComboBox cboStorerkey 
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            TabIndex        =   195
            Top             =   240
            Width           =   1605
         End
         Begin VB.TextBox txt2RouteS 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   192
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txt2RouteE 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   191
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtEarning 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   189
            Text            =   "0"
            Top             =   2160
            Width           =   1125
         End
         Begin VB.TextBox txtAR 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   187
            Text            =   "0"
            Top             =   1680
            Width           =   1125
         End
         Begin VB.TextBox txtAP 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   185
            Text            =   "0"
            Top             =   1920
            Width           =   1125
         End
         Begin VB.ComboBox cboCostkind 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            TabIndex        =   181
            Top             =   960
            Width           =   1605
         End
         Begin VB.TextBox txtSignDateE 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   178
            Top             =   2040
            Width           =   1485
         End
         Begin VB.TextBox txtSignDateS 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   177
            Top             =   2040
            Width           =   1485
         End
         Begin VB.ComboBox cboCar 
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            TabIndex        =   175
            Top             =   600
            Width           =   1605
         End
         Begin VB.TextBox txtRouteE 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   172
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox txtRouteS 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   171
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryE 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   168
            Top             =   1680
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryS 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   167
            Top             =   1680
            Width           =   1485
         End
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H00FFFFC0&
            Caption         =   "查詢"
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
            Left            =   7680
            Picture         =   "frm_OP_SDNAbnormal.frx":2A296
            Style           =   1  '圖片外觀
            TabIndex        =   160
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdReset 
            BackColor       =   &H00C0FFC0&
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
            Height          =   870
            Left            =   10080
            Picture         =   "frm_OP_SDNAbnormal.frx":2A5A0
            Style           =   1  '圖片外觀
            TabIndex        =   159
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "離開"
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
            Index           =   0
            Left            =   10080
            Picture         =   "frm_OP_SDNAbnormal.frx":2A8B2
            Style           =   1  '圖片外觀
            TabIndex        =   158
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmd2Excel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "轉Excel"
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
            Left            =   8880
            Picture         =   "frm_OP_SDNAbnormal.frx":544C4
            Style           =   1  '圖片外觀
            TabIndex        =   157
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtExternS 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   156
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtExternE 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            TabIndex        =   155
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToText 
            BackColor       =   &H00C0E0FF&
            Caption         =   "會計資料"
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
            Left            =   8880
            Picture         =   "frm_OP_SDNAbnormal.frx":557BE
            Style           =   1  '圖片外觀
            TabIndex        =   154
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtOrderkeyE 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   153
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtOrderkeyS 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   152
            Top             =   1320
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   29
            Left            =   4800
            TabIndex        =   196
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   28
            Left            =   2640
            TabIndex        =   194
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "二次路編"
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
            Left            =   120
            TabIndex        =   193
            Top             =   285
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "營收"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   26
            Left            =   4530
            TabIndex        =   190
            Top             =   2220
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "應收"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   25
            Left            =   4530
            TabIndex        =   188
            Top             =   1740
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "應付"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   24
            Left            =   4530
            TabIndex        =   186
            Top             =   1980
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "請款類別"
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
            Left            =   4560
            TabIndex        =   182
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "TMS單號"
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
            Left            =   120
            TabIndex        =   180
            Top             =   1380
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   21
            Left            =   2640
            TabIndex        =   179
            Top             =   2100
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車號"
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
            Left            =   4800
            TabIndex        =   176
            Top             =   660
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   20
            Left            =   120
            TabIndex        =   174
            Top             =   645
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   19
            Left            =   2640
            TabIndex        =   173
            Top             =   660
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "簽單日期"
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
            Left            =   120
            TabIndex        =   170
            Top             =   2085
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   5
            Left            =   2640
            TabIndex        =   169
            Top             =   1740
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   17
            Left            =   2655
            TabIndex        =   164
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單編號"
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
            Index           =   16
            Left            =   120
            TabIndex        =   163
            Top             =   1005
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "到貨日期"
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
            Left            =   120
            TabIndex        =   162
            Top             =   1740
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
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
            Index           =   4
            Left            =   2640
            TabIndex        =   161
            Top             =   1380
            Width           =   360
         End
      End
      Begin VB.Frame fra_OneOrder_Detail 
         Appearance      =   0  '平面
         BackColor       =   &H00404000&
         ForeColor       =   &H80000008&
         Height          =   3420
         Left            =   120
         TabIndex        =   26
         Top             =   5280
         Width           =   11880
         Begin VB.TextBox txt_OneOrder_SignQty 
            BackColor       =   &H0000FFFF&
            Height          =   270
            Left            =   1500
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1515
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cmb_OneOrder_RSCCode 
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNAbnormal.frx":55AC8
            Left            =   1530
            List            =   "frm_OP_SDNAbnormal.frx":55ACA
            Style           =   2  '單純下拉式
            TabIndex        =   28
            Top             =   2355
            Visible         =   0   'False
            Width           =   2490
         End
         Begin VB.ComboBox cmb_OneOrder_RBCCode 
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNAbnormal.frx":55ACC
            Left            =   1500
            List            =   "frm_OP_SDNAbnormal.frx":55ACE
            Style           =   2  '單純下拉式
            TabIndex        =   27
            Top             =   1995
            Visible         =   0   'False
            Width           =   2340
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gd_OneOrder_OrderDetail 
            Height          =   3270
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   5768
            _Version        =   393216
            ScrollBars      =   2
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox txt_Tab0_Route_No 
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
         Left            =   -62160
         TabIndex        =   124
         Top             =   3120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Frame Frame8 
         Caption         =   "計費項目"
         Height          =   2535
         Left            =   -74280
         TabIndex        =   121
         Top             =   6120
         Width           =   12975
         Begin VB.TextBox Text4 
            BackColor       =   &H0000FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   123
            Top             =   840
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab2_DelCost 
            BackColor       =   &H00FFFFC0&
            Caption         =   "刪除計費CTrl+D"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '圖片外觀
            TabIndex        =   115
            Top             =   960
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab2_AddCost 
            BackColor       =   &H00FFFFC0&
            Caption         =   "新增計費CTrl+A"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '圖片外觀
            TabIndex        =   114
            Top             =   240
            Width           =   1035
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab2_SDN_Cost 
            Height          =   2145
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   3784
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Cols            =   14
            FixedCols       =   0
            BackColorSel    =   10354595
            ForeColorSel    =   8454016
            BackColorBkg    =   -2147483644
            AllowBigSelection=   0   'False
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   14
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "訂單裝載內容"
         Height          =   3135
         Left            =   -74280
         TabIndex        =   120
         Top             =   2280
         Width           =   12975
         Begin VB.TextBox Text3 
            BackColor       =   &H0000FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   122
            Top             =   960
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab2_DelOrder 
            BackColor       =   &H00FFFFC0&
            Caption         =   "刪除訂單 Ctrl+D"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '圖片外觀
            TabIndex        =   112
            Top             =   960
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab2_AddOrder 
            BackColor       =   &H00FFFFC0&
            Caption         =   "新增訂單CTrl+A"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '圖片外觀
            TabIndex        =   111
            Top             =   240
            Width           =   1035
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab2_SDN_Detail 
            Height          =   2760
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   4868
            _Version        =   393216
            Cols            =   14
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   14
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
      End
      Begin VB.CommandButton cmd_Tab2_SelectCar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "？"
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
         Height          =   255
         Left            =   -69720
         Style           =   1  '圖片外觀
         TabIndex        =   106
         Top             =   1680
         Width           =   330
      End
      Begin VB.TextBox txt_Tab02_C_VEHICLE_ID_NO 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -70680
         TabIndex        =   105
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_Tab02_Delivery_Date 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -72960
         TabIndex        =   104
         Top             =   1680
         Width           =   1200
      End
      Begin VB.TextBox txt_Tab02_Driver 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -68160
         TabIndex        =   107
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_Tab02_Receiver 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -65880
         TabIndex        =   109
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_Tab02_C_Route_No 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -63360
         TabIndex        =   103
         Top             =   1680
         Width           =   1320
      End
      Begin VB.Frame Frame6 
         Height          =   525
         Left            =   -74280
         TabIndex        =   93
         Top             =   5520
         Width           =   5355
         Begin VB.OptionButton Op_Tab2_WT 
            Caption         =   "Option1"
            Height          =   255
            Left            =   4920
            TabIndex        =   99
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_CBM 
            Caption         =   "Option1"
            Height          =   255
            Left            =   3360
            TabIndex        =   98
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_CS 
            Caption         =   "Option1"
            Height          =   255
            Left            =   1800
            TabIndex        =   97
            Top             =   210
            Width           =   255
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Case 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   975
            TabIndex        =   96
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Volumn 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   2505
            TabIndex        =   95
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Weight 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4080
            TabIndex        =   94
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重量"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   13
            Left            =   3705
            TabIndex        =   102
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材積"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   14
            Left            =   2115
            TabIndex        =   101
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "總計：箱數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   15
            Left            =   75
            TabIndex        =   100
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '平面
         BackColor       =   &H00004000&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -74280
         TabIndex        =   86
         Top             =   480
         Width           =   12960
         Begin VB.CommandButton cmd_Tab2_AddNew 
            BackColor       =   &H00C0FFC0&
            Caption         =   "新  增"
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
            Left            =   4560
            Style           =   1  '圖片外觀
            TabIndex        =   92
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Modify 
            BackColor       =   &H00C0E0FF&
            Caption         =   "修  改"
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
            Left            =   3360
            Style           =   1  '圖片外觀
            TabIndex        =   91
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Save 
            BackColor       =   &H00C0C0FF&
            Caption         =   "存  檔"
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
            Left            =   5760
            Style           =   1  '圖片外觀
            TabIndex        =   90
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Delete 
            BackColor       =   &H000080FF&
            Caption         =   "刪  除"
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
            Left            =   8160
            Style           =   1  '圖片外觀
            TabIndex        =   89
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
            Index           =   1
            Left            =   9360
            Style           =   1  '圖片外觀
            TabIndex        =   88
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Cancel 
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
            Left            =   6960
            Style           =   1  '圖片外觀
            TabIndex        =   87
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         Height          =   525
         Left            =   -66660
         TabIndex        =   76
         Top             =   5520
         Width           =   5355
         Begin VB.OptionButton Op_Tab2_SumWT 
            Caption         =   "Option1"
            Height          =   255
            Left            =   5040
            TabIndex        =   82
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_SumCBM 
            Caption         =   "Option1"
            Height          =   255
            Left            =   3480
            TabIndex        =   81
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_SumCS 
            Caption         =   "Option1"
            Height          =   255
            Left            =   1920
            TabIndex        =   80
            Top             =   210
            Width           =   255
         End
         Begin VB.TextBox txt_Tab2_sum_Case 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1095
            TabIndex        =   79
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_sum_CBM 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   2625
            TabIndex        =   78
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_sum_WT 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4200
            TabIndex        =   77
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重量"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   12
            Left            =   3825
            TabIndex        =   85
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材積"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   10
            Left            =   2235
            TabIndex        =   84
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "小計：箱數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   7
            Left            =   195
            TabIndex        =   83
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Frame fra_Function 
         Height          =   615
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   11895
         Begin VB.CommandButton cmdUnRouteConfirm 
            BackColor       =   &H00FFFFC0&
            Caption         =   "     取消     出車確認"
            Enabled         =   0   'False
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
            Left            =   8640
            Style           =   1  '圖片外觀
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmdShipNotes 
            BackColor       =   &H00FFC0C0&
            Caption         =   "補印出貨單"
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
            Left            =   9360
            Style           =   1  '圖片外觀
            TabIndex        =   197
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdCost 
            BackColor       =   &H00FFFFC0&
            Caption         =   "運費維護"
            Enabled         =   0   'False
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
            Left            =   8040
            Style           =   1  '圖片外觀
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmdCarNOChange 
            BackColor       =   &H00C0FFC0&
            Caption         =   "車號變更"
            Enabled         =   0   'False
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
            Left            =   6720
            Style           =   1  '圖片外觀
            TabIndex        =   4
            Top             =   120
            Width           =   1245
         End
         Begin VB.ComboBox cmbOrderkey 
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
            ItemData        =   "frm_OP_SDNAbnormal.frx":55AD0
            Left            =   120
            List            =   "frm_OP_SDNAbnormal.frx":55AD2
            Style           =   2  '單純下拉式
            TabIndex        =   0
            Top             =   165
            Width           =   1455
         End
         Begin VB.CommandButton cmdNotYetOrder 
            BackColor       =   &H00C0FFFF&
            Caption         =   "待確認簽單"
            Enabled         =   0   'False
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
            Left            =   5400
            Style           =   1  '圖片外觀
            TabIndex        =   3
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FF80FF&
            Caption         =   "離開"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   0
            Left            =   10680
            Style           =   1  '圖片外觀
            TabIndex        =   7
            Top             =   120
            Width           =   1110
         End
         Begin VB.CommandButton cmd_OrderQuery 
            BackColor       =   &H00C0E0FF&
            Caption         =   "查詢"
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
            Left            =   4320
            Style           =   1  '圖片外觀
            TabIndex        =   2
            Top             =   120
            Width           =   1005
         End
         Begin VB.TextBox txt_OrderKey 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   1
            Top             =   165
            Width           =   2745
         End
      End
      Begin VB.Frame fra_OneOrder_Header 
         Appearance      =   0  '平面
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4305
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   11895
         Begin VB.TextBox txt_BranchId 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   281
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txt_Externordertype 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   279
            Top             =   800
            Width           =   1095
         End
         Begin VB.CommandButton cmdSDNBack 
            BackColor       =   &H0000FFFF&
            Caption         =   "簽單狀態"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   9600
            Style           =   1  '圖片外觀
            TabIndex        =   278
            Top             =   2280
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txt_OneOrder_ConsigneeKey1 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   276
            Top             =   1440
            Width           =   1080
         End
         Begin VB.TextBox txt_OneOrder_FullName1 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   275
            Top             =   1440
            Width           =   4680
         End
         Begin VB.TextBox txt_OneOrder_OrderKey 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   274
            Top             =   150
            Width           =   1095
         End
         Begin VB.TextBox txt_Storer 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   273
            Top             =   240
            Width           =   2040
         End
         Begin VB.TextBox txt_OneOrder_Address1 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   271
            Top             =   1740
            Width           =   4680
         End
         Begin VB.TextBox txt_Zip1 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   270
            Top             =   1740
            Width           =   1080
         End
         Begin VB.CommandButton cmdReceiptDetail 
            BackColor       =   &H00C0C0C0&
            Caption         =   "收貨明細"
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
            Left            =   9240
            Style           =   1  '圖片外觀
            TabIndex        =   269
            Top             =   3000
            Width           =   1005
         End
         Begin VB.ComboBox cboInvBack 
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
            ItemData        =   "frm_OP_SDNAbnormal.frx":55AD4
            Left            =   9360
            List            =   "frm_OP_SDNAbnormal.frx":55AD6
            TabIndex        =   20
            Top             =   3555
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_SDNNote 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   21
            ToolTipText     =   "配送異常需詳述異常發生原因"
            Top             =   3900
            Width           =   2535
         End
         Begin VB.TextBox txt_C_ROUTE_NO 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   149
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txt_ZIP 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   148
            Top             =   1140
            Width           =   1080
         End
         Begin VB.TextBox txt_OneOrder_DeliveryDate 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_RouteNo 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   143
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_StorerKey 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   141
            Top             =   240
            Width           =   1080
         End
         Begin VB.TextBox txt_Priority 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   139
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txt_TRPHandle 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Top             =   3300
            Width           =   4815
         End
         Begin VB.TextBox txt_SortingCost 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   18
            ToolTipText     =   "配送異常所衍生出之理貨費"
            Top             =   3300
            Width           =   975
         End
         Begin VB.TextBox txt_CustHandle 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Top             =   3000
            Width           =   4815
         End
         Begin VB.TextBox txt_TRPCost 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   17
            ToolTipText     =   "配送異常所衍生出之配送費"
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox txt_INVHandle 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   3900
            Width           =   4815
         End
         Begin VB.TextBox txt_TotalCost 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   19
            ToolTipText     =   "配送異常所衍生出之費用合計"
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txt_Advance 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Top             =   3600
            Width           =   4815
         End
         Begin VB.TextBox txt_OneOrder_StorerOrderKey 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   130
            Top             =   240
            Width           =   1680
         End
         Begin VB.ComboBox cmbScan 
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
            Left            =   5400
            TabIndex        =   10
            Text            =   "cmbScan"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txt_OneOrder_CustomerOrderkey1 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7560
            TabIndex        =   11
            Top             =   2400
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_Status 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            TabIndex        =   40
            ToolTipText     =   "紅色為遲交;綠色為達交"
            Top             =   1700
            Width           =   1095
         End
         Begin VB.CommandButton cmd_OneOrder_Expect 
            BackColor       =   &H000080FF&
            Caption         =   "異常訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10320
            Style           =   1  '圖片外觀
            TabIndex        =   22
            Top             =   3000
            Width           =   1485
         End
         Begin VB.CommandButton cmd_OneOrder_Deliveryok 
            BackColor       =   &H00FF8080&
            Caption         =   "正常訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10320
            Style           =   1  '圖片外觀
            TabIndex        =   12
            Top             =   2280
            Width           =   1485
         End
         Begin VB.CommandButton cmd_OneOrder_NoDelivery 
            BackColor       =   &H008080FF&
            Caption         =   "未出訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10320
            Style           =   1  '圖片外觀
            TabIndex        =   23
            ToolTipText     =   "點選""未出訂單""，系統將不于計費"
            Top             =   3600
            Width           =   1485
         End
         Begin VB.TextBox txt_OneOrder_Address 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1140
            Width           =   4680
         End
         Begin VB.TextBox txt_OneOrder_FullName 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   840
            Width           =   4680
         End
         Begin VB.TextBox txt_OneOrder_TRPCompany 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_Driver 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1740
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_VehicleID 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_OrderDate 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1100
            Width           =   1095
         End
         Begin VB.TextBox txt_OneOrder_ConsigneeKey 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   840
            Width           =   1080
         End
         Begin VB.TextBox txt_OneOrder_Description 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   540
            Width           =   5760
         End
         Begin VB.TextBox txt_OneOrder_ArriveDate 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1400
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp_OneOrder_SignDate 
            Height          =   375
            Left            =   960
            TabIndex        =   8
            Top             =   2400
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   -2147483643
            CustomFormat    =   "yyyy/MM/dd HH:mm"
            Format          =   97255427
            UpDown          =   -1  'True
            CurrentDate     =   39438
         End
         Begin MSComCtl2.DTPicker dtpSDNSendDate 
            Height          =   375
            Left            =   3600
            TabIndex        =   9
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   97255427
            UpDown          =   -1  'True
            CurrentDate     =   39438
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "分公司"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   282
            Top             =   2100
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶訂單類別"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   9480
            TabIndex        =   280
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "A2B到貨"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   30
            Left            =   2880
            TabIndex        =   277
            Top             =   1500
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "到貨地址"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   27
            Left            =   2880
            TabIndex        =   272
            Top             =   1800
            UseMnemonic     =   0   'False
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "發票回收"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   29
            Left            =   9360
            TabIndex        =   184
            Top             =   3360
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "簽單備註"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   28
            Left            =   6720
            TabIndex        =   183
            Top             =   3960
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "二次路編"
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
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   150
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "到貨地址"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   55
            Left            =   2880
            TabIndex        =   147
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "出車日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   146
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   144
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   142
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單類別"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   54
            Left            =   9840
            TabIndex        =   140
            Top             =   530
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "TMS單號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   53
            Left            =   9840
            TabIndex        =   138
            Top             =   240
            Width           =   855
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   240
            X2              =   11640
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "後續處理方式"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   47
            Left            =   120
            TabIndex        =   137
            Top             =   3360
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "異常理貨費"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   51
            Left            =   6720
            TabIndex        =   136
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶回覆處理方式"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   46
            Left            =   120
            TabIndex        =   135
            Top             =   3060
            Width           =   1560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "異常配送費"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   50
            Left            =   6720
            TabIndex        =   134
            Top             =   3060
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "庫存調整方式"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   52
            Left            =   120
            TabIndex        =   133
            Top             =   3960
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "費用合計"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   49
            Left            =   6720
            TabIndex        =   132
            Top             =   3660
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "改善方式"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   48
            Left            =   120
            TabIndex        =   131
            Top             =   3660
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   9840
            TabIndex        =   45
            Top             =   1150
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "掃描"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   45
            Left            =   4920
            TabIndex        =   129
            Top             =   2460
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "簽單回傳"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   10
            Left            =   2760
            TabIndex        =   128
            Top             =   2460
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "驗收單號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   44
            Left            =   6720
            TabIndex        =   127
            Top             =   2460
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "簽單狀態"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   9840
            TabIndex        =   50
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶簽收"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   2460
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運輸公司"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "駕駛姓名"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   47
            Top             =   1800
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送車號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   1500
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "到貨日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   9840
            TabIndex        =   44
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   2880
            TabIndex        =   43
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單備註"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   42
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單號碼"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   6960
            TabIndex        =   41
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.Frame fra_MultiOrder_Detail 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   3780
         Left            =   120
         TabIndex        =   52
         Top             =   3240
         Visible         =   0   'False
         Width           =   10920
         Begin VB.ComboBox cmb_MultiOrder_RBCCode 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNAbnormal.frx":55AD8
            Left            =   1545
            List            =   "frm_OP_SDNAbnormal.frx":55ADA
            Style           =   2  '單純下拉式
            TabIndex        =   55
            Top             =   2010
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.ComboBox cmb_MultiOrder_RSCCode 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNAbnormal.frx":55ADC
            Left            =   1575
            List            =   "frm_OP_SDNAbnormal.frx":55ADE
            Style           =   2  '單純下拉式
            TabIndex        =   54
            Top             =   2370
            Visible         =   0   'False
            Width           =   2490
         End
         Begin VB.TextBox txt_MultiOrder_SignQty 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Left            =   1545
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1770
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gd_MultiOrder_OrderDetail 
            Height          =   3510
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   6191
            _Version        =   393216
            FixedCols       =   0
            ScrollBars      =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fra_MultiOrder_Header 
         Appearance      =   0  '平面
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   0
         TabIndex        =   57
         Top             =   4080
         Visible         =   0   'False
         Width           =   10935
         Begin VB.CommandButton cmd_MultiOrder_NoDelivery 
            BackColor       =   &H008080FF&
            Caption         =   "未出訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   7635
            Style           =   1  '圖片外觀
            TabIndex        =   68
            Top             =   1470
            Width           =   1485
         End
         Begin VB.CommandButton cmd_MultiOrder_Deliveryok 
            BackColor       =   &H00FF8080&
            Caption         =   "正常訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5940
            Style           =   1  '圖片外觀
            TabIndex        =   67
            Top             =   1470
            Width           =   1485
         End
         Begin VB.CommandButton cmd_MultiOrder_Expect 
            BackColor       =   &H000080FF&
            Caption         =   "異常確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   9330
            Style           =   1  '圖片外觀
            TabIndex        =   66
            Top             =   1470
            Width           =   1485
         End
         Begin VB.TextBox txt_MultiOrder_SignDate 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            Left            =   5400
            TabIndex        =   65
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txt_MultiOrder_Status 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6735
            TabIndex        =   64
            Top             =   450
            Width           =   1215
         End
         Begin VB.TextBox txt_MultiOrder_ArriveDate 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9375
            TabIndex        =   63
            Top             =   450
            Width           =   1215
         End
         Begin VB.TextBox txt_MultiOrder_ConsigneeKey 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   885
            TabIndex        =   62
            Top             =   135
            Width           =   1575
         End
         Begin VB.TextBox txt_MultiOrder_OrderDate 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9375
            TabIndex        =   61
            Top             =   150
            Width           =   1215
         End
         Begin VB.TextBox txt_MultiOrder_StorerKey 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6735
            TabIndex        =   60
            Top             =   150
            Width           =   1020
         End
         Begin VB.TextBox txt_MultiOrder_FullName 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2490
            TabIndex        =   59
            Top             =   135
            Width           =   2775
         End
         Begin VB.TextBox txt_MultiOrder_Address 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   885
            TabIndex        =   58
            Top             =   435
            Width           =   4395
         End
         Begin MSDataGridLib.DataGrid dg_MultiOrder 
            Height          =   1185
            Left            =   60
            TabIndex        =   69
            Top             =   840
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   2090
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
         Begin MSComCtl2.DTPicker dtp_MultiOrder_SignDate 
            Height          =   375
            Left            =   7560
            TabIndex        =   126
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
            Format          =   97255427
            UpDown          =   -1  'True
            CurrentDate     =   39438
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶簽收日期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   18
            Left            =   5940
            TabIndex        =   75
            Top             =   990
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "狀態"
            Height          =   180
            Index           =   17
            Left            =   6345
            TabIndex        =   74
            Top             =   525
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   73
            Top             =   195
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "送貨日"
            Height          =   180
            Index           =   15
            Left            =   8805
            TabIndex        =   72
            Top             =   510
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單日"
            Height          =   180
            Index           =   14
            Left            =   8805
            TabIndex        =   71
            Top             =   225
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   6330
            TabIndex        =   70
            Top             =   210
            Width           =   390
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   5895
         Left            =   -74760
         TabIndex        =   198
         Top             =   1440
         Visible         =   0   'False
         Width           =   13695
         Begin VB.TextBox Text1 
            BackColor       =   &H0000FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   243
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Frame Frame3 
            Caption         =   "小計"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   525
            Left            =   8880
            TabIndex        =   233
            Top             =   3360
            Width           =   4875
            Begin VB.TextBox txt_Tab0_sum_WT 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   3720
               TabIndex        =   239
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_sum_CBM 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   2145
               TabIndex        =   238
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_sum_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   615
               TabIndex        =   237
               Top             =   165
               Width           =   840
            End
            Begin VB.OptionButton Op_SumCS_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   1440
               TabIndex        =   236
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton Op_SumCBM_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   3000
               TabIndex        =   235
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton Op_SumWT_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   4560
               TabIndex        =   234
               Top             =   210
               Width           =   255
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   3
               Left            =   195
               TabIndex        =   242
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
               Left            =   1755
               TabIndex        =   241
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "重量"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3345
               TabIndex        =   240
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame Frame2 
            Height          =   525
            Left            =   0
            TabIndex        =   221
            Top             =   3300
            Width           =   7635
            Begin VB.TextBox txt_Tab0_srcTotal_Weight 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   6480
               TabIndex        =   227
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Volumn 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   4905
               TabIndex        =   226
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Case 
               Alignment       =   2  '置中對齊
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   3375
               TabIndex        =   225
               Top             =   165
               Width           =   840
            End
            Begin VB.OptionButton Op_CS_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   4200
               TabIndex        =   224
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton Op_CBM_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   5760
               TabIndex        =   223
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton OpWT_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   7320
               TabIndex        =   222
               Top             =   210
               Width           =   255
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "總計：箱數"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   11
               Left            =   2475
               TabIndex        =   231
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "材積"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   9
               Left            =   4515
               TabIndex        =   230
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
               Left            =   6105
               TabIndex        =   229
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Lb_Route 
               Caption         =   "Route"
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
               Height          =   165
               Left            =   1200
               TabIndex        =   228
               Top             =   210
               Width           =   1215
            End
         End
         Begin VB.TextBox Text2_del 
            BackColor       =   &H0000FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   220
            Top             =   4800
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H0000FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7200
            TabIndex        =   219
            Top             =   2160
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab0_AddCost_del 
            BackColor       =   &H00FFFFC0&
            Caption         =   "新增計費"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6360
            Style           =   1  '圖片外觀
            TabIndex        =   218
            Top             =   2895
            Width           =   1035
         End
         Begin VB.TextBox txt_Tab0_SumQty 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   10080
            TabIndex        =   217
            Top             =   2895
            Width           =   1005
         End
         Begin VB.Frame Frame11 
            Height          =   1005
            Left            =   120
            TabIndex        =   199
            Top             =   240
            Width           =   13605
            Begin VB.TextBox txt_DeliveryDate_End 
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
               Left            =   2445
               TabIndex        =   209
               Top             =   630
               Width           =   1125
            End
            Begin VB.TextBox txt_DeliveryDate_Start 
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
               Left            =   1050
               TabIndex        =   208
               Top             =   630
               Width           =   1125
            End
            Begin VB.TextBox txt_RouteNo_Start 
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
               Left            =   1050
               TabIndex        =   207
               Top             =   240
               Width           =   1125
            End
            Begin VB.TextBox txt_RouteNo_End 
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
               Left            =   2445
               TabIndex        =   206
               Top             =   240
               Width           =   1125
            End
            Begin VB.CheckBox ck_confirm 
               Caption         =   "未確認簽單"
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
               Left            =   3720
               TabIndex        =   205
               Top             =   285
               Width           =   1455
            End
            Begin VB.OptionButton Op_UnCheck 
               Caption         =   "未整理"
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
               Height          =   345
               Left            =   7080
               TabIndex        =   204
               Top             =   210
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CheckBox ck_back 
               Caption         =   "未回收簽單"
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
               Left            =   3720
               TabIndex        =   203
               Top             =   675
               Width           =   1455
            End
            Begin VB.OptionButton Op_OnCheck 
               Caption         =   "已整理"
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
               Height          =   255
               Left            =   7080
               TabIndex        =   202
               Top             =   645
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txt_Tab0_C_VEHICLE_ID_NO 
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
               Left            =   5880
               TabIndex        =   201
               Top             =   240
               Width           =   1125
            End
            Begin VB.TextBox txt_Tab0_Receiver 
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
               Left            =   5880
               TabIndex        =   200
               Top             =   630
               Width           =   1125
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
               Index           =   19
               Left            =   2205
               TabIndex        =   215
               Top             =   555
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
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   20
               Left            =   150
               TabIndex        =   214
               Top             =   675
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
               Index           =   21
               Left            =   2205
               TabIndex        =   213
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label3 
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
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   22
               Left            =   150
               TabIndex        =   212
               Top             =   285
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "車   號"
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
               Index           =   39
               Left            =   5190
               TabIndex        =   211
               Top             =   285
               Width           =   600
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "領款人"
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
               Index           =   40
               Left            =   5220
               TabIndex        =   210
               Top             =   675
               Width           =   630
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_SDN_Detail_del 
            Height          =   4440
            Left            =   0
            TabIndex        =   216
            Top             =   4440
            Width           =   13605
            _ExtentX        =   23998
            _ExtentY        =   7832
            _Version        =   393216
            Cols            =   14
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   14
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_SDN_Head_del 
            Height          =   1920
            Left            =   120
            TabIndex        =   232
            Top             =   1320
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   3387
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Cols            =   14
            FixedCols       =   0
            BackColorSel    =   10354595
            ForeColorSel    =   8454016
            BackColorBkg    =   -2147483644
            AllowBigSelection=   0   'False
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   14
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "車號："
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
         Index           =   35
         Left            =   -71340
         TabIndex        =   119
         Top             =   1725
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "送貨日："
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
         Index           =   26
         Left            =   -73800
         TabIndex        =   118
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "駕駛人："
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
         Index           =   25
         Left            =   -69000
         TabIndex        =   117
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "領款人："
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
         Index           =   24
         Left            =   -66720
         TabIndex        =   116
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "路線編號："
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
         Index           =   23
         Left            =   -64440
         TabIndex        =   108
         Top             =   1725
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '不透明
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Height          =   735
         Index           =   4
         Left            =   -74280
         Top             =   1440
         Width           =   12930
      End
   End
End
Attribute VB_Name = "frm_OP_SDNAbnormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬
Private iLoop As Double              '迴圈計數

Private blShipped As Boolean         '是否已經執行過 Shipped Confirm
Private blSDNConfirm As Boolean      '是否已經執行過 SDN Confirm
Private blCanUpdate As Boolean       '是否可以執行 SDN Confirm
Private blRouteT0Change As Boolean   '是否執行明細查詢

Private rs_MultiOrder As ADODB.Recordset
Private rs_Tab1_SDN05T As ADODB.Recordset
Private rs_cost As ADODB.Recordset
Private rs_cust As ADODB.Recordset
Private intR, i, j, intC As Integer
Private a, B, C As Double            '統計利潤
Private str_DELIVERY_DATE, str_C_ROUTE_NO, str_C_VEHICLE_ID_NO, str_Driver, Str_Receiver, str_ChargeQty As String
Private str_Receivable, str_Payable, str_Premiam, str_Reason, str_AreaStart, str_AreaEnd, str_SDNStatus As String
Private str_ROUTE_NO, str_EXTERN, str_ARRIVE_DATE, str_CUST_NAME, str_SHIP_CS, str_SHIP_CBM, str_SHIP_WT As String
Private route, str_CAR_NOTES, str_SDN_NOTE, str_uom, str_SumReceivable, str_SumPayable, str_C_ROUTE_Time, str_SDN_Date As String
Private str_OnTimeDelivery, str_PODOnTime, str_RejectOrder, str_C_ROUTE_Total, str_SDN_NO, str_SDN_Name, str_CostKind As String
Private rsMain1 As ADODB.Recordset
Private rsRouteT0 As ADODB.Recordset
Private rsOrderT0 As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private Sub Ship2TMS(strOrderkey As String)

Call ReDim_Recordset(tmp_Rs)
str_SQL = "select * from sdn03t where receipt_no = '" & strOrderkey & "' and ship_qty = 0 "
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'無未回傳資料
If tmp_Rs.EOF Then tmp_Rs.Close: Exit Sub

str_SQL = "select o.route " & _
        ",o.storerkey " & _
        ",o.orderkey " & _
        ",o.updatesource " & _
        ",o.Externorderkey " & _
        ",od.ExternLineno " & _
        ",od.sku " & _
        ",od.shippedqty " & _
        ",od.editdate " & _
        "from " & strWMSDB & "..orders o join " & strWMSDB & "..orderdetail od on o.orderkey = od.orderkey " & _
        "and o.status = '9' " & _
        "where len(rtrim(isnull(o.updatesource,''))) > 0 and o.updatesource = '" & strOrderkey & "' " & _
        "and od.shippedqty > 0 "

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'無資料
If tmp_Rs.EOF Then tmp_Rs.Close: Exit Sub

Dim i As Long
tmp_Rs.MoveFirst

Tran_Level = cn.BeginTrans
Do While Not tmp_Rs.EOF

    str_SQL = "UPDATE TRP03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
             "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
             "and receipt_no ='" & tmp_Rs("updatesource") & "' and product_no = '" & tmp_Rs("sku") & "' and SHIP_QTY = 0 "
    cn.Execute str_SQL ', RowsAffect, adExecuteNoRecords
    
    str_SQL = "UPDATE SDN03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
             "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
             "and receipt_no ='" & tmp_Rs("updatesource") & "' and product_no = '" & tmp_Rs("sku") & "'  and SHIP_QTY = 0 "
    cn.Execute str_SQL ', RowsAffect, adExecuteNoRecords
               
    tmp_Rs.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0
tmp_Rs.Close

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, Me.Name)
End Sub

Private Sub cmb_OneOrder_RBCCode_Click()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：責屬
gd_OneOrder_OrderDetail.Text = cmb_OneOrder_RBCCode.List(cmb_OneOrder_RBCCode.ListIndex)
If Len(Trim(cmb_OneOrder_RBCCode.List(cmb_OneOrder_RBCCode.ListIndex))) > 0 Then
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 10) = Left(cmb_OneOrder_RBCCode.List(cmb_OneOrder_RBCCode.ListIndex), 3)
Else
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 10) = ""
End If
cmb_OneOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_OneOrder_RBCCode_LostFocus()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：責屬
cmb_OneOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_MultiOrder_RBCCode_Click()
'ㄧ張貨主單號對應多張排車系統訂單：責屬
gd_MultiOrder_OrderDetail.Text = cmb_MultiOrder_RBCCode.List(cmb_MultiOrder_RBCCode.ListIndex)
If Len(Trim(cmb_MultiOrder_RBCCode.List(cmb_MultiOrder_RBCCode.ListIndex))) > 0 Then
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 9) = Left(cmb_MultiOrder_RBCCode.List(cmb_MultiOrder_RBCCode.ListIndex), 3)
Else
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 9) = ""
End If
cmb_MultiOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_MultiOrder_RBCCode_LostFocus()
'ㄧ張貨主單號對應多張排車系統訂單：責屬
cmb_MultiOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_OneOrder_RSCCode_Click()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：異常原因
gd_OneOrder_OrderDetail.Text = cmb_OneOrder_RSCCode.List(cmb_OneOrder_RSCCode.ListIndex)
If Len(Trim(cmb_OneOrder_RSCCode.List(cmb_OneOrder_RSCCode.ListIndex))) > 0 Then
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 9) = Left(cmb_OneOrder_RSCCode.List(cmb_OneOrder_RSCCode.ListIndex), 3)
Else
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 9) = ""
End If
cmb_OneOrder_RSCCode.Visible = False
End Sub

Private Sub cmb_OneOrder_RSCCode_LostFocus()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：異常原因
cmb_OneOrder_RSCCode.Visible = False
End Sub

Private Sub cmb_multiOrder_RSCCode_Click()
'ㄧ張貨主單號對應多張排車系統訂單：異常原因
gd_MultiOrder_OrderDetail.Text = cmb_MultiOrder_RSCCode.List(cmb_MultiOrder_RSCCode.ListIndex)
If Len(Trim(cmb_MultiOrder_RSCCode.List(cmb_MultiOrder_RSCCode.ListIndex))) > 0 Then
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 8) = Left(cmb_MultiOrder_RSCCode.List(cmb_MultiOrder_RSCCode.ListIndex), 3)
Else
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 8) = ""
End If
cmb_MultiOrder_RSCCode.Visible = False
End Sub

Private Sub cmb_MultiOrder_RSCCode_LostFocus()
'ㄧ張貨主單號對應多張排車系統訂單：異常原因
cmb_MultiOrder_RSCCode.Visible = False
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
    '離開
    Unload Me
End Sub

Private Sub cmd_MultiOrder_Deliveryok_Click()
'ㄧ張貨主單號對應多張排車系統訂單：正常訂單
If Len(txt_MultiOrder_SignDate.Text) = 0 Then
   msg_text = "資料錯誤：未輸入 [客戶簽收日期]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
   msg_text = "客戶簽收日期：" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_MultiOrder_SignDate.SelStart = 0: txt_MultiOrder_SignDate.SelLength = Len(txt_MultiOrder_SignDate.Text): txt_MultiOrder_SignDate.SetFocus
   Exit Sub
End If

Tran_Level = 0
Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'更新 TRP02T
rs_MultiOrder.MoveFirst
Do While Not rs_MultiOrder.EOF
   str_SQL = "Update TRP02T Set CustSignDate = '" & Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate, 5, 2) & "/" & Right(txt_MultiOrder_SignDate, 2) & "'," & _
             "   Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '正常訂單' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("訂單編號").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '更新 TRP03T
   str_SQL = "Update TRP03T Set Sign_Qty = Ship_Qty,RSC_Code = '' , RBC_Code = '' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("訂單編號").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   rs_MultiOrder.MoveNext
Loop

cn.CommitTrans
Tran_Level = 0

Call ClearForm
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
   CreateErrorLog Me.Name & "-未出訂單", Me.Caption, "cmd_OnOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_MultiOrder_Expect_Click()
'ㄧ張貨主單號對應多張排車系統訂單：異常訂單
If Len(txt_MultiOrder_SignDate.Text) = 0 Then
   msg_text = "資料錯誤：未輸入 [客戶簽收日期]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
   msg_text = "客戶簽收日期：" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_MultiOrder_SignDate.SelStart = 0: txt_MultiOrder_SignDate.SelLength = Len(txt_MultiOrder_SignDate.Text): txt_MultiOrder_SignDate.SetFocus
   Exit Sub
End If

Dim strRBC As String
Dim strRSC As String
Dim dbSeqNo As Double
Dim dnSignQty As Double

'檢核是否選取 [異常原因] 與 [責任歸屬]
With gd_MultiOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 8    '異常代碼
        If Len(Trim(.Text)) > 0 Then
           .Col = 8: strRSC = strRSC & Trim(.Text)
           .Col = 9: strRBC = strRBC & Trim(.Text)
        End If
     Next iLoop
End With
If strRSC = "" Or strRBC = "" Then
   msg_text = "異常訂單必須選取對應之 [異常原因] 與 [責任歸屬]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Tran_Level = 0
Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass
rs_MultiOrder.MoveFirst
Do While Not rs_MultiOrder.EOF
   '更新 TRP02T
   str_SQL = "Update TRP02T Set CustSignDate = '" & Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate, 5, 2) & "/" & Right(txt_MultiOrder_SignDate, 2) & "'," & _
             "   Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '異常訂單' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("訂單編號").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   rs_MultiOrder.MoveNext
Loop

'更新 TRP03T
With gd_MultiOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 1: dbSeqNo = Val(.Text)
        .Col = 5: dnSignQty = Val(.Text)
        .Col = 8: strRSC = Trim(.Text)
        .Col = 9: strRBC = Trim(.Text)
        .Col = 0          '訂單編號
        If strRSC = "" And strRBC = "" Then
           str_SQL = "Update TRP03T Set Sign_Qty = Ship_Qty,RSC_Code = '',RBC_Code = '' " & _
                     "Where Receipt_No = '" & .Text & "' and Seq_No = " & dbSeqNo
        Else
           str_SQL = "Update TRP03T Set Sign_Qty = " & dnSignQty & ",RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "' " & _
                     "Where Receipt_No = '" & .Text & "' and Seq_No = " & dbSeqNo
        End If
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
     Next iLoop
End With

cn.CommitTrans
Tran_Level = 0

Call ClearForm
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
   CreateErrorLog Me.Name & "-異常訂單", Me.Caption, "cmd_MultiOrder_Expect_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_MultiOrder_NoDelivery_Click()
'ㄧ張貨主單號對應多張排車系統訂單：未出訂單

If Len(txt_MultiOrder_SignDate.Text) = 0 Then
   msg_text = "資料錯誤：未輸入 [客戶簽收日期]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
   msg_text = "客戶簽收日期：" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_MultiOrder_SignDate.SelStart = 0: txt_MultiOrder_SignDate.SelLength = Len(txt_MultiOrder_SignDate.Text): txt_MultiOrder_SignDate.SetFocus
   Exit Sub
End If


Dim strRBC As String
Dim strRSC As String

'檢核是否於第一項選取 [異常原因] 與 [責任歸屬]
With gd_MultiOrder_OrderDetail
        .Row = 1
        .Col = 9    '異常代碼
        If Len(Trim(.Text)) = 0 Then
           msg_text = "未出訂單，請於細項第一筆選取 [異常原因] 與 [責任歸屬]"
           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
           Exit Sub
        Else
           .Col = 9: strRSC = .Text
           .Col = 10: strRBC = .Text
        End If
End With

Tran_Level = 0
Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass
'更新 TRP02T
rs_MultiOrder.MoveFirst
Do While Not rs_MultiOrder.EOF
   str_SQL = "Update TRP02T Set CustSignDate = '" & Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate, 5, 2) & "/" & Right(txt_MultiOrder_SignDate, 2) & "'," & _
             "   Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '未出訂單' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("訂單編號").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '更新 TRP03T
   str_SQL = "Update TRP03T Set Sign_Qty = 0,RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("訂單編號").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   rs_MultiOrder.MoveNext
Loop
cn.CommitTrans
Tran_Level = 0

Call ClearForm
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
   CreateErrorLog Me.Name & "-未出訂單", Me.Caption, "cmd_MultiOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Public Sub cmd_OneOrder_Deliveryok_Click()
On Error GoTo err_Handle
'簽單是否已維護
If Len(Trim(txt_OneOrder_Status)) > 0 Then MsgBox "此簽單已經維護，無法修改!!", vbOKOnly, Me.Caption: Exit Sub

'清除特殊字元
Call myFormExCharFilter(Me)

Dim strInt As Long, blTmp As Boolean

'有異常無法點選正常簽單
With gd_OneOrder_OrderDetail
        For iLoop = 1 To .Rows - 1
            .Row = iLoop
            '異常代碼
            .Col = 9: If Len(Trim(.Text)) > 0 Then MsgBox "明細有維護異常，無法選取正常訂單!!", vbOKOnly, Me.Caption: Exit Sub
            .Col = 10: If Len(Trim(.Text)) > 0 Then MsgBox "明細有維護異常，無法選取正常訂單!!", vbOKOnly, Me.Caption: Exit Sub
            .Col = 14: If Len(Trim(.Text)) > 0 Then MsgBox "明細有維護異常，無法選取正常訂單!!", vbOKOnly, Me.Caption: Exit Sub
            .Col = 4: strInt = Val(Trim(.Text))
            .Col = 5: If Val(Trim(.Text)) <> strInt Then blTmp = True
        Next iLoop
End With

If frm_SDNConfirmNotYet.Visible = False Then '是否快速簽單確認
    '訂單量與出貨量不符
    If blTmp = True Then
        If MsgBox("訂單量與出貨量不符，是否繼續確認！", vbYesNo, "正常訂單確認") <> vbYes Then Exit Sub
    End If
    
    If blTmp = True Then MsgBox "訂單量與出貨量不符，請務必確認運費資料是否正確！", 64, "正常訂單確認"
End If

Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'更新 SDN01T
str_SQL = "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & txt_C_Route_NO.Text & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'更新 SDN02T
str_SQL = "Update SDN02T Set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "', invback = '" & cboInvBack.Text & "'," & _
          " Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '正常訂單' , CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' " & _
          ", cust_handle = '" & txt_CustHandle.Text & "' " & _
          ", TRP_Handle = '" & txt_TRPHandle.Text & "' " & _
          ", Advance = '" & txt_Advance.Text & "' " & _
          ", INV_Handle = '" & txt_INVHandle.Text & "' " & _
          ", TRP_Cost = '" & txt_TRPCost.Text & "' " & _
          ", Sorting_Cost = '" & txt_SortingCost.Text & "' " & _
          ", Total_Cost = '" & txt_TotalCost.Text & "' " & _
          ", SDN_NOTE = '" & txt_SDNNote.Text & "' " & _
          "Where Receipt_No = '" & RTrim(txt_OneOrder_OrderKey.Text) & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'是否快速簽單確認
If frm_SDNConfirmNotYet.Visible = True Then cn.Execute "update sdn02t set sdn_note = '快速簽單確認' Where Receipt_No = '" & RTrim(txt_OneOrder_OrderKey.Text) & "'", RowsAffect, adExecuteNoRecords

'更新 SDN03T
Dim dbSeqNo, dbShipQty, dnSignQty, strRSC, strRBC, strResponsible As String
With gd_OneOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 0: dbSeqNo = .Text
        .Col = 5: dbShipQty = Val(.Text)
        .Col = 6: dnSignQty = Val(.Text)
        .Col = 9: strRSC = Trim(.Text)
        .Col = 10: strRBC = Trim(.Text)
        .Col = 14: strResponsible = Trim(.Text)
        str_SQL = "Update SDN03T Set Ship_Qty = " & dbShipQty & ",Sign_Qty =  " & dbShipQty & ",RSC_Code = '',RBC_Code = '',Responsible = '' " & _
                     "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' and Seq_No = '" & dbSeqNo & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
     Next iLoop
End With

''運費計算
'cn.Execute "exec gs_Cost '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'更新簽單狀態
If rsOrderT0 Is Nothing = False Then
    If rsOrderT0.RecordCount > 0 And rsOrderT0.EOF = False And rsOrderT0("TMS單號") = txt_OneOrder_OrderKey.Text Then
        rsOrderT0("驗收單號") = txt_OneOrder_CustomerOrderkey1
        rsOrderT0("狀態") = "正常訂單"
    End If
End If

Call cmd_OrderQuery_Click
Screen.MousePointer = vbDefault
cmbOrderkey.ListIndex = 0

'是否快速簽單確認
If frm_SDNConfirmNotYet.Visible = False Then
    cmbOrderkey.SetFocus
    
    '運費計算
'    Call cmdCost_Click 'marked by gemini @20111223
End If

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-正常訂單", Me.Caption, "cmd_OnOrder_Deliveryok_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_OneOrder_Expect_Click()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：未出訂單
'If Len(dtp_OneOrder_SignDate.Value) = 0 Then
'   msg_text = "資料錯誤：未輸入 [客戶簽收日期]"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'End If
'If Fun_ChkDateFormat(txt_OneOrder_SignDate.Text) = 1 Then
'   msg_text = "客戶簽收日期：" & funRtn_msg
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   txt_OneOrder_SignDate.SelStart = 0: txt_OneOrder_SignDate.SelLength = Len(txt_OneOrder_SignDate.Text): txt_OneOrder_SignDate.SetFocus
'   Exit Sub
'End If
On Error GoTo err_Handle
If Len(Trim(txt_OneOrder_Status)) > 0 Then MsgBox "此簽單已經維護，無法修改!!", vbOKOnly, Me.Caption: Exit Sub

'清除特殊字元
Call myFormExCharFilter(Me)

Dim strRBC As String, strRSC As String, dbSeqNo As String, dnSignQty As Double, dbShipQty As Double, strInt As Double, blTmp As Boolean, strResponsible As String

'檢核是否選取 [異常原因] 與 [責任歸屬]
With gd_OneOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 9    '異常代碼
        If Len(Trim(.Text)) > 0 Then
           .Col = 9: strRSC = strRSC & Trim(.Text)
           .Col = 10: strRBC = strRBC & Trim(.Text)
        End If
            .Col = 4: strInt = Val(Trim(.Text))
            .Col = 5: If Val(Trim(.Text)) <> strInt Then blTmp = True
     Next iLoop
End With

If strRSC = "" Or strRBC = "" Then
   msg_text = "異常訂單必須選取對應之 [異常原因] 與 [責任歸屬]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'訂單量與出貨量不符
If blTmp = True Then
    If MsgBox("訂單量與送貨量不符，是否繼續確認！", vbYesNo, "異常訂單確認") <> vbYes Then Exit Sub
End If

If blTmp = True Then MsgBox "訂單量與送貨量不符，請務必確認運費資料是否正確！", 64, "異常訂單確認"

Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'更新 TRP01T
str_SQL = "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & txt_C_Route_NO.Text & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'更新 SDN02T
str_SQL = "Update SDN02T " & _
            "Set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "' " & _
              ", Confirm_UserID = '" & User_id & "' " & _
              ", cust_handle = '" & txt_CustHandle.Text & "' " & _
              ", TRP_Handle = '" & txt_TRPHandle.Text & "' " & _
              ", Advance = '" & txt_Advance.Text & "' " & _
              ", INV_Handle = '" & txt_INVHandle.Text & "' " & _
              ", TRP_Cost = '" & txt_TRPCost.Text & "' " & _
              ", Sorting_Cost = '" & txt_SortingCost.Text & "' " & _
              ", Total_Cost = '" & txt_TotalCost.Text & "' " & _
              ", SDN_NOTE = '" & txt_SDNNote.Text & "' ,invback = '" & cboInvBack.Text & "' " & _
              ",Confirm_Date = getdate() ,Confirm_Notes = '異常訂單' , CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' " & _
              "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'更新 SDN03T
With gd_OneOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 0: dbSeqNo = .Text
        .Col = 5: dbShipQty = Val(.Text)
        .Col = 6: dnSignQty = Val(.Text)
        .Col = 9: strRSC = Trim(.Text)
        .Col = 10: strRBC = Trim(.Text)
        .Col = 14: strResponsible = Trim(.Text)
        If strRSC = "" And strRBC = "" Then '未異常簽收數量=訂單數量
           str_SQL = "Update SDN03T set ship_qty = " & dbShipQty & ", Sign_Qty = " & dbShipQty & ",RSC_Code = '',RBC_Code = '',Responsible = '' " & _
                     "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' and Seq_No = '" & dbSeqNo & "' "
        Else
           str_SQL = "Update SDN03T Set Sign_Qty = " & dnSignQty & ",ship_qty = " & dbShipQty & ",RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "',Responsible = '" & strResponsible & "' " & _
                     "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' and Seq_No = '" & dbSeqNo & "' "
        End If
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
     Next iLoop
End With

''運費計算
'cn.Execute "exec gs_Cost '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'讀取ini參數，是否入WMS系統
Dim objIni As New vbIniFile, strOtherOrder2WMS As String
objIni.FileName = App.Path & "/" & App.title & ".ini"

strOtherOrder2WMS = objIni.ReadData("OPTION", "OtherOrder2WMS", "YES")
Set objIni = Nothing

If UCase(strOtherOrder2WMS) = "YES" And txt_OneOrder_StorerKey <> "LABT01" Then  'WMS是否新增採購單

    '寫入WMS採購單
    If RTrim(txt_Priority.Text) = "R" Or RTrim(txt_Priority.Text) = "RC" Or RTrim(txt_Priority.Text) = "A2B" Then '退貨單、提貨配送與移倉入庫不產生WMS採購單
    Else
        Dim rsKeycount As New ADODB.Recordset, strKeycount As String, intLineNumber As Integer
        If RTrim(txt_OneOrder_StorerKey.Text) = "LMBO01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LPSI01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LCHF01" Then
            '利豐拒短收先不寫入ASN，等雪慧改好再說
            If RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Then GoTo NoDo
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select asnkey from " & strWMSDB & "..asn where buyersreference = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '判斷TMS單號是否已經寫入asn的buyersreference欄位
            
    '            If MsgBox("WMS是否產生預收採購單?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '取系統採購單號
                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='asn' ", cn
                    '單號+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'asn'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '寫入表頭
                    str_SQL = "insert into " & strWMSDB & "..asn (asnKey,StorerKey,externasnkey , sellersreference,BuyersReference,asntype,notes,buyerVAT) " & _
                    "select asnKey = '" & strKeycount & "' , s2.storerkey , rtrim(o.externorderkey) , o.consigneekey , s2.receipt_no , 'A' , description , '" & RTrim(txt_OneOrder_FullName1) & "' " & _
                    "from sdn02t s2 join orders o on s2.c_receipt_no = o.orderkey " & _
                    "Where s2.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "'"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '寫入表身
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    
                    str_SQL = "select s3.product_no , s.descr , s3.storerkey , s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) " & _
                                "from sdn03t s3 join gv_skuxpack s on s.sku = s3.product_no and s3.storerkey = s.storerkey " & _
                                "Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' " & _
                                "group by s3.product_no , s.descr , s3.storerkey, s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    'Add by Gemini @20090303
                    If tmp_Rs.EOF Then cn.RollbackTrans: MsgBox "簽收量等於出貨量，無需產生採購單！", 64, "簽單維護": GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                    intLineNumber = intLineNumber + 1
            
                    str_SQL = "insert into " & strWMSDB & "..asndetail (asnKey,asnLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered,packkey,lottable06) " & _
                            "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & RTrim(tmp_Rs("Descr")) & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "','')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "已新增WMS預約收貨採購單(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
            End If
        Else
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select * from " & strWMSDB & "..po where externpokey = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '是否已產生採購單號
            
    '            If MsgBox("是否產生預收採購單?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '取系統採購單號
                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
                    '單號+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '寫入表頭
                    str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,externpokey , sellername,selleraddress1,BuyersReference,potype,notes) " & _
                                "select poKey = '" & strKeycount & "' , storerkey , receipt_no , consigneekey , cust_name , extern , 'A' , description from sdn02t Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '寫入表身
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    
                    str_SQL = "select s3.product_no , s.descr , s3.storerkey , s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) " & _
                                "from sdn03t s3 join gv_skuxpack s on s.sku = s3.product_no and s3.storerkey = s.storerkey " & _
                                "Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' " & _
                                "group by s3.product_no , s.descr , s3.storerkey, s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    'Add by Gemini @20090303
                    If tmp_Rs.EOF Then MsgBox "簽收量等於出貨量，無需產生採購單！", 64, "簽單維護": cn.RollbackTrans: GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                    intLineNumber = intLineNumber + 1
            
                    str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered,packkey) " & _
                            "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & RTrim(tmp_Rs("Descr")) & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "已新增WMS預約收貨採購單(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
            End If
        End If
    tmp_Rs.Close
NoDo:
    End If
End If
    
Screen.MousePointer = vbDefault

''異常費用計算
'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'str_SQL = "select trp_cost = sum(trp_cost) , sorting_cost = sum(sorting_cost) from gv_ExpectCost Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
'tmp_Rs.Open str_SQL, cn
'
''更新SDN02T
'cn.Execute "update sdn02t set trp_cost = '" & tmp_Rs("trp_cost") & "',sorting_cost = '" & tmp_Rs("sorting_cost") & "',Total_Cost = '" & tmp_Rs("trp_cost") + tmp_Rs("sorting_cost") & "' Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

'更新簽單狀態
If rsOrderT0 Is Nothing = False Then
    If rsOrderT0.RecordCount > 0 And rsOrderT0.EOF = False And rsOrderT0("TMS單號") = txt_OneOrder_OrderKey.Text Then
        rsOrderT0("驗收單號") = txt_OneOrder_CustomerOrderkey1
        rsOrderT0("狀態") = "異常訂單"
    End If
End If

'Call ClearForm
Call cmd_OrderQuery_Click
cmbOrderkey.ListIndex = 0: cmbOrderkey.SetFocus

'運費計算
'Call cmdCost_Click'marked by gemini @20111223

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-異常訂單", Me.Caption, "cmd_OneOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_OneOrder_NoDelivery_Click()
On Error GoTo err_Handle
'ㄧ張貨主單號對應ㄧ張排車系統訂單：未出訂單
If Len(dtp_OneOrder_SignDate.Value) = 0 Then
   msg_text = "資料錯誤：未輸入 [客戶簽收日期]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
'If Fun_ChkDateFormat(dtp_OneOrder_SignDate.Value) = 1 Then
'   msg_text = "客戶簽收日期：" & funRtn_msg
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   txt_OneOrder_SignDate.SelStart = 0: txt_OneOrder_SignDate.SelLength = Len(txt_OneOrder_SignDate.Text): txt_OneOrder_SignDate.SetFocus
'   Exit Sub
'End If

If Len(Trim(txt_OneOrder_Status)) > 0 Then MsgBox "此簽單已經維護，無法修改!!", vbOKOnly, Me.Caption: Exit Sub

'清除特殊字元
Call myFormExCharFilter(Me)

Dim strRBC As String, strRSC As String, strResponsible As String

'檢核是否於第一項選取 [異常原因] 與 [責任歸屬]
With gd_OneOrder_OrderDetail
        .Row = 1
        .Col = 9: strRSC = .Text    '異常代碼
        .Col = 10: strRBC = .Text '責任歸屬
        .Col = 14: strResponsible = .Text '責任歸屬人
End With

If Len(Trim(strRSC)) = 0 Or Len(Trim(strRBC)) = 0 Then MsgBox "請於細項第一筆選取 [異常原因] 與 [責任歸屬]", 64, "未出訂單確認": Exit Sub

Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'更新 TRP01T
str_SQL = "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & txt_C_Route_NO.Text & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'更新 SDN02T
str_SQL = "Update SDN02T " & _
              "Set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "' " & _
              ", Confirm_UserID = '" & User_id & "' " & _
              ", cust_handle = '" & txt_CustHandle.Text & "' " & _
              ", TRP_Handle = '" & txt_TRPHandle.Text & "' " & _
              ", Advance = '" & txt_Advance.Text & "' " & _
              ", INV_Handle = '" & txt_INVHandle.Text & "' " & _
              ", TRP_Cost = '" & txt_TRPCost.Text & "' " & _
              ", Sorting_Cost = '" & txt_SortingCost.Text & "' " & _
              ", Total_Cost = '" & txt_TotalCost.Text & "' " & _
              ", SDN_NOTE = '" & txt_SDNNote.Text & "' ,invback = '" & cboInvBack.Text & "' " & _
              ",Confirm_Date = getdate() ,Confirm_Notes = '未出訂單' , CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' " & _
              "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
          
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'更新 SDN03T
str_SQL = "Update SDN03T Set Sign_Qty = 0,RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "',Responsible = '" & strResponsible & "' " & _
          "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'運費計算
cn.Execute "exec gs_Cost '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

'計費數量改為 0
cn.Execute "update sdn05t set chargeqty = 0 , sumreceivable = 0 ,sumpayable = 0 where sdn_no = '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'讀取ini參數，是否入WMS系統
Dim objIni As New vbIniFile, strOtherOrder2WMS As String
objIni.FileName = App.Path & "/" & App.title & ".ini"

strOtherOrder2WMS = objIni.ReadData("OPTION", "OtherOrder2WMS", "YES")
Set objIni = Nothing

If UCase(strOtherOrder2WMS) = "YES" And txt_OneOrder_StorerKey <> "LABT01" Then 'WMS新增採購單 1s

    '寫入WMS採購單
    If txt_Priority.Text <> "R" And txt_Priority.Text <> "RC" And txt_Priority.Text <> "A2B" Then '退貨單、提貨配送與移倉入庫不產生WMS採購單2s
        '毛寶和菲仕蘭
        
        Dim rsKeycount As New ADODB.Recordset, strKeycount As String, intLineNumber As Integer
        If RTrim(txt_OneOrder_StorerKey.Text) = "LMBO01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LPSI01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LCHF01" Then '3s
            '利豐拒短收先不寫入ASN，等雪慧改好再說
            If RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Then GoTo NoDo
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select asnkey from " & strWMSDB & "..asn where buyersreference = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '是否已產生採購單號 4s
            
    '            If MsgBox("WMS是否產生預收採購單?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '取系統採購單號

                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='asn' ", cn
                    '單號+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'asn'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '寫入表頭
                    str_SQL = "insert into " & strWMSDB & "..asn (asnKey,StorerKey,externasnkey , sellersreference,BuyersReference,asntype,notes,buyerVAT) " & _
                                "select asnKey = '" & strKeycount & "' , s2.storerkey , rtrim(o.externorderkey) , o.consigneekey , s2.receipt_no , 'A' , description ,'" & RTrim(txt_OneOrder_FullName1) & "' " & _
                                "from sdn02t s2 join orders o on s2.c_receipt_no = o.orderkey " & _
                                "Where s2.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "'"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '寫入表身
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    
                    str_SQL = "select s3.product_no , s.descr , s3.storerkey , s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) " & _
                                "from sdn03t s3 join gv_skuxpack s on s.sku = s3.product_no and s3.storerkey = s.storerkey " & _
                                "Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' " & _
                                "group by s3.product_no , s.descr , s3.storerkey, s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    'Add by Gemini @20090303
                    If tmp_Rs.EOF Then cn.RollbackTrans: MsgBox "簽收量等於出貨量，無需產生採購單！", 64, "簽單維護": GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                    intLineNumber = intLineNumber + 1
            
                    str_SQL = "insert into " & strWMSDB & "..asndetail (asnKey,asnLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered,packkey) " & _
                            "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & RTrim(tmp_Rs("Descr")) & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "已新增WMS預約收貨採購單(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
            End If  '4e
        Else    '3m
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select pokey from " & strWMSDB & "..po where externpokey = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '是否已產生採購單號5s
            
                'If MsgBox("WMS是否產生預收採購單?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '取系統採購單號
                    'Dim rsKeycount As New ADODB.Recordset, strKeycount As String, intLineNumber As Integer
                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
                    '單號+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '寫入表頭
                    str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,externpokey , sellername,selleraddress1,BuyersReference,potype,notes) " & _
                                "select poKey = '" & strKeycount & "' , storerkey , receipt_no , consigneekey , cust_name , extern , 'A' , description from sdn02t Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '寫入表身
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    str_SQL = "select s3.product_no , s3.storerkey,s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) from sdn03t s3 join gv_skuxpack s on s.storerkey = s3.storerkey and s.sku = s3.product_no Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' group by s3.product_no , s3.storerkey,s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    If tmp_Rs.EOF Then cn.RollbackTrans: MsgBox "簽收量等於出貨量，無需產生採購單！", 64, "簽單維護": GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                        intLineNumber = intLineNumber + 1
                
                        str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,StorerKey,QtyOrdered,packkey) " & _
                                "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "') "
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "已新增WMS預約收貨採購單(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
                End If '5e
NoDo:
        tmp_Rs.Close
        End If '3e
    End If  '2e
End If  '1e

Screen.MousePointer = vbDefault

'異常費用計算
'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'str_SQL = "select trp_cost = sum(trp_cost) , sorting_cost = sum(sorting_cost) from gv_ExpectCost Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
'tmp_Rs.Open str_SQL, cn

''更新SDN02T
'cn.Execute "update sdn02t set trp_cost = '" & tmp_Rs("trp_cost") & "',sorting_cost = '" & tmp_Rs("sorting_cost") & "',Total_Cost = '" & tmp_Rs("trp_cost") + tmp_Rs("sorting_cost") & "' Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

cmbOrderkey.ListIndex = 0: cmbOrderkey.SetFocus

'更新簽單狀態
If rsOrderT0 Is Nothing = False Then
    If rsOrderT0.RecordCount > 0 And rsOrderT0.EOF = False And rsOrderT0("TMS單號") = txt_OneOrder_OrderKey.Text Then
        rsOrderT0("驗收單號") = txt_OneOrder_CustomerOrderkey1
        rsOrderT0("狀態") = "未出訂單"
    End If
End If

Call cmd_OrderQuery_Click

'運費計算
'Call cmdCost_Click'marked by gemini @20111223

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-未出訂單", Me.Caption, "cmd_OnOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   
End Sub

Public Sub cmd_OrderQuery_Click()
'訂單查詢
If Trim(txt_OrderKey.Text) = "" Then Exit Sub
On Error GoTo err_Handle

Dim strOrderkey As String, strOrderType As String
strOrderkey = Trim(txt_OrderKey.Text)
strOrderType = cmbOrderkey.Text

If cmbOrderkey = "" Then
    Call ClearForm
    txt_OrderKey.Text = strOrderkey
    str_SQL = "Select Receipt_No From SDN02T (nolock) Where receipt_no = '" & strOrderkey & "' or receipt_no = '" & Format(strOrderkey, "0000000000") & "' or extern = '" & strOrderkey & "' "

ElseIf cmbOrderkey.Text = "TMS單號" Then
    strOrderkey = Format(Trim(txt_OrderKey.Text), "0000000000")
    Call ClearForm
    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
    str_SQL = "Select Receipt_No From SDN02T (nolock) Where receipt_no = '" & strOrderkey & "' "
    
ElseIf cmbOrderkey.Text = "貨主單號" Then
    Call ClearForm
    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
    str_SQL = "Select Receipt_No From SDN02T (nolock) Where extern = '" & strOrderkey & "' "
    
Else
    Exit Sub
End If

'檢查單號是否一對多
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.RecordCount = 0 Then
   tmp_Rs.Close
   msg_text = "查詢無資料!!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_OrderKey.SelStart = 0:   txt_OrderKey.SelLength = Len(txt_OrderKey.Text): txt_OrderKey.SetFocus
   Exit Sub
   
ElseIf tmp_Rs.RecordCount = 1 Then
   strOrderkey = RTrim(tmp_Rs("Receipt_No"))
   tmp_Rs.Close
   Call Display_OrderData_OneReceipNo(strOrderkey)
   
    '不允許操作 add @20120112
    cmdCost.Enabled = False: cmdSDNBack.Enabled = False
   
Else
tmp_Rs.Close

    'ㄧ筆貨主單號對應多張排車系統訂單
    frm_MulitiTMSOrder2.Show vbModal
    '   tmp_rs.Close
    '   Call Display_OrderData_MultiReceipNo(strOrderKey)

End If

'補資料
str_SQL = "select 分公司代碼=isnull(co.BranchId,' '),客戶訂單類別=isnull(o.externordertype,' ') from orders o (nolock) left join custorders co on o.orderkey = co.orderkey where o.orderkey = '" & RTrim(txt_OneOrder_OrderKey) & "'"
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.RecordCount = 1 Then
    txt_Externordertype = tmp_Rs.Fields("客戶訂單類別")
    txt_BranchId = tmp_Rs.Fields("分公司代碼")
Else
    tmp_Rs.Close
End If

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "cmd_OrderQuery_Click()")
End Sub

Public Sub cmd_OrderQuery_Click_del()
'訂單查詢
If Trim(txt_OrderKey.Text) = "" Then Exit Sub
On Error GoTo err_Handle

Dim strOrderkey As String, strOrderType As String
strOrderkey = Trim(txt_OrderKey.Text)
strOrderType = cmbOrderkey.Text

If cmbOrderkey.Text = "TMS單號" Then
    strOrderkey = Format(Trim(txt_OrderKey.Text), "0000000000")
    Call ClearForm
    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
    str_SQL = "Select Count(Distinct Receipt_No) as OrderCnt From SDN02T Where receipt_no = '" & strOrderkey & "' "

Else
    Call ClearForm
    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
    str_SQL = "Select Count(Distinct Receipt_No) as OrderCnt From SDN02T Where extern = '" & strOrderkey & "' "
End If

'檢查貨主單號在排車系統有無進行 [訂單切割]
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.Fields("OrderCnt").Value = 0 Then
   tmp_Rs.Close
   msg_text = "查詢無資料!!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_OrderKey.SelStart = 0:   txt_OrderKey.SelLength = Len(txt_OrderKey.Text): txt_OrderKey.SetFocus
   Exit Sub
ElseIf tmp_Rs.Fields("OrderCnt").Value = 1 Then
   'ㄧ筆貨主單號對應ㄧ張排車系統訂單
   tmp_Rs.Close
   Call Display_OrderData_OneReceipNo(strOrderkey)
Else
tmp_Rs.Close

'ㄧ筆貨主單號對應多張排車系統訂單
frm_MulitiTMSOrder.Show vbModal
'   tmp_rs.Close
'   Call Display_OrderData_MultiReceipNo(strOrderKey)

End If

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "cmd_OrderQuery_Click()")
End Sub

Private Sub cmd2Excel_Click()

'資料排序
Recordset2Excel "簽單查詢", rsMain1

'..在此編輯EXCEL
With MyXlsApp
  
End With

Set MyXlsApp = Nothing

End Sub

Private Sub cmdCarNOChange_Click()
'If blAdmin = False Then MsgBox "系統管理員才有權限執行此作業!", 64, "權限不足": Exit Sub
If Len(RTrim(txt_OneOrder_VehicleID.Text)) = 0 Or Len(RTrim(txt_C_Route_NO.Text)) = 0 Then Exit Sub

'此路編是否已計費
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select sum(應收總價+應付總價) from gv_sdn05tdetail where 二次路編 = '" & txt_C_Route_NO & "' "

tmp_Rs.Open str_SQL, cn
If Not tmp_Rs.EOF Then

    If tmp_Rs(0) <> 0 Then
        MsgBox "此簽單已維護運費，無法變更車號，請洽計費人員！", 16, "注意": tmp_Rs.Close: Exit Sub
    End If

End If

tmp_Rs.Close

intSDNCarChange = 1 '由異常簽單維護進入

frm_SDNCarNOFix.Show vbModal
End Sub

Private Sub cmdDeliveryokT0_Click()
On Error GoTo err_Handle

'是否有資料
If rsOrderT0 Is Nothing Then Exit Sub
If rsOrderT0.RecordCount = 0 Then Exit Sub

rsOrderT0.Filter = "狀態 = 'V'"

If rsOrderT0.RecordCount = 0 Then rsOrderT0.Filter = "": rsOrderT0.Sort = "編號": Exit Sub

rsOrderT0.MoveFirst
dgOrderT0.Col = 0

''更新請款人
'cn.Execute "update SDN01T set sdn_Date = getdate() , receiver = '" & rsRouteT0("請款人") & "' where c_route_no = '" & rsRouteT0("二次路編") & "'", RowsAffect, adExecuteNoRecords

Do While Not rsOrderT0.EOF

    '簽單是否已維護
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select * from sdn02t where confirm_notes <> '' and receipt_no = '" & rsOrderT0("TMS單號") & "' "
    tmp_Rs.Open str_SQL, cn
    If Not tmp_Rs.EOF Then MsgBox "此簽單已經維護過!!", 16, "TMS單號：" & rsOrderT0("TMS單號"): tmp_Rs.Close: Exit Sub
    tmp_Rs.Close
    
    '訂單量與出貨量不符
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select * from sdn03t where order_qty <> ship_qty and receipt_no = '" & rsOrderT0("TMS單號") & "' "
    tmp_Rs.Open str_SQL, cn
    If Not tmp_Rs.EOF Then MsgBox "訂單量與出貨量不符!!", 16, "TMS單號：" & rsOrderT0("TMS單號"): tmp_Rs.Close: Exit Sub
    tmp_Rs.Close
    
    Screen.MousePointer = 11
    Tran_Level = cn.BeginTrans
     
    '更新 SDN02T
    str_SQL = "Update SDN02T Set CustSignDate = isnull(CustSignDate,isnull(SCHEDULEDATE,Arrive_Date)), invback = 'N',sdnback = 1, " & _
              "Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '正常訂單' , CustomerOrderkey1 ='" & rsOrderT0("驗收單號") & "', Scan = 'N',SDNSendDate = '" & Format(Now, "YYYY/MM/DD") & "' " & _
              "Where Receipt_No = '" & rsOrderT0("TMS單號") & "'"
    
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '更新 SDN03T
    str_SQL = "Update SDN03T Set Sign_Qty =  ship_Qty,RSC_Code = '',RBC_Code = '' Where Receipt_No = '" & rsOrderT0("TMS單號") & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '更新簽單狀態
    rsOrderT0("狀態") = "正常訂單"
    
    ''運費計算
    cn.Execute "exec gs_Cost '" & rsOrderT0("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans: Tran_Level = 0
    Screen.MousePointer = vbDefault
    
    rsOrderT0.MoveNext
Loop

Screen.MousePointer = 0
rsOrderT0.Filter = "": rsOrderT0.Sort = "編號"

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdExit_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdOpenOrderT0_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

If dgOrderT0.DataSource Is Nothing Then Exit Sub
If Len(RTrim(rsOrderT0("TMS單號"))) < 10 Then Exit Sub

cmbOrderkey = "TMS單號"
txt_OrderKey = rsOrderT0("TMS單號")
SSTab1.Tab = 3
DoEvents: DoEvents
Call cmd_OrderQuery_Click

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

    '一次排車路線編號
        str_SQL = "select * from gv_Sdn05tDetail where 1 = 1 "
        
    Dim str_Where As String
    
    '路線編號
    str_Where = ""
    If Len(RTrim(txtRouteS.Text)) > 0 And Len(RTrim(txtRouteE.Text)) > 0 Then
        str_Where = str_Where & " and 路線編號 Between '" & RTrim(txtRouteS.Text) & "' and '" & RTrim(txtRouteE.Text) & "' "
    ElseIf Len(RTrim(txtRouteS.Text)) > 0 And Len(RTrim(txtRouteE.Text)) = 0 Then
        str_Where = str_Where & " and 路線編號 = '" & RTrim(txtRouteS.Text) & "' "
    ElseIf Len(RTrim(txtRouteS.Text)) = 0 And Len(RTrim(txtRouteE.Text)) > 0 Then
        str_Where = str_Where & " and 路線編號 = '" & RTrim(txtRouteE.Text) & "' "
    End If
    
    '二次路編
    If Len(RTrim(txt2RouteS.Text)) > 0 And Len(RTrim(txt2RouteE.Text)) > 0 Then
        str_Where = str_Where & " and 二次路編 Between '" & RTrim(txt2RouteS.Text) & "' and '" & RTrim(txt2RouteE.Text) & "' "
    ElseIf Len(RTrim(txt2RouteS.Text)) > 0 And Len(RTrim(txt2RouteE.Text)) = 0 Then
        str_Where = str_Where & " and 二次路編 = '" & RTrim(txt2RouteS.Text) & "' "
    ElseIf Len(RTrim(txt2RouteS.Text)) = 0 And Len(RTrim(txt2RouteE.Text)) > 0 Then
        str_Where = str_Where & " and 二次路編 = '" & RTrim(txt2RouteE.Text) & "' "
    End If

    '貨主單號
    If Len(RTrim(txtExternS.Text)) > 0 And Len(RTrim(txtExternE.Text)) > 0 Then
        str_Where = str_Where & " and 貨主單號 Between '" & RTrim(txtExternS.Text) & "' and '" & RTrim(txtExternE.Text) & "' "
    ElseIf Len(RTrim(txtExternS.Text)) > 0 And Len(RTrim(txtExternE.Text)) = 0 Then
        str_Where = str_Where & " and 貨主單號 = '" & RTrim(txtExternS.Text) & "' "
    ElseIf Len(RTrim(txtExternS.Text)) = 0 And Len(RTrim(txtExternE.Text)) > 0 Then
        str_Where = str_Where & " and 貨主單號 = '" & RTrim(txtExternE.Text) & "' "
    End If
        
    'TMS單號
    If Len(RTrim(txtOrderkeyS.Text)) > 0 And Len(RTrim(txtOrderkeyE.Text)) > 0 Then
        txtOrderkeyS.Text = Format(txtOrderkeyS.Text, "0000000000"): txtOrderkeyE.Text = Format(txtOrderkeyE.Text, "0000000000")
        str_Where = str_Where & " and TMS單號 Between '" & RTrim(txtOrderkeyS.Text) & "' and '" & RTrim(txtOrderkeyE.Text) & "' "
    ElseIf Len(RTrim(txtOrderkeyS.Text)) > 0 And Len(RTrim(txtOrderkeyE.Text)) = 0 Then
        txtOrderkeyS.Text = Format(txtOrderkeyS.Text, "0000000000")
        str_Where = str_Where & " and TMS單號 = '" & RTrim(txtOrderkeyS.Text) & "' "
    ElseIf Len(RTrim(txtOrderkeyS.Text)) = 0 And Len(RTrim(txtOrderkeyE.Text)) > 0 Then
        txtOrderkeyS.Text = Format(txtOrderkeyS.Text, "0000000000")
        str_Where = str_Where & " and TMS單號 = '" & RTrim(txtOrderkeyE.Text) & "' "
    End If
    
    '到貨日期
    If Len(RTrim(txtDeliveryS.Text)) > 0 And Len(RTrim(txtDeliveryE.Text)) > 0 Then
        str_Where = str_Where & " and 到貨日 Between '" & RTrim(txtDeliveryS.Text) & "' and '" & RTrim(txtDeliveryE.Text) & "' "
    ElseIf Len(RTrim(txtDeliveryS.Text)) > 0 And Len(RTrim(txtDeliveryE.Text)) = 0 Then
        str_Where = str_Where & " and 到貨日 = '" & RTrim(txtDeliveryS.Text) & "' "
    ElseIf Len(RTrim(txtDeliveryS.Text)) = 0 And Len(RTrim(txtDeliveryE.Text)) > 0 Then
        str_Where = str_Where & " and 到貨日 = '" & RTrim(txtDeliveryE.Text) & "' "
    End If
    
    '簽單日期
    If Len(RTrim(txtSignDateS.Text)) > 0 And Len(RTrim(txtSignDateE.Text)) > 0 Then
        str_Where = str_Where & " and isnull(convert(varchar(8),簽單日,112),'') Between '" & RTrim(txtSignDateS.Text) & "' and '" & RTrim(txtSignDateE.Text) & "' "
    ElseIf Len(RTrim(txtSignDateS.Text)) > 0 And Len(RTrim(txtSignDateE.Text)) = 0 Then
        str_Where = str_Where & " and isnull(convert(varchar(8),簽單日,112),'') = '" & Len(RTrim(txtSignDateS.Text)) & "' "
    ElseIf Len(RTrim(txtSignDateS.Text)) = 0 And Len(RTrim(txtSignDateE.Text)) > 0 Then
        str_Where = str_Where & " and isnull(convert(varchar(8),簽單日,112),'') = '" & Len(RTrim(txtSignDateE.Text)) & "' "
    End If
    
    '貨主
    If Len(RTrim(cboStorerkey.Text)) > 0 Then str_Where = str_Where & " and 貨主 = '" & RTrim(cboStorerkey.Text) & "' "
    
    '車號
    If Len(RTrim(cboCar.Text)) > 0 Then str_Where = str_Where & " and 車號 = '" & RTrim(cboCar.Text) & "' "

    '請款類別
    If Len(RTrim(cboCostkind.Text)) > 0 Then str_Where = str_Where & " and 請款類別 = '" & RTrim(cboCostkind.Text) & "' "
    
    str_SQL = str_SQL & str_Where & " Order by 到貨日,二次路編 "
    
    On Error GoTo err_Handle
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "查詢結果：無符合搜尋條件之排車資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Set dgMain1.DataSource = Nothing: Set rsMain1 = Nothing
        Exit Sub
    End If
    
    Call Replication_Recordset(tmp_Rs, rsMain1)
    
    txtAR.Text = 0: txtAP.Text = 0
    
    Do While Not rsMain1.BOF
    
        txtAR.Text = txtAR.Text + rsMain1("應收總價")
        txtAP.Text = txtAP.Text + rsMain1("應付總價")
        rsMain1.MovePrevious
    Loop
    
        txtAR.Text = Round(txtAR.Text, 0): txtAP.Text = Round(txtAP.Text, 0)
        txtEarning.Text = txtAR.Text - txtAP.Text
    
    Set dgMain1.DataSource = rsMain1
    tmp_Rs.Close
    
    DoEvents: DoEvents
'    rsMain1.MoveFirst
    SetDataGridColWidth Me.Caption, dgMain1
    Screen.MousePointer = 0

    Exit Sub
    
err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "送貨簽單確認-簽單查詢", Me.Caption, "cmd_Tab1_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_AddCost_Click()
    '新增一行
    dg_Tab2_SDN_Cost.Col = 2
    dg_Tab2_SDN_Cost.Rows = dg_Tab2_SDN_Cost.Rows + 1
    dg_Tab2_SDN_Cost.Row = dg_Tab2_SDN_Cost.Row + 1
    NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
End Sub

Private Sub cmd_Tab2_AddNew_Click()
    Call Clear_CardData
    cmd_Tab2_Save.Enabled = True
    cmd_Tab2_Cancel.Enabled = True
    cmd_Tab2_AddNew.Enabled = False
    cmd_Tab2_Modify.Enabled = False
    cmd_Tab2_Delete.Enabled = False
    txt_Tab02_C_VEHICLE_ID_NO.Enabled = True
    txt_Tab02_Driver.Enabled = True
    txt_Tab02_Receiver.Enabled = True
    txt_Tab02_Delivery_Date.Enabled = True
    cmd_Tab2_SelectCar.Enabled = True
    txt_Tab02_Delivery_Date.SetFocus
End Sub

Private Sub cmd_Tab2_AddOrder_Click()
    '新增一行
    dg_Tab2_SDN_Detail.Col = 2
    dg_Tab2_SDN_Detail.Rows = dg_Tab2_SDN_Detail.Rows + 1
    dg_Tab2_SDN_Detail.Row = dg_Tab2_SDN_Detail.Row + 1
    NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
End Sub

Private Sub cmd_Tab2_Cancel_Click()
    '卡鐘資料 >> 取消
    Call Clear_CardData
    cmd_Tab2_Cancel.Enabled = False
    cmd_Tab2_Save.Enabled = False
    cmd_Tab2_AddNew.Enabled = True
    cmd_Tab2_Modify.Enabled = True
    cmd_Tab2_Delete.Enabled = True
    cmd_Tab2_SelectCar.Enabled = False
    txt_Tab02_C_VEHICLE_ID_NO.Enabled = False
    txt_Tab02_Driver.Enabled = False
    txt_Tab02_Receiver.Enabled = False
    txt_Tab02_Delivery_Date.Enabled = False
End Sub

Private Sub cmd_Tab2_DelCost_Click()
    '刪除一行
    If dg_Tab2_SDN_Cost.Rows > 2 Then
        dg_Tab2_SDN_Cost.Rows = dg_Tab2_SDN_Cost.Rows - 1
        dg_Tab2_SDN_Cost.Row = dg_Tab2_SDN_Cost.Rows - 1
        NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
    End If
End Sub

Private Sub cmd_Tab2_DelOrder_Click()
    '刪除一行
    If dg_Tab2_SDN_Detail.Rows > 2 Then
        dg_Tab2_SDN_Detail.Rows = dg_Tab2_SDN_Detail.Rows - 1
        dg_Tab2_SDN_Detail.Row = dg_Tab2_SDN_Detail.Rows - 1
        NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
    End If
End Sub

Private Sub cmd_Tab2_Save_Click()
    If Len(Trim(txt_Tab02_Delivery_Date.Text)) = 0 Then
        msg_text = "必須輸入出車日"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(txt_Tab02_C_VEHICLE_ID_NO.Text)) = 0 Then
        msg_text = "必須輸入車號"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(Trim(txt_Tab02_Receiver.Text))) = 0 Then
        msg_text = "必須輸入領款人"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    For i = 1 To dg_Tab2_SDN_Detail.Rows - 1
        dg_Tab2_SDN_Detail.Row = i
        dg_Tab2_SDN_Detail.Col = 2: str_EXTERN = Trim(dg_Tab2_SDN_Detail.Text)
        dg_Tab2_SDN_Detail.Col = 3: str_CUST_NAME = Trim(dg_Tab2_SDN_Detail.Text)
'        If Len(Trim(str_EXTERN)) = 0 Or Len(Trim(str_CUST_NAME)) = 0 Then
'            msg_text = "裝載明細資料不齊"
'            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'            Exit Sub
'        End If
        str_SQL = "select * from SDN02T where EXTERN='" & str_EXTERN & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If Not tmp_Rs.EOF Then
            tmp_Rs.Close
            msg_text = "客戶單號重複"
            MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
        tmp_Rs.Close
    End If
    Next
    On Error GoTo err_Handle
    '取得路編
    str_SQL = "select isnull(max(C_Route_No),0) from Logictown.dbo.SDN01T where left(C_Route_No,2)='WD'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    str_C_ROUTE_NO = "WD" & StrPadLeft(Val(Right(Trim(tmp_Rs.Fields(0)), 8)) + 1, 8, 0)
    tmp_Rs.Close
    cn.BeginTrans
        '存表頭,SDN01T
        str_DELIVERY_DATE = Trim(txt_Tab02_Delivery_Date.Text)
        str_C_VEHICLE_ID_NO = Trim(txt_Tab02_C_VEHICLE_ID_NO.Text)
        str_Driver = Trim(txt_Tab02_Driver.Text)
        Str_Receiver = Trim(txt_Tab02_Receiver.Text)
        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,Receiver,SDNStatus,AddUser)" & _
            "Values ( '" & str_DELIVERY_DATE & "','" & str_C_ROUTE_NO & "','" & str_C_VEHICLE_ID_NO & "','" & str_Driver & "','" & Str_Receiver & "', " & _
            "'" & str_SDNStatus & "','" & User_id & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '存表身:裝載明細,SDN02T
        For i = 1 To dg_Tab2_SDN_Detail.Rows - 1
            dg_Tab2_SDN_Detail.Row = i
            dg_Tab2_SDN_Detail.Col = 2: str_EXTERN = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 3: str_CUST_NAME = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 4: str_SHIP_CS = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 5: str_SHIP_CBM = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 6: str_SHIP_WT = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 7: str_CAR_NOTES = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 8: str_SDN_NOTE = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 1
            str_SDNStatus = 0
'            If Trim(dg_SDN_Detail.Text) = "Ｖ" Then str_SDNStatus = 1
'            dg_SDN_Detail.Col = 0
'            If Trim(dg_SDN_Detail.Text) = "Ｖ" Then str_SDNStatus = 2
            If Len(str_EXTERN) = 0 And Len(str_CUST_NAME) = 0 And Len(str_SHIP_CS) = 0 And Len(str_SHIP_CBM) = 0 And Len(str_SHIP_WT) = 0 And Len(str_CAR_NOTES) = 0 And Len(str_SDN_NOTE) = 0 Then
                '無資料不存檔
            Else
                str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,RECEIPT_NO) " & _
                    "Values ( '" & str_C_ROUTE_NO & "','" & str_C_ROUTE_NO & "','" & str_EXTERN & "','" & str_DELIVERY_DATE & "','" & str_CUST_NAME & "', " & _
                    "'" & str_SHIP_CS & "','" & str_SHIP_CBM & "','" & str_SHIP_WT & "','" & str_CAR_NOTES & "','" & str_SDNStatus & "','" & str_SDN_NOTE & "','CT" & str_EXTERN & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
            
        Next
        '存表身:計費項目,SDN05T
        For i = 1 To dg_Tab2_SDN_Cost.Rows - 1
            dg_Tab2_SDN_Cost.Row = i
            dg_Tab2_SDN_Cost.Col = 1: str_SDN_Name = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 2: str_SDN_NO = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 3: str_AreaStart = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 4: str_AreaEnd = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 5: str_uom = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 6: str_ChargeQty = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 7: str_Receivable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 8: str_Payable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 9: str_Premiam = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 10: str_Reason = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 11: str_SumReceivable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 12: str_SumPayable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 13: str_CostKind = Trim(dg_Tab2_SDN_Cost.Text)
            If Len(str_SDN_NO) = 0 And Len(str_AreaEnd) = 0 And Len(str_AreaStart) = 0 And Len(str_SumPayable) = 0 And Len(str_SumReceivable) = 0 And Len(str_Reason) = 0 And Len(str_Premiam) = 0 And Len(str_Payable) = 0 And Len(str_Receivable) = 0 And Len(str_ChargeQty) = 0 And Len(str_uom) = 0 Then
                '無資料不存檔
            Else
                str_SQL = "Insert into SDN05T (C_ROUTE_NO,Uom,ChargeQty,Receivable,Payable,Premiam,Reason,SumReceivable,SumPayable,AreaStart,AreaEnd,SDN_NO,SDN_Name,CostKind) " & _
                    "Values ( '" & str_C_ROUTE_NO & "','" & str_uom & "','" & str_ChargeQty & "','" & str_Receivable & "','" & str_Payable & "', " & _
                    "'" & str_Premiam & "','" & str_Reason & "','" & str_SumReceivable & "','" & str_SumPayable & "','" & str_AreaStart & "','" & str_AreaEnd & "','" & str_SDN_NO & "','" & str_SDN_Name & "','" & str_CostKind & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
        Next
    cn.CommitTrans
    txt_Tab02_C_Route_No.Text = str_C_ROUTE_NO
    cmd_Tab2_Cancel.Enabled = False
    cmd_Tab2_Save.Enabled = False
    cmd_Tab2_AddNew.Enabled = True
    cmd_Tab2_Modify.Enabled = False
    cmd_Tab2_Delete.Enabled = False
    cmd_Tab2_SelectCar.Enabled = False
    txt_Tab02_C_VEHICLE_ID_NO.Enabled = False
    txt_Tab02_Driver.Enabled = False
    txt_Tab02_Receiver.Enabled = False
    txt_Tab02_Delivery_Date.Enabled = False
    Exit Sub
    
err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "送貨簽單確認-存檔", Me.Caption, "cmd_Tab0_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_SelectCar_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab2_SelectCar.Name & "2")
End Sub

Private Sub cmdNotYetOrder_Click()
    
    txt_OrderKey = ""
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select * from sdn02t where len(rtrim(isnull(confirm_notes,''))) = 0 "
    tmp_Rs.Open str_SQL, cn
    If tmp_Rs.EOF Then MsgBox "無待確認簽單!!", vbOKOnly, Me.cmdNotYetOrder.Caption: Exit Sub
    
    frm_SDNConfirmNotYet.Show vbModal
    
End Sub


Private Sub cmdCost_Click()

If Len(RTrim(txt_OneOrder_OrderKey.Text)) = 0 Then Exit Sub

frm_Cost.Show vbModal
End Sub

Private Sub cmdQueryT0_Click()

On Error GoTo err_Handle
Dim str_Where As String

'出車日期
If Len(RTrim(txtDeliveryDateST0.Text)) > 0 And Len(RTrim(txtDeliveryDateET0.Text)) > 0 Then
    str_Where = str_Where & " and convert(Char(8), t1.delivery_date, 112) Between '" & RTrim(txtDeliveryDateST0.Text) & "' and '" & RTrim(txtDeliveryDateET0.Text) & "' "
ElseIf Len(RTrim(txtDeliveryDateST0.Text)) > 0 And Len(RTrim(txtDeliveryDateET0.Text)) = 0 Then
    str_Where = str_Where & " and convert(Char(8), t1.delivery_date, 112) = '" & RTrim(txtDeliveryDateST0.Text) & "' "
ElseIf Len(RTrim(txtDeliveryDateST0.Text)) = 0 And Len(RTrim(txtDeliveryDateET0.Text)) > 0 Then
    str_Where = str_Where & " and convert(Char(8), t1.delivery_date, 112) = '" & RTrim(txtDeliveryDateET0.Text) & "' "
End If

'二次路編
If Len(RTrim(txtRouteST0.Text)) > 0 And Len(RTrim(txtRouteET0.Text)) > 0 Then
    str_Where = str_Where & " and t1.c_route_no Between '" & RTrim(txtRouteST0.Text) & "' and '" & RTrim(txtRouteET0.Text) & "' "
ElseIf Len(RTrim(txtRouteST0.Text)) > 0 And Len(RTrim(txtRouteET0.Text)) = 0 Then
    str_Where = str_Where & " and t1.c_route_no = '" & RTrim(txtRouteST0.Text) & "' "
ElseIf Len(RTrim(txtRouteST0.Text)) = 0 And Len(RTrim(txtRouteET0.Text)) > 0 Then
    str_Where = str_Where & " and t1.c_route_no = '" & RTrim(txtRouteET0.Text) & "' "
End If

If Len(RTrim(cboStorerT0)) > 0 Then str_Where = str_Where & "and t2.storerkey = '" & RTrim(cboStorerT0) & "' "
If Len(RTrim(cboCarT0)) > 0 Then str_Where = str_Where & "and t1.c_vehicle_id_no = '" & RTrim(cboCarT0) & "' "

str_SQL = "select distinct 選取 = ' ' " & _
        ",出車日期 = convert(Char(8), t1.delivery_date, 112) " & _
        ",二次路編 = t1.c_route_no " & _
        ",車牌號碼 = rtrim(t1.c_vehicle_id_no) " & _
        ",駕駛人 = rtrim(t1.driver) " & _
        ",請款人 = rtrim(isnull(t1.receiver,'')) " & _
        ",新增 = rtrim(t1.adduser) " & _
        ",新增時間 = t1.adddate " & _
        "From sdn01t t1 join sdn02t t2 on t1.c_route_no = t2.c_route_no where t2.storerkey <> 'LTHL01' "
        
str_SQL = str_SQL & str_Where & "order by 出車日期,二次路編"

Screen.MousePointer = 11

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "查詢結果：無符合搜尋條件之路編資料"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Set dgRouteT0.DataSource = Nothing: Set rsRouteT0 = Nothing
    Set dgOrderT0.DataSource = Nothing: Set rsOrderT0 = Nothing
    Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rsRouteT0)
tmp_Rs.Close: rsRouteT0.MoveFirst

Set dgRouteT0.DataSource = rsRouteT0

Call dgRouteT0_RowColChange(1, 1)

SetDataGridColWidth Me.Caption, dgRouteT0
Screen.MousePointer = 0
blRouteT0Change = True

Exit Sub
    
err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "簽單確認-簽單查詢", Me.Caption, "cmdQueryT0_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdReceiptDetail_Click()
On Error GoTo err_Handle
If Len(RTrim(txt_OneOrder_StorerOrderKey)) = 0 Then Exit Sub

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)

'毛寶菲仕蘭帶ordertype+externorderkey , 菲仕蘭帶externorderkey
If RTrim(txt_OneOrder_StorerKey.Text) = "LMBO01" Then
    str_SQL = "exec es_SDNReceiptDetail '" & RTrim(txt_OneOrder_StorerOrderKey) & "' "
ElseIf RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Then
    str_SQL = "exec es_SDNReceiptDetail '" & RTrim(txt_OneOrder_StorerOrderKey) & "' "
Else
'其他
    str_SQL = "exec gs_SDNReceiptDetail '" & txt_OneOrder_OrderKey & "' "
End If
tmp_Rs.Open str_SQL, cn
If tmp_Rs.EOF Then MsgBox "查無資料!", 64, Me.Caption: tmp_Rs.Close: Exit Sub

'轉Excel
Recordset2Excel "簽單查詢", tmp_Rs

'..在此編輯EXCEL
With MyXlsApp
  
End With

Set MyXlsApp = Nothing
tmp_Rs.Close

Exit Sub
err_Handle:
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdReset_Click()
Call ClearForm_AllField(Me)
txtDeliveryS = Format(Now - 30, "YYYYMMDD")
txtDeliveryE = Format(Now + 7, "YYYYMMDD")
End Sub

Private Sub cmdSaveToText_Click()
If rsMain1 Is Nothing Then Exit Sub: If rsMain1.EOF Then Exit Sub
End Sub

Private Sub cmdSDNBack_Click()

If Len(RTrim(txt_OneOrder_OrderKey.Text)) = 0 Then Exit Sub

On Error GoTo err_Handle

Tran_Level = cn.BeginTrans

str_SQL = "update sdn02t set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "', invback = '" & cboInvBack.Text & "' ,sdnback = '1' " & ",SDNSendDate = '" & Format(dtpSDNSendDate.Value, "YYYY/MM/DD") & "', CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' where receipt_no = '" & txt_OneOrder_OrderKey & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

cmdSDNBack.BackColor = vbGreen
cmdSDNBack.Caption = "簽單已回"

'Call cmdCost_Click 不允許操作 @20120112
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdShipNotes_Click()

If Len(RTrim(txt_OneOrder_StorerKey)) = 0 Then Exit Sub
If MsgBox("是否補印出貨單?", vbOKCancel, "列印") <> vbOK Then Exit Sub

On Error GoTo err_Handle

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)

Dim rs_Access As New ADODB.Recordset
Call AccessDB_Connect
strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)
Tran_Level = cnAccess.BeginTrans

If txt_OneOrder_StorerKey = "LVTL01" And Left(txt_C_Route_NO, 1) <> "R" Then

    'VTL出貨單
    str_SQL = "select * from gv_ReportShipNotesVTL Where 佰事達單號 = '" & txt_OneOrder_OrderKey & "' "
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then '無資料時無須列印

    str_SQL = "Delete From VTL出貨單"
    cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
    rs_Access.Open "VTL出貨單", cnAccess, adOpenStatic, adLockOptimistic
    With tmp_Rs
        .MoveFirst
        Do While Not .EOF
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
           rs_Access.Fields("備註").Value = .Fields("備註").Value
           rs_Access.Fields("USER").Value = User_Name
           rs_Access.Update
           .MoveNext
        Loop
    
    End With
    cnAccess.CommitTrans: Tran_Level = 0
    Call DB_Disconnect(cnAccess)
    MSAccessAP.DoCmd.OpenReport "VTL出貨單", acViewPreview
    MSAccessAP.DoCmd.Maximize
    MSAccessAP.Visible = True
    
    End If
    
ElseIf Left(txt_C_Route_NO, 1) <> "R" Then '其他出貨單

    str_SQL = "Select * From gv_ReportShipNotes where 佰事達單號 = '" & txt_OneOrder_OrderKey & "' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If Not tmp_Rs.EOF Then '無資料時無須列印
            str_SQL = "Delete From VLL出貨單"
            cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Call ReDim_Recordset(rs_Access)
            rs_Access.Open "VLL出貨單", cnAccess, adOpenStatic, adLockOptimistic
            With tmp_Rs
                .MoveFirst
                Do While Not .EOF
                   rs_Access.AddNew
'                   rs_Access.Fields("編號").Value = .Fields("編號").Value
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
                   .MoveNext
                Loop
            End With
            
            cnAccess.CommitTrans: Tran_Level = 0
            Call DB_Disconnect(cnAccess)
            MSAccessAP.DoCmd.OpenReport "VLL出貨單", acViewPreview
            MSAccessAP.DoCmd.Maximize
            MSAccessAP.Visible = True

    End If
Else '退貨單

    str_SQL = "select 訂單類別 = case o2t.priority when 'RC' then '提貨入庫單' when 'A2B' then '提貨配送單' else case when o2t.storerkey = 'LTKK01' and substring(o2t.extern,3,2) = '12' then '退貨單(換貨)' else '退貨單' end end " & _
            ", 貨主名稱 =  (select rtrim(t16.c_name) from trp16m t16 where t16.storerkey = o2t.storerkey ) " & _
            ", 路線編號 = o2t.route_no , 參考路編 = o.ContainerType " & _
            ", 出車日期 = convert(char(8) , o1t.delivery_date , 112) " & _
            ", 收貨日期 = convert(char(8) , o2t.arrive_date , 112) " & _
            ", 車號 = o2t.vehicle_id_no , 駕駛 = t9m.driver " & _
            ", TMS單號 = o2t.receipt_no + '(補)' , 貨主單號 = o2t.extern " & _
            ", 客戶訂單號碼 = o.customerorderkey " & _
            ", 客戶名稱 = t1m.short_name , 客戶地址 = t1m.address ,電話 = t1m.phone, 客戶需求 = t1m.notes " & _
            ", 到貨客戶 = case when len(rtrim(o.b_company)) = 0 then '' else '貨送：' + rtrim(t1ma.short_name) + '-'+ rtrim(t1ma.address) + ' ' + rtrim(t1ma.phone) end " & _
            ", 項次 = rtrim(o3t.seq_no) , 貨號 = Rtrim(o3t.Product_No)  " & _
            ", 品名 = sp.descr " & _
            ", 箱數 =isnull(case when sp.casecnt = 0 then 0 else floor(o3t.order_qty/sp.Casecnt) end ,0) ,大包裝 = isnull(rtrim(sp.busr3),'箱') " & _
            ", 個數 =isnull(case when sp.casecnt = 0 then o3t.order_qty else cast(o3t.order_qty as int)%cast(sp.Casecnt as int) end ,0) , 小包裝 = isnull(rtrim(sp.busr1),'個') " & _
            ", 備註 = case when len(cast(o.notes as varchar(1000))) > 0 or len(cast(od.notes as varchar(1000))) > 0 then cast(o.notes as varchar(1000)) + '_' + cast(od.notes as varchar(1000)) else ' ' end  , 總個數= o3t.order_qty " & _
            ", 排車者 = Case When Isnull(o1t.C_Route_No,'') = '' Then Isnull(Rtrim(o1t.AddWho),'') else Rtrim(o1t.AddWho) End " & _
            "from ort01t o1t join ort02t o2t on o1t.route_no = o2t.route_no " & _
            "join ort03t o3t on o3t.receipt_no = o2t.receipt_no " & _
            "join orders o on o.orderkey = o2t.receipt_no " & _
            "left join trp01m t1m on o2t.consigneekey = t1m.consigneekey and t1m.storerkey = o2t.storerkey " & _
            "left join trp01m t1ma on o.b_company = t1ma.consigneekey and t1ma.storerkey = o.storerkey  " & _
            "left join trp09m t9m on t9m.vehicle_id_no = o2t.vehicle_id_no " & _
            "join orderdetail od on od.orderkey = o.orderkey and od.orderlinenumber = o3t.seq_no  " & _
            "join gv_skuxpack sp on sp.sku = od.sku and sp.storerkey = o2t.storerkey " & _
            "where left(o2t.route_no,1) = 'R' and o2t.receipt_no ='" & txt_OneOrder_OrderKey & "' "
    
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If Not tmp_Rs.EOF Then '無資料時無須列印

        cnAccess.Execute "Delete From 退貨簽收單", RowsAffect, adExecuteNoRecords
        rs_Access.Open "退貨簽收單", cnAccess, adOpenStatic, adLockOptimistic
        tmp_Rs.MoveFirst
        Do While Not tmp_Rs.EOF
        
            rs_Access.AddNew
            For i = 0 To tmp_Rs.Fields.Count - 1
             rs_Access.Fields(i).Value = RTrim(tmp_Rs.Fields(i).Value)
            Next i
            rs_Access.Update
        
        tmp_Rs.MoveNext
        
        Loop
        
        cnAccess.CommitTrans: Tran_Level = 0
        Call DB_Disconnect(cnAccess)
        MSAccessAP.DoCmd.OpenReport "退貨簽收單", acViewPreview
        MSAccessAP.DoCmd.Maximize
        MSAccessAP.Visible = True
    End If
End If

tmp_Rs.Close

'更新列印次數
str_SQL = "Update Ort01T Set VLListCount = VLListCount + 1 ,VLListPrintDate = getdate() " & _
          "Where Route_No = '" & txt_OneOrder_RouteNo & "' or C_Route_No = '" & txt_C_Route_NO & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "Update TRP01T Set VLListCount = VLListCount + 1,VLListPrintDate = getdate() " & _
          "Where Route_No = '" & txt_OneOrder_RouteNo & "' or C_Route_No = '" & txt_C_Route_NO & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

Screen.MousePointer = 0
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit:      Set MSAccessAP = Nothing
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-出貨單-列印", Me.Caption, "cmdShipNotes_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmdTKPremiamAR_Click()
If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strCarno As String, strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "選取 = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("選取") = "V" Then

    If rsRouteT0("車牌號碼") = "000-31" Or rsRouteT0("車牌號碼") = "001-36" Or rsRouteT0("車牌號碼") = "000-70" Or rsRouteT0("車牌號碼") = "000-67" Or rsRouteT0("車牌號碼") = "001-23" Then MsgBox "TK議價應收分攤計算，無法選取車牌號碼(" & rsRouteT0("車牌號碼") & ")!", 16, "注意": GoTo EndProc
    If Len(Trim(strCarno)) > 0 And strCarno <> rsRouteT0("車牌號碼") Then MsgBox "車牌號碼不同!", 16, "注意": GoTo EndProc
    strCarno = rsRouteT0("車牌號碼")
    strRoute = strRoute & rsRouteT0("二次路編") & "','"

End If

rsRouteT0.MoveNext
Loop

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and 二次路編 in ('" & strRoute & "') and 貨主 = 'LTKK01' ", cn
If tmp_Rs.EOF Then MsgBox "查無資料或尚未維護運費!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'原始資料轉Excel
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String, dbAR As Double, dbPremiam

rsTmp.Filter = "計費時間 = '1900-01-01 00:00:00.000'"
If Not rsTmp.EOF Then
    Recordset2Excel "TK議價應收分攤", rsTmp
    MyXlsApp.Visible = True
    Set MyXlsApp = Nothing
    rsTmp.Filter = ""
    rsTmp.Close
    MsgBox "部份訂單未維護運費，TK議價應收分攤計算終止!", 16, Me.Caption
    GoTo EndProc
End If

rsTmp.Filter = ""

'取總應收金額
rsTmp.MoveFirst
Do While Not rsTmp.EOF

If Left(rsTmp("備註"), 4) <> "二次配送" And Left(rsTmp("備註"), 3) <> "不分攤" And rsTmp("請款代碼") <> "Cancel" And rsTmp("請款代碼") <> "I" And rsTmp("請款代碼") <> "R" Then dbAR = dbAR + rsTmp("應收總價")

rsTmp.MoveNext
Loop

dbPremiam = InputBox("不列入TK議價應收分攤條件如下：" & vbCr & vbLf & "1.代碼前六碼:" & vbCr & vbLf & "2.計費代碼:Cancel,I,R" & vbCr & vbLf & "3.計費類別:" & vbCr & vbLf & "4.備註開頭:不分攤，二次配送", "請輸入議價金額(輸入0元或按取消可中止計算)", 0, 0)

If Val(dbPremiam) = 0 Then
        Recordset2Excel "TK議價應收分攤", rsTmp
        MyXlsApp.Visible = True
        Set MyXlsApp = Nothing
        GoTo EndProc
End If

Tran_Level = cn.BeginTrans

'計算議價
str_SQL = "Update sdn05t " & _
          "Set sumreceivable = sdn05t.sumreceivable / " & dbAR & " * " & dbPremiam & _
          ",receivable = sdn05t.sumreceivable / " & dbAR & " * " & dbPremiam & " / sdn05t.chargeqty " & _
          ",note = '專車價(" & dbPremiam & ")' + '_' + sdn05t.note " & _
          "from sdn05t join sdn02t s2 on s2.receipt_no = sdn05t.sdn_no and s2.storerkey = 'LTKK01' " & _
          "where sdn05t.c_route_no in ('" & strRoute & "')  " & _
          "and sdn05t.sumreceivable > 0 " & _
          "and sdn05t.costcode not in ('I','R','Cancel') " & _
          "and left(sdn05t.Note,3) <> '不分攤' " & _
          "and left(sdn05t.Note,4) <> '二次配送' "
          
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and 二次路編 in ('" & strRoute & "') and 貨主 = 'LTKK01' ", cn
If tmp_Rs.EOF Then MsgBox "查無資料!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

 Recordset2Excel "TK議價應收分攤", rsTmp
'在此編輯EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

.Visible = True: End With

Set MyXlsApp = Nothing

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdUnRouteConfirm_Click()
    On Error GoTo err_Handle
    
    If Left(txt_C_Route_NO.Text, 1) = "N" Then MsgBox "打散重組路編(N開頭路編)，無法取消出車確認，請由出車確認模組中執行!(已確認路編==>路編刪除)", 64, Trim(txt_C_Route_NO) & "==>未出車確認": Exit Sub
    If Len(RTrim(txt_OneOrder_VehicleID.Text)) = 0 Or Len(RTrim(txt_C_Route_NO.Text)) = 0 Then Exit Sub
    If MsgBox("此路編將回復未出車確認狀態，該路編所有訂單運費與簽單確認將一併刪除，是否繼續?", vbOKCancel, Trim(txt_C_Route_NO) & "==>未出車確認") <> vbOK Then Exit Sub
    
    '確保路編仍存在(出車確認模組中，已確認簽單裡沒被刪除路編)
    Call cmd_OrderQuery_Click
    
    Tran_Level = cn.BeginTrans
    cn.Execute "exec gs_UnRouteConfirm '" & Trim(txt_C_Route_NO) & "' ", RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans: Tran_Level = 0
    
    Call cmd_OrderQuery_Click
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dg_SDN_Head_Scroll()
    Text1.Visible = False
End Sub

Private Sub cmdPremiamAP_Click()

If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strCarno As String, strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "選取 = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("選取") = "V" Then

    If rsRouteT0("車牌號碼") = "000-31" Or rsRouteT0("車牌號碼") = "001-36" Or rsRouteT0("車牌號碼") = "000-70" Or rsRouteT0("車牌號碼") = "000-67" Or rsRouteT0("車牌號碼") = "001-23" Then MsgBox "議價應付計算，無法選取車牌號碼(" & rsRouteT0("車牌號碼") & ")!", 16, "注意": GoTo EndProc
    If Len(Trim(strCarno)) > 0 And strCarno <> rsRouteT0("車牌號碼") Then MsgBox "車牌號碼不同!", 16, "注意": GoTo EndProc
    strCarno = rsRouteT0("車牌號碼")
    strRoute = strRoute & rsRouteT0("二次路編") & "','"

End If

rsRouteT0.MoveNext
Loop

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and 二次路編 in ('" & strRoute & "')", cn
If tmp_Rs.EOF Then MsgBox "查無資料!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'原始資料轉Excel
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String, dbAP As Double, dbPremiam

rsTmp.Filter = "計費時間 = '1900-01-01 00:00:00.000'"
If Not rsTmp.EOF Then
    Recordset2Excel "議價應付分攤", rsTmp
    MyXlsApp.Visible = True
    Set MyXlsApp = Nothing
    rsTmp.Filter = ""
    rsTmp.Close
    MsgBox "部份訂單未維護運費，議價應付分攤計算終止!", 16, Me.Caption
    GoTo EndProc
End If

rsTmp.Filter = ""

'取總應付金額
rsTmp.MoveFirst
Do While Not rsTmp.EOF

If rsTmp("請款代碼") <> "forklift" And rsTmp("請款代碼") <> "Cancel" And rsTmp("請款代碼") <> "repalletis" And Left(rsTmp("請款代碼"), 6) <> "000-31" And Left(rsTmp("請款代碼"), 6) <> "001-36" And Left(rsTmp("請款代碼"), 6) <> "000-70" And Left(rsTmp("請款代碼"), 6) <> "000-67" And Left(rsTmp("請款代碼"), 6) <> "001-23" And Left(rsTmp("請款代碼"), 6) <> "Cancel" And rsTmp("請款類別") <> "轉運" And Left(rsTmp("備註"), 3) <> "不分攤" Then dbAP = dbAP + rsTmp("應付總價")

rsTmp.MoveNext
Loop

If dbAP = 0 Then MsgBox "應付總金額為0，無法進行分攤作業，應付分攤終止！", 16, Me.Caption: GoTo EndProc

dbPremiam = InputBox("不列入A段分攤條件如下：" & vbCr & vbLf & "1.代碼前六碼:000-31,001-36,000-70,000-67,001-23,Cancel" & vbCr & vbLf & "2.計費代碼:forklift,RePalletIs" & vbCr & vbLf & "3.計費類別:轉運" & vbCr & vbLf & "4.備註開頭:不分攤", "請輸入議價金額(輸入0元或按取消可中止計算)", 0, 0)

If Val(dbPremiam) = 0 Then
        Recordset2Excel "議價應付分攤", rsTmp
        MyXlsApp.Visible = True
        Set MyXlsApp = Nothing
        GoTo EndProc
End If

Tran_Level = cn.BeginTrans

'議價歸零
cn.Execute "Update sdn05t set Premiam = 0 where c_route_no in ('" & strRoute & "') ", RowsAffect, adExecuteNoRecords

'計算議價
str_SQL = "Update sdn05t " & _
          "Set Premiam = sumpayable / " & dbAP & " * " & dbPremiam & _
          ",note = note + '_" & strCarno & " 專車價(" & dbPremiam & ")' " & _
          "where c_route_no in ('" & strRoute & "') " & _
          "and left(costcode,6) not in ('000-31','001-36','000-70','000-67','001-23','Cancel') " & _
          "and costcode not in ('forklift','repalletis') " & _
          "and costkind <> ('轉運') " & _
          "and left(note,3) <> '不分攤'"
          
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and 二次路編 in ('" & strRoute & "')", cn
If tmp_Rs.EOF Then MsgBox "查無資料!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

 Recordset2Excel "議價應付分攤", rsTmp
'在此編輯EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

.Visible = True: End With

Set MyXlsApp = Nothing

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub dg_Tab2_SDN_Cost_Click()
    If dg_Tab2_SDN_Cost.Col < 13 Then
        NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
    End If
End Sub

Private Sub dg_Tab2_SDN_Detail_Click()
    If dg_Tab2_SDN_Detail.Col = 0 Or dg_Tab2_SDN_Detail.Col = 1 Then
'        If Len(dg_SDN_Detail.Text) = 0 Then
'            dg_Tab2_SDN_Detail.Text = "Ｖ"
'        Else
'            dg_Tab2_SDN_Detail.Text = ""
'        End If
    End If
    If dg_Tab2_SDN_Detail.Col = 9 Then
        If Len(dg_Tab2_SDN_Detail.Text) = 0 Then
            dg_Tab2_SDN_Detail.Text = "Ｖ"
            dg_Tab2_SDN_Detail.Col = 4
            txt_Tab2_sum_Case.Text = Val(txt_Tab2_sum_Case.Text) + Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 5
            txt_Tab2_sum_CBM.Text = Val(txt_Tab2_sum_CBM.Text) + Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 6
            txt_Tab2_sum_WT.Text = Val(txt_Tab2_sum_WT.Text) + Val(dg_Tab2_SDN_Detail.Text)
        Else
            dg_Tab2_SDN_Detail.Text = ""
            dg_Tab2_SDN_Detail.Col = 4
            txt_Tab2_sum_Case.Text = Val(txt_Tab2_sum_Case.Text) - Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 5
            txt_Tab2_sum_CBM.Text = Val(txt_Tab2_sum_CBM.Text) - Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 6
            txt_Tab2_sum_WT.Text = Val(txt_Tab2_sum_WT.Text) - Val(dg_Tab2_SDN_Detail.Text)
        End If
        dg_Tab2_SDN_Detail.Col = 9
    End If
    If dg_Tab2_SDN_Detail.Col > 1 And dg_Tab2_SDN_Detail.Col < 9 Then
        NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
    End If
End Sub

Private Sub dg_Tab2_SDN_Detail_Scroll()
    Text3.Visible = False
End Sub

Private Sub dgMain1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain1

'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain1_HeadClick(ByVal ColIndex As Integer)
If dgMain1.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMain1.Sort = dgMain1.Columns(ColIndex).Caption & " DESC"
    dgMain1.ClearSelCols
    intColumnIndex = 255

Else
    rsMain1.Sort = dgMain1.Columns(ColIndex).Caption
    dgMain1.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgOrderT0_HeadClick(ByVal ColIndex As Integer)
Dim dg As Object, rs As Object
Set dg = dgOrderT0: Set rs = rsOrderT0

If dg.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rs.Sort = dg.Columns(ColIndex).Caption & " DESC"
    dg.ClearSelCols
    intColumnIndex = 255

Else
    rs.Sort = dg.Columns(ColIndex).Caption
    dg.ClearSelCols
    intColumnIndex = ColIndex

End If
End Sub

Private Sub dgRouteT0_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgRouteT0
If dg Is Nothing Then Exit Sub

'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgRouteT0_HeadClick(ByVal ColIndex As Integer)

Dim dg As Object, rs As Object
Set dg = dgRouteT0: Set rs = rsRouteT0

If dg.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rs.Sort = dg.Columns(ColIndex).Caption & " DESC"
    dg.ClearSelCols
    intColumnIndex = 255

Else
    rs.Sort = dg.Columns(ColIndex).Caption
    dg.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgRouteT0_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'是否有資料
If rsRouteT0 Is Nothing Then Exit Sub
If rsRouteT0.RecordCount = 0 Then Exit Sub

If blRouteT0Change = False Then Exit Sub

'選取
If dgRouteT0.Col = 1 Then

    rsRouteT0("選取") = IIf(rsRouteT0("選取") = "V", " ", "V")
    dgRouteT0.Col = 0

End If

'同一行選取
If LastRow = Empty Then Exit Sub

Screen.MousePointer = 11

Frame13.Caption = rsRouteT0("二次路編")

'str_SQL = "select 路線編號 " & _
'            ",到貨日期 " & _
'            ",訂單編號 = 貨主單號 " & _
'            ",狀態 = case when len(rtrim(狀態)) = 0 then (select case when sum(order_qty - ship_qty) <> 0 then '出貨不符' else ' ' end from sdn03t where receipt_no = 訂單編號) else 狀態 end " & _
'            ",驗收單號 = 客戶驗收單號 " & _
'            ",TMS單號 = 訂單編號 " & _
'            ",車牌號碼 " & _
'            ",駕駛人 " & _
'            ",一次車號 " & _
'            ",貨主名稱 " & _
'            ",訂單類別 " & _
'            ",客戶名稱 " & _
'            ",訂單備註 = 說明 " & _
'            ",送貨地址 " & _
'            ",簽單回傳日期 " & _
'            ",貨主 " & _
'            "From SDNConfirm_OrderDate_One where 1 = 1 and 二次路編 = '" & rsRouteT0("二次路編") & "' "
'
'If Len(RTrim(cboStorerT0)) > 0 Then str_SQL = str_SQL & "and 貨主 = '" & RTrim(cboStorerT0) & "' "
'If Len(RTrim(cboCarT0)) > 0 Then str_SQL = str_SQL & "and 車牌號碼 = '" & RTrim(cboCarT0) & "' order by 貨主,貨主單號"


str_SQL = "select 路線編號 = t02t.Route_No,到貨日期 = rtrim(t02t.Arrive_Date) ,訂單編號 = Rtrim(t02t.Extern) " & _
    ",狀態 = case when len(rtrim(Isnull(Rtrim(t02t.Confirm_Notes),''))) = 0 then (select case when sum(order_qty - ship_qty) <> 0 then '出貨不符' else ' ' end from sdn03t where receipt_no = t02t.Receipt_No) else Isnull(Rtrim(t02t.Confirm_Notes),'') end " & _
    ",驗收單號=Isnull(rtrim(t02t.CustomerOrderkey1),''),TMS單號 = rtrim(t02t.Receipt_No),車牌號碼 = Rtrim(t01t.c_Vehicle_ID_No) " & _
    ",駕駛人 = Rtrim(t01t.driver),一次車號 = Rtrim(t02t.Vehicle_ID_No),貨主名稱 =rtrim(t16.c_name),訂單類別 = rtrim(t02t.priority) " & _
    ",客戶名稱 = Rtrim(Isnull(t1m.Short_Name,'')),訂單備註 = Rtrim(Isnull(t02t.Description,'')),送貨地址 = Rtrim(Isnull(t1m.Address,'')) " & _
    ",簽單回傳日期=isnull(t02t.SDNSendDate,getdate()),貨主 = Rtrim(t02t.StorerKey) " & _
    "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
    "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
    "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
    "left join trp09m t9m (nolock) on t9m.vehicle_id_no = t01t.c_vehicle_id_no " & _
    "left join trp08m t8m (nolock) on t8m.company_code = t9m.trp_company_code " & _
    "where 1 = 1 and t02t.c_Route_No = '" & rsRouteT0("二次路編") & "' "
    
    If Len(RTrim(cboStorerT0)) > 0 Then str_SQL = str_SQL & "and t02t.StorerKey = '" & RTrim(cboStorerT0) & "' "
    If Len(RTrim(cboCarT0)) > 0 Then str_SQL = str_SQL & "and t01t.c_Vehicle_ID_No = '" & RTrim(cboCarT0) & "' "
    str_SQL = str_SQL & "order by t02t.StorerKey,t02t.Extern "

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "查詢結果：無符合搜尋條件之訂單資料"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Set dgOrderT0.DataSource = Nothing: Set rsOrderT0 = Nothing
    Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rsOrderT0)
tmp_Rs.Close: rsOrderT0.MoveFirst

Set dgOrderT0.DataSource = rsOrderT0

SetDataGridColWidth Me.Caption, dgOrderT0

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgOrderT0_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgOrderT0
If dg Is Nothing Then Exit Sub

'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgOrderT0_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

txtCustomerOrderkey.Visible = False

''同一行選取
'If LastRow = Empty Then Exit Sub

'是否有資料
If rsOrderT0 Is Nothing Then Exit Sub
If rsOrderT0.RecordCount = 0 Then Exit Sub
If rsOrderT0.EOF Then Exit Sub
If Len(RTrim(rsOrderT0("狀態"))) > 1 Then Exit Sub

Screen.MousePointer = 11

If dgOrderT0.Col = 4 And (rsOrderT0("貨主") = "LTKK01" Or rsOrderT0("貨主") = "LVTL01" Or rsOrderT0("貨主") = "LNSL01") Then
    If rsOrderT0("狀態") = "V" Then
        rsOrderT0("狀態") = " "
    Else
        rsOrderT0("狀態") = "V"
    End If
dgOrderT0.Col = 0
End If

'驗收單號
If dgOrderT0.Col = 5 Then Call CustomerOrderkey

Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub CustomerOrderkey()

With dgOrderT0
    txtCustomerOrderkey.Height = .RowHeight + 10
    If .Col = 5 Then
        If .Columns(.Col).Left > 0 Then
                
                txtCustomerOrderkey.Move .Left + .Columns(.Col).Left + 15, .Top + .RowTop(.Row), .Columns(.Col).Width
                If txtCustomerOrderkey.Left + txtCustomerOrderkey.Width > .Left + .Width Then '如果欄位超出DataGrid的顯示範圍的處理
                    txtCustomerOrderkey.Width = txtCustomerOrderkey.Width + .Left + .Width - txtCustomerOrderkey.Left - txtCustomerOrderkey.Width
                End If
                txtCustomerOrderkey.Text = rsOrderT0("驗收單號")  '更新儲存格的值

                txtCustomerOrderkey.Visible = True
        Else '如果用捲軸捲動出了DataGrid的顯示範圍，值會小於0
            txtCustomerOrderkey.Visible = False
        End If
    Else
        txtCustomerOrderkey.Visible = False
    End If
End With

End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "送貨簽單確認"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 10000
dbsrcFormWidth = 15000

Me.Height = 9700: Me.Width = 15000
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
'Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2
'ㄧ張貨主單號對應ㄧ張排車系統訂單
Call SetGridFormat_OneOrder_OrderDetail

'ㄧ張貨主單號對應多張排車系統訂單
Call SetGridFormat_MultiOrder_OrderDetail

'取得所有 [狀況代碼] From LogicTown.dbo.TRP05M
str_SQL = "SELECT Rtrim(RSC_CODE)+' ' as RSC_CODE,RTRIM(isnull(DESCRIPTION,'')) AS 'Descr' FROM TRP05M Order by RSC_CODE"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then Exit Sub
iLoop = 0
cmb_OneOrder_RSCCode.Clear: cmb_MultiOrder_RSCCode.Clear
cmb_OneOrder_RSCCode.AddItem ""
cmb_MultiOrder_RSCCode.AddItem ""
Do While Not tmp_Rs.EOF
   cmb_OneOrder_RSCCode.AddItem tmp_Rs.Fields("RSC_CODE") & "  " & tmp_Rs.Fields("descr")
   cmb_MultiOrder_RSCCode.AddItem tmp_Rs.Fields("RSC_CODE") & "  " & tmp_Rs.Fields("descr")
   tmp_Rs.MoveNext
   iLoop = iLoop + 1
Loop
tmp_Rs.Close
Call ComboBox_SetWidth(cmb_OneOrder_RSCCode, 30)

'取得所有 [責任歸屬代碼] From LogicTown.dbo.TRP06M
str_SQL = "SELECT Rtrim(RBC_CODE)+' ' as 'RBC_CODE',RTRIM(isnull(Description,'')) AS 'Descr' FROM dbo.TRP06M Order by RBC_CODE"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then Exit Sub
iLoop = 0
cmb_OneOrder_RBCCode.Clear: cmb_OneOrder_RBCCode.AddItem ""
cmb_MultiOrder_RBCCode.Clear: cmb_MultiOrder_RBCCode.AddItem ""
Do While Not tmp_Rs.EOF
   cmb_OneOrder_RBCCode.AddItem tmp_Rs.Fields("RBC_CODE") & "  " & tmp_Rs.Fields("descr")
   cmb_MultiOrder_RBCCode.AddItem tmp_Rs.Fields("RBC_CODE") & "  " & tmp_Rs.Fields("descr")
   tmp_Rs.MoveNext
   iLoop = iLoop + 1
Loop

'設定dg_grid之格式
'Call SetGridFormat_SDN_Head
'Call SetGridFormat_SDN_Detail
'Call SetGridFormat_SDN_Cost
Call SetGridFormat_Tab2_SDN_Detail
Call SetGridFormat_Tab2_SDN_Cost
SSTab1.Tab = 3
Op_UnCheck.Visible = False
Op_OnCheck.Value = True
txt_DeliveryDate_Start.Text = Format(Now, "yyyymmdd")

cmbScan.AddItem "Y"
cmbScan.AddItem "N"
cmbScan.ListIndex = 0

cboInvBack.AddItem "Y"
cboInvBack.AddItem "N"

cmbOrderkey.AddItem ""
cmbOrderkey.AddItem "TMS單號"
cmbOrderkey.AddItem "貨主單號"
cmbOrderkey.ListIndex = 0

tmp_Rs.Close

'取車號
str_SQL = "select distinct vehicle_id_no= rtrim(vehicle_id_no) from trp09m order by vehicle_id_no "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then tmp_Rs.MoveFirst

Do While Not tmp_Rs.EOF
    cboCar.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    cboCarT0.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    tmp_Rs.MoveNext
Loop
cboCarT0 = ""
tmp_Rs.Close

'貨主
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(rtrim(storerkey)) as storerkey from trp16M", cn, adOpenKeyset, adLockPessimistic

If Not tmp_Rs.EOF Then
    
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboStorerkey.AddItem tmp_Rs("storerkey")
        cboStorerT0.AddItem tmp_Rs("storerkey")
        tmp_Rs.MoveNext
    Next
    tmp_Rs.Close: Set tmp_Rs = Nothing
    cboStorerkey = ""
    cboStorerT0 = ""
End If

'取請款類別
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select distinct costkind from trp17m order by costkind "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then

    If Not tmp_Rs.EOF Then tmp_Rs.MoveFirst
    
    Do While Not tmp_Rs.EOF
        cboCostkind.AddItem RTrim(tmp_Rs("costkind"))
        tmp_Rs.MoveNext
    Loop

End If

tmp_Rs.Close

txtDeliveryS = Format(Now - 30, "YYYYMMDD")
txtDeliveryE = Format(Now + 7, "YYYYMMDD")
txtDeliveryDateST0 = Format(Now - 1, "YYYYMMDD")
txtDeliveryDateET0 = Format(Now + 2, "YYYYMMDD")

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'攔截整個表單鍵盤按鍵事件
'用途：使用者按下 Esc 則不傳回任何資料
'且關閉日期選取視窗
mvDate.Visible = False
If KeyCode = vbKeyEscape Then
   
   txt_OneOrder_SignQty.Visible = False
   cmb_OneOrder_RBCCode.Visible = False
   cmb_OneOrder_RSCCode.Visible = False
   
   txt_MultiOrder_SignQty.Visible = False
   cmb_MultiOrder_RBCCode.Visible = False
   cmb_MultiOrder_RSCCode.Visible = False
   
End If

End Sub

Private Sub Form_Resize()
On Error GoTo err_Handle

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化
If SSTab1.Height < 7000 And Me.ScaleHeight < 7000 Then Exit Sub

SSTab1.Height = Me.ScaleHeight: SSTab1.Width = Me.ScaleWidth

Frame10.Width = SSTab1.Width - 240: dgMain1.Width = Frame10.Width - 240
Frame12.Width = SSTab1.Width - 240: dgRouteT0.Width = Frame12.Width - 1440
Frame13.Width = SSTab1.Width - 240: dgOrderT0.Width = Frame13.Width - 240
fra_OneOrder_Detail.Width = SSTab1.Width - 240: gd_OneOrder_OrderDetail.Width = fra_OneOrder_Detail.Width - 240

Frame10.Height = SSTab1.Height - Frame10.Top - 120: dgMain1.Height = Frame10.Height - 360
dgRouteT0.Height = Frame12.Height - 360
Frame13.Height = SSTab1.Height - Frame14.Top - Frame14.Height - Frame12.Height - 120
If Frame13.Height > 840 Then dgOrderT0.Height = Frame13.Height - 840
fra_OneOrder_Detail.Height = SSTab1.Height - fra_OneOrder_Detail.Top - 120: gd_OneOrder_OrderDetail.Height = fra_OneOrder_Detail.Height - 360

Exit Sub
err_Handle:
'Call ErrorMsgbox(Me.Caption & "_Form_Resize", err.Number, err.Description, "")
End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_OP_SDNAbnormal = Nothing
Set rsMain1 = Nothing
Set rsRouteT0 = Nothing
Set rsOrderT0 = Nothing
End Sub

Private Sub ClearForm()
'清除所有欄位值
Call ClearForm_AllField(Me)
blSDNConfirm = False
blCanUpdate = False
cmdCost.Enabled = False

'ㄧ張貨主單號對應ㄧ張排車系統訂單
cmd_OneOrder_Deliveryok.Enabled = False
cmd_OneOrder_Expect.Enabled = False
cmd_OneOrder_NoDelivery.Enabled = False
Call SetGridFormat_OneOrder_OrderDetail

'ㄧ張貨主單號對應多張排車系統訂單
cmd_MultiOrder_Deliveryok.Enabled = False
cmd_MultiOrder_Expect.Enabled = False
cmd_MultiOrder_NoDelivery.Enabled = False
Call SetGridFormat_MultiOrder_OrderDetail
Set dg_MultiOrder.DataSource = Nothing
Set rs_MultiOrder = Nothing

End Sub

Private Sub Display_OrderData_OneReceipNo(ByVal strExtern As String)
'ㄧ張貨主單號對應ㄧ張排車系統訂單：訂單資料查詢
Screen.MousePointer = vbHourglass
fra_OneOrder_Header.Visible = True
fra_OneOrder_Detail.Visible = True
fra_MultiOrder_Header.Visible = False
fra_MultiOrder_Detail.Visible = False
cmdCost.Enabled = True
txt_OneOrder_Status.BackColor = "&H80000000"

On Error GoTo err_Handle
'str_SQL = "Select * From SDNConfirm_OrderDate_One Where 訂單編號 = '" & strExtern & "'"

str_SQL = "select 二次路編 = t02t.c_Route_No " & _
            ",路線編號 = t02t.Route_No,車牌號碼 = Rtrim(t01t.c_Vehicle_ID_No),一次車號 = Rtrim(t02t.Vehicle_ID_No) " & _
            ",駕駛人 = Rtrim(t01t.driver),貨運公司 = Isnull(Rtrim(t8m.Short_Name),'') " & _
            ",出車日期 = convert(varchar,t01t.Delivery_Date,112),貨主 = Rtrim(t02t.StorerKey) " & _
            ",貨主名稱 =rtrim(t16.c_name),說明 = Rtrim(Isnull(t02t.Description,'')),客戶編號 = Rtrim(t02t.ConsigneeKey) " & _
            ",客戶名稱 = Rtrim(Isnull(t1m.Short_Name,'')),郵遞區號 = Rtrim(Isnull(t1m.zip,'')),送貨地址 = Rtrim(Isnull(t1m.Address,'')) " & _
            ",訂單類別 = rtrim(t02t.priority),訂單編號 = rtrim(t02t.Receipt_No),訂單日期 = rtrim(t02t.Receipt_Date) " & _
            ",到貨日期 = rtrim(t02t.Arrive_Date),貨主單號 = Rtrim(t02t.Extern) " & _
            ",簽單日期 = isnull(t02t.CustSignDate,isnull(t02t.SCHEDULEDATE,Arrive_Date)),系統日期 = Convert(varchar,Getdate(),112) " & _
            ",狀態 = Isnull(Rtrim(t02t.Confirm_Notes),''),客戶驗收單號=Isnull(rtrim(t02t.CustomerOrderkey1),'') " & _
            ",掃描=Isnull(rtrim(t02t.Scan),''),簽單回傳日期=isnull(t02t.SDNSendDate,getdate()),客戶回覆處理方式=Isnull(rtrim(t02t.CUST_Handle),'') " & _
            ",後續處理方式=Isnull(rtrim(t02t.TRP_Handle),''),改善方式=Isnull(rtrim(t02t.Advance),''),庫存調整方式=Isnull(rtrim(t02t.INV_Handle),'') " & _
            ",配送費=Isnull(rtrim(t02t.TRP_Cost),0),理貨費=Isnull(rtrim(t02t.Sorting_Cost),0),異常費用合計=Isnull(rtrim(t02t.Total_Cost),'') " & _
            ",簽單備註=Isnull(rtrim(t02t.sdn_note),''),入庫完成=Isnull(rtrim(ExpectReceiptOK),''),發票回收=invBack " & _
            ",到貨 = isnull(ontimedelivery,0),簽單已回 = t02t.sdnback " & _
            "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
            "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
            "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
            "left join trp09m t9m (nolock) on t9m.vehicle_id_no = t01t.c_vehicle_id_no " & _
            "left join trp08m t8m (nolock) on t8m.company_code = t9m.trp_company_code " & _
            "where t02t.Receipt_No = '" & strExtern & "' "

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

txt_C_Route_NO.Text = tmp_Rs.Fields("二次路編").Value
txt_OneOrder_RouteNo.Text = tmp_Rs.Fields("路線編號").Value
txt_OneOrder_VehicleID.Text = tmp_Rs.Fields("車牌號碼").Value & "_" & tmp_Rs.Fields("一次車號").Value
txt_OneOrder_Driver.Text = tmp_Rs.Fields("駕駛人").Value & ""
txt_OneOrder_TRPCompany.Text = tmp_Rs.Fields("貨運公司").Value & ""
txt_OneOrder_DeliveryDate.Text = tmp_Rs.Fields("出車日期").Value & ""
txt_OneOrder_StorerKey.Text = tmp_Rs.Fields("貨主").Value & ""
txt_Storer.Text = tmp_Rs.Fields("貨主名稱").Value & ""
txt_OneOrder_Description.Text = tmp_Rs.Fields("說明").Value
txt_OneOrder_ConsigneeKey.Text = tmp_Rs.Fields("客戶編號").Value
txt_OneOrder_FullName.Text = tmp_Rs.Fields("客戶名稱").Value
txt_OneOrder_Address.Text = tmp_Rs.Fields("送貨地址").Value
txt_OneOrder_OrderKey.Text = tmp_Rs.Fields("訂單編號").Value
txt_OneOrder_OrderDate.Text = tmp_Rs.Fields("訂單日期").Value & ""
txt_OneOrder_ArriveDate.Text = tmp_Rs.Fields("到貨日期").Value
txt_OneOrder_Status.Text = tmp_Rs.Fields("狀態").Value
txt_OneOrder_StorerOrderKey.Text = tmp_Rs.Fields("貨主單號").Value
txt_OneOrder_CustomerOrderkey1.Text = tmp_Rs.Fields("客戶驗收單號").Value & ""
cmbScan.Text = tmp_Rs.Fields("掃描").Value: If cmbScan.Text <> "Y" Then cmbScan.Text = "N"
dtpSDNSendDate.Value = tmp_Rs.Fields("簽單回傳日期").Value
If tmp_Rs("到貨") = "5" Then txt_OneOrder_Status.BackColor = "&H00C0C0FF"
If tmp_Rs("到貨") = "9" Then txt_OneOrder_Status.BackColor = "&H00C0FFC0"
dtp_OneOrder_SignDate.Value = tmp_Rs.Fields("簽單日期").Value
txt_CustHandle.Text = tmp_Rs.Fields("客戶回覆處理方式").Value
txt_TRPHandle.Text = tmp_Rs.Fields("後續處理方式").Value
txt_Advance.Text = tmp_Rs.Fields("改善方式").Value
txt_INVHandle.Text = tmp_Rs.Fields("庫存調整方式").Value
txt_TRPCost.Text = tmp_Rs.Fields("配送費").Value
txt_SortingCost.Text = tmp_Rs.Fields("理貨費").Value
txt_TotalCost.Text = tmp_Rs.Fields("異常費用合計").Value
txt_SDNNote.Text = tmp_Rs("簽單備註")
txt_Priority.Text = tmp_Rs.Fields("訂單類別").Value
txt_ZIP.Text = tmp_Rs.Fields("郵遞區號").Value
cboInvBack.Text = tmp_Rs.Fields("發票回收").Value

If tmp_Rs.Fields("簽單已回").Value = "1" Then
    cmdSDNBack.BackColor = vbGreen
    cmdSDNBack.Caption = "簽單已回"
Else
    cmdSDNBack.BackColor = vbRed
    cmdSDNBack.Caption = "簽單未回"
End If

blShipped = True    '無法判斷訂單揀貨量是否已更新(Ship_Qty)
blCanUpdate = True
blSDNConfirm = False
blCanUpdate = True

tmp_Rs.Close

'取提貨地址
If txt_Priority = "A2B" Then

Call ReDim_Recordset(tmp_Rs)
str_SQL = "select 客戶編號=rtrim(t1.consigneekey) , 客戶名稱=rtrim(t1.full_name) , 郵遞區號=rtrim(t1.zip) , 送貨地址=rtrim(t1.address) from orders o join trp01m t1 on t1.storerkey = o.storerkey and t1.consigneekey = o.b_company where o.orderkey = '" & txt_OneOrder_OrderKey & "' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    If Not tmp_Rs.EOF Then
        txt_OneOrder_ConsigneeKey1 = tmp_Rs.Fields("客戶編號").Value
        txt_OneOrder_FullName1 = tmp_Rs.Fields("客戶名稱").Value
        txt_Zip1 = tmp_Rs.Fields("郵遞區號").Value
        txt_OneOrder_Address1 = tmp_Rs.Fields("送貨地址").Value
    End If
    tmp_Rs.Close
End If

'簽單維護期限
If Val(txt_OneOrder_ArriveDate) > lngDueDate Then
    txt_OneOrder_Status.Enabled = True
    cmdCarNOChange.Enabled = True: cmdUnRouteConfirm.Enabled = True: cmdShipNotes.Enabled = True
Else
    txt_OneOrder_Status.Enabled = False
    cmdCarNOChange.Enabled = False: cmdUnRouteConfirm.Enabled = False: cmdShipNotes.Enabled = False
End If

'更新出貨量
Call Ship2TMS(strExtern)

'ㄧ張貨主單號對應ㄧ張排車系統訂單：訂單名細
Call SetGridFormat_OneOrder_OrderDetail
Dim tmpI As Double
'str_SQL = "Select 項次,貨號,品名,單位,訂單量,送貨量,簽單量,異常原因,責任歸屬,異常碼,責屬碼,箱包轉換率,單位重量,單位材積,責屬人 " & _
'          "From SDNConfirm_OrderDetail_OneOrder "
          
'        If cmbOrderkey.Text = "TMS單號" Then
'            str_SQL = str_SQL & "Where 訂單號碼 = '" & strExtern & "' Order by 項次"
'        Else
'            str_SQL = str_SQL & "Where 貨主單號 = '" & strExtern & "' Order by 項次"
'        End If

'str_SQL = "exec SDN_OrderDetail '" & strExtern & "' "

str_SQL = "Select Rtrim(t02t.Extern) As 貨主單號 " & _
    ",rtrim(t02t.receipt_no) as 訂單號碼 " & _
    ",t03t.Seq_No as 項次 " & _
    ",Rtrim(t03t.Product_No) as 貨號 " & _
    ",Rtrim(Isnull(sku.Descr,'')) as 品名 " & _
    ",單位=isnull(sku.busr1,'EA') " & _
    ",t03t.Order_Qty as 訂單量 " & _
    ",Isnull(t03t.Ship_Qty,0) as 送貨量 " & _
    ",Isnull(t03t.Sign_Qty,0) as 簽單量 " & _
    ",Isnull(Rtrim(t03t.RSC_Code) + '  ' + Rtrim(t05m.Description),'  ') as 異常原因 " & _
    ",Isnull(Rtrim(t03t.RBC_Code) + '  ' + Rtrim(t06m.Description),'  ') as 責任歸屬 " & _
    ",Isnull(Rtrim(t03t.RSC_Code),'  ') as 異常碼 " & _
    ",Isnull(Rtrim(t03t.RBC_Code),'  ') as 責屬碼 " & _
    ",箱包轉換率 = isnull(sku.casecnt,'') " & _
    ",單位重量 = round(isnull(sku.stdgrosswgt,0),9) " & _
    ",單位材積 = round(isnull(sku.stdcube,0),9) " & _
    ",責屬人 = isnull(responsible,'') " & _
    "From SDN02T t02t (nolock) join SDN03T t03t (nolock) on t03t.Receipt_No = t02t.Receipt_No " & _
    "join gv_SKUxpack sku on sku.StorerKey = t03t.StorerKey and sku.SKU = t03t.Product_No " & _
    "Left join TRP05M t05m on t05m.RSC_Code = t03t.RSC_Code " & _
    "Left join TRP06M t06m on t06m.RBC_Code = t03t.RBC_Code " & _
    "where t02t.receipt_no = '" & strExtern & "'  order by t03t.Seq_No "
          
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
tmpI = 1
Do While Not tmp_Rs.EOF
   With gd_OneOrder_OrderDetail
        If .Rows < (tmpI + 1) Then .Rows = .Rows + 1
        .Row = tmpI
        .Col = 0: .Text = RTrim(tmp_Rs.Fields("項次").Value)
        .Col = 1: .Text = tmp_Rs.Fields("貨號").Value
        .Col = 2: .Text = tmp_Rs.Fields("品名").Value
        .Col = 3: .Text = RTrim(tmp_Rs.Fields("單位").Value)
        .Col = 4: .Text = tmp_Rs.Fields("訂單量").Value
        .Col = 5: .Text = tmp_Rs.Fields("送貨量").Value
        'mark by gemini
'         If blCanUpdate Then    '尚未執行 SDN Confirmed
'            .Col = 6: .Text = 0   '簽單量
'            .Col = 7: .Text = ""
'            .Col = 8: .Text = ""
'            .Col = 9: .Text = ""
'            .Col = 10: .Text = ""
'         Else
            .Col = 6: .Text = tmp_Rs.Fields("簽單量").Value
            .Col = 7: .Text = tmp_Rs.Fields("異常原因").Value
            .Col = 8: .Text = tmp_Rs.Fields("責任歸屬").Value
            .Col = 9: .Text = tmp_Rs.Fields("異常碼").Value
            .Col = 10: .Text = tmp_Rs.Fields("責屬碼").Value
            .Col = 11: .Text = tmp_Rs.Fields("箱包轉換率").Value
            .Col = 12: .Text = tmp_Rs.Fields("單位重量").Value
            .Col = 13: .Text = tmp_Rs.Fields("單位材積").Value
            .Col = 14: .Text = tmp_Rs.Fields("責屬人").Value
            
'        End If
   End With
   tmp_Rs.MoveNext
   tmpI = tmpI + 1
Loop
tmp_Rs.Close

If blCanUpdate Then
    cmd_OneOrder_Deliveryok.Enabled = True
    cmd_OneOrder_Expect.Enabled = True
    cmd_OneOrder_NoDelivery.Enabled = True
Else
    cmd_OneOrder_Deliveryok.Enabled = False
    cmd_OneOrder_Expect.Enabled = False
    cmd_OneOrder_NoDelivery.Enabled = False
End If

'特定人員可執行 SDN Confirm 重新存檔
'權限設定儲存於 CodeLKUP ListName = [SDNRECONDURM]
'  Code：User_ID   Short：權限設定 1-可以重複執行
If (Not blCanUpdate) And CheckSDNReConfirm(User_id) Then
    '允許重複執行 SDN Confirm 之使用者，開放 SDN Confirm 日期的修改
    cmd_OneOrder_Deliveryok.Enabled = True
    cmd_OneOrder_Expect.Enabled = True
    cmd_OneOrder_NoDelivery.Enabled = True
    blCanUpdate = True
End If

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-訂單查詢", Me.Caption, "Form 內部 SubProgram Display_OrderData_OneReceiptNo", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub SetGridFormat_OneOrder_OrderDetail()
'名稱：SetGridFormatt_OneOrder_OrderDetail
'類別：副程式
'功能：清除並設定 [SDN Confirm] 表單 [ㄧ張貨主單號對應ㄧ張排車系統訂單] 訂單名細顯示格式
'參數：傳入值：無
Dim sub_var1 As Integer, sub_var2 As Integer
gd_OneOrder_OrderDetail.Visible = False
With gd_OneOrder_OrderDetail
     .Rows = 2: .FixedRows = 1: .Cols = 15
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
     .ColWidth(1) = 1800
     .ColWidth(2) = 3000
     .ColWidth(3) = 500
     .ColWidth(4) = 750
     .ColWidth(5) = 750
     .ColWidth(6) = 750
     .ColWidth(7) = 1500
     .ColWidth(8) = 1000
     .ColWidth(9) = 600
     .ColWidth(10) = 600
     .ColWidth(11) = 500
     .ColWidth(12) = 600
     .ColWidth(13) = 600
     .ColWidth(14) = 1000
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "項次"
     .Col = 1: .Text = "貨號"
     .Col = 2: .Text = "中文品名"
     .Col = 3: .Text = "單位"
     .Col = 4: .Text = "訂單量"
     .Col = 5: .Text = "送貨量"
     .Col = 6: .Text = "簽單量"
     .Col = 7: .Text = "異常原因"
     .Col = 8: .Text = "責任歸屬"
     .Col = 9: .Text = "異常碼"
     .Col = 10: .Text = "責屬碼"
     .Col = 11: .Text = "每箱"
     .Col = 12: .Text = "單位重"
     .Col = 13: .Text = "單位材"
     .Col = 14: .Text = "責屬人"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignLeftCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .ColAlignment(10) = flexAlignCenterCenter
     .ColAlignment(11) = flexAlignRightCenter
     .ColAlignment(12) = flexAlignLeftCenter
     .ColAlignment(13) = flexAlignLeftCenter
     .ColAlignment(14) = flexAlignLeftCenter
     .Rows = 2
     .Row = 0
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .Text = ""
     Next sub_var1
     
End With
gd_OneOrder_OrderDetail.Visible = True
End Sub

Private Sub HideGridUseObject_OneOrder()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：隱藏 [簽單量] [異常原因] [責屬] 控制項
txt_OneOrder_SignQty.Visible = False
cmb_OneOrder_RBCCode.Visible = False
cmb_OneOrder_RSCCode.Visible = False
End Sub

Private Sub gd_MultiOrder_OrderDetail_Click()
'ㄧ張貨主單號對應多張排車系統訂單：訂單名細選取
Dim SelectedCol As Integer, SelectedRow As Integer
If Not blCanUpdate Then Exit Sub
Call HideGridUseObject_MultiOrder
On Error Resume Next
With gd_MultiOrder_OrderDetail
     SelectedCol = .Col: SelectedRow = .Row
     Select Case SelectedCol
       Case 5      '簽單量
           txt_MultiOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_MultiOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_MultiOrder_SignQty.Left = txt_MultiOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_MultiOrder_SignQty.Top = txt_MultiOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_MultiOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_MultiOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_MultiOrder_SignQty.Text = .Text
           txt_MultiOrder_SignQty.Visible = True
           txt_MultiOrder_SignQty.SelStart = 0: txt_MultiOrder_SignQty.SelLength = Len(txt_MultiOrder_SignQty.Text)
           txt_MultiOrder_SignQty.SetFocus
           
       Case 6      '異常原因
           cmb_MultiOrder_RSCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_MultiOrder_RSCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_MultiOrder_RSCCode.Left = cmb_MultiOrder_RSCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_MultiOrder_RSCCode.Top = cmb_MultiOrder_RSCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_MultiOrder_RSCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_MultiOrder_RSCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_MultiOrder_RSCCode.ListCount - 1
                  If Left(cmb_MultiOrder_RSCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_MultiOrder_RSCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_MultiOrder_RSCCode.Visible = True
           cmb_MultiOrder_RSCCode.SetFocus
           SendKeys "%{DOWN}"
           
       Case 7      '權責區分：責屬
           cmb_MultiOrder_RBCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_MultiOrder_RBCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_MultiOrder_RBCCode.Left = cmb_MultiOrder_RBCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_MultiOrder_RBCCode.Top = cmb_MultiOrder_RBCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_MultiOrder_RBCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_MultiOrder_RBCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_MultiOrder_RBCCode.ListCount - 1
                  If Left(cmb_MultiOrder_RBCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_MultiOrder_RBCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_MultiOrder_RBCCode.Visible = True
           cmb_MultiOrder_RBCCode.SetFocus
           SendKeys "%{DOWN}"
           
     End Select
End With

End Sub

Private Sub gd_OneOrder_OrderDetail_Click()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：訂單名細選取
Dim SelectedCol As Integer, SelectedRow As Integer
If Not blCanUpdate Then Exit Sub
cmb_OneOrder_RSCCode.Visible = False: cmb_OneOrder_RSCCode.Visible = False
Call HideGridUseObject_OneOrder
On Error Resume Next
With gd_OneOrder_OrderDetail
     SelectedCol = .Col: SelectedRow = .Row
     Select Case SelectedCol
        Case 5      '出貨量
           txt_OneOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_OneOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_OneOrder_SignQty.Left = txt_OneOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_OneOrder_SignQty.Top = txt_OneOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_OneOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_OneOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_OneOrder_SignQty.Text = .Text
           txt_OneOrder_SignQty.Visible = True
           txt_OneOrder_SignQty.SelStart = 0: txt_OneOrder_SignQty.SelLength = Len(txt_OneOrder_SignQty.Text)
           txt_OneOrder_SignQty.SetFocus
       Case 6      '簽單量
           txt_OneOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_OneOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_OneOrder_SignQty.Left = txt_OneOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_OneOrder_SignQty.Top = txt_OneOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_OneOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_OneOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_OneOrder_SignQty.Text = .Text
           txt_OneOrder_SignQty.Visible = True
           txt_OneOrder_SignQty.SelStart = 0: txt_OneOrder_SignQty.SelLength = Len(txt_OneOrder_SignQty.Text)
           txt_OneOrder_SignQty.SetFocus
           
       Case 7      '異常原因
       DoEvents: DoEvents
           cmb_OneOrder_RSCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_OneOrder_RSCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_OneOrder_RSCCode.Left = cmb_OneOrder_RSCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_OneOrder_RSCCode.Top = cmb_OneOrder_RSCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_OneOrder_RSCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_OneOrder_RSCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_OneOrder_RSCCode.ListCount - 1
                  If Left(cmb_OneOrder_RSCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_OneOrder_RSCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_OneOrder_RSCCode.Visible = True
           cmb_OneOrder_RSCCode.SetFocus
           SendKeys "%{DOWN}"
           
       Case 8      '權責區分：責屬
       DoEvents: DoEvents
           cmb_OneOrder_RBCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_OneOrder_RBCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_OneOrder_RBCCode.Left = cmb_OneOrder_RBCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_OneOrder_RBCCode.Top = cmb_OneOrder_RBCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_OneOrder_RBCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_OneOrder_RBCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_OneOrder_RBCCode.ListCount - 1
                  If Left(cmb_OneOrder_RBCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_OneOrder_RBCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_OneOrder_RBCCode.Visible = True
           cmb_OneOrder_RBCCode.SetFocus
           SendKeys "%{DOWN}"
           
       Case 14      '權責區分：責屬人
           txt_OneOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_OneOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_OneOrder_SignQty.Left = txt_OneOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_OneOrder_SignQty.Top = txt_OneOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_OneOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_OneOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_OneOrder_SignQty.Text = .Text
           txt_OneOrder_SignQty.Visible = True
           txt_OneOrder_SignQty.SelStart = 0: txt_OneOrder_SignQty.SelLength = Len(txt_OneOrder_SignQty.Text)
           txt_OneOrder_SignQty.SetFocus
           
     End Select
End With

End Sub

Private Function CheckSDNReConfirm(ByVal strUserID As String) As Boolean
'特定人員可執行 SDN Confirm 重新存檔
'權限設定儲存於 CodeLKUP ListName = [SDNRECONDURM]
'  Code：User_ID   Short：權限設定 1-可以重複執行
CheckSDNReConfirm = False
str_SQL = "Select Isnull(Rtrim(Short),'0') as CheckFlag From CodeLKUP " & _
          "Where ListName = 'SDNRECONFIRM' AND Code = '" & strUserID & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   If tmp_Rs.Fields("CheckFlag").Value = "1" Then
      CheckSDNReConfirm = True
   End If
End If
tmp_Rs.Close
End Function

'Private Sub Lb_Route_Change()
'    cmdOTQtyFix.Enabled = False
'    If Left(Lb_Route.Caption, 1) = "R" Then cmdOTQtyFix.Enabled = True
'End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")

'日期選取
Select Case mvDate.Tag
    Case "客戶簽單日期‧多車"
         txt_MultiOrder_SignDate.Text = Format(mvDate.Value, "yyyymmdd")
'    Case "客戶簽單日期‧一車"
'         txt_OneOrder_SignDate.Text = Format(mvDate.Value, "yyyymmdd") by gemini
    Case "Tab0-出車日起"
         txt_DeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
    Case "Tab0-出車日迄"
         txt_DeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
    Case "Tab2-出車日"
         txt_Tab02_Delivery_Date.Text = Format(mvDate.Value, "yyyymmdd")

End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

'Private Sub Op_CBM_Click()
'    dg_SDN_Detail.Row = 1
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix((Val(Trim(txt_Tab0_srcTotal_Volumn.Text)) * 10) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "材積"
'    dg_SDN_Cost.Col = 2: dg_SDN_Detail.Col = 2: dg_SDN_Cost.Text = Trim(dg_SDN_Detail.Text)
'    Call Cost_SumAll
'End Sub
'
'Private Sub Op_CS_Click()
'    dg_SDN_Detail.Row = 1
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(txt_Tab0_srcTotal_Case.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "箱數"
'    dg_SDN_Cost.Col = 2: dg_SDN_Detail.Col = 2: dg_SDN_Cost.Text = Trim(dg_SDN_Detail.Text) 'Fix((a * b) + 0.5)
'    Call Cost_SumAll
'End Sub

Private Sub Op_OnCheck_Click()
    ck_confirm.Visible = True
    ck_confirm.Value = 1
    ck_back.Visible = True
  
End Sub

'Private Sub Op_SumCBM_Click()
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix((Val(Trim(Me.txt_Tab0_sum_CBM.Text)) * 10) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "材積"
'    dg_SDN_Cost.Col = 2: dg_SDN_Cost.Text = Trim(txt_Tab0_Route_No.Text)
'    Call Cost_SumAll
'End Sub
'
'Private Sub Op_SumCS_Click()
'
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(Me.txt_Tab0_sum_Case.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "箱數"
'    dg_SDN_Cost.Col = 2: dg_SDN_Cost.Text = Trim(txt_Tab0_Route_No.Text)
'    Call Cost_SumAll
'End Sub
'
'
'Private Sub Op_SumWT_Click()
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(Me.txt_Tab0_sum_WT.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "重量"
'    dg_SDN_Cost.Col = 2: dg_SDN_Cost.Text = Trim(txt_Tab0_Route_No.Text)
'    Call Cost_SumAll
'End Sub

Private Sub Op_Tab2_CBM_Click()
    dg_Tab2_SDN_Cost.Row = 1
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_srcTotal_Volumn.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "材積"
    dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Detail.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(dg_Tab2_SDN_Detail.Text)
    Call Cost_Tab2_SumAll
End Sub

Private Sub Op_Tab2_CS_Click()
    dg_Tab2_SDN_Cost.Row = 1
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_srcTotal_Case.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "箱數"
    dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Detail.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(dg_Tab2_SDN_Detail.Text)
    Call Cost_Tab2_SumAll
End Sub

Private Sub Op_Tab2_SumCBM_Click()
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(Me.txt_Tab2_sum_CBM.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "材積"
    'dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_Route_NO.Text)
    Call Cost_SumAll
End Sub

Private Sub Op_Tab2_SumCS_Click()
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(Me.txt_Tab2_sum_Case.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "箱數"
    'dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_Route_NO.Text)
    Call Cost_SumAll
End Sub

Private Sub Op_Tab2_SumWT_Click()
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(Me.txt_Tab0_sum_WT.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "重量"
    'dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_Route_NO.Text)
    Call Cost_SumAll
End Sub

Private Sub Op_Tab2_WT_Click()
    dg_Tab2_SDN_Cost.Row = 1
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_srcTotal_Weight.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "重量"
    dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Detail.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(dg_Tab2_SDN_Detail.Text)
    Call Cost_Tab2_SumAll
End Sub

Private Sub Op_UnCheck_Click()
    ck_confirm.Visible = False
    ck_confirm.Value = 0
    ck_back.Visible = False
    ck_back.Value = 0
End Sub

'Private Sub OpWT_Click()
'    dg_SDN_Detail.Row = 1
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(txt_Tab0_srcTotal_Weight.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "重量"
'    dg_SDN_Cost.Col = 2: dg_SDN_Detail.Col = 2: dg_SDN_Cost.Text = Trim(dg_SDN_Detail.Text)
'    Call Cost_SumAll
'End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.mvDate.Visible = False
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_DeliveryDate_End_Click()
    'Tab0-出車日迄
    If Trim(txt_DeliveryDate_End.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_DeliveryDate_End.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_DeliveryDate_End.Text, 4) & "/" & Mid(txt_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_DeliveryDate_End.Text, 2))
       End If
    End If
    mvDate.Left = txt_DeliveryDate_End.Left + txt_DeliveryDate_End.Left
    mvDate.Top = txt_DeliveryDate_End.Top + txt_DeliveryDate_End.Top + txt_DeliveryDate_End.Height
    mvDate.Tag = "Tab0-出車日迄"
    mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_Start_Click()
    'Tab0-出車日起
    If Trim(txt_DeliveryDate_Start.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_DeliveryDate_Start.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_DeliveryDate_Start.Text, 2))
       End If
    End If
    mvDate.Left = txt_DeliveryDate_Start.Left + txt_DeliveryDate_Start.Left
    mvDate.Top = txt_DeliveryDate_Start.Top + txt_DeliveryDate_Start.Top + txt_DeliveryDate_Start.Height
    mvDate.Tag = "Tab0-出車日起"
    mvDate.Visible = True
End Sub

Private Sub txt_ExternOrderKey_KeyPress(KeyAscii As Integer)
'貨主單號
Select Case KeyAscii
       Case vbKeyReturn
            cmd_OrderQuery.SetFocus
End Select
End Sub


Private Sub txt_MultiOrder_SignDate_Click()
'客戶簽單日期：ㄧ張貨主單號對應多張排車系統訂單
If Trim(txt_MultiOrder_SignDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate.Text, 5, 2) & "/" & Right(txt_MultiOrder_SignDate.Text, 2))
   End If
End If
mvDate.Left = fra_MultiOrder_Header.Left + txt_MultiOrder_SignDate.Left - (mvDate.Width - txt_MultiOrder_SignDate.Width)
mvDate.Top = fra_MultiOrder_Header.Top + txt_MultiOrder_SignDate.Top + txt_MultiOrder_SignDate.Height
mvDate.Tag = "客戶簽單日期‧多車"
mvDate.Visible = True
End Sub

Private Sub txt_OneOrder_OrderKey_KeyPress(KeyAscii As Integer)
'訂單編號：ㄧ張貨主單號對應一張排車系統訂單
KeyAscii = 0
End Sub

Private Sub txt_OneOrder_SignQty_Change()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：簽單量
gd_OneOrder_OrderDetail.Text = txt_OneOrder_SignQty.Text
End Sub

Private Sub txt_OneOrder_SignQty_KeyDown(KeyCode As Integer, Shift As Integer)
'ㄧ張貨主單號對應ㄧ張排車系統訂單：簽單量
If KeyCode = vbKeyReturn Then
   txt_OneOrder_SignQty.Visible = False
End If
End Sub

Private Sub txt_OneOrder_SignQty_KeyPress(KeyAscii As Integer)
'ㄧ張貨主單號對應ㄧ張排車系統訂單：簽單量

If gd_OneOrder_OrderDetail.Col <> 14 Then
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
   End Select
End If
End Sub

Private Sub txt_OneOrder_SignQty_LostFocus()
'ㄧ張貨主單號對應ㄧ張排車系統訂單：簽單量
   txt_OneOrder_SignQty.Visible = False
End Sub

Private Sub txt_MultiOrder_SignQty_Change()
'ㄧ張貨主單號對應多張排車系統訂單：簽單量
gd_MultiOrder_OrderDetail.Text = txt_MultiOrder_SignQty.Text
End Sub

Private Sub txt_multiOrder_SignQty_KeyDown(KeyCode As Integer, Shift As Integer)
'ㄧ張貨主單號對應多張排車系統訂單：簽單量
If KeyCode = vbKeyReturn Then
   txt_MultiOrder_SignQty.Visible = False
End If
End Sub

Private Sub txt_multiOrder_SignQty_KeyPress(KeyAscii As Integer)
'ㄧ張貨主單號對應多張排車系統訂單：簽單量
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
   End Select
End Sub

Private Sub txt_multiOrder_SignQty_LostFocus()
'ㄧ張貨主單號對應多張排車系統訂單：簽單量
   txt_MultiOrder_SignQty.Visible = False
End Sub

Private Sub Display_OrderData_MultiReceipNo(ByVal strExtern As String)
'ㄧ張貨主單號對應多張排車系統訂單：訂單資料查詢
Screen.MousePointer = vbHourglass
fra_OneOrder_Header.Visible = False
fra_OneOrder_Detail.Visible = False
fra_MultiOrder_Header.Visible = True
fra_MultiOrder_Detail.Visible = True

'取得排車系統訂單 TRP02T
On Error GoTo err_Handle
str_SQL = "Select 路線編號,車牌號碼,駕駛人,貨運公司,出車日期,客戶編號,客戶名稱,送貨地址,貨主,訂單編號,訂單日期,出貨日期,簽單日期,系統日期,狀態 " & _
          "From SDNConfirm_OrderDate_Multi Where 貨主單號 = '" & strExtern & "'"
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
txt_MultiOrder_StorerKey.Text = tmp_Rs.Fields("貨主").Value
txt_MultiOrder_ConsigneeKey.Text = tmp_Rs.Fields("客戶編號").Value
txt_MultiOrder_FullName.Text = tmp_Rs.Fields("客戶名稱").Value
txt_MultiOrder_Address.Text = tmp_Rs.Fields("送貨地址").Value
txt_MultiOrder_OrderDate.Text = tmp_Rs.Fields("訂單日期").Value
txt_MultiOrder_ArriveDate.Text = tmp_Rs.Fields("出貨日期").Value
txt_MultiOrder_Status.Text = tmp_Rs.Fields("狀態").Value
blShipped = True    '無法判斷訂單揀貨量是否已更新(Ship_Qty)
blCanUpdate = True
If Len(Trim(tmp_Rs.Fields("簽單日期").Value)) > 0 Then
   txt_MultiOrder_SignDate.Text = tmp_Rs.Fields("簽單日期").Value
   blSDNConfirm = True
   blCanUpdate = False
Else
   txt_MultiOrder_SignDate.Text = tmp_Rs.Fields("出貨日期").Value
   blSDNConfirm = False
   blCanUpdate = True
End If
'Reset 訂單編號-路編列表
Call CreateRS_MultiOrder_RouteDate
Do While Not tmp_Rs.EOF
   rs_MultiOrder.AddNew
   rs_MultiOrder.Fields("路線編號").Value = tmp_Rs.Fields("路線編號").Value
   rs_MultiOrder.Fields("車牌號碼").Value = tmp_Rs.Fields("車牌號碼").Value
   rs_MultiOrder.Fields("駕駛人").Value = tmp_Rs.Fields("駕駛人").Value
   rs_MultiOrder.Fields("貨運公司").Value = tmp_Rs.Fields("貨運公司").Value
   rs_MultiOrder.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
   rs_MultiOrder.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close

With dg_MultiOrder
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_MultiOrder.MoveFirst
Set dg_MultiOrder.DataSource = rs_MultiOrder
With dg_MultiOrder
    .RowHeight = 250
    .Columns(0).Width = 1000         '路線編號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000         '車牌號碼
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 800          '駕駛人
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000         '貨運公司
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 1100         '訂單編號
    .Columns(4).Alignment = dbgCenter
End With

'ㄧ張貨主單號對應多張排車系統訂單：訂單名細
Call SetGridFormat_MultiOrder_OrderDetail
Dim tmpI As Double
str_SQL = "Select 項次,貨號,品名,訂單量,送貨量,簽單量,異常原因,責任歸屬,異常碼,責屬碼,訂單編號 " & _
          "From SDNConfirm_OrderDetail_MultiOrder " & _
          "Where 貨主單號 = '" & strExtern & "' Order by 項次"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
tmpI = 1
Do While Not tmp_Rs.EOF
   With gd_MultiOrder_OrderDetail
        If .Rows < (tmpI + 1) Then .Rows = .Rows + 1
        .Row = tmpI
        .Col = 0: .Text = tmp_Rs.Fields("訂單編號").Value
        .Col = 1: .Text = tmp_Rs.Fields("項次").Value
        .Col = 2: .Text = tmp_Rs.Fields("貨號").Value
        .Col = 3: .Text = tmp_Rs.Fields("品名").Value
        .Col = 4: .Text = tmp_Rs.Fields("送貨量").Value
         If blCanUpdate Then      '尚未執行 SDN Confirmed
            .Col = 5: .Text = 0   '簽單量
            .Col = 6: .Text = ""  '異常原因
            .Col = 7: .Text = ""  '責任歸屬
            .Col = 8: .Text = ""  '異常原因代碼
            .Col = 9: .Text = ""  '責任歸屬代碼
         Else
            .Col = 5: .Text = tmp_Rs.Fields("簽單量").Value
            .Col = 6: .Text = tmp_Rs.Fields("異常原因").Value
            .Col = 7: .Text = tmp_Rs.Fields("責任歸屬").Value
            .Col = 8: .Text = tmp_Rs.Fields("異常碼").Value
            .Col = 9: .Text = tmp_Rs.Fields("責屬碼").Value
        End If
   End With
   tmp_Rs.MoveNext
   tmpI = tmpI + 1
Loop
tmp_Rs.Close

If blCanUpdate Then
    cmd_MultiOrder_Deliveryok.Enabled = True
    cmd_MultiOrder_Expect.Enabled = True
    cmd_MultiOrder_NoDelivery.Enabled = True
Else
    cmd_MultiOrder_Deliveryok.Enabled = False
    cmd_MultiOrder_Expect.Enabled = False
    cmd_MultiOrder_NoDelivery.Enabled = False
End If

'特定人員可執行 SDN Confirm 重新存檔
'權限設定儲存於 CodeLKUP ListName = [SDNRECONDURM]
'  Code：User_ID   Short：權限設定 1-可以重複執行
If (Not blCanUpdate) And CheckSDNReConfirm(User_id) Then
    '允許重複執行 SDN Confirm 之使用者，開放 SDN Confirm 日期的修改
    cmd_MultiOrder_Deliveryok.Enabled = True
    cmd_MultiOrder_Expect.Enabled = True
    cmd_MultiOrder_NoDelivery.Enabled = True
    blCanUpdate = True
End If

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單查詢", Me.Caption, "Form 內部 SubProgram Display_OrderData_MultiReceiptNo", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub CreateRS_MultiOrder_RouteDate()
Call ReDim_Recordset(rs_MultiOrder)
With rs_MultiOrder
     .Fields.Append "路線編號", adVarChar, 10
     .Fields.Append "車牌號碼", adVarChar, 20
     .Fields.Append "駕駛人", adVarChar, 60
     .Fields.Append "貨運公司", adVarChar, 60
     .Fields.Append "訂單編號", adVarChar, 120
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With

End Sub

Private Sub SetGridFormat_MultiOrder_OrderDetail()
'名稱：SetGridFormatt_MultiOrder_OrderDetail
'類別：副程式
'功能：清除並設定 [SDN Confirm] 表單 [ㄧ張貨主單號對應多張排車系統訂單] 訂單名細顯示格式
'參數：傳入值：無
Dim sub_var1 As Integer, sub_var2 As Integer
gd_MultiOrder_OrderDetail.Visible = False
With gd_MultiOrder_OrderDetail
     .Rows = 2: .FixedRows = 1: .Cols = 11
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
     .ColWidth(0) = 1200
     .ColWidth(1) = 500
     .ColWidth(2) = 1000
     .ColWidth(3) = 2700
     .ColWidth(4) = 750
     .ColWidth(5) = 750
     .ColWidth(6) = 1850
     .ColWidth(7) = 1400
     .ColWidth(8) = 1000
     .ColWidth(9) = 1000
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "訂單編號"
     .Col = 1: .Text = "項次"
     .Col = 2: .Text = "貨號"
     .Col = 3: .Text = "中文品名"
     .Col = 4: .Text = "送貨量"
     .Col = 5: .Text = "簽單量"
     .Col = 6: .Text = "異常原因"
     .Col = 7: .Text = "責任歸屬"
     .Col = 8: .Text = "異常碼"
     .Col = 9: .Text = "責屬碼"
     '設定列表之文字對齊
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .Rows = 2
     .Row = 0
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .Text = ""
     Next sub_var1
     
End With
gd_MultiOrder_OrderDetail.Visible = True
End Sub

Private Sub HideGridUseObject_MultiOrder()
'ㄧ張貨主單號對應多張排車系統訂單：隱藏 [簽單量] [異常原因] [責屬] 控制項
txt_MultiOrder_SignQty.Visible = False
cmb_MultiOrder_RBCCode.Visible = False
cmb_MultiOrder_RSCCode.Visible = False
End Sub


Private Sub SetGridFormat_Tab2_SDN_Detail()
'回傳欲設定之路線編號資料
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Tab2_SDN_Detail.Visible = False
With dg_Tab2_SDN_Detail
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
     .ColWidth(0) = 450
     .ColWidth(1) = 450
     .ColWidth(2) = 1500
     .ColWidth(3) = 2000
     .ColWidth(4) = 700
     .ColWidth(5) = 700
     .ColWidth(6) = 700
     .ColWidth(7) = 900
     .ColWidth(8) = 3500
     .ColWidth(9) = 450

     '設定列表之標題:送貨日,路線編號,車號 ,駕駛人,公司,領款人,應收單價,應付單價,其他金額,原因, 起點,迄點
     '二次排車,路線編號,客戶單號,日期,指送客戶,箱數,材積,重量,多車
     .Row = 0
     .Col = 0: .Text = "確認"
     .Col = 1: .Text = "回收"
     .Col = 2: .Text = "客戶單號"
     .Col = 3: .Text = "指送客戶"
     .Col = 4: .Text = "箱數"
     .Col = 5: .Text = "材積"
     .Col = 6: .Text = "重量"
     .Col = 7: .Text = "多車"
     .Col = 8: .Text = "備註"
     .Col = 9: .Text = "小計"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter


     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Tab2_SDN_Detail.Visible = True
End Sub


Private Sub SetGridFormat_Tab2_SDN_Cost()
'回傳欲設定之路線編號資料
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Tab2_SDN_Cost.Visible = False
With dg_Tab2_SDN_Cost
     .Rows = 2: .Cols = 14
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
     .ColWidth(0) = 800
     .ColWidth(1) = 1500
     .ColWidth(2) = 1000
     .ColWidth(3) = 700
     .ColWidth(4) = 700
     .ColWidth(5) = 600
     .ColWidth(6) = 800
     .ColWidth(7) = 700
     .ColWidth(8) = 700
     .ColWidth(9) = 800
     .ColWidth(10) = 1500
     .ColWidth(11) = 700
     .ColWidth(12) = 700
     .ColWidth(13) = 1000
     '設定列表之標題:送貨日,路線編號,車號 ,駕駛人,公司,領款人,計費數量,應收單價,應付單價,其他金額,原因, 起點,迄點
     .Row = 0
     .Col = 0: .Text = "代碼"
     .Col = 1: .Text = "客戶"
     .Col = 2: .Text = "單號"
     .Col = 3: .Text = "起點"
     .Col = 4: .Text = "迄點"
     .Col = 5: .Text = "單位"
     .Col = 6: .Text = "計費數量"
     .Col = 7: .Text = "應收單"
     .Col = 8: .Text = "應付單"
     .Col = 9: .Text = "其他金額"
     .Col = 10: .Text = "原因"
     .Col = 11: .Text = "實收"
     .Col = 12: .Text = "實付"
     .Col = 13: .Text = "請款類別"
     '設定列表之文字對齊
     
     .ColAlignment(0) = flexAlignLeftCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignLeftCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignRightCenter
     .ColAlignment(10) = flexAlignLeftCenter
     .ColAlignment(11) = flexAlignRightCenter
     .ColAlignment(12) = flexAlignRightCenter
     .ColAlignment(13) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Tab2_SDN_Cost.Visible = True
End Sub

Public Sub NextPositionTab2Detail(ByVal r As Integer, ByVal C As Integer)     '移動文字方塊
    On Error GoTo NextError
    Text3.Width = dg_Tab2_SDN_Detail.CellWidth                     '寬度
    Text3.Height = dg_Tab2_SDN_Detail.CellHeight                   '高度
    Text3.Left = dg_Tab2_SDN_Detail.Left + dg_Tab2_SDN_Detail.ColPos(C) + 30 '左側
    Text3.Top = dg_Tab2_SDN_Detail.Top + dg_Tab2_SDN_Detail.RowPos(r)     '上方
    Text3.Text = dg_Tab2_SDN_Detail.Text       '將MSFlexGrid目前作用儲存格內容放置於文字方塊
    Text3.Visible = True                '將文字方塊顯示於畫面上
    Text3.SetFocus                      '將游標移至文字方塊
    Exit Sub
NextError:
    MsgBox err.Description
End Sub

Public Sub NextPositionTab2Cost(ByVal r As Integer, ByVal C As Integer)     '移動文字方塊
    On Error GoTo NextError
    Text4.Width = dg_Tab2_SDN_Cost.CellWidth                     '寬度
    Text4.Height = dg_Tab2_SDN_Cost.CellHeight                   '高度
    Text4.Left = dg_Tab2_SDN_Cost.Left + dg_Tab2_SDN_Cost.ColPos(C) + 30 '左側
    Text4.Top = dg_Tab2_SDN_Cost.Top + dg_Tab2_SDN_Cost.RowPos(r)     '上方
    Text4.Text = dg_Tab2_SDN_Cost.Text       '將MSFlexGrid目前作用儲存格內容放置於文字方塊
    Text4.Visible = True                '將文字方塊顯示於畫面上
    Text4.SetFocus                      '將游標移至文字方塊
    Exit Sub
NextError:
    MsgBox err.Description
End Sub

Private Sub Text1_LostFocus()
    On Error GoTo TextError
        Text1.Visible = False
    Exit Sub
TextError:
    MsgBox err.Description
End Sub


Private Sub Text3_LostFocus()
    On Error GoTo TextError
        Text3.Visible = False
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Text3_Change()  '將文字方塊內容寫至對應儲存格
    On Error GoTo TextError
    dg_Tab2_SDN_Detail.Text = Text3.Text   '將文字方塊內容寫至對應儲存格
    If dg_Tab2_SDN_Detail.Col = 4 Or dg_Tab2_SDN_Detail.Col = 5 Or dg_Tab2_SDN_Detail.Col = 6 Then
        Call Tab2Detail_Sum
    End If
    Exit Sub
 
TextError:
    MsgBox err.Description
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    On Error GoTo TextError
    If KeyAscii = vbKeyReturn Then                '在按下Enter時，決定下個grid的位置
        If dg_Tab2_SDN_Detail.Col < 8 Then
            dg_Tab2_SDN_Detail.Col = dg_Tab2_SDN_Detail.Col + 1
            NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
        End If

    End If
    'i = KeyAscii
    If KeyAscii = 1 Then 'Ctrl+A
        Call cmd_Tab2_AddOrder_Click
    End If
    If KeyAscii = 4 Then 'Ctrl+D
        Call cmd_Tab2_DelOrder_Click
    End If
    If KeyAscii = 26 Then 'Ctrl+Z
        dg_Tab2_SDN_Cost.Row = 1
        dg_Tab2_SDN_Cost.Col = 0
        NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
    End If
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Text4_LostFocus()
    On Error GoTo TextError
        Text4.Visible = False
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Text4_Change()  '將文字方塊內容寫至對應儲存格
    On Error GoTo TextError
    dg_Tab2_SDN_Cost.Text = Text4.Text   '將文字方塊內容寫至對應儲存格
    Call Cost_Tab2_Sum
    Exit Sub
 
TextError:
    MsgBox err.Description
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    On Error GoTo TextError
    If KeyAscii = vbKeyReturn Then                '在按下Enter時，決定下個grid的位置
        If dg_Tab2_SDN_Cost.Col = 0 Then
            If Len(Trim(dg_Tab2_SDN_Cost.Text)) > 0 Then
                Call Confirm_Recordset_Closed(tmp_Rs)
                str_SQL = "SELECT RTRIM(CostCode) AS 代碼,CostName as 客戶名稱,Receivable as 應收單價,Payable as 應付單價,AreaStart as 起點,AreaEnd as 迄點,CostKind as 請款類別  " & _
                          "From TRP17M where CostCode='" & Trim(Text4.Text) & "'"
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If Not tmp_Rs.EOF Then
                    dg_Tab2_SDN_Cost.Col = 1: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(1).Value)
                    dg_Tab2_SDN_Cost.Col = 3: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(4).Value)
                    dg_Tab2_SDN_Cost.Col = 4: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(5).Value)
                    dg_Tab2_SDN_Cost.Col = 7: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(2).Value)
                    dg_Tab2_SDN_Cost.Col = 8: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(3).Value)
                    dg_Tab2_SDN_Cost.Col = 13: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(6).Value)
                    dg_Tab2_SDN_Cost.Col = 1
                    Call Cost_Tab2_SumAll
                End If
                tmp_Rs.Close
            End If
        End If
        If dg_Tab2_SDN_Cost.Col < 12 Then
            dg_Tab2_SDN_Cost.Col = dg_Tab2_SDN_Cost.Col + 1
            NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
        End If
    End If
    'i = KeyAscii
    If KeyAscii = 1 Then 'Ctrl+A
        Call cmd_Tab2_AddCost_Click
    End If
    If KeyAscii = 4 Then 'Ctrl+D
        Call cmd_Tab2_DelCost_Click
    End If
    If KeyAscii = 26 Then 'Ctrl+Z
    
    End If
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Cost_Tab2_Sum()  '統計實收與實付
    intR = dg_Tab2_SDN_Cost.Col
    '統計實收
    If dg_Tab2_SDN_Cost.Col = 6 Or dg_Tab2_SDN_Cost.Col = 7 Then
        dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 7: B = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 11: dg_Tab2_SDN_Cost.Text = Round(a * B)
    End If
    dg_Tab2_SDN_Cost.Col = intR
    '統計實付
    If dg_Tab2_SDN_Cost.Col = 6 Or dg_Tab2_SDN_Cost.Col = 8 Or dg_Tab2_SDN_Cost.Col = 9 Then
        dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 8: B = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 9: C = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 12: dg_Tab2_SDN_Cost.Text = Round(a * B + C, 0)
    End If
    dg_Tab2_SDN_Cost.Col = intR
End Sub

Private Sub Cost_SumAll()  '統計實收與實付
'    intR = dg_SDN_Cost.Col
'    '統計實收
'    dg_SDN_Cost.Col = 6: a = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 7: B = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 11: dg_SDN_Cost.Text = Round(a * B, 0)
'    dg_SDN_Cost.Col = intR
'    '統計實付
'    dg_SDN_Cost.Col = 6: a = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 8: B = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 9: C = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 12: dg_SDN_Cost.Text = Round(a * B + C, 0)
'    dg_SDN_Cost.Col = intR
End Sub

Private Sub Cost_Tab2_SumAll()  '統計實收與實付
'    intR = dg_SDN_Cost.Col
'    '統計實收
'    dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 7: B = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 11: dg_Tab2_SDN_Cost.Text = Round(a * B, 0)
'    dg_Tab2_SDN_Cost.Col = intR
'    '統計實付
'    dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 8: B = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 9: C = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 12: dg_Tab2_SDN_Cost.Text = Round(a * B + C, 0)
'    dg_Tab2_SDN_Cost.Col = intR
End Sub

Private Sub Tab2Detail_Sum()  '統計實收與實付
    intC = dg_Tab2_SDN_Detail.Col
    intR = dg_Tab2_SDN_Detail.Row
    txt_Tab2_srcTotal_Case.Text = 0
    txt_Tab2_srcTotal_Volumn.Text = 0
    txt_Tab2_srcTotal_Weight.Text = 0
    For i = 1 To dg_Tab2_SDN_Detail.Rows - 1
        dg_Tab2_SDN_Detail.Row = i
        dg_Tab2_SDN_Detail.Col = 4
        txt_Tab2_srcTotal_Case.Text = Val(dg_Tab2_SDN_Detail.Text) + Val(txt_Tab2_srcTotal_Case.Text)
        dg_Tab2_SDN_Detail.Col = 5
        txt_Tab2_srcTotal_Volumn.Text = Val(dg_Tab2_SDN_Detail.Text) + Val(txt_Tab2_srcTotal_Volumn.Text)
        dg_Tab2_SDN_Detail.Col = 6
        txt_Tab2_srcTotal_Weight.Text = Val(dg_Tab2_SDN_Detail.Text) + Val(txt_Tab2_srcTotal_Weight.Text)
    Next
    dg_Tab2_SDN_Detail.Col = intC
    dg_Tab2_SDN_Detail.Row = intR
End Sub

Private Sub txt_Tab02_C_VEHICLE_ID_NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '在按下Enter時，決定下個grid的位置
        txt_Tab02_C_VEHICLE_ID_NO.SetFocus
    End If
End Sub

Private Sub txt_Tab02_Delivery_Date_Click()
    'Tab2-出車日起
    If Trim(txt_Tab02_Delivery_Date.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_Tab02_Delivery_Date.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_Tab02_Delivery_Date.Text, 4) & "/" & Mid(txt_Tab02_Delivery_Date.Text, 5, 2) & "/" & Right(txt_Tab02_Delivery_Date.Text, 2))
       End If
    End If
    mvDate.Left = txt_Tab02_Delivery_Date.Left + txt_Tab02_Delivery_Date.Left
    mvDate.Top = txt_Tab02_Delivery_Date.Top + txt_Tab02_Delivery_Date.Top + txt_Tab02_Delivery_Date.Height
    mvDate.Tag = "Tab2-出車日"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab02_Delivery_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '在按下Enter時，決定下個grid的位置
        txt_Tab02_C_VEHICLE_ID_NO.SetFocus
    End If
End Sub

Private Sub txt_Tab02_Driver_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '在按下Enter時，決定下個grid的位置
        txt_Tab02_Receiver.SetFocus
    End If
End Sub

Private Sub txt_Tab02_Receiver_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '在按下Enter時，決定下個grid的位置
        dg_Tab2_SDN_Detail.Col = 2
        dg_Tab2_SDN_Detail.Row = 1
        NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
    End If
    
End Sub

Private Sub Clear_CardData()
    txt_Tab02_Delivery_Date.Text = ""
    txt_Tab02_C_VEHICLE_ID_NO.Text = ""
    txt_Tab02_Driver.Text = ""
    txt_Tab02_Receiver.Text = ""
    txt_Tab02_C_Route_No = ""
    txt_Tab2_srcTotal_Case.Text = ""
    txt_Tab2_srcTotal_Volumn.Text = ""
    txt_Tab2_srcTotal_Weight.Text = ""
    txt_Tab2_sum_Case.Text = ""
    txt_Tab2_sum_CBM.Text = ""
    txt_Tab2_sum_WT.Text = ""
    dg_Tab2_SDN_Detail.Rows = 2
    dg_Tab2_SDN_Detail.Row = 1
    For i = 0 To dg_Tab2_SDN_Detail.Cols - 1
        dg_Tab2_SDN_Detail.Col = i
        dg_Tab2_SDN_Detail.Text = ""
    Next
    dg_Tab2_SDN_Cost.Rows = 2
    dg_Tab2_SDN_Cost.Row = 1
    For i = 0 To dg_Tab2_SDN_Cost.Cols - 1
        dg_Tab2_SDN_Cost.Col = i
        dg_Tab2_SDN_Cost.Text = ""
    Next
End Sub

Private Sub Tab0_SumQty(cr As String)
    str_SQL = "select isnull(sum(ChargeQty),0) from SDN05T where C_ROUTE_NO ='" & cr & "' and Uom in ('重量','材積')"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    txt_Tab0_SumQty.Text = tmp_Rs.Fields(0).Value
    tmp_Rs.Close
End Sub

Private Sub txtCustomerOrderkey_Change()
    If txtCustomerOrderkey.Visible = False Then Exit Sub
    rsOrderT0("驗收單號") = txtCustomerOrderkey.Text
End Sub

Private Sub txtCustomerOrderkey_Click()
    txtCustomerOrderkey.SetFocus: txtCustomerOrderkey.SelStart = 0: txtCustomerOrderkey.SelLength = Len(txtCustomerOrderkey.Text)
End Sub

Private Sub txtCustomerOrderkey_GotFocus()
    txtCustomerOrderkey.SelStart = 0: txtCustomerOrderkey.SelLength = Len(txtCustomerOrderkey.Text)
End Sub

Private Sub txtCustomerOrderkey_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   txtCustomerOrderkey.Visible = False
End If
End Sub

Private Sub txtCustomerOrderkey_LostFocus()
   txtCustomerOrderkey.Visible = False
   dgOrderT0.Col = 0
End Sub

Private Sub txtDeliveryE_Click()
    Set objMvdateTarget = txtDeliveryE
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryS_Click()
    Set objMvdateTarget = txtDeliveryS
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtSignDateS_Click()
    Set objMvdateTarget = txtSignDateS
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub txtSignDateE_Click()
    Set objMvdateTarget = txtSignDateE
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtSignDateS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txtSignDateE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryDateST0_Click()
    Set objMvdateTarget = txtDeliveryDateST0
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryDateET0_Click()
    Set objMvdateTarget = txtDeliveryDateET0
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryDateST0_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryDateET0_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub
