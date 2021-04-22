VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_TRPTrack 
   Caption         =   "到貨追蹤表"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "細明體"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_Report_TRPTrack.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   14070
   Begin MSComCtl2.DTPicker dtpDeliveryTime 
      Height          =   375
      Left            =   9720
      TabIndex        =   29
      Top             =   3480
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
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   201785347
      UpDown          =   -1  'True
      CurrentDate     =   39438
   End
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3360
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
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
      StartOfWeek     =   201785345
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   11
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14055
      Begin VB.CommandButton cmdImportExternShipKey 
         BackColor       =   &H00C0C0FF&
         Caption         =   "託運單號匯入"
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
         Left            =   12840
         Style           =   1  '圖片外觀
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDeliveryE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
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
         TabIndex        =   32
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
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
         TabIndex        =   31
         Top             =   240
         Width           =   1485
      End
      Begin VB.CheckBox chkShowWH 
         Caption         =   "顯示七天內裝載點"
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5760
         TabIndex        =   30
         ToolTipText     =   "顯示裝載點，查詢需較久的時間"
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3480
         Style           =   1  '項目包含核取方塊
         TabIndex        =   28
         ToolTipText     =   "到貨狀態"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   4560
         Style           =   1  '項目包含核取方塊
         TabIndex        =   27
         ToolTipText     =   "訂單狀態"
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox List4 
         Columns         =   2
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   120
         Style           =   1  '項目包含核取方塊
         TabIndex        =   4
         ToolTipText     =   "單別"
         Top             =   1320
         Width           =   3405
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   6360
         Style           =   1  '項目包含核取方塊
         TabIndex        =   5
         ToolTipText     =   "貨主"
         Top             =   240
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Columns         =   3
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   8880
         Style           =   1  '項目包含核取方塊
         TabIndex        =   6
         ToolTipText     =   "區碼"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPreView 
         BackColor       =   &H00C0FFFF&
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
         Height          =   870
         Left            =   6720
         Picture         =   "frm_Report_TRPTrack.frx":0342
         Style           =   1  '圖片外觀
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FF8080&
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
         Left            =   5640
         Picture         =   "frm_Report_TRPTrack.frx":064C
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
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
         TabIndex        =   0
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtRouteE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
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
         TabIndex        =   3
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtRouteS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
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
         TabIndex        =   2
         Top             =   960
         Width           =   1485
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
         Left            =   10440
         Picture         =   "frm_Report_TRPTrack.frx":0956
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   1200
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
         Left            =   11640
         Picture         =   "frm_Report_TRPTrack.frx":1C50
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   1200
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
         Left            =   11640
         Picture         =   "frm_Report_TRPTrack.frx":2B862
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   240
         Width           =   1065
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
         Left            =   10440
         Picture         =   "frm_Report_TRPTrack.frx":2BB74
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   240
         Width           =   1065
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog 
         Left            =   13320
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   285
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
         Index           =   2
         Left            =   2640
         TabIndex        =   33
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "A2B.提貨配送"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   10
         Left            =   1560
         TabIndex        =   26
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "RC.提貨入庫"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   9
         Left            =   1560
         TabIndex        =   25
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "R.退貨訂單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   8
         Left            =   1560
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "A.麒麟轉倉"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "I.正常訂單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1035
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
         TabIndex        =   19
         Top             =   660
         Width           =   360
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
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   960
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
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1005
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
         Index           =   1
         Left            =   2655
         TabIndex        =   16
         Top             =   1020
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   15
      Top             =   6030
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "狀態"
            TextSave        =   "狀態"
            Object.ToolTipText     =   "狀態"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   18177
            MinWidth        =   2646
            Object.ToolTipText     =   "資料筆數"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.ToolTipText     =   "使用者"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Report_TRPTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmdImportExternShipKey_Click()
msg_title = "托運單號匯入"

On Error Resume Next
Dim strFileName As String, strFieldName As String, k As Integer, j As Integer, i As Integer, arrTmp

With dlgCommonDialog
    .DialogTitle = "托運單號匯入"
    .CancelError = True
    .InitDir = App.Path
    'ToDo: 設定通用對話方塊控制項的旗標及屬性
    .Filter = "*.csv|*.csv"
    .ShowOpen
    strFileName = .FileName
    
    If err.Number = cdlCancel Then strFileName = "": Exit Sub
    
    If Len(strFileName) = 0 Then Exit Sub

End With

On Error GoTo err_Handle
If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "托運單號匯入": Exit Sub '找不到檔案

Screen.MousePointer = 11

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

'    '尋找指定工作表
'    For i = 1 To .Sheets.Count
'        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '選定工作表
'        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '取第一個工作表名稱
'        .Sheets(.Sheets(i).Name).Select
'    Next
'
'    '找不到指定工作表，選用第一個
'    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(1).Select
.Sheets(1).Select
    
'    For i = 1 To .Sheets.Count
'        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '選定工作表
'
'    Next
'
'    If RTrim(.Sheets(i).Name) <> strSheetName Then
'        '找不到用第一個
'
''        MsgBox "找不到 " & strSheetName & "工作表！", vbOKOnly + vbInformation, "Excel2Recordset"
''        .Quit: Set MyXlsApp = Nothing
''        Exit Sub
'    End If
    
    k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
'    If strFieldName = "" Then
'        '取欄位名稱
        For i = 1 To 255
            If Len(RTrim(.Cells(1, i) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(RTrim(.Cells(1, i))) & Chr(9)
        Next i
        k = 2 '由第二列開始匯入
'    End If
    
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '寫入Recordset
    Do While Len(RTrim(.Cells(k, 3))) > 0 'Or Len(RTrim(.Cells(k, 2))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    MyXlsApp.Quit: Set MyXlsApp = Nothing
    
endsub:

.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
Dim intLine As Integer
intLine = 1

Do While Not rsTmp.EOF
    str_SQL = "UPDATE Orders set ExternShipKey='" & RTrim(rsTmp("十碼貨號")) & "' " & _
         "where EXTERNOrderkey='" & RTrim(rsTmp("訂單號碼")) & "' and type <> '刪單' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

   If RowsAffect > 0 Then intLine = intLine + 1
rsTmp.MoveNext
Loop

rsTmp.Close: Set rsTmp = Nothing
Screen.MousePointer = 0

MsgBox "更新 " & intLine & "筆託運單號!", 64, msg_title

Exit Sub
err_Handle:
Dim str As String
If MyXlsApp Is Nothing = False Then MyXlsApp.Quit: Set MyXlsApp = Nothing

If err.Number = 3367 Then
    str = "欄位名稱( " & arrTmp(i) & ")重複！"
    
ElseIf err.Number = -2147217887 Then
    str = "請確認儲存格(" & Chr(64 + j) & k & ")，資料是否有誤！"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdPreView_Click()

Dim i As Integer, j As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub
Screen.MousePointer = 11

'資料寫入 Access 資料庫
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From 到貨追蹤表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "到貨追蹤表", cnAccess, adOpenStatic, adLockOptimistic

With rsMain
    .MoveFirst
    Do While Not .EOF
       rs_Access.AddNew
       For i = 0 To .Fields.Count - 1
           rs_Access.Fields(i).Value = .Fields(i).Value
       Next i
       rs_Access.Update
       .MoveNext
    Loop
    .MoveFirst
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
End With

strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
With MSAccessAP
    .Visible = False
    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
    .DoCmd.Maximize
    
    '寫入USER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    .Reports(Me.Caption).[User_id].Caption = User_id
    .DoCmd.Close

    .DoCmd.OpenReport "到貨追蹤表", acViewPreview
    .Visible = True

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdPrint_Click()
Dim i As Integer, j As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub
Screen.MousePointer = 11

'資料寫入 Access 資料庫
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From 到貨追蹤表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "到貨追蹤表", cnAccess, adOpenStatic, adLockOptimistic

With rsMain
    .MoveFirst
    Do While Not .EOF
       rs_Access.AddNew
       For i = 0 To .Fields.Count - 1
           rs_Access.Fields(i).Value = .Fields(i).Value
       Next i
       rs_Access.Update
       .MoveNext
    Loop
    .MoveFirst
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
End With

strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
With MSAccessAP
    .Visible = False
    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
    
    '寫入USER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    .Reports(Me.Caption).[User_id].Caption = User_id
    .DoCmd.Close
    
    '直接列印至印表機
    .Visible = False
    .DoCmd.OpenReport "到貨追蹤表", acViewNormal
    .CloseCurrentDatabase
    .Quit: Set MSAccessAP = Nothing

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd2Excel_Click()
'資料排序
Recordset2Excel Me.Caption, rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
                
    End With
End If
Set MyXlsApp = Nothing

End Sub
Private Sub cmdQuery_Click()
Dim chc_Route As String, chc_DeliveryDate As String, i As Integer, strSelected As String, strSectionKey As String, chc_DeliveryDate1 As String, chc_DeliveryDate2 As String, strViewName As String
strViewName = "TRPTrack" & Replace(strComputerName, "-", "")
'If Len(RTrim(txtDeliveryS)) = 0 And Len(RTrim(txtDeliveryE)) = 0 Then MsgBox "請輸入日期區間!", 64, Me.Caption: Exit Sub
If Len(RTrim(txtDeliveryDateS)) = 0 And Len(RTrim(txtDeliveryDateS)) = 0 Then MsgBox "請輸入到貨日期區間!", 64, Me.Caption: Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"

cn.Execute "if object_id ('tempdb..##" & strViewName & "') is not null drop table ##" & strViewName, RowsAffect, adExecuteNoRecords

'未轉入
str_SQL = "select 單別 = RTrim(o.Priority),狀態 = '未轉入                ',區碼 = rtrim(t1m.area_code),二次車號 = '                  ',二次駕駛人 = '                  ',二次路編 = '          ',一次車號 = '                  ',一次駕駛人 = '                  ',一次路編 = '          ',訂單日期 = o.orderdate,到貨日期 = o.DeliveryDate " & _
",簽單日期 = '                    ',POD天數 = '          ',貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS單號 = rtrim(o.orderkey),訂單號碼 = rtrim(o.externorderkey) " & _
",客戶名稱 = rtrim(t1m.short_name),縣市 = rtrim(t2m.description),訂單板數 = round(sum(case when s.pallet = 0 then 0 else od.originalqty/s.pallet end),3),訂單箱數 = sum(case when s.casecnt = 0 then 0 else floor(od.originalqty/s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then od.originalqty else cast(od.originalqty as int) % cast(s.casecnt as int) end),總個數 = sum(od.originalqty),預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(od.originalqty /s.casecnt) end) " & _
",入庫完成 = '     ',訂單重量 = sum(od.originalqty*s.stdgrosswgt),訂單材積 = sum(od.originalqty*s.stdcube),訂單備註 = cast(o.notes as varchar(1000)),預估到貨 = '                    ' " & _
",達交 = '     ',遲交 = ' ',訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) into ##" & strViewName & _
" from orders o (nolock) join orderdetail od (nolock) on o.orderkey = od.orderkey and o.b_phone2 is null and isnull(o.type,'') <> '刪單' " & _
"join gv_skuxpack s(nolock) on s.sku = od.sku and s.storerkey = o.storerkey join trp16m t16 on t16.storerkey = o.storerkey " & _
"left join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end and t1m.storerkey = o.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),o.deliverydate,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),o.priority,o.orderkey,o.externorderkey,o.DeliveryDate,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,cast(o.notes as varchar(1000)),t2m.description,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'未排
str_SQL = "insert into ##" & strViewName & " select 單別 = RTrim(w2.Priority),狀態 = '未排' " & _
",區碼 = rtrim(t1m.area_code),二次車號 = '',二次駕駛人 = '',二次路編 = '          ',一次車號 = '',一次駕駛人 = '',一次路編 = '          ',訂單日期 = o.orderdate,到貨日 = w2.arrive_date " & _
",簽單日期 = '',POD天數 = '',貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS單號 = rtrim(w2.receipt_no) " & _
",訂單號碼 = rtrim(w2.extern),客戶名稱 = rtrim(t1m.short_name),縣市 = rtrim(t2m.description),訂單板數 = round(sum(case when s.pallet = 0 then 0 else w3.order_qty/s.pallet end),3) " & _
",訂單箱數 = sum(case when s.casecnt = 0 then 0 else floor(w3.order_qty/s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then w3.order_qty else cast(w3.order_qty as int) % cast(s.casecnt as int) end) " & _
",總個數 = sum(w3.order_qty),預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(w3.order_qty /s.casecnt) end),入庫完成 = '' " & _
",訂單重量 = sum(w3.order_qty*s.stdgrosswgt),訂單材積 = sum(w3.order_qty*s.stdcube),訂單備註 = rtrim(w2.description),預估到貨 = '' " & _
",達交 = ' ',遲交 = ' ',訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from trp03w w3 (nolock) join trp02w w2 (nolock) on w2.receipt_no = w3.receipt_no join orders o (nolock) on o.orderkey = w2.c_receipt_no " & _
"join gv_skuxpack s(nolock) on s.sku = w3.product_no and s.storerkey = w2.storerkey join trp16m t16 on t16.storerkey = w2.storerkey " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else w2.consigneekey end and t1m.storerkey = w2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),w2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),w2.priority,w2.receipt_no,w2.extern ,w2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,w2.DESCRIPTION,o.adddate,o.addwho"

'For i = 0 To List5.ListCount - 1
'    If List5.Selected(i) Then If List5.List(i) = "未排" Then cn.Execute str_SQL, RowsAffect, adExecuteNoRecords: Exit For
'Next

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'未排
str_SQL = "insert into ##" & strViewName & " select 單別 = RTrim(w2.Priority),狀態 = '未排',區碼 = rtrim(t1m.area_code),二次車號 = '',二次駕駛人 = '',二次路編 = '          ',一次車號 = '',一次駕駛人 = '',一次路編 = '          ' " & _
",訂單日期 = o.orderdate,到貨日 = w2.arrive_date,簽單日期 = '',POD天數 = '',貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS單號 = rtrim(w2.receipt_no) " & _
",訂單號碼 = rtrim(w2.extern),客戶名稱 = rtrim(t1m.short_name),縣市 = rtrim(t2m.description),訂單板數 = round(sum(case when s.pallet = 0 then 0 else w3.order_qty/s.pallet end),3) " & _
",訂單箱數 = sum(case when s.casecnt = 0 then 0 else floor(w3.order_qty/s.casecnt) end),訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then w3.order_qty else cast(w3.order_qty as int) % cast(s.casecnt as int) end) " & _
",總個數 = sum(w3.order_qty),預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(w3.order_qty /s.casecnt) end) " & _
",入庫完成 = '',訂單重量 = sum(w3.order_qty*s.stdgrosswgt),訂單材積 = sum(w3.order_qty*s.stdcube),訂單備註 = rtrim(w2.description) " & _
",預估到貨 = '',達交 = ' ',遲交 = ' ',訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from ort03w w3 (nolock) join ort02w w2 (nolock) on w2.receipt_no = w3.receipt_no join orders o (nolock) on o.orderkey = w2.c_receipt_no " & _
"join gv_skuxpack s(nolock) on s.sku = w3.product_no and s.storerkey = w2.storerkey join trp16m t16 on t16.storerkey = w2.storerkey " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else w2.consigneekey end and t1m.storerkey = w2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),w2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),w2.priority,w2.receipt_no,w2.extern ,w2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,w2.DESCRIPTION,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'保留
str_SQL = "insert into ##" & strViewName & " select 單別 = RTrim(t2.Priority),狀態 = '保留',區碼 = rtrim(t1m.area_code),二次車號 = '',二次駕駛人 = '',二次路編 = '          ',一次車號 = '',一次駕駛人 = '',一次路編 = '          ' " & _
",訂單日期 = o.orderdate,到貨日 = t2.arrive_date,簽單日期 = '',POD天數 = '',貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS單號 = rtrim(t2.receipt_no),訂單號碼 = rtrim(t2.extern),客戶名稱 = rtrim(t1m.short_name),縣市 = rtrim(t2m.description) " & _
",訂單板數 = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3),訂單箱數 = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),總個數 = sum(t3.order_qty) " & _
",預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end),入庫完成 = '',訂單重量 = sum(t3.order_qty*s.stdgrosswgt) " & _
",訂單材積 = sum(t3.order_qty*s.stdcube),訂單備註 = rtrim(t2.description),預估到貨 = '',達交 = ' ',遲交 = ' ' " & _
",訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from trp03t t3 (nolock) join trp02t t2 (nolock) on t2.receipt_no = t3.receipt_no and t2.route_no = 'D' " & _
"join orders o (nolock) on o.orderkey = t2.c_receipt_no join gv_skuxpack s on s.sku = t3.product_no and s.storerkey = t2.storerkey " & _
"join trp16m t16(nolock) on t16.storerkey = t2.storerkey join trp01m t1m on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end and t1m.storerkey = t2.storerkey " & _
"left join trp02m t2m(nolock) on t2m.zip = t1m.zip " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),t2.priority,t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,o.adddate,o.addwho"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'保留
str_SQL = "insert into ##" & strViewName & " select 單別 = RTrim(t2.Priority),狀態 = '保留',區碼 = rtrim(t1m.area_code),二次車號 = '',二次駕駛人 = '',二次路編 = '          ',一次車號 = '',一次駕駛人 = '',一次路編 = '          ' " & _
",訂單日期 = o.orderdate,到貨日 = t2.arrive_date,簽單日期 = '',POD天數 = '',貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS單號 = rtrim(t2.receipt_no),訂單號碼 = rtrim(t2.extern),客戶名稱 = rtrim(t1m.short_name),縣市 = rtrim(t2m.description) " & _
",訂單板數 = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3),訂單箱數 = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),總個數 = sum(t3.order_qty) " & _
",預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end),入庫完成 = '',訂單重量 = sum(t3.order_qty*s.stdgrosswgt) " & _
",訂單材積 = sum(t3.order_qty*s.stdcube),訂單備註 = rtrim(t2.description),預估到貨 = '',達交 = ' ',遲交 = ' ',訂單來源 = isnull(o.updatesource,'') " & _
",訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from ort03t t3 (nolock) join ort02t t2 (nolock) on t2.receipt_no = t3.receipt_no and t2.route_no = 'D' join orders o (nolock) on o.orderkey = t2.c_receipt_no " & _
"join gv_skuxpack s(nolock) on s.sku = t3.product_no and s.storerkey = t2.storerkey join trp16m t16 on t16.storerkey = t2.storerkey " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end and t1m.storerkey = t2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),t2.priority,t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'已排
str_SQL = "insert into ##" & strViewName & " Select 單別 = RTrim(t2.Priority),狀態 = '已排',區碼 = rtrim(t1m.area_code) " & _
",二次車號 = isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No) ,二次駕駛人 = Rtrim(Isnull(t09m.Driver,'')),二次路編 = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),一次車號 = rtrim(isnull((select top 1 t9.VEHICLE_ID_NO  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),駕駛人 = RTRIM(ISNULL((select top 1 t9.DRIVER  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),一次路編 = RTRIM(ISNULL(t2.route_no,'')),訂單日期 = o.orderdate " & _
",到貨日期 = t2.arrive_date,簽單日期 = '',POD天數 = '',貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS單號 = rtrim(t2.receipt_no),訂單號碼 = rtrim(t2.extern),客戶名稱 = rtrim(t1m.short_name),縣市 = rtrim(t2m.description)" & _
",訂單板數 = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3),訂單箱數 = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),總個數 = sum(t3.order_qty) " & _
",預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end),入庫完成 = '',訂單重量 = sum(t3.order_qty*s.stdgrosswgt) " & _
",訂單材積 = sum(t3.order_qty*s.stdcube),訂單備註 = rtrim(t2.description),預估到貨 = isnull(convert(char(20),t2.scheduledate,120),''),達交 = ' ',遲交 = ' ' " & _
",訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From TRP01T t01t  (nolock) join trp02t t2 (nolock) on t2.Route_No = t01t.Route_No and left(t01t.route_no,1) = 'F' " & _
"join trp03t t3 (nolock) on t3.receipt_no = t2.receipt_no join orders o (nolock) on o.orderkey = t2.c_receipt_no " & _
"join gv_skuxpack s (nolock) on s.storerkey = t2.storerkey and s.sku = t3.product_no " & _
"join trp01m t1m (nolock) on t1m.storerkey = t2.storerkey and t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end " & _
"join trp16m t16 (nolock) on t16.storerkey = t2.storerkey join TRP09M t09m on t09m.Vehicle_ID_No = isnull(t01t.C_Vehicle_ID_No,t2.Vehicle_ID_No) " & _
"join trp02m t2m (nolock) on t2m.zip = t1m.zip join TRP05T t05t (nolock) on t05t.Route_No = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO) and isnull(t05t.sdnstatus,'0') = '0' " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),isnull(convert(char(20),t2.scheduledate,120),''),t2.priority,isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No),t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,t09m.DRIVER,o.adddate,o.addwho,t2.VEHICLE_ID_NO,t2.ROUTE_NO "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'已排
str_SQL = "insert into ##" & strViewName & " select 單別 = RTrim(t2.Priority),狀態 = '已排',區碼 = rtrim(t1m.area_code) " & _
",二次車號 = isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No) ,二次駕駛人 = Rtrim(Isnull(t09m.Driver,'')),二次路編 = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),一次車號 = rtrim(isnull((select top 1 t9.VEHICLE_ID_NO  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),駕駛人 = RTRIM(ISNULL((select top 1 t9.DRIVER  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),一次路編 = RTRIM(ISNULL(t2.route_no,'')),訂單日期 = o.orderdate,到貨日期 = t2.arrive_date,簽單日期 = '',POD天數 = '' " & _
",貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS單號 = rtrim(t2.receipt_no),訂單號碼 = rtrim(t2.extern),客戶名稱 = rtrim(t1m.short_name) " & _
",縣市 = rtrim(t2m.description),訂單板數 = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3) " & _
",訂單箱數 = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),總個數 = sum(t3.order_qty) " & _
",預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end) " & _
",入庫完成 = '',訂單重量 = sum(t3.order_qty*s.stdgrosswgt),訂單材積 = sum(t3.order_qty*s.stdcube),訂單備註 = rtrim(t2.description),預估到貨 = '',達交 = ' ' " & _
",遲交 = ' ',訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From ort01T t01t (nolock) join ort02t t2 (nolock) on t2.Route_No = t01t.Route_No and left(t01t.route_no,1) = 'R' join ort03t t3 (nolock) on t3.receipt_no = t2.receipt_no " & _
"join orders o (nolock) on o.orderkey = t2.c_receipt_no join gv_skuxpack s on s.storerkey = t2.storerkey and s.sku = t3.product_no " & _
"join trp01m t1m (nolock) on t1m.storerkey = t2.storerkey and t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end join trp16m t16 on t16.storerkey = t2.storerkey " & _
"join TRP09M t09m (nolock) on t09m.Vehicle_ID_No = isnull(t01t.C_Vehicle_ID_No,t2.Vehicle_ID_No) join trp02m t2m on t2m.zip = t1m.zip " & _
"join ort05T t05t (nolock) on t05t.Route_No = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO) and isnull(t05t.sdnstatus,'0') = '0' " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),t2.priority,isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No),t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,t09m.DRIVER,o.adddate,o.addwho,t2.VEHICLE_ID_NO,t2.ROUTE_NO "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'待重組
str_SQL = "insert into ##" & strViewName & " select 單別 = RTrim(o.Priority),狀態 = '待重組',區碼 = rtrim(t1m.area_code),二次車號 = '',二次駕駛人 = '',二次路編 = '          ',一次車號 = '',一次駕駛人 = '',一次路編 = '          ' " & _
",訂單日期 = o.orderdate,到貨日期 = cast(s2.arrive_date as datetime),簽單日期 = '',POD天數 = '',貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS單號 = rtrim(s2.receipt_no),訂單號碼 = rtrim(s2.extern),客戶名稱 = rtrim(t1m.short_name),縣市 = rtrim(t2m.description) " & _
",訂單板數 = round( sum(case when isnull(s.pallet,0) = 0 then 0 else s3.order_qty /s.pallet end) ,3),訂單箱數 = sum(case when isnull(s.casecnt,0) = 0 then 0 else floor(s3.order_qty /s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then s3.order_qty else cast(s3.order_qty as int) % cast(s.casecnt as int) end),總個數 = sum(s3.order_qty) " & _
",預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(s3.order_qty /s.casecnt) end),入庫完成 = '' " & _
",訂單重量 = round( sum(s3.order_qty * s.stdgrosswgt),3),訂單材積 = round( sum( s3.order_qty * s.stdcube),3),訂單備註 = cast(o.notes as varchar(1000)) " & _
",預估到貨 = '',達交 = ' ',遲交 = ' ',訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From sdn02w s2 (nolock) join sdn03w s3 (nolock) on s3.receipt_no = s2.receipt_no join gv_skuxpack s on s.storerkey = s3.storerkey and s.sku = s3.product_no " & _
"join orders o (nolock) on o.orderkey = s2.c_receipt_no " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end and t1m.storerkey = o.storerkey join trp16m t16 on t16.storerkey = o.storerkey " & _
"left join trp02m t2m(nolock) on t2m.zip = t1m.zip " & _
"where convert(char(8),s2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),o.priority,cast(o.notes as varchar(1000)),s2.receipt_no,t16.storerkey,t16.short_name,t1m.area_code,s2.arrive_date,s2.extern,t1m.short_name,t2m.DESCRIPTION,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'出車
str_SQL = "set nocount on insert into ##" & strViewName & " Select 單別 = RTrim(s2.Priority),狀態 = '出車' + '-' + case when s2.sdnback = 0 then '簽單未回' else rtrim(s2.confirm_notes) end " & _
",區碼 = rtrim(t1m.area_code),二次車號 = rtrim(s1.c_Vehicle_ID_No),二次駕駛人 = Rtrim(Isnull(s1.Driver,'')),二次路編 = rtrim(s1.c_route_No) ,一次車號 = rtrim(isnull(s2.VEHICLE_ID_NO,'')),駕駛人 = RTRIM(ISNULL(t9.driver,'')),一次路編 = RTRIM(ISNULL(s2.route_no,'')) " & _
",訂單日期 = o.orderdate,到貨日期 = cast(s2.arrive_date as datetime),簽單日期 = isnull(convert(char(10),s2.sdnsenddate,20),'') " & _
",POD天數 = case when rtrim(s2.confirm_notes) = '' then '' when s2.sdnsenddate is null then '' else cast(datediff(dd,cast(s2.arrive_date as datetime),s2.sdnsenddate+1) as varchar(4)) end " & _
",貨主 = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS單號 = rtrim(s2.receipt_no),訂單號碼 = rtrim(s2.extern),客戶名稱 = rtrim(t1m.short_name) " & _
",縣市 = rtrim(t2m.description),訂單板數 = round( sum(case when isnull(s.pallet,0) = 0 then 0 else s3.order_qty /s.pallet end) ,3),訂單箱數 = sum(case when isnull(s.casecnt,0) = 0 then 0 else floor(s3.order_qty /s.casecnt) end) " & _
",訂單個數 = sum(case when isnull(s.casecnt,0) = 0 then s3.order_qty else cast(s3.order_qty as int) % cast(s.casecnt as int) end),總個數 = sum(s3.order_qty),預估件數 = sum(case when s.casecnt = 0 then 1 else ceiling(s3.order_qty /s.casecnt) end) " & _
",入庫完成 = isnull(s2.invback,''),訂單重量 = round( sum(s3.order_qty * s.stdgrosswgt),3),訂單材積 = round( sum(s3.order_qty * s.stdcube),3),訂單備註 = rtrim(s2.description) " & _
",預估到貨 = isnull(convert(char(20),isnull(s2.scheduledate,s2.custsigndate),120),''),達交 = case when s2.ontimedelivery = 9 then 'V' else ' ' end,遲交 = case when s2.ontimedelivery = 5 then 'V' else ' ' end " & _
",訂單來源 = isnull(o.updatesource,''),訂單新增時間 = o.adddate,訂單新增人員 = rtrim(o.addwho),託運單號 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From sdn02t s2 (nolock) join sdn01T s1(nolock) on s1.c_route_no = s2.c_route_no " & _
"join sdn03t s3 (nolock) on s3.receipt_no = s2.receipt_no  join orders o (nolock) on o.orderkey = s2.c_receipt_no join gv_skuxpack s on s.storerkey = s3.storerkey and s.sku = s3.product_no join trp01m t1m on t1m.consigneekey =case when o.priority ='A2B' then o.b_company else s2.consigneekey end and t1m.storerkey = s2.storerkey " & _
"join trp09m t9 (nolock) on t9.Vehicle_ID_No = s2.Vehicle_ID_No left join trp16m t16 on t16.storerkey = s2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where s2.arrive_date between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),s2.sdnback,o.orderdate,isnull(o.updatesource,''),isnull(convert(char(20),isnull(s2.scheduledate,s2.custsigndate),120),''),s2.confirm_notes,s2.OnTimeDelivery,s2.PRIORITY,t2m.DESCRIPTION,s2.receipt_no,t16.storerkey,t16.short_name,t1m.area_code,s1.c_Vehicle_ID_No,s1.Driver, s1.c_route_No,s2.arrive_date,s2.extern,t1m.short_name,s2.description,isnull(s2.invback,''),s2.sdnsenddate,o.adddate,o.addwho,rtrim(isnull(s2.VEHICLE_ID_NO,'')),RTRIM(ISNULL(t9.driver,'')),RTRIM(ISNULL(s2.route_no,'')) set nocount off"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "select * from ##" & strViewName & " Where 1 = 1 "
'str_SQL = "select * from gv_TRPTrack Where 1 = 1 "

'取貨主
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) Then strSelected = strSelected & "'" & List2.List(i) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and 貨主 in ( " & strSelected & "'') "

'取區碼
strSelected = ""
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then strSelected = strSelected & "'" & Trim(List1.List(i)) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and 區碼 in ( " & strSelected & "'') "

'取達交
strSelected = ""
For i = 0 To List3.ListCount - 1
    If List3.Selected(i) And Trim(List3.List(i)) = "未達" Then strSelected = strSelected & "(rtrim(遲交) = '' and rtrim(達交) = '') or "
    If List3.Selected(i) And Trim(List3.List(i)) = "遲交" Then strSelected = strSelected & "rtrim(遲交) = 'V' or "
    If List3.Selected(i) And Trim(List3.List(i)) = "達交" Then strSelected = strSelected & "rtrim(達交) = 'V' or "
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and (" & strSelected & " 1 = 0) "

'取單別
strSelected = ""
For i = 0 To List4.ListCount - 1
    If List4.Selected(i) Then strSelected = strSelected & "'" & mySplit(Trim(List4.List(i)), "_", 0) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and 單別 in ( " & strSelected & "'') "

'取狀態
strSelected = ""
For i = 0 To List5.ListCount - 1
    If List5.Selected(i) Then strSelected = strSelected & "'" & Trim(List5.List(i)) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and 狀態 in ( " & strSelected & "'') "

'路線編號
chc_Route = ""
If Len(txtRouteS.Text) > 0 And Len(txtRouteE.Text) > 0 Then
   chc_Route = "and 二次路編 between '" & txtRouteS.Text & "' and '" & txtRouteE.Text & "' "
ElseIf Len(txtRouteS.Text) > 0 And Len(txtRouteE.Text) = 0 Then
   chc_Route = "and 二次路編 = '" & txtRouteS.Text & "' "
ElseIf Len(txtRouteS.Text) = 0 And Len(txtRouteE.Text) > 0 Then
   chc_Route = "and 二次路編 = '" & txtRouteE.Text & "' "
End If

'到貨日期
chc_DeliveryDate = ""
chc_DeliveryDate1 = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(char(8),到貨日期,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
   chc_DeliveryDate1 = "and convert(char(8),o.deliverydate,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_DeliveryDate = "and convert(char(8),到貨日期,112) = '" & txtDeliveryDateS.Text & "' "
   chc_DeliveryDate1 = "and convert(char(8),o.deliverydate,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(char(8),到貨日期,112) = '" & txtDeliveryDateE.Text & "' "
   chc_DeliveryDate1 = "and convert(char(8),o.deliverydate,112) = '" & txtDeliveryDateE.Text & "' "
End If

'出車日期
chc_DeliveryDate2 = ""
If Len(txtDeliveryS.Text) > 0 And Len(txtDeliveryE.Text) > 0 Then
   chc_DeliveryDate2 = "and '20' + substring(二次路編,2,6) between '" & txtDeliveryS.Text & "' and '" & txtDeliveryE.Text & "' "
ElseIf Len(txtDeliveryS.Text) > 0 And Len(txtDeliveryE.Text) = 0 Then
   chc_DeliveryDate2 = "and '20' + substring(二次路編,2,6) = '" & txtDeliveryS.Text & "' "
ElseIf Len(txtDeliveryS.Text) = 0 And Len(txtDeliveryE.Text) > 0 Then
   chc_DeliveryDate2 = "and '20' + substring(二次路編,2,6) = '" & txtDeliveryE.Text & "' "
End If

'組合字串
str_SQL = str_SQL & chc_Route & chc_DeliveryDate & chc_DeliveryDate2 & "order by 到貨日期,區碼,二次車號,二次路編,貨主,訂單號碼 "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
cn.CommandTimeout = 600
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close

rsMain.MoveFirst

If chkShowWH = 1 Then

    '取配置資料
    Dim rsTmp As New ADODB.Recordset
    '    str_SQL = "select distinct sectionkey ,o.updatesource from " & strWMSDB & "..orders o join " & strWMSDB & "..pickdetail p on p.orderkey = o.orderkey and convert(char(8),o.deliverydate,112) > convert(char(8),getdate()-7,112) join " & strWMSDB & "..loc l on l.loc = p.loc order by sectionkey "
    str_SQL = "select distinct sectionkey ,o.updatesource from " & strWMSDB & "..orders o join " & strWMSDB & "..pickdetail p on p.orderkey = o.orderkey " & chc_DeliveryDate1 & " join " & strWMSDB & "..loc l on l.loc = p.loc order by sectionkey "
    tmp_Rs.Open str_SQL, cn
    Call Replication_Recordset(tmp_Rs, rsTmp)
    tmp_Rs.Close
    
    Do While Not rsMain.EOF
    
        If rsMain("到貨日期") > Format(Now() - 8, "yyyymmdd") Then
    
            rsTmp.Filter = "(updatesource = '" & rsMain("TMS單號") & "')"
        
            strSectionKey = ""
        
            If rsTmp.EOF Then
                rsMain("裝載點") = "未配置"
            Else
                Do While Not rsTmp.EOF
                    If UCase(RTrim(rsTmp("sectionkey"))) <> "FACILITY" Then strSectionKey = strSectionKey & RTrim(rsTmp("sectionkey")) & ";"
                    rsTmp.MoveNext
                Loop
                rsMain("裝載點") = strSectionKey
            End If
        
            If rsMain("單別") = "R" Or rsMain("單別") = "RC" Or rsMain("單別") = "A2B" Then rsMain("裝載點") = ""
        
            rsTmp.Filter = ""
        End If
        rsMain.MoveNext
    Loop
    rsTmp.Close
    
End If

Set dgMain.DataSource = rsMain

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True
cn.Execute "if object_id ('tempdb..##" & strViewName & "') is not null drop table ##" & strViewName, RowsAffect, adExecuteNoRecords

Exit Sub
err_Handle:
cn.Execute "if object_id ('tempdb..##" & strViewName & "') is not null drop table ##" & strViewName, RowsAffect, adExecuteNoRecords
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub



Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
If dg.Col > 0 Then If rsMain.Fields(dg.Col).Name = "預估到貨" Then dg.Columns(ColIndex).Width = dtpDeliveryTime.Width

'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
dtpDeliveryTime.Visible = False
If dgMain.DataSource Is Nothing Then Exit Sub
If rsMain.RecordCount = 0 Then Exit Sub
If rsMain.EOF Then Exit Sub
If dgMain.Col = -1 Then Exit Sub
If Left(rsMain("狀態"), 2) <> "出車" Then Exit Sub

With dgMain

'到貨時間
If rsMain.Fields(.Col).Name = "預估到貨" Then

    dtpDeliveryTime.Visible = True
    dtpDeliveryTime.Move .Left + .Columns(.Col).Left + Frame2.Left + 15, .Top + .RowTop(.Row) + Frame2.Top, .Columns(.Col).Width
    
    If dtpDeliveryTime.Left + dtpDeliveryTime.Width - Frame2.Left > .Left + .Width Then '如果欄位超出DataGrid的顯示範圍的處理
        dtpDeliveryTime.Width = dtpDeliveryTime.Width + .Left + .Width - dtpDeliveryTime.Left - dtpDeliveryTime.Width
    End If
    dtpDeliveryTime.Value = IIf(RTrim(rsMain("預估到貨")) = "", Now, rsMain("預估到貨"))

Else
    dtpDeliveryTime.Visible = False
End If

'達交
If rsMain.Fields(.Col).Name = "達交" Then
    If Trim(rsMain("達交")) = "" And rsMain("遲交") = "V" Then
        If MsgBox("達交確認?", vbOKCancel, "狀態變更") = vbOK Then rsMain("達交") = "V": rsMain("遲交") = " ": cn.Execute "update sdn02t set ontimedelivery = 9 where receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    ElseIf Trim(rsMain("達交")) = "V" Then
        If MsgBox("達交取消確認?", vbOKCancel, "狀態變更") = vbOK Then rsMain("達交") = " ": cn.Execute "update sdn02t set ontimedelivery = 0 where receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    Else
        If MsgBox("達交確認?", vbOKCancel, "狀態變更") = vbOK Then rsMain("達交") = "V": cn.Execute "update sdn02t set ontimedelivery = 9 where receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    End If
    .Col = 18
End If

'遲交
If rsMain.Fields(.Col).Name = "遲交" Then
    If Trim(rsMain("遲交")) = "" And rsMain("達交") = "V" Then
        If MsgBox("遲交確認?", vbOKCancel, "狀態變更") = vbOK Then rsMain("遲交") = "V": rsMain("達交") = " ": cn.Execute "update sdn02t set ontimedelivery = 5 where receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    ElseIf Trim(rsMain("遲交")) = "V" Then
        If MsgBox("遲交取消確認?", vbOKCancel, "狀態變更") = vbOK Then rsMain("遲交") = " ": cn.Execute "update sdn02t set ontimedelivery = 0 where receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    Else
        If MsgBox("遲交確認?", vbOKCancel, "狀態變更") = vbOK Then rsMain("遲交") = "V": cn.Execute "update sdn02t set ontimedelivery = 5 where receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    End If
    .Col = 18
End If

End With

End Sub

Private Sub dtpDeliveryTime_LostFocus()
     If MsgBox("預估到貨時間變更?", vbOKCancel, "確認") = vbOK Then
        rsMain("預估到貨") = Format(dtpDeliveryTime, "yyyy-mm-dd HH:MM")
        cn.Execute "update sdn02t set scheduledate = '" & rsMain("預估到貨") & "' where receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
        dtpDeliveryTime.Visible = False
     End If
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    dgMain.Height = Frame2.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth - 120
    dgMain.Width = Frame2.Width - 240
End If

End Sub

Private Sub cmdReset_Click()
Dim i As Integer

'重設
Call ClearForm_AllField(Me)

txtDeliveryDateS = Format(Now() - 7, "YYYYMMDD")
txtDeliveryDateE = Format(Now(), "YYYYMMDD")

End Sub

Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)

If dgMain.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMain.Sort = dgMain.Columns(ColIndex).Caption & " DESC"
    dgMain.ClearSelCols
    intColumnIndex = 255

Else
    rsMain.Sort = dgMain.Columns(ColIndex).Caption
    dgMain.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgmain_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub cmdExit_Click()
Unload Me '結束此程序
'End 結束應用程式
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 筆資料列"
StatusBar.Panels(3).Text = User_id

'貨主
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
'tmp_rs.Open "select distinct(貨主)  from gv_TRPTrack order by 貨主 ", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.Open "select distinct rtrim(storerkey) + '_' + rtrim(short_name) as 貨主 from trp16m order by rtrim(storerkey) + '_' + rtrim(short_name)", cn, adOpenKeyset, adLockPessimistic


If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        List2.AddItem RTrim(tmp_Rs("貨主"))
        tmp_Rs.MoveNext
    Next
    tmp_Rs.Close: Set tmp_Rs = Nothing
End If
    
'區域
Dim rsTmp As New ADODB.Recordset
With rsTmp
    .CursorLocation = 3
    .Open "select area_code from trp03m order by area_code ", cn
    
If Not rsTmp.EOF Then
    .MoveFirst
    For i = 0 To .RecordCount - 1
        List1.AddItem rsTmp("area_code")
        .MoveNext
    Next
    .Close: Set rsTmp = Nothing
End If

End With

'達交
List3.AddItem "未達"
List3.AddItem "達交"
List3.AddItem "遲交"

'單別
List4.AddItem "C_越庫訂單"
List4.AddItem "I_正常訂單"
List4.AddItem "A_轉倉"
List4.AddItem "R_退貨訂單"
List4.AddItem "RC_提貨入庫"
List4.AddItem "A2B_提貨配送"

'狀態
List5.AddItem "未轉入" ': List5.Selected(0) = True
List5.AddItem "未排" ': List5.Selected(1) = True
List5.AddItem "保留" ': List5.Selected(2) = True
List5.AddItem "已排" ': List5.Selected(3) = True
List5.AddItem "待重組" ': List5.Selected(4) = True
List5.AddItem "出車-簽單未回" ': List5.Selected(5) = True
List5.AddItem "出車-正常訂單"
List5.AddItem "出車-異常訂單"
List5.AddItem "出車-未出訂單"

txtDeliveryDateS = Format(Now() - 7, "YYYYMMDD")
txtDeliveryDateE = Format(Now(), "YYYYMMDD")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub
Private Sub txtDeliveryS_Click()
Set objMvdateTarget = txtDeliveryS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryE_Click()
Set objMvdateTarget = txtDeliveryE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryS_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateS_Click()
Set objMvdateTarget = txtDeliveryDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryDateE_Click()
Set objMvdateTarget = txtDeliveryDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
