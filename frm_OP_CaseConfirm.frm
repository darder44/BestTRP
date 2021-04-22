VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frm_OP_CaseConfirm 
   Caption         =   "出貨件數確認"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
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
   ScaleHeight     =   7020
   ScaleWidth      =   8370
   Visible         =   0   'False
   WindowState     =   2  '最大化
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3720
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
      StartOfWeek     =   228130817
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
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
      Height          =   2895
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8295
      Begin VB.CheckBox chkScan 
         Caption         =   "掃描模式"
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
         Left            =   3000
         TabIndex        =   36
         Top             =   600
         Width           =   1425
      End
      Begin VB.CheckBox chk_PrintAddress 
         BackColor       =   &H80000004&
         Caption         =   "8x4地址條(cm)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   2520
         Width           =   2010
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1065
         TabIndex        =   30
         Top             =   2145
         Width           =   3015
         Begin VB.OptionButton optPrintAll 
            Caption         =   "全部"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   33
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optPrintNO 
            Caption         =   "未列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optPrintYes 
            Caption         =   "已列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   31
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdHCT 
         BackColor       =   &H00FFFF00&
         Caption         =   "轉HCT"
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
         Left            =   120
         Style           =   1  '圖片外觀
         TabIndex        =   29
         Top             =   2040
         Width           =   945
      End
      Begin VB.CommandButton cmdPrintReport 
         BackColor       =   &H00C0FFC0&
         Caption         =   "地址標籤"
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
         Left            =   5880
         Picture         =   "frm_OP_CaseConfirm.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   1320
         Width           =   1065
      End
      Begin VB.ComboBox cboCar 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1485
      End
      Begin VB.ComboBox cboStorerkey 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1485
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0E0FF&
         Caption         =   "件數確認"
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
         Left            =   4680
         Picture         =   "frm_OP_CaseConfirm.frx":030A
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1065
         TabIndex        =   24
         Top             =   1740
         Width           =   3375
         Begin VB.OptionButton optYes 
            Caption         =   "已確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   7
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optNo 
            Caption         =   "未確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "全部"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   0
            Width           =   735
         End
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
         Height          =   990
         Left            =   5880
         Picture         =   "frm_OP_CaseConfirm.frx":2004
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '置中對齊
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
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   2
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
         Alignment       =   2  '置中對齊
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
         MaxLength       =   8
         TabIndex        =   3
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyS 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1320
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyE 
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
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1320
         Width           =   1485
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
         Height          =   990
         Left            =   7080
         Picture         =   "frm_OP_CaseConfirm.frx":32FE
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   1320
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
         Height          =   990
         Left            =   7080
         Picture         =   "frm_OP_CaseConfirm.frx":2CF10
         Style           =   1  '圖片外觀
         TabIndex        =   14
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
         Height          =   990
         Left            =   4680
         Picture         =   "frm_OP_CaseConfirm.frx":2D222
         Style           =   1  '圖片外觀
         TabIndex        =   10
         ToolTipText     =   "到貨日期180天內"
         Top             =   240
         Width           =   1065
      End
      Begin VB.CheckBox chkPrintPreView 
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
         Left            =   3000
         TabIndex        =   34
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton cmdOTUpdate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "毛寶件數批次確認"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3240
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "配送車號"
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
         Left            =   120
         TabIndex        =   28
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "貨主編號"
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
         TabIndex        =   27
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "確認狀態"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1725
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
         Index           =   23
         Left            =   2655
         TabIndex        =   23
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
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
         Index           =   22
         Left            =   120
         TabIndex        =   22
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
         Index           =   3
         Left            =   2655
         TabIndex        =   21
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "訂單號碼"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1365
         Width           =   960
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   7920
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   0
      TabIndex        =   17
      Top             =   2880
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   390
      Left            =   0
      TabIndex        =   26
      Top             =   6630
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   688
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
            Object.Width           =   8123
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
Attribute VB_Name = "frm_OP_CaseConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Public Sub chkScan_Click()
    
If chkScan.Value = 1 Then
    cboStorerKey.Enabled = False
    cboCar.Enabled = False
    txtDeliveryDateS.Enabled = False
    txtDeliveryDateE.Enabled = False
    txtExternOrderkeyE.Enabled = False
'    Frame4.Enabled = False--mark by Gemini 4 @20200117 RQ_2020020403
'    Frame3.Enabled = False--mark by Gemini 4 @20200117 RQ_2020020403
'    txtExternOrderkeyS.SetFocus
    txtExternOrderkeyS.SelStart = 0: txtExternOrderkeyS.SelLength = Len(txtExternOrderkeyS)
'    optAll.Value = 1
'    optPrintAll.Value = 1
Else
    cboStorerKey.Enabled = 1
    cboCar.Enabled = 1
    txtDeliveryDateS.Enabled = 1
    txtDeliveryDateE.Enabled = 1
    txtExternOrderkeyE.Enabled = 1
'    Frame4.Enabled = 1--mark by Gemini 4 @20200117 RQ_2020020403
'    Frame3.Enabled = 1--mark by Gemini 4 @20200117 RQ_2020020403
End If

End Sub

Private Sub cmd2Excel_Click()

If dgMain.DataSource Is Nothing Then Exit Sub

Recordset2Excel Me.Caption, rsMain
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmdHCT_Click()

If rsMain Is Nothing Then Exit Sub
If rsMain.EOF Then Exit Sub

'取已選取資料
rsMain.Filter = "(＊ = 'V')"
If rsMain.RecordCount = 0 Then rsMain.Filter = 0: MsgBox "請選取欲轉出的資料。", 64, "轉HCT": rsMain.Sort = "編號": Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11: cmdHCT.Enabled = False: dgMain.Enabled = False

Dim i As Integer, strFileName1 As String, strCheck As String

'轉文字檔
If Dir("C:\HCT\BEST2HCT", vbDirectory) = "" Then MkDirs "C:\HCT\BEST2HCT"
strFileName1 = "HCT" & Format(Now, "yyyymmddhhMMss") & ".csv"

Open "C:\HCT\BEST2HCT\" & strFileName1 For Output As #1

'交易開始
Tran_Level = cn.BeginTrans

Print #1, "客戶名稱"; ","; "聯絡人_內附文件"; ","; "聯絡電話1"; ","; "貨主單號"; ","; "客戶地址_貨主單號"; ","; "出貨件數"; ","; "訂單備註"

rsMain.MoveFirst

Do While Not rsMain.EOF
    Print #1, myExCharFilter(rsMain("客戶名稱")); ","; myExCharFilter(rsMain("聯絡人")); ","; myExCharFilter(rsMain("聯絡電話1")); ","; myExCharFilter(rsMain("訂單號碼")); ","; myExCharFilter(RTrim(rsMain("客戶地址"))); ","; rsMain("出貨件數"); ","; myExCharFilter(rsMain("訂單備註")) & "_" & myExCharFilter(rsMain("訂單號碼"))
    
'更新為已回傳
str_SQL = "update trp02t " & _
            "set otprinttimes = otprinttimes + 1 " & _
            ", otprintdate = getdate() " & _
            "where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "update ort02t " & _
            "set otprinttimes = otprinttimes + 1 " & _
            ", otprintdate = getdate() " & _
            "where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

rsMain("列印次數") = rsMain("列印次數") + 1
rsMain("列印時間") = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
rsMain.MoveNext
Loop

'關閉檔案
Close

cn.CommitTrans: Tran_Level = 0

Screen.MousePointer = 0: cmdHCT.Enabled = True: dgMain.Enabled = True
MsgBox "資料轉出完成!!" & vbCrLf & "檔案儲存 C:\HCT\BEST2HCT\" & strFileName1, vbOKOnly, Me.Caption

rsMain.Filter = 0: rsMain.Sort = "編號"

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdOK_Click()

If rsMain Is Nothing Then Exit Sub
strOtQtyFixOrderkey = rsMain("TMS單號")
frm_OTQtyFix.Show vbModal

'更新Datagrid
Call UpdateDatagrid

End Sub

Public Sub UpdateDatagrid()

'更新Datagrid
str_SQL = "select 異動日期 = t2.OTConfirmdate " & _
", 出貨件數 = isnull(t2.otqty,0) " & _
", 異動人員 = isnull(t2.otconfirmuser,'') " & _
"from trp02t t2 " & _
"where t2.Receipt_no = '" & rsMain("TMS單號") & "' " & _
"union select 異動日期 = t2.OTConfirmdate " & _
", 出貨件數 = isnull(t2.otqty,0) " & _
", 異動人員 = isnull(t2.otconfirmuser,'') " & _
"from TRP02w t2 " & _
"where t2.Receipt_no = '" & rsMain("TMS單號") & "' " & _
"union select 異動日期 = t2.OTConfirmdate " & _
", 出貨件數 = isnull(t2.otqty,0) " & _
", 異動人員 = isnull(t2.otconfirmuser,'') " & _
"from ort02t t2 " & _
"where t2.Receipt_no = '" & rsMain("TMS單號") & "' " & _
"union select 異動日期 = t2.OTConfirmdate " & _
", 出貨件數 = isnull(t2.otqty,0) " & _
", 異動人員 = isnull(t2.otconfirmuser,'') " & _
"from ort02w t2 " & _
"where t2.Receipt_no = '" & rsMain("TMS單號") & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn

rsMain("異動日期") = Format(tmp_Rs("異動日期"), "yyyy-mm-dd hh:MM:ss") & ""
rsMain("出貨件數") = tmp_Rs("出貨件數")
rsMain("異動人員") = tmp_Rs("異動人員") & ""
rsMain.Update

End Sub

Private Sub cmdOTUpdate_Click()

On Error GoTo err_Handle

Dim strOrderPath As String, strFileName As String, i As Long, j As Long, LngOTQty As Long
Dim rs As New ADODB.Recordset

With dlgCommonDialog
    .FileName = ""
    .DialogTitle = "選取件數維護資料"
    .CancelError = False
    .InitDir = "C:\"
    'ToDo: 設定通用對話方塊控制項的旗標及屬性
    .Filter = "件數檔案 (*.xls)|*.xls"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    strFileName = .FileName
End With

Screen.MousePointer = 11

MsgBox "匯入時，請關閉Excel檔案!!", vbInformation + vbOKOnly, "件數批次確認"

Call Excel2Recordset(strFileName, "件數維護", "", rs)
j = rs.RecordCount
rs.MoveFirst
Do While Not rs.EOF
    str_SQL = "update trp02t set otqty = " & Val(rs("整箱件數")) + Val(rs("零散件數")) & ",otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where storerkey = 'LMBO01' and receipt_no ='" & myFilter(Trim(rs("TMS單號"))) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    If RowsAffect = 1 Then i = i + RowsAffect: LngOTQty = LngOTQty + Val(rs("零散件數"))

    
'    str_SQL = "update trp02w set otqty = " & Val(rs("整箱件數")) + Val(rs("零散件數")) & ",otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where receipt_no ='" & myFilter(Trim(rs("TMS單號"))) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'    i = i + RowsAffect
    
'    str_SQL = "update ort02t set otqty = " & Val(rs("整箱件數")) + Val(rs("零散件數")) & ",otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where receipt_no ='" & myFilter(Trim(rs("TMS單號"))) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'    i = i + RowsAffect
    
'    str_SQL = "update ort02w set otqty = " & Val(rs("整箱件數")) + Val(rs("零散件數")) & ",otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where receipt_no ='" & myFilter(Trim(rs("TMS單號"))) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'    i = i + RowsAffect
        
    rs.MoveNext
Loop

Set rs = Nothing

MsgBox "共更新 " & i & "筆資料!總零散件數: " & LngOTQty & " 件", 64, "件數維護完成"
Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Public Sub cmdPrintReport_Click()
Dim i As Integer, j As Integer, k As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub

'取已選取資料
If RTrim(strOtQtyFixOrderkey) <> "" Then
    rsMain.Filter = "(TMS單號 = " & strOtQtyFixOrderkey & ")"
Else
    rsMain.Filter = "(＊ = 'V')"
End If

If rsMain.RecordCount = 0 Then rsMain.Filter = 0: MsgBox "請選取欲列印之資料。", 64, "列印": rsMain.Sort = "編號": Exit Sub

Screen.MousePointer = 11
    Dim rs_Access As New ADODB.Recordset
    Dim MSAccessAP As New access.Application
    
'判斷4*8地址條則呼叫另外一個報表
If cboStorerKey = "LLFA01" Then
    
    '資料寫入 Access 資料庫
    Call AccessDB_Connect
    cnAccess.BeginTrans
    Tran_Level = cn.BeginTrans
        
    cnAccess.Execute "Delete From 出貨件數", RowsAffect, adExecuteNoRecords
    
    rs_Access.Open "出貨件數", cnAccess, adOpenStatic, adLockOptimistic
    
    rsMain.MoveFirst
    
    Do While Not rsMain.EOF
        For j = 1 To rsMain("出貨件數") '一件寫入一筆
            rs_Access.AddNew
            
            For i = 0 To rsMain.Fields.Count - 1 '寫入每個欄位
                rs_Access.Fields(i).Value = rsMain.Fields(i).Value
            Next i
            
            rs_Access.Fields(i).Value = j
            rs_Access.Fields(i + 1).Value = rsMain("出貨件數")
            rs_Access.Update
        Next j
        
        'TRP02T更新為已回傳
        str_SQL = "update trp02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'TRP02W更新為已回傳
        str_SQL = "update TRP02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'ORT02T更新為已回傳
        str_SQL = "update ort02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'ORT02W更新為已回傳
        str_SQL = "update ort02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
           
        rsMain("列印次數") = rsMain("列印次數") + 1
        rsMain("列印時間") = Format(Now, "yyyy/mm/dd hh:mm:ss")
        
       rsMain.MoveNext
    Loop
    
    cn.CommitTrans: Tran_Level = 0
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
    
    strAccessDBFileName_FullPath = GetAccessDBFileName

    With MSAccessAP
        .OpenCurrentDatabase (strAccessDBFileName_FullPath)
        
        If chkPrintPreView.Value = vbChecked Then
        '預覽列印
             .DoCmd.OpenReport "利豐地址標籤", acViewPreview
            .DoCmd.Maximize
            .Visible = True
        Else
        '直接列印至印表機
            .Visible = False
            .DoCmd.OpenReport "利豐地址標籤", acViewNormal
            .CloseCurrentDatabase
            .Quit
            Set MSAccessAP = Nothing
        End If
    
    End With
    
ElseIf chk_PrintAddress.Value = 0 Then
    
    '資料寫入 Access 資料庫
    Call AccessDB_Connect
    cnAccess.BeginTrans
    Tran_Level = cn.BeginTrans
        
    cnAccess.Execute "Delete From 出貨件數", RowsAffect, adExecuteNoRecords
    
    rs_Access.Open "出貨件數", cnAccess, adOpenStatic, adLockOptimistic
    
    rsMain.MoveFirst
    
    Do While Not rsMain.EOF
        For j = 1 To rsMain("出貨件數") '一件寫入一筆
            rs_Access.AddNew
            
            For i = 0 To rsMain.Fields.Count - 1 '寫入每個欄位
                rs_Access.Fields(i).Value = rsMain.Fields(i).Value
            Next i
            
            rs_Access.Fields(i).Value = j
            rs_Access.Fields(i + 1).Value = rsMain("出貨件數")
            rs_Access.Update
        Next j
        
        'TRP02T更新為已回傳
        str_SQL = "update trp02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'TRP02W更新為已回傳
        str_SQL = "update TRP02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'ORT02T更新為已回傳
        str_SQL = "update ort02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'ORT02W更新為已回傳
        str_SQL = "update ort02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
           
        rsMain("列印次數") = rsMain("列印次數") + 1
        rsMain("列印時間") = Format(Now, "yyyy/mm/dd hh:mm:ss")
        
       rsMain.MoveNext
    Loop
    
    cn.CommitTrans: Tran_Level = 0
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
    
    strAccessDBFileName_FullPath = GetAccessDBFileName

    With MSAccessAP
        .OpenCurrentDatabase (strAccessDBFileName_FullPath)
        
        If chkPrintPreView.Value = vbChecked Then
        '預覽列印
             .DoCmd.OpenReport "出貨件數", acViewPreview
            .DoCmd.Maximize
            .Visible = True
        Else
        '直接列印至印表機
            .Visible = False
            .DoCmd.OpenReport "出貨件數", acViewNormal
            .CloseCurrentDatabase
            .Quit
            Set MSAccessAP = Nothing
        End If
    
    End With
Else
    '資料寫入 Access 資料庫
    Call AccessDB_Connect
    cnAccess.BeginTrans
    Tran_Level = cn.BeginTrans
        
    cnAccess.Execute "Delete From 出貨件數", RowsAffect, adExecuteNoRecords
    

    rs_Access.Open "出貨件數", cnAccess, adOpenStatic, adLockOptimistic
    
    rsMain.MoveFirst
    
    Do While Not rsMain.EOF
        For j = 1 To rsMain("出貨件數") '一件寫入一筆
            rs_Access.AddNew
            
            For i = 0 To rsMain.Fields.Count - 1 '寫入每個欄位
                rs_Access.Fields(i).Value = rsMain.Fields(i).Value
            Next i
            
            rs_Access.Fields(i).Value = j
            rs_Access.Fields(i + 1).Value = rsMain("出貨件數")
            rs_Access.Update
        Next j
        
        'TRP02T更新為已回傳
        str_SQL = "update trp02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'TRP02W更新為已回傳
        str_SQL = "update TRP02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'ORT02T更新為已回傳
        str_SQL = "update ort02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        'ORT02W更新為已回傳
        str_SQL = "update ort02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
           
        rsMain("列印次數") = rsMain("列印次數") + 1
        rsMain("列印時間") = Format(Now, "yyyy/mm/dd hh:mm:ss")
        
       rsMain.MoveNext
    Loop
    
    cn.CommitTrans: Tran_Level = 0
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
    
    strAccessDBFileName_FullPath = GetAccessDBFileName

    With MSAccessAP
        .OpenCurrentDatabase (strAccessDBFileName_FullPath)
        
        If chkPrintPreView.Value = vbChecked Then
        '預覽列印
            .DoCmd.OpenReport "地址條8x4", acViewPreview
            .DoCmd.Maximize
            .Visible = True
        Else
        '直接列印至印表機
            .Visible = False
            .DoCmd.OpenReport "地址條8x4", acViewNormal
            .CloseCurrentDatabase
            .Quit
            Set MSAccessAP = Nothing
        End If
    
    End With

End If

rsMain.Filter = 0
rsMain.Sort = "編號"
Screen.MousePointer = 0
strOtQtyFixOrderkey = ""
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_DeliveryDate As String, chc_ExternOrderkey, chc_Status As String, chc_Storerkey As String, chc_Carno As String, chc_Print As String

'str_SQL = "select * from gv_OTQtyConfirm where 1 = 1 "

'到貨日期
chc_DeliveryDate = ""
If Len(RTrim(txtDeliveryDateS.Text)) > 0 And Len(RTrim(txtDeliveryDateE.Text)) > 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateS.Text)) > 0 And Len(RTrim(txtDeliveryDateE.Text)) = 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateS.Text)) = 0 And Len(RTrim(txtDeliveryDateE.Text)) > 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) = '" & txtDeliveryDateE.Text & "' "
End If

'貨主單號
chc_ExternOrderkey = ""
If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chc_ExternOrderkey = "and o.externorderkey between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
   chc_ExternOrderkey = "and o.externorderkey = '" & txtExternOrderkeyS.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chc_ExternOrderkey = "and o.externorderkey = '" & txtExternOrderkeyE.Text & "' "
End If

'件數狀態
chc_Status = ""
If optNo = True Then chc_Status = "and len(isnull(convert(char(20),t2.OTconfirmdate,120),'')) = 0 "
If optYes = True Then chc_Status = "and len(isnull(convert(char(20),t2.OTconfirmdate,120),'')) > 0 "

'列印狀態
chc_Print = ""
If optPrintNO = True Then chc_Print = "and t2.otprinttimes = 0 "
If optPrintYes = True Then chc_Print = "and t2.otprinttimes > 0 "

'貨主
chc_Storerkey = ""
If Len(RTrim(cboStorerKey.Text)) > 0 Then chc_Storerkey = " and o.storerkey = '" & RTrim(cboStorerKey.Text) & "' "

'車號
chc_Carno = ""
If Len(RTrim(cboCar.Text)) > 0 Then chc_Carno = " and isnull(t1.c_vehicle_id_no,t2.vehicle_id_no) = '" & RTrim(cboCar.Text) & "' "

'組合字串
'TRP02T
str_SQL = "select [＊] = ' ' " & _
",貨主編號 = rtrim(o.storerkey) " & _
",訂單號碼 = rtrim(o.externorderkey) " & _
",TMS單號 = t2.receipt_no " & _
",配送車號 = isnull(t1.c_vehicle_id_no,t2.vehicle_id_no) " & _
",到貨日期 = convert(char(8),o.deliveryDate,112) " & _
",客戶名稱 = rtrim(isnull(t1m.short_name,o.c_company)) " & _
",聯絡人 = rtrim(isnull(isnull(t1m.contact,o.c_contact1),'')) " & _
",聯絡電話1 = rtrim(isnull(isnull(t1m.phone,o.c_phone1),'')) " & _
",客戶地址 = isnull(isnull(t1m.address,rtrim(o.c_address1) + rtrim(o.c_address2)),'') " & _
",出貨件數 = isnull(t2.otqty,0) " & _
",出貨箱數 = sum(case when sp.casecnt = 0 then 0 else floor(t3.ship_qty/sp.casecnt) end) " & _
",出貨個數 = sum(case when sp.casecnt = 0 then t3.ship_qty else cast(t3.ship_qty as int)%cast(sp.casecnt as int) end) " & _
",訂單備註 = rtrim(t2.description) " & _
",客戶需求 = isnull(cast(t1m.notes as varchar(300)),'') " & _
",信速站所 = case when o.storerkey = 'LLFA01' then rtrim(isnull(t2m.area_code,'')) else rtrim(isnull(t2m.dcode,'')) end " & _
",代收貨款 = o.cash+o.bill,列印次數 = t2.otprinttimes,列印時間 = isnull(convert(char(20),t2.otprintdate,120),'') " & _
",異動日期 = isnull(convert(char(20),t2.OTconfirmdate,120),'') " & _
",異動人員 = isnull(t2.OTconfirmuser,'') " & _
"from trp02t t2 join orders o on o.orderkey = t2.c_receipt_no " & _
"join trp03t t3 on t3.receipt_no = t2.receipt_no " & _
"join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey " & _
"join trp01m t1m on t1m.storerkey = t2.storerkey and t1m.consigneekey = t2.consigneekey " & _
"left join trp02m t2m on t2m.zip = t1m.zip left join trp01t t1 on t1.route_no = t2.route_no where 1 = 1 "
str_SQL = str_SQL & chc_DeliveryDate & chc_ExternOrderkey & chc_Status & chc_Print & chc_Storerkey & chc_Carno & _
"group by t2m.area_code,t2.otqty,o.storerkey,o.externorderkey,t2.receipt_no,t1.c_vehicle_id_no,t2.vehicle_id_no,convert(char(8),o.deliveryDate,112),t1m.short_name,o.c_company,t1m.contact,o.c_contact1,t1m.phone,o.c_phone1,t1m.address,o.c_address1,o.c_address2,t2.description,isnull(cast(t1m.notes as varchar(300)),''),t2.otprinttimes,t2.otprintdate,t2.OTconfirmdate,t2.OTconfirmuser,t2m.dcode,o.cash,o.bill "

'TRP02W
str_SQL = str_SQL & "union select [＊] = ' ' " & _
",貨主編號 = rtrim(o.storerkey) " & _
",訂單號碼 = rtrim(o.externorderkey) " & _
",TMS單號 = t2.receipt_no " & _
",配送車號 = '未排車' " & _
",到貨日期 = convert(char(8),o.deliveryDate,112) " & _
",客戶名稱 = rtrim(isnull(t1m.short_name,o.c_company)) " & _
",聯絡人 = rtrim(isnull(isnull(t1m.contact,o.c_contact1),'')) " & _
",聯絡電話1 = rtrim(isnull(isnull(t1m.phone,o.c_phone1),'')) " & _
",客戶地址 = isnull(isnull(t1m.address,rtrim(o.c_address1) + rtrim(o.c_address2)),'') " & _
",出貨件數 = isnull(t2.otqty,0) " & _
",出貨箱數 = sum(case when sp.casecnt = 0 then 0 else floor(t3.order_qty/sp.casecnt) end) " & _
",出貨個數 = sum(case when sp.casecnt = 0 then t3.order_qty else cast(t3.order_qty as int)%cast(sp.casecnt as int) end) " & _
",訂單備註 = rtrim(t2.description) " & _
",客戶需求 = isnull(cast(t1m.notes as varchar(300)),'') " & _
",信速站所 = case when o.storerkey = 'LLFA01' then rtrim(isnull(t2m.area_code,'')) else rtrim(isnull(t2m.dcode,'')) end " & _
",代收貨款 = o.cash+o.bill,列印次數 = t2.otprinttimes,列印時間 = isnull(convert(char(20),t2.otprintdate,120),'') " & _
",異動日期 = isnull(convert(char(20),t2.OTconfirmdate,120),'') " & _
",異動人員 = isnull(t2.OTconfirmuser,'') " & _
"from TRP02w t2 join orders o on o.orderkey = t2.c_receipt_no " & _
"join TRP03w t3 on t3.receipt_no = t2.receipt_no " & _
"join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey " & _
"join trp01m t1m on t1m.storerkey = t2.storerkey and t1m.consigneekey = t2.consigneekey " & _
"left join trp02m t2m on t2m.zip = t1m.zip where 1 = 1 "
str_SQL = str_SQL & chc_DeliveryDate & chc_ExternOrderkey & chc_Status & chc_Print & chc_Storerkey & _
"group by t2m.area_code,t2.otqty,o.storerkey,o.externorderkey,t2.receipt_no,convert(char(8),o.deliveryDate,112),t1m.short_name,o.c_company,t1m.contact,o.c_contact1,t1m.phone,o.c_phone1,t1m.address,o.c_address1,o.c_address2,t2.description,isnull(cast(t1m.notes as varchar(300)),''),t2.otprinttimes,t2.otprintdate,t2.OTconfirmdate,t2.OTconfirmuser,t2m.dcode,o.cash,o.bill "

'ORT02T
str_SQL = str_SQL & "union select [＊] = ' ' " & _
",貨主編號 = rtrim(o.storerkey) " & _
",訂單號碼 = rtrim(o.externorderkey) " & _
",TMS單號 = t2.receipt_no " & _
",配送車號 = isnull(t1.c_vehicle_id_no,t2.vehicle_id_no) " & _
",到貨日期 = convert(char(8),o.deliveryDate,112) " & _
",客戶名稱 = rtrim(isnull(t1m.short_name,o.c_company)) " & _
",聯絡人 = rtrim(isnull(isnull(t1m.contact,o.c_contact1),'')) " & _
",聯絡電話1 = rtrim(isnull(isnull(t1m.phone,o.c_phone1),'')) " & _
",客戶地址 = isnull(isnull(t1m.address,rtrim(o.c_address1) + rtrim(o.c_address2)),'') " & _
",出貨件數 = isnull(t2.otqty,0) " & _
",出貨箱數 = sum(case when sp.casecnt = 0 then 0 else floor(t3.ship_qty/sp.casecnt) end) " & _
",出貨個數 = sum(case when sp.casecnt = 0 then t3.ship_qty else cast(t3.ship_qty as int)%cast(sp.casecnt as int) end) " & _
",訂單備註 = rtrim(t2.description) " & _
",客戶需求 = isnull(cast(t1m.notes as varchar(300)),'') " & _
",信速站所 = case when o.storerkey = 'LLFA01' then rtrim(isnull(t2m.area_code,'')) else rtrim(isnull(t2m.dcode,'')) end " & _
",代收貨款 = o.cash+o.bill,列印次數 = t2.otprinttimes,列印時間 = isnull(convert(char(20),t2.otprintdate,120),'') " & _
",異動日期 = isnull(convert(char(20),t2.OTconfirmdate,120),'') " & _
",異動人員 = isnull(t2.OTconfirmuser,'') " & _
"from ort02t t2 join orders o on o.orderkey = t2.c_receipt_no " & _
"join ort03t t3 on t3.receipt_no = t2.receipt_no " & _
"join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey " & _
"join trp01m t1m on t1m.storerkey = t2.storerkey and t1m.consigneekey = t2.consigneekey " & _
"left join trp02m t2m on t2m.zip = t1m.zip left join trp01t t1 on t1.route_no = t2.route_no where 1 = 1 "
str_SQL = str_SQL & chc_DeliveryDate & chc_ExternOrderkey & chc_Status & chc_Print & chc_Storerkey & chc_Carno & _
"group by t2m.area_code,t2.otqty,o.storerkey,o.externorderkey,t2.receipt_no,t1.c_vehicle_id_no,t2.vehicle_id_no,convert(char(8),o.deliveryDate,112),t1m.short_name,o.c_company,t1m.contact,o.c_contact1,t1m.phone,o.c_phone1,t1m.address,o.c_address1,o.c_address2,t2.description,isnull(cast(t1m.notes as varchar(300)),''),t2.otprinttimes,t2.otprintdate,t2.OTconfirmdate,t2.OTconfirmuser,t2m.dcode,o.cash,o.bill "

'ORT02W
str_SQL = str_SQL & "union select [＊] = ' ' " & _
",貨主編號 = rtrim(o.storerkey) " & _
",訂單號碼 = rtrim(o.externorderkey) " & _
",TMS單號 = t2.receipt_no " & _
",配送車號 = '未排車' " & _
",到貨日期 = convert(char(8),o.deliveryDate,112) " & _
",客戶名稱 = rtrim(isnull(t1m.short_name,o.c_company)) " & _
",聯絡人 = rtrim(isnull(isnull(t1m.contact,o.c_contact1),'')) " & _
",聯絡電話1 = rtrim(isnull(isnull(t1m.phone,o.c_phone1),'')) " & _
",客戶地址 = isnull(isnull(t1m.address,rtrim(o.c_address1) + rtrim(o.c_address2)),'') " & _
",出貨件數 = isnull(t2.otqty,0) " & _
",出貨箱數 = sum(case when sp.casecnt = 0 then 0 else floor(t3.order_qty/sp.casecnt) end) " & _
",出貨個數 = sum(case when sp.casecnt = 0 then t3.order_qty else cast(t3.order_qty as int)%cast(sp.casecnt as int) end) " & _
",訂單備註 = rtrim(t2.description) " & _
",客戶需求 = isnull(cast(t1m.notes as varchar(300)),'') " & _
",信速站所 = case when o.storerkey = 'LLFA01' then rtrim(isnull(t2m.area_code,'')) else rtrim(isnull(t2m.dcode,'')) end " & _
",代收貨款 = o.cash+o.bill,列印次數 = t2.otprinttimes,列印時間 = isnull(convert(char(20),t2.otprintdate,120),'') " & _
",異動日期 = isnull(convert(char(20),t2.OTconfirmdate,120),'') " & _
",異動人員 = isnull(t2.OTconfirmuser,'') " & _
"from ort02w t2 join orders o on o.orderkey = t2.c_receipt_no " & _
"join ort03w t3 on t3.receipt_no = t2.receipt_no " & _
"join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey " & _
"join trp01m t1m on t1m.storerkey = t2.storerkey and t1m.consigneekey = t2.consigneekey " & _
"left join trp02m t2m on t2m.zip = t1m.zip where 1 = 1 "
str_SQL = str_SQL & chc_DeliveryDate & chc_ExternOrderkey & chc_Status & chc_Print & chc_Storerkey & _
"group by t2m.area_code,t2.otqty,o.storerkey,o.externorderkey,t2.receipt_no,convert(char(8),o.deliveryDate,112),t1m.short_name,o.c_company,t1m.contact,o.c_contact1,t1m.phone,o.c_phone1,t1m.address,o.c_address1,o.c_address2,t2.description,isnull(cast(t1m.notes as varchar(300)),''),t2.otprinttimes,t2.otprintdate,t2.OTconfirmdate,t2.OTconfirmuser,t2m.dcode,o.cash,o.bill "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic

If tmp_Rs.EOF = True Then
    Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption
    If chkScan.Value = 1 Then txtExternOrderkeyS.SelStart = 0: txtExternOrderkeyS.SelLength = Len(txtExternOrderkeyS)
    Exit Sub
End If

tmp_Rs.Sort = "到貨日期,訂單號碼"
Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

'    .ColumnHeaders = True        '標題行顯示
'    .RowHeight = 300
'    .Columns(0).Alignment = dbgCenter
'    .Columns(10).Alignment = dbgRight

End With
SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True

If chkScan.Value = 1 Then Call cmdOK_Click

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain_DblClick()
Call cmdOK_Click
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

With dgMain

If .DataSource Is Nothing Then Exit Sub
'If LastRow = Empty Then Exit Sub
If .Row = -1 Or .Col <> 1 Then Exit Sub
On Error GoTo err_Handle

If .Col = 1 Then
    If UCase(dgMain) <> "V" And Val(rsMain("出貨件數")) > 0 Then '未選取與件數大於0
        dgMain = "V"
    Else
        dgMain = " "
    
    End If
.Col = 0
End If

End With
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    dgMain.Height = Frame2.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth
    dgMain.Width = Frame2.Width - 240
End If

End Sub

Private Sub cmdReset_Click()

'重設
Call ClearForm_AllField(Me)
optNo.Value = True
optPrintNO.Value = True

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

If KeyAscii = 13 Then Call cmdOK_Click

End Sub

Private Sub cmdExit_Click()
Unload Me '結束此程序
'End 結束應用程式
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle

optNo.Value = True
optPrintNO.Value = True
StatusBar.Panels(2).Text = "0 筆資料列"
StatusBar.Panels(3).Text = User_id

Dim i As Integer

'取車號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select distinct vehicle_id_no from trp09m order by vehicle_id_no "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then tmp_Rs.MoveFirst

Do While Not tmp_Rs.EOF
    cboCar.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    tmp_Rs.MoveNext
Loop
cboCar.AddItem "未排車"

cboCar = ""

tmp_Rs.Close

'貨主
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(storerkey) from trp16M", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.MoveFirst
For i = 0 To tmp_Rs.RecordCount - 1
    cboStorerKey.AddItem RTrim(tmp_Rs("storerkey"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close: Set tmp_Rs = Nothing
cboStorerKey.ListIndex = 0

txtDeliveryDateS = Format(Now + 1, "YYYYMMDD")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
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
'Private Sub txtOrderDateS_Click()
'Set objMvdateTarget = txtOrderDateS
'mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
'mvDate.Visible = True: mvDate.Value = Now
'
'End Sub
'Private Sub txtOrderDateE_Click()
'Set objMvdateTarget = txtOrderDateE
'mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
'mvDate.Visible = True: mvDate.Value = Now
'
'End Sub
'Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then mvDate.Visible = False
'End Sub
'Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then mvDate.Visible = False
'End Sub
Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub

Private Sub txtExternOrderkeys_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdQuery_Click
End Sub
