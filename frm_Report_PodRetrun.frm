VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Report_PodRetrun 
   Caption         =   "POD回傳"
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
   WindowState     =   2  '最大化
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   240
      TabIndex        =   15
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
      StartOfWeek     =   121831425
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
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmd_SendMail 
         BackColor       =   &H00C0C0FF&
         Caption         =   "轉Excel發送"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7320
         Picture         =   "frm_Report_PodRetrun.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_Msg 
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
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.TextBox txt_UnReciept 
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
         TabIndex        =   35
         ToolTipText     =   "有拒短收，尚未全收的訂單號碼"
         Top             =   1680
         Width           =   3285
      End
      Begin VB.CheckBox Check4 
         Caption         =   "D"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "C"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "B"
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "A"
         Height          =   255
         Left            =   4680
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox cb_all 
         Caption         =   "回傳全選"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   2040
         Value           =   1  '核取
         Width           =   1335
      End
      Begin VB.OptionButton optIn 
         Caption         =   "退貨單"
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
         Left            =   2160
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optOut 
         Caption         =   "出貨單"
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
         Left            =   1200
         TabIndex        =   28
         Top             =   2040
         Value           =   -1  'True
         Width           =   975
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
         TabIndex        =   24
         Top             =   360
         Value           =   1  '核取
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdOTUpdate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "毛寶批次POD回傳"
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
         Style           =   1  '圖片外觀
         TabIndex        =   7
         ToolTipText     =   "只針對毛寶"
         Top             =   1320
         Width           =   1065
      End
      Begin VB.ComboBox cboStorerkey 
         Enabled         =   0   'False
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
         Top             =   480
         Width           =   1485
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
         TabIndex        =   20
         Top             =   2340
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton optYes 
            Caption         =   "已回傳"
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
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optNo 
            Caption         =   "未回傳"
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
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
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
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.CommandButton cmd2Excel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "轉出回檔"
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
         Picture         =   "frm_Report_PodRetrun.frx":08CA
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
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
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
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
         Picture         =   "frm_Report_PodRetrun.frx":1BC4
         Style           =   1  '圖片外觀
         TabIndex        =   12
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
         Picture         =   "frm_Report_PodRetrun.frx":2B7D6
         Style           =   1  '圖片外觀
         TabIndex        =   11
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
         Picture         =   "frm_Report_PodRetrun.frx":2BAE8
         Style           =   1  '圖片外觀
         TabIndex        =   9
         ToolTipText     =   "到貨日期180天內"
         Top             =   240
         Width           =   1065
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
         TabIndex        =   23
         Top             =   540
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "未全收的"
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
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   1380
         Visible         =   0   'False
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
         TabIndex        =   16
         Top             =   1365
         Visible         =   0   'False
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
      TabIndex        =   13
      Top             =   2520
      Width           =   8295
      Begin TabDlg.SSTab SSTab1 
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6165
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Header"
         TabPicture(0)   =   "frm_Report_PodRetrun.frx":2BDF2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dgMain_Header"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "frm_Report_PodRetrun.frx":2BE0E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgMain_Detail"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgMain_Header 
            Height          =   2295
            Left            =   120
            TabIndex        =   26
            Top             =   480
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
         Begin MSDataGridLib.DataGrid dgMain_Detail 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   27
            Top             =   480
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
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   390
      Left            =   0
      TabIndex        =   22
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
            Object.Width           =   8149
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
Attribute VB_Name = "frm_Report_PodRetrun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blRouteT0Change As Boolean
Private rsMainHeader As ADODB.Recordset
Private rsMainDetail As ADODB.Recordset
Private rs_Receipt As ADODB.Recordset
Private rsMainReceitDetail As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private fso As Scripting.FileSystemObject
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cb_all_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

If cb_all.Value = 1 Then
    '全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("是否回傳") = "V"
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '取銷全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("是否回傳") = " "
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
End If

End Sub

Private Sub Check1_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'清除
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("是否回傳") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check1.Value = 1 Then
    '全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("分分司代號")) = "A" Then
            rsMainHeader.Fields("是否回傳") = "V"
        Else
            rsMainHeader.Fields("是否回傳") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '取銷全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("是否回傳") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub


Private Sub Check2_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'清除
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("是否回傳") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check2.Value = 1 Then
    '全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("分分司代號")) = "B" Then
            rsMainHeader.Fields("是否回傳") = "V"
        Else
            rsMainHeader.Fields("是否回傳") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '取銷全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("是否回傳") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub


Private Sub Check3_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'清除
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("是否回傳") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check3.Value = 1 Then
    '全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("分分司代號")) = "C" Then
            rsMainHeader.Fields("是否回傳") = "V"
        Else
            rsMainHeader.Fields("是否回傳") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '取銷全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("是否回傳") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub


Private Sub Check4_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'清除
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("是否回傳") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check4.Value = 1 Then
    '全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("分分司代號")) = "D" Then
            rsMainHeader.Fields("是否回傳") = "V"
        Else
            rsMainHeader.Fields("是否回傳") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '取銷全選
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("是否回傳") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub

Private Sub cmd_SendMail_Click()

On Error GoTo err_Handle
Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, str_Date As String, strLMBO01Mail As String, strAddAttachment As String
'讀取ini參數
Dim objIni As New vbIniFile
str_Date = Format(Now(), "YYYY/MM/DD hh:mm:ss")
'objIni.FileName = App.Path & "/" & App.title & ".ini"
'
'strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
'strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
'strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
'strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
'strSubject = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Subject", "")
'strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
'strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
'strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
'strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")

'直接指定
strFrom = "autoreport@bestlog.com.tw"
strTo = "wanfang@maobao.com.tw;kellychen@maobao.com.tw;sara@maobao.com.tw;betty@maobao.com.tw;dennisren@maobao.com.tw;sharon@maobao.com.tw;alisa@maobao.com.tw;tia@maobao.com.tw;julia@maobao.com.tw;yolanda@maobao.com.tw"
'strTo = "gemini@bestlog.com.tw"
strCC = "tina.h@bestlog.com.tw;joane@bestlog.com.tw"
strSubject = str_Date & " POD Feedback Notice"
strTextbody = txt_Msg & Chr(13) & Chr(10) & "The letter sent automatically by the system, do not directly reply.Thanks" & Chr(13) & Chr(10) & "Time:" & str_Date
strEmailID = "autoreport"
strEmailPW = "bestauto"
strAlways = "NO"

If UCase(RTrim(strAlways)) <> "YES" Then strAlways = "NO"
Set objIni = Nothing

If Len(RTrim(strFrom)) > 0 Then '有寄件者
    strLMBO01Mail = "YES"
End If

If strLMBO01Mail = "YES" Then
Screen.MousePointer = 11
'傳送郵件
    Dim objEmail As Object
    Set objEmail = CreateObject("CDO.Message")

    objEmail.From = strFrom
    objEmail.To = strTo
    objEmail.CC = strCC   ' 副本
    objEmail.BCC = strBCC ' 密件副本
    objEmail.Subject = strSubject
    objEmail.TextBody = strTextbody
    'objEmail.AddAttachment strAddAttachment

    objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
    objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    'SMTP 伺服器需要驗證時
    If Len(RTrim(strEmailID)) > 0 Then
        objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
        objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
    End If
    objEmail.Configuration.Fields.Update
    objEmail.Send

    Set objEmail = Nothing

End If

Exit Sub

err_Handle:
Screen.MousePointer = 0
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd2Excel_Click()
On Error GoTo LogOnError
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub
    
    Screen.MousePointer = 11
    Dim FileName As String, txtpath As String, bl_Check As Boolean, str_Date As String, str_Orderkey As String
    Dim bl_CheckA As Boolean, bl_CheckB As Boolean, bl_CheckC As Boolean, bl_CheckD As Boolean, bl_CheckE As Boolean, bl_CheckSNRT As Boolean, bl_CheckOther As Boolean
    bl_CheckA = False: bl_CheckB = False: bl_CheckC = False: bl_CheckD = False: bl_CheckE = False: bl_CheckSNRT = False: bl_CheckOther = False
    
    If rsMainHeader.RecordCount = 0 Then Exit Sub
    bl_Check = False
    '檢查是否有勾選，並檢查有哪些分公司別
    Do While Not rsMainHeader.EOF
        If rsMainHeader.Fields("是否回傳") = "V" Then
            bl_Check = True
            If RTrim(rsMainHeader.Fields("分類")) = "SNRT" Then bl_CheckSNRT = True
            If RTrim(rsMainHeader.Fields("分類")) = "Other" Then bl_CheckOther = True
            If RTrim(rsMainHeader.Fields("分類")) = "A" Then bl_CheckA = True
            If RTrim(rsMainHeader.Fields("分類")) = "B" Then bl_CheckB = True
            If RTrim(rsMainHeader.Fields("分類")) = "C" Then bl_CheckC = True
            If RTrim(rsMainHeader.Fields("分類")) = "D" Then bl_CheckD = True
            If Len(RTrim(rsMainHeader.Fields("分類"))) = 0 Then bl_CheckE = True
        End If
        rsMainHeader.MoveNext
    Loop
    
    If bl_Check = False Then MsgBox "沒有勾選回傳資料，請確認再回傳!", vbCritical + vbOKOnly, "回傳檢查": Screen.MousePointer = 0: Exit Sub
    

    str_Date = Format(Now, "YYMMDDHHNNSS"): str_Orderkey = ""
        
    If bl_CheckSNRT = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\小北大潤發", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\小北大潤發", str_Date, "SNRT")
    End If
    
    If bl_CheckOther = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\其他", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\其他", str_Date, "Other")
    End If
    
    If bl_CheckA = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\總公司", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\總公司", str_Date, "A")
    End If
    
    If bl_CheckB = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\北區", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\北區", str_Date, "B")
    End If
    
    If bl_CheckC = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\南區", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\南區", str_Date, "C")
    End If
    
    If bl_CheckD = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\中區", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\中區", str_Date, "D")
    End If

    If bl_CheckE = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\異常", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\異常", str_Date, "")
    End If
    
    Screen.MousePointer = 0:
    rsMainHeader.Filter = ""

Exit Sub

LogOnError:
'    rsMainHeader.Close
'    rsMainDetail.Close
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Function MBOrs2txt(Str_Path1 As String, Str_Path2 As String, str_Date As String, Str_Company As String)
        Dim FileName As String, str_Orderkey As String, txtpath As String
'Str_Path1 = 本機備份路徑
'Str_Path2 = FTP備份路徑
'Str_RoutePath = 本機路編資料備份路徑
'Str_FtpRoutePath = FTP路編資料備份路徑

On Error GoTo LogOnError
            Dim ReturnOrders As Double, ReturnOrderdetail As Double, ReturnSignqty As Double
            Dim Str_RoutePath As String, Str_FtpRoutePath As String
'            Str_RoutePath = "C:\BEST\LMBO01\POD\配送路編"
'            Str_FtpRoutePath = "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\配送路編"
            ReturnOrders = 0: ReturnOrderdetail = 0: ReturnSignqty = 0
            rsMainHeader.MoveFirst
            '訂單主檔轉文字檔
            rsMainHeader.MoveFirst
            'strOrderNo = ""
            Set fso = New FileSystemObject
            FileName = Str_Company & "_rtb" & str_Date & ".txt"
            If Str_Company = "" Then FileName = "E_rtb" & str_Date & ".txt"
            If Dir(Str_Path1, vbDirectory) = "" Then MkDirs Str_Path1
            txtpath = Str_Path1 & "\" & FileName
            Open txtpath For Append As #1
            Do While Not rsMainHeader.EOF
                If rsMainHeader.Fields("是否回傳") = "V" And (rsMainHeader.Fields("分類") = Str_Company Or Len(Str_Company) = 0) Then
                    ReturnOrders = ReturnOrders + 1
                    str_Orderkey = str_Orderkey & "'" & RTrim(rsMainHeader.Fields("TMS單號")) & "',"
                    If Not IsNull(rsMainHeader.Fields(3)) Then Print #1, StrPadRightC(rsMainHeader.Fields(3), 1); Else Print #1, StrPadRightC(" ", 1);  '分分司代號
                    If Not IsNull(rsMainHeader.Fields(4)) Then Print #1, StrPadRightC(rsMainHeader.Fields(4), 8); Else Print #1, StrPadRightC(" ", 8);  '訂單號碼
                    If Not IsNull(rsMainHeader.Fields(5)) Then Print #1, StrPadRightC(rsMainHeader.Fields(5), 7); Else Print #1, StrPadRightC(" ", 7);  '第單日期
                    If Not IsNull(rsMainHeader.Fields(6)) Then Print #1, StrPadRightC(rsMainHeader.Fields(6), 10); Else Print #1, StrPadRightC(" ", 10);    '發票號碼
                    If Not IsNull(rsMainHeader.Fields(7)) Then Print #1, StrPadRightC(rsMainHeader.Fields(7), 2); Else Print #1, StrPadRightC(" ", 2);  '發票號碼檢查碼
                    If Not IsNull(rsMainHeader.Fields(8)) Then Print #1, StrPadRightC(rsMainHeader.Fields(8), 7); Else Print #1, StrPadRightC(" ", 7);  '發票日期
                    If Not IsNull(rsMainHeader.Fields(9)) Then Print #1, StrPadRightC(rsMainHeader.Fields(9), 8); Else Print #1, StrPadRightC(" ", 8);  '客戶編號
                    If Not IsNull(rsMainHeader.Fields(10)) Then Print #1, StrPadRightC(rsMainHeader.Fields(10), 50); Else Print #1, StrPadRightC(" ", 50);  '客戶名稱
                    If Not IsNull(rsMainHeader.Fields(11)) Then Print #1, StrPadRightC(rsMainHeader.Fields(11), 3); Else Print #1, StrPadRightC(" ", 3);    '葉代代號
                    If Not IsNull(rsMainHeader.Fields(12)) Then Print #1, StrPadRightC(rsMainHeader.Fields(12), 2); Else Print #1, StrPadRightC(" ", 2);    '下貨收現
                    If Not IsNull(rsMainHeader.Fields(13)) Then Print #1, StrPadRightC(rsMainHeader.Fields(13), 70); Else Print #1, StrPadRightC(" ", 70);  '送貨地址
                    If Not IsNull(rsMainHeader.Fields(14)) Then Print #1, StrPadRightC(rsMainHeader.Fields(14), 1); Else Print #1, StrPadRightC(" ", 1);    '聯式
                    If Not IsNull(rsMainHeader.Fields(15)) Then Print #1, StrPadRightC(rsMainHeader.Fields(15), 8); Else Print #1, StrPadRightC(" ", 8);    '統一編號
                    If Not IsNull(rsMainHeader.Fields(16)) Then Print #1, StrPadLeft(rsMainHeader.Fields(16), 8); Else Print #1, StrPadLeft(" ", 8);    '折讓金額
                    If Not IsNull(rsMainHeader.Fields(17)) Then Print #1, StrPadLeft(rsMainHeader.Fields(17), 8); Else Print #1, StrPadLeft(" ", 8);    '數量折讓金額
                    If Not IsNull(rsMainHeader.Fields(18)) Then Print #1, StrPadLeft(rsMainHeader.Fields(18), 8); Else Print #1, StrPadLeft(" ", 8);    '特別折讓金額
                    If Not IsNull(rsMainHeader.Fields(19)) Then Print #1, StrPadLeft(rsMainHeader.Fields(19), 8); Else Print #1, StrPadLeft(" ", 8);    '現金折讓
                    If Not IsNull(rsMainHeader.Fields(20)) Then Print #1, StrPadLeft(rsMainHeader.Fields(20), 10); Else Print #1, StrPadLeft(" ", 10);  '貨款
                    If Not IsNull(rsMainHeader.Fields(21)) Then Print #1, StrPadLeft(rsMainHeader.Fields(21), 10); Else Print #1, StrPadLeft(" ", 10);  '稅前金額
                    If Not IsNull(rsMainHeader.Fields(22)) Then Print #1, StrPadLeft(rsMainHeader.Fields(22), 8); Else Print #1, StrPadLeft(" ", 8);    '稅額
                    If Not IsNull(rsMainHeader.Fields(23)) Then Print #1, StrPadRightC(rsMainHeader.Fields(23), 70); Else Print #1, StrPadRightC(" ", 70);  '備註
                    If Not IsNull(rsMainHeader.Fields(24)) Then Print #1, StrPadRightC(rsMainHeader.Fields(24), 25); Else Print #1, StrPadRightC(" ", 25);  '客戶訂單編號
                    If Not IsNull(rsMainHeader.Fields(25)) Then Print #1, StrPadRightC(rsMainHeader.Fields(25), 1); Else Print #1, StrPadRightC(" ", 1);    '隨貨附發票碼
                    If Not IsNull(rsMainHeader.Fields(26)) Then Print #1, StrPadRightC(rsMainHeader.Fields(26), 1); Else Print #1, StrPadRightC(" ", 1);    '隨貨附訂單碼
                    If Not IsNull(rsMainHeader.Fields(27)) Then Print #1, StrPadRightC(rsMainHeader.Fields(27), 1); Else Print #1, StrPadRightC(" ", 2);    '計算物流費
                    If Not IsNull(rsMainHeader.Fields(28)) Then Print #1, StrPadRightC(rsMainHeader.Fields(28), 1); Else Print #1, StrPadRightC(" ", 1);    '送貨否
                    If Not IsNull(rsMainHeader.Fields(29)) Then Print #1, StrPadRightC(rsMainHeader.Fields(29), 2); Else Print #1, StrPadRightC(" ", 2);    '訂單種類
                    If Not IsNull(rsMainHeader.Fields(30)) Then Print #1, StrPadRightC(rsMainHeader.Fields(30), 1); Else Print #1, StrPadRightC(" ", 1);    '實收量處理MARK
                    If Not IsNull(rsMainHeader.Fields(31)) Then Print #1, StrPadRightC(rsMainHeader.Fields(31), 12); Else Print #1, StrPadRightC(" ", 12);  '聯絡人
                    If Not IsNull(rsMainHeader.Fields(32)) Then Print #1, StrPadRightC(rsMainHeader.Fields(32), 20); Else Print #1, StrPadRightC(" ", 20);  '電話
                    If Not IsNull(rsMainHeader.Fields(33)) Then Print #1, StrPadRightC(rsMainHeader.Fields(33), 12); Else Print #1, StrPadRightC(" ", 12);  '葉代姓名
                    If Not IsNull(rsMainHeader.Fields(34)) Then Print #1, StrPadRightC(rsMainHeader.Fields(34), 12); Else Print #1, StrPadRightC(" ", 12);  '主管姓名
                    If Not IsNull(rsMainHeader.Fields(35)) Then Print #1, StrPadRightC(rsMainHeader.Fields(35), 50); Else Print #1, StrPadRightC(" ", 50);  '指送客戶
                    If Not IsNull(rsMainHeader.Fields(36)) Then Print #1, StrPadRightC(rsMainHeader.Fields(36), 7); Else Print #1, StrPadRightC(" ", 7);    '預計日期
                    If Not IsNull(rsMainHeader.Fields(37)) Then Print #1, StrPadLeft(rsMainHeader.Fields(37), 8); Else Print #1, StrPadLeft(" ", 8);    '運費
                    If Not IsNull(rsMainHeader.Fields(38)) Then Print #1, StrPadRightC(rsMainHeader.Fields(38), 1); Else Print #1, StrPadRightC(" ", 1);    '付款方式
                    If Not IsNull(rsMainHeader.Fields(39)) Then Print #1, StrPadRightC(rsMainHeader.Fields(39), 20); Else Print #1, StrPadRightC(" ", 20);  '業務手機
                    If Not IsNull(rsMainHeader.Fields(40)) Then Print #1, StrPadRightC(rsMainHeader.Fields(40), 1); Else Print #1, StrPadRightC(" ", 1);    '是否為電子發票
                    If Not IsNull(rsMainHeader.Fields(41)) Then Print #1, StrPadRightC(rsMainHeader.Fields(41), 6); Else Print #1, StrPadRightC(" ", 6);    '總重量
                    If Not IsNull(rsMainHeader.Fields(42)) Then Print #1, StrPadRightC(rsMainHeader.Fields(42), 4); Else Print #1, StrPadRightC(" ", 4);    '信卡後4碼
                    If Not IsNull(rsMainHeader.Fields(43)) Then Print #1, StrPadLeft(rsMainHeader.Fields(43), 10); Else Print #1, StrPadLeft(" ", 10);  '代收貨款
                    If Not IsNull(rsMainHeader.Fields(44)) Then Print #1, StrPadRightC(rsMainHeader.Fields(44), 1); Else Print #1, StrPadRightC(" ", 1);    '發票列印方式
                    If Not IsNull(rsMainHeader.Fields(45)) Then Print #1, StrPadRightC(rsMainHeader.Fields(45), 20); Else Print #1, StrPadRightC(" ", 20);  '電話2
                    If Not IsNull(rsMainHeader.Fields(46)) Then Print #1, StrPadRightC(rsMainHeader.Fields(46), 8); Else Print #1, StrPadRightC(" ", 8);    '統計對像
                    If Not IsNull(rsMainHeader.Fields(47)) Then Print #1, StrPadRightC(rsMainHeader.Fields(47), 3); Else Print #1, StrPadRightC(" ", 3);    '縣市別
                    If Not IsNull(rsMainHeader.Fields(48)) Then Print #1, StrPadRightC(rsMainHeader.Fields(48), 3); Else Print #1, StrPadRightC(" ", 3);    '行政區
                    If Not IsNull(rsMainHeader.Fields(49)) Then Print #1, StrPadRightC(rsMainHeader.Fields(49), 2); Else Print #1, StrPadRightC(" ", 2);    '樓層
                    If Not IsNull(rsMainHeader.Fields(50)) Then Print #1, StrPadRightC(rsMainHeader.Fields(50), 1); Else Print #1, StrPadRightC(" ", 1);    '越庫訂單
                    If Not IsNull(rsMainHeader.Fields(51)) Then Print #1, StrPadRightC(rsMainHeader.Fields(51), 12); Else Print #1, StrPadRightC(" ", 12);  '提貨倉
                    If Not IsNull(rsMainHeader.Fields(52)) Then Print #1, StrPadRightC(rsMainHeader.Fields(52), 10); Else Print #1, StrPadRightC(" ", 10);  '稅區/稅率
                    If Not IsNull(rsMainHeader.Fields(53)) Then Print #1, StrPadRightC(rsMainHeader.Fields(53), 40); Else Print #1, StrPadRightC(" ", 40);  '客戶簡稱
                    If Not IsNull(rsMainHeader.Fields(54)) Then Print #1, StrPadRightC(rsMainHeader.Fields(54), 7); Else Print #1, StrPadRightC(" ", 7);  '實際到貨日
                    If Not IsNull(rsMainHeader.Fields(55)) Then Print #1, StrPadRightC(rsMainHeader.Fields(55), 10); Else Print #1, StrPadRightC(" ", 10);  '關聯訂單號碼
                    Print #1, vbCrLf;
                End If
                rsMainHeader.MoveNext
            Loop
            Close #1
            rsMainHeader.MoveFirst

            '訂單明細檔轉文字檔
            rsMainDetail.MoveFirst
            'strOrderNo = ""
            Set fso = New FileSystemObject
            FileName = Str_Company & "_rdb" & str_Date & ".txt"
            If Str_Company = "" Then FileName = "E_rdb" & str_Date & ".txt"
            If Dir(Str_Path1, vbDirectory) = "" Then MkDirs Str_Path1
            txtpath = Str_Path1 & "\" & FileName
            Open txtpath For Append As #1
            Do While Not rsMainHeader.EOF
              If rsMainHeader.Fields("是否回傳") = "V" Then
                rsMainDetail.Filter = "TMS單號 = '" & rsMainHeader.Fields("TMS單號") & "'"
                rsMainDetail.MoveFirst
                '排除數量有差異的品項
                Do While Not rsMainDetail.EOF
                '如果未出訂單則要回傳，如果不是全未出訂單則不回傳有差異品項的細項 edit by Eric 20141001 Phil通知
                    If rsMainHeader.Fields("TMS單號") = rsMainDetail.Fields("TMS單號") And rsMainHeader.Fields("是否回傳") = "V" And (rsMainHeader.Fields("分類") = Str_Company Or Len(Str_Company) = 0) Then
                        ReturnOrderdetail = ReturnOrderdetail + 1
                        ReturnSignqty = ReturnSignqty + Val(rsMainDetail.Fields(7))
                        If Not IsNull(rsMainDetail.Fields(3)) Then Print #1, StrPadRightC(rsMainDetail.Fields(3), 8); Else Print #1, StrPadRightC(" ", 8);  '訂單號碼
                        If Not IsNull(rsMainDetail.Fields(4)) Then Print #1, StrPadRightC(rsMainDetail.Fields(4), 16); Else Print #1, StrPadRightC(" ", 16);    '產品編號
                        If Not IsNull(rsMainDetail.Fields(5)) Then Print #1, StrPadRightC(rsMainDetail.Fields(5), 60); Else Print #1, StrPadRightC(" ", 60);    '產品名稱
                        If Not IsNull(rsMainDetail.Fields(6)) Then Print #1, StrPadLeft(rsMainDetail.Fields(6), 10); Else Print #1, StrPadLeft(" ", 10);    '訂貨量
                        If Not IsNull(rsMainDetail.Fields(8)) Then Print #1, StrPadLeft(rsMainDetail.Fields(8), 8); Else Print #1, StrPadLeft(" ", 8);  '單價(未稅)
                        If Not IsNull(rsMainDetail.Fields(9)) Then Print #1, StrPadLeft(rsMainDetail.Fields(9), 10); Else Print #1, StrPadLeft(" ", 10);    '訂貨金額(未稅)
                        If Not IsNull(rsMainDetail.Fields(10)) Then Print #1, StrPadLeft(rsMainDetail.Fields(10), 8); Else Print #1, StrPadLeft(" ", 8);  '單價(含稅)
                        If Not IsNull(rsMainDetail.Fields(11)) Then Print #1, StrPadLeft(rsMainDetail.Fields(11), 10); Else Print #1, StrPadLeft(" ", 10);  '訂貨金額(含稅)
                        If Not IsNull(rsMainDetail.Fields(7)) Then Print #1, StrPadLeft(rsMainDetail.Fields(7), 10); Else Print #1, StrPadLeft(" ", 10);  '訂貨量-實收量
                        If Not IsNull(rsMainDetail.Fields(12)) Then Print #1, StrPadRightC(rsMainDetail.Fields(12), 25); Else Print #1, StrPadRightC(" ", 25);  '國際條碼
                        If Not IsNull(rsMainDetail.Fields(13)) Then Print #1, StrPadLeft(rsMainDetail.Fields(13), 7); Else Print #1, StrPadLeft(" ", 7);    '行號
                        If Not IsNull(rsMainDetail.Fields(14)) Then Print #1, StrPadRightC(rsMainDetail.Fields(14), 2); Else Print #1, StrPadRightC(" ", 2);    '單位
                        If Not IsNull(rsMainDetail.Fields(15)) Then Print #1, StrPadRightC(rsMainDetail.Fields(15), 2); Else Print #1, StrPadRightC(" ", 2);    '訂單種類
                        If Not IsNull(rsMainDetail.Fields(16)) Then Print #1, StrPadRightC(rsMainDetail.Fields(16), 1); Else Print #1, StrPadRightC(" ", 1);    '發票明細列印否
                        If Not IsNull(rsMainDetail.Fields(17)) Then Print #1, StrPadRightC(rsMainDetail.Fields(17), 20); Else Print #1, StrPadRightC(" ", 20);  '允收期
                        Print #1, vbCrLf;
                    End If
                    rsMainDetail.MoveNext
                Loop
              End If
                rsMainHeader.MoveNext
                rsMainDetail.MoveFirst
            Loop
            Close #1
            rsMainHeader.MoveFirst
            rsMainDetail.MoveFirst

'Mark by Eric 20141210，小北和大潤發放至北區修改，一併移除此功能
'    '轉出配送路編
'        str_Orderkey = Mid(str_Orderkey, 1, Len(str_Orderkey) - 1)
'        Call Confirm_Recordset_Closed(tmp_Rs)
''        '補資料版
''        str_SQL = "select 訂單種類=co.ordertype,訂單號碼=co.externorderkey,路編編號=s2.c_route_no,送貨否=co.DeliveryCode,堆高機費用 =isnull(sum(sumreceivable),0) " & _
''                    "from sdn02t s2 join custorders co on s2.c_receipt_no = co.orderkey " & _
''                    "left join  sdn05t s5 on s5.sdn_no = s2.receipt_no and s5.costcode = 'forklift' " & _
''                    "where convert(char(8),co.adddate,112) between '20140801' and '20140901' " & _
''                    "group by co.OrderType,co.externorderkey,s2.c_route_no,co.DeliveryCode"
''
'        '正常版
'        str_SQL = "select 訂單種類=co.ordertype,訂單號碼=co.externorderkey,路編編號=s2.c_route_no,送貨否=co.DeliveryCode,堆高機費用 =isnull(sum(sumreceivable),0) " & _
'                    "from sdn02t s2 join custorders co on s2.c_receipt_no = co.orderkey " & _
'                    "left join  sdn05t s5 on s5.sdn_no = s2.receipt_no and s5.costcode = 'forklift' " & _
'                    "where s2.c_receipt_no in (" & str_Orderkey & ") " & _
'                    "group by co.OrderType,co.externorderkey,s2.c_route_no,co.DeliveryCode "
'
'        tmp_Rs.Open str_SQL, cn
'
'            '轉文字檔
'            tmp_Rs.MoveFirst
'            Set fso = New FileSystemObject
'            FileName = Str_Company & "_配送路編" & Str_Date & ".txt"
'            If Str_Company = "" Then FileName = "E_配送路編" & Str_Date & ".txt"
'            If Dir(Str_RoutePath, vbDirectory) = "" Then MkDirs Str_RoutePath
'            txtpath = Str_RoutePath & "\" & FileName
'            Open txtpath For Append As #1
'            Do While Not tmp_Rs.EOF
'                '查出配送路編
'                If Not IsNull(tmp_Rs.Fields(0)) Then Print #1, StrPadRightC(tmp_Rs.Fields(0), 2); Else Print #1, StrPadRightC(" ", 2);  '訂單種類
'                If Not IsNull(tmp_Rs.Fields(1)) Then Print #1, StrPadRightC(tmp_Rs.Fields(1), 8); Else Print #1, StrPadRightC(" ", 8);    '訂單號碼
'                If Not IsNull(tmp_Rs.Fields(2)) Then Print #1, StrPadRightC(tmp_Rs.Fields(2), 10); Else Print #1, StrPadRightC(" ", 10);    '路線編號
'                If Not IsNull(tmp_Rs.Fields(3)) Then Print #1, StrPadRightC(tmp_Rs.Fields(3), 1); Else Print #1, StrPadLeft(" ", 1);    '送貨否
'                If Not IsNull(tmp_Rs.Fields(4)) Then Print #1, StrPadLeft(tmp_Rs.Fields(4), 9); Else Print #1, StrPadLeft(" ", 9);    '堆高機費用
'                Print #1, vbCrLf;
'                tmp_Rs.MoveNext
'            Loop
'            Close #1
'            tmp_Rs.Close

'紀錄此回傳的訂單筆數，訂單明細數，訂單總回傳量
txt_Msg.Text = txt_Msg.Text & Str_Company & "_rdb" & str_Date & ".txt : " & ReturnOrders & " Orders; " & ReturnOrderdetail & " Detail;Total Qty = " & ReturnSignqty & Chr(13) & Chr(10)

'備份到FTP
'備份檔案
If Dir(Str_Path2, vbDirectory) = "" Then
    MkDirs Str_Path2
    If Len(Str_Company) = 0 Then Str_Company = "E"
    FileCopy Str_Path1 & "\" & Str_Company & "_rtb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rtb" & str_Date & ".txt"
    FileCopy Str_Path1 & "\" & Str_Company & "_rdb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rdb" & str_Date & ".txt"
    'FileCopy Str_RoutePath & "\" & Str_Company & "_配送路編" & Str_Date & ".txt", Str_FtpRoutePath & "\" & Str_Company & "_配送路編" & Str_Date & ".txt"
Else
    If Len(Str_Company) = 0 Then Str_Company = "E"
    FileCopy Str_Path1 & "\" & Str_Company & "_rtb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rtb" & str_Date & ".txt"
    FileCopy Str_Path1 & "\" & Str_Company & "_rdb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rdb" & str_Date & ".txt"
    'FileCopy Str_RoutePath & "\" & Str_Company & "_配送路編" & Str_Date & ".txt", Str_FtpRoutePath & "\" & Str_Company & "_配送路編" & Str_Date & ".txt"
End If

Exit Function
LogOnError:
'    rsMainHeader.Close
'    rsMainDetail.Close

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Function
Private Sub cmdOK_Click()
'
'If rsMain Is Nothing Then Exit Sub
'strOtQtyFixOrderkey = rsMain("TMS單號")
'frm_OTQtyFix.Show vbModal
'
''更新Datagrid
'Call UpdateDatagrid

End Sub

Private Sub cmdOTUpdate_Click()

On Error GoTo err_Handle

'Dim Str_Date As String
'Str_Date = Format(Now(), "yyyymmddhhmmss")
'Call MBOrs2txt("C:\BEST\LMBO01\POD\配送路編", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\配送路編", Str_Date, "A")
Dim bl_Check1 As Boolean
bl_Check1 = False
txt_UnReciept.Text = ""
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

    Dim x As Integer, bl_Check As Boolean
    Screen.MousePointer = 11
    bl_Check = False
        
    '關閉datagrid
    dgMain_Header.Visible = False
    dgMain_Detail.Visible = False
    
    rsMainHeader.MoveFirst
    '檢查有無拆單，有拆單的話是否兩張都已回
     Do While Not rsMainHeader.EOF
                If rsMainHeader.Fields("是否回傳") = "V" Then
                '檢查有無全收
                    str_SQL = "select  receipt_no,extern,sdnback from sdn02t where storerkey = 'LMBO01' and extern = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "'"
                    Call Confirm_Recordset_Closed(tmp_Rs)
                    tmp_Rs.CursorLocation = adUseClient
                    tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                        If RTrim(tmp_Rs.Fields("sdnback")) = "0" Then
                            Screen.MousePointer = 0
                            dgMain_Header.Visible = True
                            dgMain_Detail.Visible = True
                            MsgBox "TMS單號:" & RTrim(tmp_Rs.Fields("receipt_no")) & ",貨主單號:" & RTrim(tmp_Rs.Fields("extern")) & ",仍有拆單的部分沒有回來，無法回傳。請確認!", vbOKOnly + vbCritical, "拆單已回檢查"
                            tmp_Rs.Close
                            Exit Sub
                        End If
                        tmp_Rs.MoveNext
                    Loop
                    tmp_Rs.Close
                End If
            rsMainHeader.MoveNext
        Loop
        
'    rsMainHeader.MoveFirst
'    '出貨部份，如果有異常簽單則要檢查有無全收，未出訂單則要判斷是否有出貨，有的話要全收，沒有的話不管
'    If optOut.Value = True Then
'        Do While Not rsMainHeader.EOF
'            If rsMainHeader.Fields("是否回傳") = "V" And rsMainHeader.Fields("狀態") = "異常訂單" Then
'            '檢查有無全收
'                str_SQL = "select externreceiptkey from " & strWMSDB & "..receipt where storerkey = 'LMBO01' and status = '9' and receipttype = 'A' and externreceiptkey = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "'"
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.CursorLocation = adUseClient
'                tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If tmp_Rs.EOF = True Then
'                '未全收
'                txt_UnReciept.Text = txt_UnReciept.Text & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & ","
'                rsMainHeader.Fields("是否回傳") = " " '取消回傳
'                End If
'                tmp_Rs.Close
'            End If
'            If rsMainHeader.Fields("是否回傳") = "V" And rsMainHeader.Fields("狀態") = "未出訂單" Then
'            '先檢查有無配置，有則要判斷有無全收
'                str_SQL = "select shippedqty=isnull(sum(shippedqty),0) from " & strWMSDB & "..orderdetail where storerkey = 'LMBO01' and status = '9' and externorderkey = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "'"
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.CursorLocation = adUseClient
'                tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If Val(tmp_Rs.Fields("shippedqty")) > 0 Then
'                '有出貨，檢查有無全收
'                str_SQL = "select externreceiptkey from " & strWMSDB & "..receipt where storerkey = 'LMBO01' and status = '9' and receipttype = 'A' and externreceiptkey = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "'"
'                Call Confirm_Recordset_Closed(rs_Receipt)
'                rs_Receipt.CursorLocation = adUseClient
'                rs_Receipt.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If rs_Receipt.EOF = True Then
'                '未全收
'                    txt_UnReciept.Text = txt_UnReciept.Text & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & ","
'                    rsMainHeader.Fields("是否回傳") = " " '取消回傳
'                End If
'                rs_Receipt.Close
'                End If
'                tmp_Rs.Close
'            End If
'
'            rsMainHeader.MoveNext
'        Loop
'    End If
'
'    If optIn.Value = True Then
' Do While Not rsMainHeader.EOF
'            If rsMainHeader.Fields("是否回傳") = "V" Then
'            '檢查有無全收
'                str_SQL = "select externreceiptkey from " & strWMSDB & "..receipt where storerkey = 'LMBO01' and status = '9' and externreceiptkey = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "'"
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.CursorLocation = adUseClient
'                tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If tmp_Rs.EOF = True Then
'                '未全收
'                txt_UnReciept.Text = txt_UnReciept.Text & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & ","
'                rsMainHeader.Fields("是否回傳") = " " '取消回傳
'                End If
'                tmp_Rs.Close
'            End If
'            rsMainHeader.MoveNext
'        Loop
'    End If
    rsMainHeader.MoveFirst
    '檢查有勾選回傳
    Do While Not rsMainHeader.EOF
        If rsMainHeader.Fields("是否回傳") = "V" Then
            bl_Check = True
        End If
        rsMainHeader.MoveNext
    Loop
    
    If bl_Check = False Then Screen.MousePointer = 0: dgMain_Header.Visible = True:     dgMain_Detail.Visible = False: Exit Sub
    rsMainHeader.Filter = "是否回傳 = 'V'"
    
    '比對訂單量是否等於簽單量
        rsMainHeader.MoveFirst
        rsMainDetail.MoveFirst
    Do While Not rsMainHeader.EOF
       If rsMainHeader.Fields("是否回傳") = "V" Then
       rsMainDetail.Filter = "TMS單號 = '" & rsMainHeader.Fields("TMS單號") & "'"
       rsMainDetail.MoveFirst
            Do While Not rsMainDetail.EOF
                If rsMainHeader.Fields("TMS單號") = rsMainDetail.Fields("TMS單號") Then
                    If Abs(Val(rsMainDetail.Fields("訂貨量"))) <> Val(rsMainDetail.Fields("訂貨量-實收量")) Then
                        x = MsgBox("訂單號碼:" & rsMainDetail.Fields("TMS單號") & "項次:" & rsMainDetail.Fields("TMS單號項次") & ":訂貨量<>簽收量，請確認是否更新回傳?", vbQuestion + vbYesNo, "數量檢查")
                        If x = 6 Then
                            '記續
                        Else
                            rsMainHeader.Fields("是否回傳") = " "
                            GoTo next1
                            Exit Sub
                        End If
                    End If
                End If
                rsMainDetail.MoveNext
            Loop
       End If
next1:
        rsMainDetail.MoveFirst
        rsMainHeader.MoveNext
    Loop
'
'    rsMainHeader.MoveFirst
'    '檢查有無回傳了
'    Do While Not rsMainHeader.EOF
'        If rsMainHeader.Fields("是否回傳") = "V" Then
'            str_SQL = "select returnstatus from sdn02t where c_receipt_no = '" & rsMainHeader.Fields("TMS單號") & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.CursorLocation = adUseClient
'            tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
'
'            If tmp_Rs.EOF = True Then
'                Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'            Else
'                Do While Not tmp_Rs.EOF
'                    If RTrim(tmp_Rs.Fields("returnstatus")) = "1" Or RTrim(tmp_Rs.Fields("returnstatus")) = "2" Then MsgBox rsMainHeader.Fields("TMS單號") & "有已回傳資料,回傳中止", vbCritical + vbOKOnly, "回傳檢查": tmp_Rs.Close:    Screen.MousePointer = 0: Exit Sub
'                    tmp_Rs.MoveNext
'                Loop
'            End If
'            tmp_Rs.Close
'        End If
'        rsMainHeader.MoveNext
'    Loop
'
'    '檢查拆單是否都出車確認了
'    rsMainHeader.MoveFirst
'    Do While Not rsMainHeader.EOF
'      If rsMainHeader.Fields("是否回傳") = "V" Then
'            str_SQL = "exec es_CheckConfirm '" & rsMainHeader.Fields("TMS單號") & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'            If Not tmp_Rs.EOF Then
'                MsgBox "TMS單號:" & rsMainHeader.Fields("TMS單號") & "中有單號:" & tmp_Rs.Fields("receipt_no") & "未出車，請確認後，再進行回傳作業，回傳終止", vbCritical + vbOKOnly, "回傳檢查"
'                tmp_Rs.Close: Screen.MousePointer = 0: Exit Sub
'            End If
'
'        End If
'                    rsMainHeader.MoveNext
'    Loop
'    rsMainHeader.MoveFirst: tmp_Rs.Close
    
'    '檢查回傳的TMS單號中拆的單，是否已經回來了
'    rsMainHeader.MoveFirst
'    Do While Not rsMainHeader.EOF
'      If rsMainHeader.Fields("是否回傳") = "V" Then
'        str_SQL = "select 拆單TMS = s2.receipt_no,簽單狀態 = s2.sdnback from sdn02t s2 where s2.c_receipt_no = '" & rsMainHeader.Fields("TMS單號") & "'"
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        tmp_Rs.MoveFirst
'        Do While Not tmp_Rs.EOF
'            If tmp_Rs.Fields("簽單狀態") = "0" Then
'                MsgBox "TMS單號:" & rsMainHeader.Fields("TMS單號") & "中有拆單單號:" & tmp_Rs.Fields("拆單TMS") & "未回，請確認後，再進行回傳作業，回傳終止", vbCritical + vbOKOnly, "回傳檢查"
'                tmp_Rs.Close: Screen.MousePointer = 0:
'                dgMain_Header.Visible = True
'                dgMain_Detail.Visible = True
'                Exit Sub
'            End If
'            tmp_Rs.MoveNext
'        Loop
'        End If
'        rsMainHeader.MoveNext
'    Loop
'
    Tran_Level = cn.BeginTrans:
'
'    '檢查CO訂單的receiptdetail實收量是否有超收
'    rsMainHeader.MoveFirst
'    If RTrim(rsMainHeader.Fields("訂單種類")) = "CO" Or RTrim(rsMainHeader.Fields("訂單種類")) = "SC" Then
'        Do While Not rsMainHeader.EOF
'          If rsMainHeader.Fields("是否回傳") = "V" Then
'            str_SQL = "select 訂單號碼 = isnull(a.externasnkey,r.externreceiptkey),品號=isnull(a.sku,r.sku),通知量 =sum(isnull(a.qty,0)) ,實收量 = sum(isnull(r.qty,0)) " & _
'                        "from ( " & _
'                        "select a.externasnkey,ad.sku,qty=sum(isnull(ad.qtyordered,0)) " & _
'                        "From " & strWMSDB & "..asn a join " & strWMSDB & "..asndetail ad on a.asnkey = ad.asnkey " & _
'                        "where a.externasnkey = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "' " & _
'                        "group by a.externasnkey,ad.sku " & _
'                        ") a full join " & _
'                        "( " & _
'                        "select r.externreceiptkey,rd.sku,qty=sum(isnull(rd.qtyreceived,0)) " & _
'                        "from " & strWMSDB & "..receipt r join " & strWMSDB & "..receiptdetail rd on r.receiptkey = rd.receiptkey " & _
'                        "where r.externreceiptkey = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "' and r.status = '9' " & _
'                        "group by r.externreceiptkey,rd.sku " & _
'                        ") r on a.externasnkey = r.externreceiptkey and a.sku = r.sku " & _
'                        "group by isnull(a.externasnkey,r.externreceiptkey),isnull(a.sku,r.sku) "
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'              tmp_Rs.MoveFirst
'              Do While Not tmp_Rs.EOF
'                If Val(tmp_Rs.Fields("通知量")) < Val(tmp_Rs.Fields("實收量")) Then
'                    '產生差異表
'                    tmp_Rs.MoveFirst
'                    Recordset2Excel "差異表", tmp_Rs
'                    Screen.MousePointer = 0
'                    cn.RollbackTrans: Tran_Level = 0
'                    dgMain_Header.Visible = True
'                    dgMain_Detail.Visible = True
'                    MsgBox "訂單號碼:" & RTrim(tmp_Rs.Fields("訂單號碼")) & " 品號:" & RTrim(tmp_Rs.Fields("品號")) & " 通知量:" & RTrim(tmp_Rs.Fields("通知量")) & " <> 實收量:" & RTrim(tmp_Rs.Fields("實收量")) & "，無法回傳", vbCritical + vbOKOnly, "回傳檢查"
'                    tmp_Rs.Close
'                    Exit Sub
'                End If
'                tmp_Rs.MoveNext
'              Loop
'              tmp_Rs.Close
'              cn.Execute "update " & strWMSDB & "..asn set status = '9' where externasnkey = '" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "'", RowsAffect, adExecuteNoRecords
'          End If
'              rsMainHeader.MoveNext
'        Loop
'
'    End If
    rsMainHeader.MoveFirst
    rsMainDetail.MoveFirst
    
    If rsMainHeader.EOF Then
        MsgBox "查無資料可供轉檔！", vbOKOnly + vbInformation, Me.Caption
        Screen.MousePointer = 0:
        cn.RollbackTrans: Tran_Level = 0
    Else
        Do While Not rsMainHeader.EOF
          If rsMainHeader.Fields("是否回傳") = "V" Then
                '檢查returnstatus = 2的則不可以更新成1
                str_SQL = "select returnstatus from sdn02t where c_receipt_no = '" & rsMainHeader.Fields("TMS單號") & "'"
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If tmp_Rs.Fields("returnstatus") = "2" Then Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0: dgMain_Header.Visible = True: dgMain_Detail.Visible = True: MsgBox "TMS單號:" & rsMainHeader.Fields("TMS單號") & "已回傳，無法再回傳!回傳終止", vbOKOnly + vbCritical, "回傳檢查": tmp_Rs.Close: Exit Sub
                tmp_Rs.Close
                '更新retrunstatus
                str_SQL = "update sdn02t set returnstatus = '1' where c_receipt_no = '" & rsMainHeader.Fields("TMS單號") & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
            rsMainHeader.MoveNext
        Loop
    End If
    rsMainHeader.MoveFirst
    
    '回傳到FTP
    Call cmd2Excel_Click
    cn.CommitTrans: Tran_Level = 0
    cmdOTUpdate.Enabled = False
    dgMain_Header.Visible = True
    dgMain_Detail.Visible = True
    'Send Mail通知客戶
    Call cmd_SendMail_Click
    Screen.MousePointer = 0: rsMainHeader.Filter = "": rsMainDetail.Filter = "": rsMainHeader.Close: rsMainDetail.Close: txt_Msg = ""
    MsgBox "簽收量已回傳^_^並已mail通知客戶。", vbOKOnly, "POD回傳成功"
    
Exit Sub
err_Handle:
Screen.MousePointer = 0
    dgMain_Header.Visible = True
    dgMain_Detail.Visible = True
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Public Sub cmdPrintReport_Click()
'Dim i As Integer, j As Integer, k As Integer
'On Error GoTo err_Handle
'
'If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub
'
''取已選取資料
'If RTrim(strOtQtyFixOrderkey) <> "" Then
'    rsMain.Filter = "(TMS單號 = " & strOtQtyFixOrderkey & ")"
'Else
'    rsMain.Filter = "(＊ = 'V')"
'End If
'
'If rsMain.RecordCount = 0 Then rsMain.Filter = 0: MsgBox "請選取欲列印之資料。", 64, "列印": rsMain.Sort = "編號": Exit Sub
'
'Screen.MousePointer = 11
'
''資料寫入 Access 資料庫
'Call AccessDB_Connect
'cnAccess.BeginTrans
'Tran_Level = cn.BeginTrans
'
'cnAccess.Execute "Delete From 出貨件數", RowsAffect, adExecuteNoRecords
'
'Dim rs_Access As New ADODB.Recordset
'rs_Access.Open "出貨件數", cnAccess, adOpenStatic, adLockOptimistic
'
'rsMain.MoveFirst
'
'Do While Not rsMain.EOF
'    For j = 1 To rsMain("出貨件數") '一件寫入一筆
'        rs_Access.AddNew
'
'        For i = 0 To rsMain.Fields.Count - 1 '寫入每個欄位
'            rs_Access.Fields(i).Value = rsMain.Fields(i).Value
'        Next i
'
'        rs_Access.Fields(i).Value = j
'        rs_Access.Fields(i + 1).Value = rsMain("出貨件數")
'        rs_Access.Update
'    Next j
'
'    'TRP02T更新為已回傳
'    str_SQL = "update trp02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    'TRP02W更新為已回傳
'    str_SQL = "update TRP02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    'ORT02T更新為已回傳
'    str_SQL = "update ort02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    'ORT02W更新為已回傳
'    str_SQL = "update ort02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS單號")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    rsMain("列印次數") = rsMain("列印次數") + 1
'    rsMain("列印時間") = Format(Now, "yyyy/mm/dd hh:mm:ss")
'
'   rsMain.MoveNext
'Loop
'
'cn.CommitTrans: Tran_Level = 0
'cnAccess.CommitTrans
'
'Call DB_Disconnect(cnAccess)
'
'strAccessDBFileName_FullPath = GetAccessDBFileName
'Dim MSAccessAP As New access.Application
'With MSAccessAP
'    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
'
'    If chkPrintPreView.Value = vbChecked Then
'    '預覽列印
'         .DoCmd.OpenReport "出貨件數", acViewPreview
'        .DoCmd.Maximize
'        .Visible = True
'    Else
'    '直接列印至印表機
'        .Visible = False
'        .DoCmd.OpenReport "出貨件數", acViewNormal
'        .CloseCurrentDatabase
'        .Quit
'        Set MSAccessAP = Nothing
'End If
'
'End With
'rsMain.Filter = 0
'rsMain.Sort = "編號"
'Screen.MousePointer = 0
'strOtQtyFixOrderkey = ""
'Exit Sub
'
'err_Handle:
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

Set dgMain_Header.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆訂單資料列"
Set dgMain_Detail.DataSource = Nothing: 'StatusBar.Panels(2).Text = "0 筆資料列"

Dim chc_DeliveryDate As String, chc_ExternOrderkey, chc_Status As String, chc_Storerkey As String, chc_Carno As String, chc_Print As String, str_WhereExternorderkey As String
str_WhereExternorderkey = ""
''先檢查有無asn標記不用回傳的資料
'If optIn = True Then
'    str_SQL = "update s2 set s2.returnstatus = '3' " & _
'        "from " & strWMSDB & "..asn a join custorders co on a.externasnkey = rtrim(co.ordertype) + rtrim(co.externorderkey) " & _
'        "join sdn02t s2 on s2.c_receipt_no = co.orderkey " & _
'        "where a.asntype = 'R' and a.status = 2  and s2.returnstatus <> 3 "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'End If

'到貨日期
chc_DeliveryDate = ""
If Len(RTrim(txtDeliveryDateS.Text)) > 0 And Len(RTrim(txtDeliveryDateE.Text)) > 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateS.Text)) > 0 And Len(RTrim(txtDeliveryDateE.Text)) = 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateS.Text)) = 0 And Len(RTrim(txtDeliveryDateE.Text)) > 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) = '" & txtDeliveryDateE.Text & "' "
End If
'
''貨主單號
'chc_ExternOrderkey = ""
'If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
'   chc_ExternOrderkey = "and o.externorderkey between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
'ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
'   chc_ExternOrderkey = "and o.externorderkey = '" & txtExternOrderkeyS.Text & "' "
'ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
'   chc_ExternOrderkey = "and o.externorderkey = '" & txtExternOrderkeyE.Text & "' "
'End If

'件數狀態
chc_Status = ""
If optNo = True Then chc_Status = chc_Status & "and s2.ReturnStatus = 0 "
If optYes = True Then chc_Status = chc_Status & "and s2.ReturnStatus <> 0 "
If optOut = True Then chc_Status = chc_Status & "and co.ordertype not in ('CO','SC') "
If optIn = True Then chc_Status = chc_Status & "and co.ordertype in ('CO','SC') "


'貨主
chc_Storerkey = ""
If Len(RTrim(cboStorerkey.Text)) > 0 Then chc_Storerkey = " and o.storerkey = '" & RTrim(cboStorerkey.Text) & "' "
chc_Storerkey = "LMBO01"


If optOut.Value = True Then
'出貨部份只挑出簽單量=訂單通知量的簽單。 edit by Eric 20141003，Phil通知,20150122宜花東不回傳。
        str_SQL = "select externorderkey = rtrim(co.ordertype)+rtrim(co.ExternOrderkey) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
                    "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
                    "join orders o on co.orderkey = o.orderkey " & _
                    "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
                    "where s2.storerkey = '" & chc_Storerkey & "' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
                    "and co.address not like  '%台東%' and address not like '%花蓮%' and address not like'%宜蘭%' " & _
                    "and co.City not in ('260','261','262','263','264','265','266','267','268','269','270','272','290','950','951','952','953','954','955','956','957','958','959','961','962','963','964','965','966','970','971','972','973','974','975','976','977','978','979','981','982','983') " & _
                    "and co.Administration not in ('015','016','017') " & _
                    "group by co.ExternOrderkey,co.ordertype " & _
                    "having sum(s3.sign_qty) = sum(cast(cod.originalqty as float)) "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = adUseClient
        tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic   '
        
        If tmp_Rs.EOF = True Then
            tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
        Else
            '將所有符合條件的訂單號碼串起來
            tmp_Rs.MoveFirst
            Do While Not tmp_Rs.EOF
                str_WhereExternorderkey = str_WhereExternorderkey & "'" & RTrim(tmp_Rs.Fields("externorderkey")) & "',"
                tmp_Rs.MoveNext
            Loop
            str_WhereExternorderkey = Mid(str_WhereExternorderkey, 1, Len(str_WhereExternorderkey) - 1)
            str_WhereExternorderkey = "(" & str_WhereExternorderkey & ")"
            tmp_Rs.Close: Set tmp_Rs = Nothing
        End If
Else
'退貨部份
'實收量=訂單通知量的
        str_SQL = "select externorderkey = rtrim(co.ordertype)+rtrim(co.ExternOrderkey) ,通知量=sum(cast(cod.originalqty as float)), " & _
                  "實收量 = (select isnull(sum(rd.QtyReceived),0) from " & strWMSDB & "..receipt r join " & strWMSDB & "..receiptdetail rd on r.receiptkey = rd.receiptkey and r.status = '9' and r.storerkey = 'LMBO01' and r.receipttype = 'R' where r.externreceiptkey =  rtrim(co.ordertype)+rtrim(co.ExternOrderkey)) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
                    "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
                    "join orders o on co.orderkey = o.orderkey " & _
                    "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
                    "where s2.storerkey = '" & chc_Storerkey & "' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
                    "group by co.ExternOrderkey,co.ordertype"
                    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = adUseClient
        tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic   '
        
        If tmp_Rs.EOF = True Then
            tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
        Else
            '將所有符合條件的訂單號碼串起來
            tmp_Rs.MoveFirst
            Do While Not tmp_Rs.EOF
                If Abs(Val(RTrim(tmp_Rs.Fields("實收量")))) = Abs(Val(RTrim(tmp_Rs.Fields("通知量")))) Then
                    str_WhereExternorderkey = str_WhereExternorderkey & "'" & RTrim(tmp_Rs.Fields("externorderkey")) & "',"
                End If
                tmp_Rs.MoveNext
            Loop
            If str_WhereExternorderkey = "" Then
                tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
            Else
                str_WhereExternorderkey = Mid(str_WhereExternorderkey, 1, Len(str_WhereExternorderkey) - 1)
                str_WhereExternorderkey = "(" & str_WhereExternorderkey & ")"
                tmp_Rs.Close: Set tmp_Rs = Nothing
            End If
        End If
End If


'組合字串
str_SQL = "select distinct " & _
        "'是否回傳' = ' ' " & _
        ",TMS單號=co.orderkey " & _
        ",分分司代號=co.BranchId,訂單號碼=co.ExternOrderkey,訂單日期=isnull(rtrim(cast(cast(convert(char(4),co.OrderDate,112) as int ) - 1911 as char)) + right(convert(char(8),co.OrderDate,112),4),'') " & _
        ",發票號碼=co.Invoice,發票號碼檢查碼=co.InvoiceCheck,發票日期=isnull(rtrim(cast(cast(convert(char(4),co.InvoiceDate,112) as int ) - 1911 as char)) + right(convert(char(8),co.InvoiceDate,112),4),'') " & _
        ",客戶編號=co.Consigneekey,客戶名稱=co.Full_Name,業代代號=co.SalesCode,下貨收現=co.COD,送貨地址=co.Address,聯式=co.Coupled,統一編號=co.VAT " & _
        ",折讓金額=cast(co.Allowance as float),數量折讓金額=cast(co.QuantityAllowance as float),特別折讓金額=cast(co.SpecialAllowance as float),現金折讓=cast(co.CashAllowance as float) " & _
        ",貨款=cast(co.Amount as float),稅前金額=cast(co.NetAmount as float),稅額=cast(co.Tax as float),備註=co.Notes,客戶訂單編號=co.CustOrderkey,隨貨附發票碼=co.InvoiceCode " & _
        ",隨貨附訂單碼=co.OrderCode,計算物流費=co.LogisticsCode,送貨否=co.DeliveryCode,訂單種類=co.OrderType,實收量處理MARK=co.PaidMARK,連絡人=co.Contact " & _
        ",電話=co.Phone1,業代姓名=co.SalesName,主管姓名=co.LeaderName,指送客戶=co.Address2,預計日期=isnull(rtrim(cast(cast(convert(char(4),co.DeliveryDate,112) as int ) - 1911 as char)) + right(convert(char(8),co.DeliveryDate,112),4),'') " & _
        ",運費=co.Freight,付款方式=co.Payment,業務手機=co.SalesPhone,是否為電子發票=co.EInvoiceMark,總重量=cast(co.TotalWeight as float),信卡後4碼=co.Credit_Last4 " & _
        ",代收貨款=cast(co.Cash as float),發票列印方式=co.InvoicePrint,電話2=co.Phone2,統計對象=co.ExternNumber,縣市別=co.City,行政區=co.Administration,樓層=co.Stairs " & _
        ",越庫訂單=co.CrossCode,提貨倉=co.Storage,'稅區/稅率'=co.InvoiceArea,客戶簡稱=co.short_name,實際到貨日=isnull(rtrim(cast(cast(convert(char(4),o.DeliveryDate,112) as int ) - 1911 as char)) + right(convert(char(8),o.DeliveryDate,112),4),''),關聯訂單=rtrim(co.connectorderkey) " & _
        ",狀態 = isnull((select top 1 sdn.confirm_notes from sdn02t sdn where sdn.c_receipt_no  = s2.c_receipt_no and sdn.confirm_notes in ('異常訂單','未出訂單')),'正常訂單') " & _
        ",分類=case when co.ExternNumber in ('10000545','10020700') then 'SNRT' else 'Other' end " & _
        "from sdn02t s2 join CustOrders co on s2.c_receipt_no = co.orderkey " & _
        "join orders o on co.orderkey = o.orderkey " & _
        "where s2.storerkey = '" & chc_Storerkey & "' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
        "and rtrim(co.ordertype) + rtrim(co.externorderkey) in " & str_WhereExternorderkey & _
        "order by co.orderkey "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic   '

If tmp_Rs.EOF = True Then
    tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
Else
    Call Replication_Recordset(tmp_Rs, rsMainHeader)
    tmp_Rs.Close: Set tmp_Rs = Nothing
    
    Set dgMain_Header.DataSource = rsMainHeader: dgMain_Header.Visible = False
    rsMainHeader.MoveFirst
    
    With dgMain_Header
    Set dgMain_Header.DataSource = rsMainHeader
    
    '    .ColumnHeaders = True        '標題行顯示
    '    .RowHeight = 300
    '    .Columns(0).Alignment = dbgCenter
    '    .Columns(10).Alignment = dbgRight
    
    End With
End If
SetDataGridColWidth Me.Caption, dgMain_Header



'明細
If optIn = True Then '退貨
    str_SQL = "select " & _
            "TMS單號=co.orderkey " & _
            ",TMS單號項次 = cod.orderlinenumber " & _
            ",訂單號碼=cod.ExternOrderkey " & _
            ",產品編號=cod.Sku " & _
            ",產品名稱=cod.Descr " & _
            ",訂貨量=cast(cod.OriginalQty as float) " & _
            ",'訂貨量-實收量'= abs(cast(cod.OriginalQty as float)) " & _
            ",'單價(未稅)'=cast(cod.UnitNetPrice as float) " & _
            ",'訂貨金額(未稅)'=cast(cod.NetPrice as float) " & _
            ",'單價(含稅)'=cast(cod.UnitGrossPrice as float) " & _
            ",'訂貨金額(含稅)'=cod.GrossPrice " & _
            ",國際條碼=cod.BarCode " & _
            ",行號=cast(cod.Externlineno as float) " & _
            ",單位=cod.UOM " & _
            ",訂單種類=cod.Ordertype " & _
            ",發票明細列印否=cod.InvoicePCode " & _
            ",允收期=cod.Acceptance " & _
            "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
            "join CustOrders co on s2.c_receipt_no = co.orderkey join orders o on co.orderkey = o.orderkey  " & _
            "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
            "where s2.storerkey = 'LMBO01' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
            "and rtrim(co.ordertype) + rtrim(co.externorderkey) in " & str_WhereExternorderkey & _
            "group by co.orderkey ,cod.orderlinenumber ,cod.ExternOrderkey ,cod.Sku ,cod.Descr ,cast(cod.OriginalQty as float) ,cast(cod.UnitNetPrice as float) , " & _
            "cast(cod.NetPrice as float) ,cod.GrossPrice,cast(cod.UnitGrossPrice as float) ,cod.BarCode ,cast(cod.Externlineno as float) ,cod.UOM ,cod.Ordertype ,cod.InvoicePCode ,cod.Acceptance,co.OrderType,co.ExternOrderkey,abs(cast(cod.OriginalQty as float)) order by co.orderkey,cod.sku"
            
Else
    str_SQL = "select " & _
            "TMS單號=co.orderkey " & _
            ",TMS單號項次 = cod.orderlinenumber " & _
            ",訂單號碼=cod.ExternOrderkey " & _
            ",產品編號=cod.Sku " & _
            ",產品名稱=cod.Descr " & _
            ",訂貨量=cast(cod.OriginalQty as float) " & _
            ",'訂貨量-實收量'=sum(s3.sign_qty) " & _
            ",'單價(未稅)'=cast(cod.UnitNetPrice as float) " & _
            ",'訂貨金額(未稅)'=cast(cod.NetPrice as float) " & _
            ",'單價(含稅)'=cast(cod.UnitGrossPrice as float) " & _
            ",'訂貨金額(含稅)'=cod.GrossPrice " & _
            ",國際條碼=cod.BarCode " & _
            ",行號=cast(cod.Externlineno as float) " & _
            ",單位=cod.UOM " & _
            ",訂單種類=cod.Ordertype " & _
            ",發票明細列印否=cod.InvoicePCode " & _
            ",允收期=cod.Acceptance " & _
            "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
            "join CustOrders co on s2.c_receipt_no = co.orderkey join orders o on co.orderkey = o.orderkey " & _
            "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
            "where s2.storerkey = 'LMBO01' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
            "and rtrim(co.ordertype) + rtrim(co.externorderkey) in " & str_WhereExternorderkey & _
            "group by co.orderkey ,cod.orderlinenumber ,cod.ExternOrderkey ,cod.Sku ,cod.Descr ,cast(cod.OriginalQty as float) ,cast(cod.UnitNetPrice as float) , " & _
            "cast(cod.NetPrice as float) ,cod.GrossPrice,cast(cod.UnitGrossPrice as float) ,cod.BarCode ,cast(cod.Externlineno as float) ,cod.UOM ,cod.Ordertype ,cod.InvoicePCode ,cod.Acceptance order by co.orderkey"
            
End If
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset

If tmp_Rs.EOF = True Then
    tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
Else
    Call Replication_Recordset(tmp_Rs, rsMainDetail)
    tmp_Rs.Close: Set tmp_Rs = Nothing
    
    Set dgMain_Detail.DataSource = rsMainDetail: dgMain_Detail.Visible = False
    rsMainDetail.MoveFirst
    
    With dgMain_Detail
    Set dgMain_Detail.DataSource = rsMainDetail
    
    '    .ColumnHeaders = True        '標題行顯示
    '    .RowHeight = 300
    '    .Columns(0).Alignment = dbgCenter
    '    .Columns(10).Alignment = dbgRight
    
    End With
    
    Call cb_all_Click
    cmdOTUpdate.Enabled = True
End If
'
'Dim str_Orderkey As String, str_externorderkey As String, Int_qty As Integer, bl_next As Boolean, Str_Sku As String
'str_externorderkey = "": str_Orderkey = "": Int_qty = 0: bl_next = True: Str_Sku = ""
'If optIn = True Then '退貨
'    '抓出所有收退未回傳的實收資料
'    Do While Not rsMainHeader.EOF
'        str_externorderkey = str_externorderkey & "'" & RTrim(rsMainHeader.Fields("訂單種類")) & RTrim(rsMainHeader.Fields("訂單號碼")) & "',"
'        rsMainHeader.MoveNext
'    Loop
'    rsMainHeader.MoveFirst
'    rsMainDetail.MoveFirst
'    str_externorderkey = Mid(str_externorderkey, 1, Len(str_externorderkey) - 1)
'    '抓出此批次的實收量
'    Call Confirm_Recordset_Closed(rsMainReceitDetail)
'    rsMainReceitDetail.CursorLocation = adUseClient
'    str_SQL = "select r.externreceiptkey,rd.sku,rd.QtyReceived from " & strWMSDB & "..receipt r join " & strWMSDB & "..receiptdetail rd on r.receiptkey = rd.receiptkey where r.externreceiptkey in (" & str_externorderkey & ") order by r.externreceiptkey,rd.sku"
'    rsMainReceitDetail.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
'
'    '以要回傳的資料為主，補進實收量
'    Do While Not rsMainDetail.EOF
'        '篩選貨主單號有無實收量
'        If bl_next = True Then
'            rsMainReceitDetail.Filter = "externreceiptkey = '" & RTrim(rsMainDetail.Fields("訂單種類")) & RTrim(rsMainDetail.Fields("訂單號碼")) & "'"
'            If rsMainReceitDetail.RecordCount = 0 Then rsMainDetail.Fields("訂貨量-實收量") = 0: GoTo next1         '如果此貨主單號沒有資料，則跳下一筆明細
'
'            rsMainReceitDetail.Filter = "externreceiptkey = '" & RTrim(rsMainDetail.Fields("訂單種類")) & RTrim(rsMainDetail.Fields("訂單號碼")) & "' and sku = '" & RTrim(rsMainDetail.Fields("產品編號")) & "'" '有則挑選特定品號
'            If rsMainReceitDetail.RecordCount = 0 Then rsMainDetail.Fields("訂貨量-實收量") = 0: GoTo next1         '如果此貨主單號沒有品號資料，則跳下一筆明細
'        End If
'
'        If Str_Sku <> RTrim(rsMainDetail.Fields("產品編號")) Then Str_Sku = RTrim(rsMainDetail.Fields("產品編號")): Int_qty = Val(RTrim(rsMainReceitDetail.Fields("QtyReceived")))
'        If Int_qty >= Abs(Val(rsMainDetail.Fields("訂貨量"))) Then
'            rsMainDetail.Fields("訂貨量-實收量") = Abs(Val(rsMainDetail.Fields("訂貨量")))
'            Int_qty = Int_qty - Abs(Val(rsMainDetail.Fields("訂貨量")))
'        Else
'            rsMainDetail.Fields("訂貨量-實收量") = Int_qty
'            Int_qty = 0
'        End If
'
'        If Int_qty = 0 Then bl_next = True Else bl_next = False
'next1:
'        rsMainDetail.MoveNext
'    Loop
'End If


rsMainHeader.MoveFirst
rsMainDetail.MoveFirst

SetDataGridColWidth Me.Caption, dgMain_Detail
StatusBar.Panels(2).Text = rsMainDetail.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain_Detail.Visible = True:: dgMain_Header.Visible = True
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
'Dim dg As Object: Set dg = dgMain
''無資料或欄寬太小，不存寬度
'If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
'SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub



'Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'
'With dgMain
'
'If .DataSource Is Nothing Then Exit Sub
''If LastRow = Empty Then Exit Sub
'If .Row = -1 Or .Col <> 1 Then Exit Sub
'On Error GoTo err_Handle
'
'If .Col = 1 Then
'    If UCase(dgMain) <> "V" And Val(rsMain("出貨件數")) > 0 Then '未選取與件數大於0
'        dgMain = "V"
'    Else
'        dgMain = " "
'
'    End If
'.Col = 0
'End If
'
'End With
'Exit Sub
'
'err_Handle:
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
'End Sub

Private Sub dgMain_Header_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'是否有資料
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.RecordCount = 0 Then Exit Sub

'If blRouteT0Change = False Then Exit Sub

'選取
If dgMain_Header.Col = 1 Then

    If rsMainHeader("是否回傳") = " " Then
    
        rsMainHeader("是否回傳") = "V"
    Else
        rsMainHeader("是否回傳") = " "
    End If
    
    dgMain_Header.Col = 0

End If

''檢查數量
'If rsMainHeader("是否回傳") = "V" Then
'            'rsMainDetail.Filter = "TMS單號 = '" & rsMainHeader.Fields("TMS單號") & "'"
'            rsMainDetail.MoveFirst
'
'            Do While Not rsMainDetail.EOF
'                If rsMainHeader.Fields("TMS單號") = rsMainDetail.Fields("TMS單號") Then
'                    If Abs(Val(rsMainDetail.Fields("訂貨量"))) <> Val(rsMainDetail.Fields("訂貨量-實收量")) Then
'                        x = MsgBox("訂單號碼:" & rsMainDetail.Fields("TMS單號") & "項次:" & rsMainDetail.Fields("TMS單號項次") & ":訂貨量<>簽收量，請確認是否更新回傳?", vbQuestion + vbYesNo, "數量檢查")
'                        If x = 6 Then
'                            '記續
'                        Else
'                            rsMainHeader("是否回傳") = " "
'                            Screen.MousePointer = 0
'                            Exit Sub
'                        End If
'                    End If
'                End If
'                rsMainDetail.MoveNext
'            Loop
'End If
'rsMainDetail.MoveFirst

'同一行選取
If LastRow = Empty Then Exit Sub

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    SSTab1.Height = Frame2.Height - 360
    dgMain_Header.Height = SSTab1.Height - 360
    dgMain_Detail.Height = SSTab1.Height - 360
    
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth
    SSTab1.Width = Frame2.Width - 240
    dgMain_Header.Width = SSTab1.Width - 240
    dgMain_Detail.Width = SSTab1.Width - 240
    
End If

End Sub

Private Sub cmdReset_Click()

'重設
Call ClearForm_AllField(Me)
optNo.Value = True
'optPrintNO.Value = True

End Sub

'Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)
'
'If dgMain.Row = -1 Then Exit Sub
'If intColumnIndex = ColIndex Then
'    rsMain.Sort = dgMain.Columns(ColIndex).Caption & " DESC"
'    dgMain.ClearSelCols
'    intColumnIndex = 255
'
'Else
'    rsMain.Sort = dgMain.Columns(ColIndex).Caption
'    dgMain.ClearSelCols
'    intColumnIndex = ColIndex
'
'End If
'
'End Sub
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

StatusBar.Panels(2).Text = "0 筆資料列"
StatusBar.Panels(3).Text = User_id

Dim i As Integer


'貨主
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(storerkey) from trp16M", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.MoveFirst
For i = 0 To tmp_Rs.RecordCount - 1
    cboStorerkey.AddItem RTrim(tmp_Rs("storerkey"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close: Set tmp_Rs = Nothing
cboStorerkey.Text = "LMBO01"

txtDeliveryDateS = Format(Now - 2, "YYYYMMDD")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMainHeader = Nothing
Set rsMainDetail = Nothing
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
