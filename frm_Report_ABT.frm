VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Report_ABT 
   Caption         =   "ABT需求報表"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "細明體"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   13665
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   7440
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2520
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
      StartOfWeek     =   102039553
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   8
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   " 回單檢核表"
      TabPicture(0)   =   "frm_Report_ABT.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frm_Report_ABT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frm_Report_ABT.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "代收貨款"
      TabPicture(3)   =   "frm_Report_ABT.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "訂單明細表"
      TabPicture(4)   =   "frm_Report_ABT.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame10"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frm_Report_ABT.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame12"
      Tab(5).Control(1)=   "Frame11"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   " "
      TabPicture(6)   =   "frm_Report_ABT.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame14"
      Tab(6).Control(1)=   "Frame13"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   " "
      TabPicture(7)   =   "frm_Report_ABT.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame15"
      Tab(7).Control(1)=   "Frame16"
      Tab(7).ControlCount=   2
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   121
         Top             =   660
         Width           =   13695
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
            Left            =   10080
            Picture         =   "frm_Report_ABT.frx":00E0
            Style           =   1  '圖片外觀
            TabIndex        =   139
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
            Left            =   12480
            Picture         =   "frm_Report_ABT.frx":03EA
            Style           =   1  '圖片外觀
            TabIndex        =   138
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
            Left            =   12480
            Picture         =   "frm_Report_ABT.frx":06FC
            Style           =   1  '圖片外觀
            TabIndex        =   137
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
            Left            =   11280
            Picture         =   "frm_Report_ABT.frx":2A30E
            Style           =   1  '圖片外觀
            TabIndex        =   136
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtOrderDateS 
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
            TabIndex        =   135
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateE 
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
            TabIndex        =   134
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.ComboBox Combo1 
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
            Left            =   1200
            Style           =   2  '單純下拉式
            TabIndex        =   133
            Top             =   240
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
            TabIndex        =   132
            Top             =   960
            Visible         =   0   'False
            Width           =   1485
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
            TabIndex        =   131
            Top             =   960
            Width           =   1485
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
            Left            =   10080
            Picture         =   "frm_Report_ABT.frx":2B608
            Style           =   1  '圖片外觀
            TabIndex        =   130
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
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
            Left            =   4680
            Style           =   1  '項目包含核取方塊
            TabIndex        =   129
            ToolTipText     =   "區碼"
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmdSaveToText 
            BackColor       =   &H00C0E0FF&
            Caption         =   "檢核表"
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
            Left            =   11280
            Picture         =   "frm_Report_ABT.frx":2B912
            Style           =   1  '圖片外觀
            TabIndex        =   128
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CheckBox optNormal 
            Caption         =   "正常簽單"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox optAbnormal 
            Caption         =   "異常簽單"
            Height          =   255
            Left            =   1200
            TabIndex        =   126
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox optNotYet 
            Caption         =   "未確認簽單"
            Height          =   255
            Left            =   2280
            TabIndex        =   125
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
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
            ItemData        =   "frm_Report_ABT.frx":2BC1C
            Left            =   1200
            List            =   "frm_Report_ABT.frx":2BC26
            Style           =   2  '單純下拉式
            TabIndex        =   124
            Top             =   1680
            Visible         =   0   'False
            Width           =   2325
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
            ItemData        =   "frm_Report_ABT.frx":2BC4E
            Left            =   6480
            List            =   "frm_Report_ABT.frx":2BC50
            Style           =   1  '項目包含核取方塊
            TabIndex        =   123
            ToolTipText     =   "貨運公司"
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ListBox List3 
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
            ItemData        =   "frm_Report_ABT.frx":2BC52
            Left            =   9000
            List            =   "frm_Report_ABT.frx":2BC54
            Style           =   1  '項目包含核取方塊
            TabIndex        =   122
            ToolTipText     =   "訂單類別"
            Top             =   240
            Visible         =   0   'False
            Width           =   975
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
            Left            =   2655
            TabIndex        =   146
            Top             =   660
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "維護日期"
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
            TabIndex        =   145
            Top             =   645
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "區域"
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
            Left            =   360
            TabIndex        =   144
            Top             =   300
            Width           =   480
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
            Index           =   30
            Left            =   120
            TabIndex        =   143
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
            Left            =   2640
            TabIndex        =   142
            Top             =   1020
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "需作完出車確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   5
            Left            =   2880
            TabIndex        =   141
            Top             =   240
            Width           =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "排序"
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
            Left            =   360
            TabIndex        =   140
            Top             =   1740
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame Frame15 
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   80
         Top             =   720
         Width           =   8295
         Begin VB.CheckBox chkT7 
            Caption         =   "僅WH-Y地段"
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtDeliveryDateST7 
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
            TabIndex        =   87
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateET7 
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
            TabIndex        =   86
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToTextT7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "轉文字檔"
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":2BC56
            Style           =   1  '圖片外觀
            TabIndex        =   85
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT7 
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":2BF60
            Style           =   1  '圖片外觀
            TabIndex        =   84
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT7 
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
            Left            =   4680
            Picture         =   "frm_Report_ABT.frx":2D25A
            Style           =   1  '圖片外觀
            TabIndex        =   83
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
            Index           =   7
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":2D564
            Style           =   1  '圖片外觀
            TabIndex        =   82
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdResetT7 
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
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":57176
            Style           =   1  '圖片外觀
            TabIndex        =   81
            Top             =   240
            Width           =   1065
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
            Left            =   2640
            TabIndex        =   89
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "庫存日期"
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
            TabIndex        =   88
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame16 
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
         Left            =   -74880
         TabIndex        =   78
         Top             =   2880
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT7 
            Height          =   2295
            Left            =   120
            TabIndex        =   79
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
      Begin VB.Frame Frame14 
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
         Left            =   -74880
         TabIndex        =   76
         Top             =   2880
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT6 
            Height          =   2295
            Left            =   120
            TabIndex        =   77
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
      Begin VB.Frame Frame13 
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   66
         Top             =   720
         Width           =   8295
         Begin VB.CheckBox chkT6 
            Caption         =   "僅WH-Y地段"
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   96
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdResetT6 
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
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":57488
            Style           =   1  '圖片外觀
            TabIndex        =   73
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
            Index           =   6
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":5779A
            Style           =   1  '圖片外觀
            TabIndex        =   72
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT6 
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
            Left            =   4680
            Picture         =   "frm_Report_ABT.frx":813AC
            Style           =   1  '圖片外觀
            TabIndex        =   71
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT6 
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":816B6
            Style           =   1  '圖片外觀
            TabIndex        =   70
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdSaveToTextT6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "轉文字檔"
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":829B0
            Style           =   1  '圖片外觀
            TabIndex        =   69
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtDeliveryDateET6 
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
            TabIndex        =   68
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST6 
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
            TabIndex        =   67
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "排除品號第一碼英文開頭的商品"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   18
            Left            =   165
            TabIndex        =   98
            Top             =   720
            Width           =   3360
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
            Index           =   15
            Left            =   120
            TabIndex        =   75
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
            Index           =   14
            Left            =   2640
            TabIndex        =   74
            Top             =   1020
            Width           =   360
         End
      End
      Begin VB.Frame Frame11 
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   56
         Top             =   660
         Width           =   8295
         Begin VB.CheckBox chkT5 
            Caption         =   "僅WH-Y地段"
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   97
            Top             =   1320
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtDeliveryDateST5 
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
            TabIndex        =   63
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateET5 
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
            TabIndex        =   62
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToTextT5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "轉文字檔"
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":82CBA
            Style           =   1  '圖片外觀
            TabIndex        =   61
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT5 
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":82FC4
            Style           =   1  '圖片外觀
            TabIndex        =   60
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT5 
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
            Left            =   4680
            Picture         =   "frm_Report_ABT.frx":842BE
            Style           =   1  '圖片外觀
            TabIndex        =   59
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
            Index           =   5
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":845C8
            Style           =   1  '圖片外觀
            TabIndex        =   58
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdResetT5 
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
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":AE1DA
            Style           =   1  '圖片外觀
            TabIndex        =   57
            Top             =   240
            Width           =   1065
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
            Index           =   13
            Left            =   2640
            TabIndex        =   65
            Top             =   1020
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
            Index           =   12
            Left            =   120
            TabIndex        =   64
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame12 
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
         Left            =   -74880
         TabIndex        =   54
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT5 
            Height          =   2295
            Left            =   120
            TabIndex        =   55
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
         TabIndex        =   42
         Top             =   2820
         Width           =   8295
         Begin TabDlg.SSTab SSTab1 
            Height          =   3735
            Left            =   0
            TabIndex        =   147
            Top             =   120
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   6588
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "其他明細表"
            TabPicture(0)   =   "frm_Report_ABT.frx":AE4EC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "dgMainT4"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "杏一明細表"
            TabPicture(1)   =   "frm_Report_ABT.frx":AE508
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "dgMainT4_1"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "維康明細表"
            TabPicture(2)   =   "frm_Report_ABT.frx":AE524
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "dgMainT4_2"
            Tab(2).ControlCount=   1
            Begin MSDataGridLib.DataGrid dgMainT4 
               Height          =   2295
               Left            =   120
               TabIndex        =   148
               Top             =   360
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
            Begin MSDataGridLib.DataGrid dgMainT4_1 
               Height          =   2295
               Left            =   -74880
               TabIndex        =   149
               Top             =   360
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
            Begin MSDataGridLib.DataGrid dgMainT4_2 
               Height          =   2295
               Left            =   -74880
               TabIndex        =   150
               Top             =   360
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
      Begin VB.Frame Frame9 
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
         Height          =   2175
         Left            =   120
         TabIndex        =   43
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtNotCarNo 
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
            TabIndex        =   151
            Top             =   600
            Width           =   3165
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
            Left            =   6360
            TabIndex        =   115
            Top             =   1680
            Width           =   735
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
            Left            =   4560
            TabIndex        =   114
            Top             =   1680
            Width           =   855
         End
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
            Left            =   5400
            TabIndex        =   113
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox cboCarT4 
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
            TabIndex        =   111
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtRouteST4 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   104
            Top             =   2040
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtRouteET4 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   103
            Top             =   2040
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateET4 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
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
            TabIndex        =   102
            Top             =   1680
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST4 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
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
            TabIndex        =   101
            Top             =   1680
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryST4 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
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
            TabIndex        =   100
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryET4 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
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
            TabIndex        =   99
            Top             =   1320
            Width           =   1485
         End
         Begin VB.CommandButton cmdResetT4 
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
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":AE540
            Style           =   1  '圖片外觀
            TabIndex        =   50
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
            Index           =   4
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":AE852
            Style           =   1  '圖片外觀
            TabIndex        =   49
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT4 
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
            Left            =   4680
            Picture         =   "frm_Report_ABT.frx":D8464
            Style           =   1  '圖片外觀
            TabIndex        =   48
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT4 
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":D876E
            Style           =   1  '圖片外觀
            TabIndex        =   47
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdSaveToTextT4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "轉文字檔"
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
            Picture         =   "frm_Report_ABT.frx":D9A68
            Style           =   1  '圖片外觀
            TabIndex        =   46
            Top             =   960
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtAddDateET4 
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
            TabIndex        =   45
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtAddDateST4 
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
            TabIndex        =   44
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "，"
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
            Left            =   360
            TabIndex        =   153
            Top             =   660
            Width           =   8280
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "排除車號"
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
            Left            =   120
            TabIndex        =   152
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "件數確認"
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
            Left            =   4560
            TabIndex        =   116
            Top             =   1320
            Width           =   960
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
            Index           =   26
            Left            =   120
            TabIndex        =   112
            Top             =   300
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
            Index           =   25
            Left            =   5175
            TabIndex        =   110
            Top             =   2100
            Visible         =   0   'False
            Width           =   360
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
            Index           =   24
            Left            =   2640
            TabIndex        =   109
            Top             =   2085
            Visible         =   0   'False
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
            Index           =   23
            Left            =   120
            TabIndex        =   108
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
            Index           =   22
            Left            =   2640
            TabIndex        =   107
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
            Index           =   21
            Left            =   2640
            TabIndex        =   106
            Top             =   1380
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
            Index           =   19
            Left            =   120
            TabIndex        =   105
            Top             =   1365
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "接單日期"
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
            Left            =   120
            TabIndex        =   52
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
            Index           =   10
            Left            =   2640
            TabIndex        =   51
            Top             =   1020
            Width           =   360
         End
      End
      Begin VB.Frame Frame8 
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT3 
            Height          =   2295
            Left            =   120
            TabIndex        =   39
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
      Begin VB.Frame Frame7 
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   28
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtOrderDateST3 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
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
            TabIndex        =   118
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateET3 
            Alignment       =   2  '置中對齊
            BeginProperty Font 
               Name            =   "細明體"
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
            TabIndex        =   117
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST3 
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
            TabIndex        =   35
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateET3 
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
            TabIndex        =   34
            Top             =   1320
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToTextT3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "轉文字檔"
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":D9D72
            Style           =   1  '圖片外觀
            TabIndex        =   33
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT3 
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":DA07C
            Style           =   1  '圖片外觀
            TabIndex        =   32
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT3 
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
            Left            =   4680
            Picture         =   "frm_Report_ABT.frx":DB376
            Style           =   1  '圖片外觀
            TabIndex        =   31
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
            Index           =   3
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":DB680
            Style           =   1  '圖片外觀
            TabIndex        =   30
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdResetT3 
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
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":105292
            Style           =   1  '圖片外觀
            TabIndex        =   29
            Top             =   240
            Width           =   1065
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
            Index           =   29
            Left            =   2640
            TabIndex        =   120
            Top             =   1020
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
            Index           =   28
            Left            =   120
            TabIndex        =   119
            Top             =   1365
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
            Index           =   9
            Left            =   2640
            TabIndex        =   37
            Top             =   1380
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單日期"
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
            TabIndex        =   36
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame6 
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
         Left            =   -74880
         TabIndex        =   24
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT2 
            Height          =   2295
            Left            =   120
            TabIndex        =   25
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
      Begin VB.Frame Frame5 
         BackColor       =   &H80000004&
         Caption         =   "配送異常表"
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   15
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtDeliveryDateET2 
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
            TabIndex        =   91
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST2 
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
            TabIndex        =   90
            Top             =   1320
            Width           =   1485
         End
         Begin VB.CommandButton cmdResetT2 
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
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":1055A4
            Style           =   1  '圖片外觀
            TabIndex        =   26
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtSdnDateST2 
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
            TabIndex        =   21
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtSdnDateET2 
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
            TabIndex        =   20
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToTextT2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "轉文字檔"
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":1058B6
            Style           =   1  '圖片外觀
            TabIndex        =   19
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT2 
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":105BC0
            Style           =   1  '圖片外觀
            TabIndex        =   18
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT2 
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
            Left            =   4680
            Picture         =   "frm_Report_ABT.frx":106EBA
            Style           =   1  '圖片外觀
            TabIndex        =   17
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
            Index           =   1
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":1071C4
            Style           =   1  '圖片外觀
            TabIndex        =   16
            Top             =   1200
            Width           =   1065
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
            Index           =   20
            Left            =   120
            TabIndex        =   93
            Top             =   1365
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
            Index           =   7
            Left            =   2640
            TabIndex        =   92
            Top             =   1380
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
            Index           =   6
            Left            =   2640
            TabIndex        =   23
            Top             =   1020
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
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000004&
         Caption         =   "請付款資料明細"
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   6
         Top             =   660
         Width           =   8295
         Begin VB.CheckBox optRepack 
            Caption         =   "加工計費"
            Height          =   255
            Left            =   3360
            TabIndex        =   94
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox optTMS 
            Caption         =   "運輸請款"
            Height          =   255
            Left            =   1200
            TabIndex        =   41
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox optWMS 
            Caption         =   "倉儲請款"
            Height          =   255
            Left            =   2280
            TabIndex        =   40
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton cmdResetT1 
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
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":130DD6
            Style           =   1  '圖片外觀
            TabIndex        =   27
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
            Index           =   2
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":1310E8
            Style           =   1  '圖片外觀
            TabIndex        =   14
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT1 
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
            Left            =   4680
            Picture         =   "frm_Report_ABT.frx":15ACFA
            Style           =   1  '圖片外觀
            TabIndex        =   11
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT1 
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":15B004
            Style           =   1  '圖片外觀
            TabIndex        =   10
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdSaveToTextT1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "轉文字檔"
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
            Left            =   5880
            Picture         =   "frm_Report_ABT.frx":15C2FE
            Style           =   1  '圖片外觀
            TabIndex        =   9
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtDeliveryDateET1 
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
            TabIndex        =   8
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST1 
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
            TabIndex        =   7
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "計費區間"
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
            TabIndex        =   13
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
            Index           =   4
            Left            =   2640
            TabIndex        =   12
            Top             =   1020
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
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
         Left            =   -74880
         TabIndex        =   4
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT1 
            Height          =   2295
            Left            =   120
            TabIndex        =   5
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
         Left            =   -74880
         TabIndex        =   2
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMain 
            Height          =   2295
            Left            =   120
            TabIndex        =   3
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
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7605
      Width           =   13665
      _ExtentX        =   24104
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
            Object.Width           =   17489
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
Attribute VB_Name = "frm_Report_ABT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private rsMainT1 As ADODB.Recordset
Private rsMainT2 As ADODB.Recordset
Private rsMainT3 As ADODB.Recordset
Private rsMainT4 As ADODB.Recordset
Private rsMainT4_1 As ADODB.Recordset
Private rsMainT4_2 As ADODB.Recordset
Private rsMainT5 As ADODB.Recordset
Private rsMainT6 As ADODB.Recordset
Private rsMainT7 As ADODB.Recordset

Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub cmdExit_Click(Index As Integer)
Unload Me '結束此程序
'End 結束應用程式
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set rsMain = Nothing
Set rsMainT1 = Nothing
Set rsMainT2 = Nothing
Set rsMainT3 = Nothing
Set rsMainT4 = Nothing
Set rsMainT5 = Nothing
Set rsMainT6 = Nothing
Set rsMainT7 = Nothing

End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 筆資料列"
StatusBar.Panels(3).Text = User_id

SSTab.Tab = 0

'取車號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select distinct vehicle_id_no from trp05t order by vehicle_id_no "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then tmp_Rs.MoveFirst

Do While Not tmp_Rs.EOF
    cboCarT4.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    tmp_Rs.MoveNext
Loop
cboCarT4.AddItem "待排車"

cboCarT4 = ""

tmp_Rs.Close

'貨主
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
'tmp_Rs.Open "select distinct(storerkey) from trp16M where storerkey = 'LABT01' ", cn, adOpenKeyset, adLockPessimistic
'
'If Not tmp_Rs.EOF Then
'    tmp_Rs.MoveFirst
'    For i = 0 To tmp_Rs.RecordCount - 1
'        Combo1.AddItem tmp_Rs("storerkey")
'        tmp_Rs.MoveNext
'    Next
'    Combo1.ListIndex = 0
'End If
'tmp_Rs.Close

Combo1.AddItem "信速"
Combo1.AddItem "中區"
Combo1.AddItem "德迅"
Combo1.ListIndex = 0
    
''區域
'With tmp_Rs
'    .Open "select area_code from trp03m order by area_code ", cn
'
'    If Not .EOF Then
'        .MoveFirst
'        For i = 0 To .RecordCount - 1
'            List1.AddItem RTrim(tmp_Rs("area_code"))
'            .MoveNext
'        Next
'
'    End If
'    .Close
'
''貨運公司
'    .Open "select company_code,short_name from trp08m order by company_code ", cn
'
'If Not .EOF Then
'    .MoveFirst
'    For i = 0 To .RecordCount - 1
'        List2.AddItem RTrim(tmp_Rs("company_code")) & "_" & RTrim(tmp_Rs("short_name"))
'        .MoveNext
'    Next
'End If
'.Close
'
''單別
'    .Open "select distinct rtrim(isnull(priority,'')) as Priority from sdn02t order by priority ", cn
'
'If Not .EOF Then
'    .MoveFirst
'    For i = 0 To .RecordCount - 1
'        List3.AddItem RTrim(tmp_Rs("Priority"))
'        .MoveNext
'    Next
'End If
'.Close
'
'End With

Combo2.ListIndex = 0
optNormal = 1
optAbnormal = 1
txtDeliveryDateS = Format(Now - 1, "YYYYMMDD")
txtOrderDateST3 = Format(Now - 1, "yyyymmdd")
'txtOrderDateET3 = Format(Now, "yyyymmdd")
txtAddDateST4 = Format(Now, "yyyymmdd")
'txtDeliveryDateET4 = Format(Now + 1, "yyyymmdd")
txtDeliveryDateST5 = Format(Now, "yyyymm") + "01"
txtDeliveryDateET5 = Format(Now + 1, "yyyymmdd")
txtDeliveryDateST6 = Format(Now, "yyyymm") + "01"
txtDeliveryDateET6 = Format(Now + 1, "yyyymmdd")
txtDeliveryDateST7 = Format(Now, "yyyymm") + "01"
txtDeliveryDateET7 = Format(Now, "yyyymmdd")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub SSTab_Click(PreviousTab As Integer)

Me.mvDate.Visible = False
If Len(Trim(SSTab.Caption)) = 0 Then SSTab.Tab = PreviousTab: Exit Sub

StatusBar.Panels(2).Text = "0 筆資料列"
If SSTab.Tab = 0 And (rsMain Is Nothing) = False Then StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
If SSTab.Tab = 1 And (rsMainT1 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT1.RecordCount & " 筆資料列"
If SSTab.Tab = 2 And (rsMainT2 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT2.RecordCount & " 筆資料列"
If SSTab.Tab = 3 And (rsMainT3 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT3.RecordCount & " 筆資料列"
If SSTab.Tab = 4 And (rsMainT4 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT4.RecordCount & " 筆資料列"
If SSTab.Tab = 5 And (rsMainT5 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT5.RecordCount & " 筆資料列"
If SSTab.Tab = 6 And (rsMainT6 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT6.RecordCount & " 筆資料列"
If SSTab.Tab = 7 And (rsMainT7 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT7.RecordCount & " 筆資料列"
    
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    SSTab.Height = Me.ScaleHeight - StatusBar.Height
    Frame2.Height = SSTab.Height - Frame1.Height - Frame1.Top - 120: dgMain.Height = Frame2.Height - 360
    Frame4.Height = SSTab.Height - Frame3.Height - Frame1.Top - 120: dgMainT1.Height = Frame4.Height - 360
    Frame6.Height = SSTab.Height - Frame5.Height - Frame1.Top - 120: dgMainT2.Height = Frame6.Height - 360
    Frame8.Height = SSTab.Height - Frame7.Height - Frame1.Top - 120: dgMainT3.Height = Frame8.Height - 360
    Frame10.Height = SSTab.Height - Frame9.Height - Frame1.Top - 120: dgMainT4.Height = Frame10.Height - 360
    SSTab1.Height = SSTab.Height - Frame9.Height - Frame1.Top - 240: dgMainT4.Height = Frame10.Height - 480: dgMainT4_1.Height = Frame10.Height - 480: dgMainT4_2.Height = Frame10.Height - 480:
    Frame12.Height = SSTab.Height - Frame11.Height - Frame1.Top - 120: dgMainT5.Height = Frame12.Height - 360
    Frame14.Height = SSTab.Height - Frame13.Height - Frame1.Top - 120: dgMainT6.Height = Frame14.Height - 360
    Frame16.Height = SSTab.Height - Frame15.Height - Frame1.Top - 120: dgMainT7.Height = Frame16.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab.Width = Me.ScaleWidth
    Frame2.Width = SSTab.Width - 360: dgMain.Width = Frame2.Width - 240
    Frame4.Width = SSTab.Width - 360: dgMainT1.Width = Frame4.Width - 240
    Frame6.Width = SSTab.Width - 360: dgMainT2.Width = Frame6.Width - 240
    Frame8.Width = SSTab.Width - 360: dgMainT3.Width = Frame8.Width - 240
    Frame10.Width = SSTab.Width - 360: dgMainT4.Width = Frame10.Width - 240
    SSTab1.Width = SSTab.Width - 360: dgMainT4.Width = Frame10.Width - 240: dgMainT4_1.Width = Frame10.Width - 240: dgMainT4_2.Width = Frame10.Width - 240
    Frame12.Width = SSTab.Width - 360: dgMainT5.Width = Frame12.Width - 240
    Frame14.Width = SSTab.Width - 360: dgMainT6.Width = Frame14.Width - 240
    Frame16.Width = SSTab.Width - 360: dgMainT7.Width = Frame16.Width - 240
End If

End Sub

Private Sub cmdReset_Click()

'重設
txtDeliveryDateS = "": txtDeliveryDateE = ""

End Sub

Private Sub cmdResetT1_Click()
'重設
txtDeliveryDateST1 = "": txtDeliveryDateET1 = ""
End Sub

Private Sub cmdResetT2_Click()
'重設
txtSdnDateST2 = "": txtSdnDateET2 = ""
txtDeliveryDateST2 = "": txtDeliveryDateET2 = ""
End Sub

Private Sub cmdResetT3_Click()
'重設
txtOrderDateST3 = "": txtOrderDateET3 = ""
txtDeliveryDateST3 = "": txtDeliveryDateET3 = ""
End Sub
Private Sub cmdResetT4_Click()
'重設
cboCarT4 = ""
txtAddDateST4 = "": txtAddDateET4 = ""
txtDeliveryDateST4 = "": txtDeliveryDateET4 = ""
txtDeliveryST4 = "": txtDeliveryET4 = ""
txtRouteST4 = "": txtRouteET4 = ""
End Sub
Private Sub cmdResetT5_Click()
'重設
txtDeliveryDateST5 = "": txtDeliveryDateET5 = ""
End Sub
Private Sub cmdResetT6_Click()
'重設
txtDeliveryDateST6 = "": txtDeliveryDateET6 = ""
End Sub
Private Sub cmdResetT7_Click()
'重設
txtDeliveryDateST7 = "": txtDeliveryDateET7 = ""
End Sub
Private Sub cmd2Excel_Click()

'資料排序
Recordset2Excel "LABT回單檢核表", rsMain
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub
Private Sub cmd2ExcelT1_Click()
If optTMS + optWMS + optRepack = 0 Then MsgBox "請選擇請款報表類別！", vbOKOnly, Me.Caption: Exit Sub
If rsMainT1 Is Nothing Then MsgBox "無資料可供轉檔！", vbOKOnly + vbInformation, "Save2Excel": Exit Sub

MsgBox "系統進行大量資料轉Excel時，請勿操作其他Excel作業，以免資料轉出錯誤！", vbOKOnly + vbInformation, "Save2Excel"

On Error GoTo err_Handle
Screen.MousePointer = 11
Dim rsTmp As New ADODB.Recordset
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    
    If Dir(App.Path & "\XLT\雀巢請付款明細.xlt") = "" Then '找不到本機範例檔
        
        '取範例檔路徑
        Dim objIni As vbIniFile, strXltPath As String
        Set objIni = New vbIniFile
        
        With objIni
        
            .FileName = striniFileName_FullPath
            strXltPath = RTrim(.ReadData("EXCEL", "XLTPATH", "")) '不支援中文資料夾名稱
            
        End With
        Set objIni = Nothing

    End If

    '無指定路徑使用本機路徑
    If Len(RTrim(strXltPath)) = 0 Then strXltPath = App.Path & "\XLT"
    
    '尋找本機範例檔
    If Dir(strXltPath & "\雀巢請付款明細.xlt") <> "" Then
        
        '開啟範例檔
        .Workbooks.Open (strXltPath & "\雀巢請付款明細.xlt")
    Else
        '新增Excel
        .Workbooks.Add
    End If
    
.ActiveWorkbook.Author = User_id

'TMS請款報表
If optTMS = 1 Then

    '雀巢計費明細資料
    '尋找工作表
    strSheet = "運費明細資料"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then .Sheets.Add: .ActiveSheet.Name = strSheet

    Call WriteOut_RunLog("運輸請款：1/5.運費明細資料..")
    rsMainT1.MoveFirst
    Call OffLineRecordset(rsMainT1, rsTmp)
    
    '寫入標題列
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close

    '日報表
    '尋找工作表
    strSheet = "日報表"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
'    str_SQL = "select * from gv_Charge where 貨主 = 'LNSL01' and 載貨日期 between '" & txtDeliveryDateST1 & "' and '" & txtDeliveryDateET1 & "' order by 載貨日期,車號,請款類別 "
    str_SQL = "exec gs_Charge 'LNSL01' , '" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    Call WriteOut_RunLog("運輸請款：2/5.轉出日報表..")
    tmp_Rs.CursorLocation = adUseClient
    tmp_Rs.Open str_SQL, cn
    tmp_Rs.Sort = "載貨日期,車號,請款類別"
    Call OffLineRecordset(tmp_Rs, rsTmp)
    tmp_Rs.Sort = ""
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'訂單配送
    Screen.MousePointer = 11
    '尋找工作表
    strSheet = "訂單配送"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01ShipCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("運輸請款：3/5.轉出配送費...")
    tmp_Rs.Open str_SQL, cn
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close

'提貨
    Screen.MousePointer = 11
    '尋找工作表
    strSheet = "提貨"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01RCCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("運輸請款：4/5.轉出提貨費....")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'退貨配送
    Screen.MousePointer = 11
    '尋找工作表
    strSheet = "退貨"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01returnCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("運輸請款：5/5.轉出配送費...")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
        
End If

'WMS請款報表
If optWMS = 1 Then
    '進貨
    Screen.MousePointer = 11
        '尋找工作表
        strSheet = "進貨"
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
        Next
    
        '找不到新增工作表
        If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
        
        str_SQL = "exec gs_LNSL01ReceiptDetailCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
    
        Call WriteOut_RunLog("倉儲請款：1/5.轉出進貨資料")
        tmp_Rs.Open str_SQL, cn
        
        Call Replication_Recordset(tmp_Rs, rsTmp)
    
        '寫入標題列
        k = 65: j = 1
        For i = 0 To rsTmp.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
            '欄位超過26
            If Chr(65 + l) = "Z" Then
                If strCol = "" Then
                    strCol = "A"
                Else
                    strCol = Chr(Asc(strCol) + 1)
                End If
            End If
        Next i
    
        .Range("A2").CopyFromRecordset rsTmp
    
        rsTmp.Close

'出貨理貨
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "出貨理貨"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01PickingCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("倉儲請款：2/5.轉出出貨理貨費資料")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp
    rsTmp.Close

'提貨明細
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "提貨明細"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01RCDetailCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("倉儲請款：3/5.轉出提貨明細資料")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp
rsTmp.Close

'退貨
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "退貨明細"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
    str_SQL = "exec gs_LNSL01ReturnReceiptDetailCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("倉儲請款：4/5.轉出退貨明細資料")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp

rsTmp.Close

'收貨明細-參考
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "收貨明細-參考"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
    str_SQL = "exec gs_LNSL01ReceiptDetail '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("倉儲請款：5/5.轉出進貨明細參考資料")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp
    rsTmp.Close
End If

'加工計費報表
If optRepack = 1 Then
    '進貨
    Screen.MousePointer = 11
        '尋找工作表
        strSheet = "NPP加工計費"
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
        Next
    
        '找不到新增工作表
        If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
        
        str_SQL = "exec gs_LNSL01repackcharge01 '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
    
        Call WriteOut_RunLog("加工計費：1/3.轉出NPP加工計費")
        tmp_Rs.Open str_SQL, cn
        
        Call Replication_Recordset(tmp_Rs, rsTmp)
    
        '寫入標題列
        k = 65: j = 1
        For i = 0 To rsTmp.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
            '欄位超過26
            If Chr(65 + l) = "Z" Then
                If strCol = "" Then
                    strCol = "A"
                Else
                    strCol = Chr(Asc(strCol) + 1)
                End If
            End If
        Next i
    
        .Range("A2").CopyFromRecordset rsTmp
    
        rsTmp.Close

'非NPP加工計費資料
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "非NPP加工計費"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01repackcharge02 '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("加工計費：2/3.轉出非NPP加工計費")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp
    rsTmp.Close

'一般加工計費
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "加工計費明細"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_Repackcharge 'LNSL01','" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("加工計費：3/3.轉出加工計費明細")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp
rsTmp.Close

End If

.Visible = True: End With

Call Unload_RunLogForm
Set MyXlsApp = Nothing
Screen.MousePointer = 0

Exit Sub

err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmd2ExcelT2_Click()

'資料排序
Recordset2Excel "LNSL01簽單回傳", rsMainT2
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT3_Click()

'資料排序
Recordset2Excel "LABT01代收貨款", rsMainT3
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT4_Click()

'資料排序
Screen.MousePointer = 11
dgMainT4.Visible = False
dgMainT4_1.Visible = False
dgMainT4_2.Visible = False
Recordset2Excel_ABT "其他訂單明細", rsMainT4
Recordset2Excel_ABT "杏一訂單明細", rsMainT4_1
Recordset2Excel_ABT "維康訂單明細", rsMainT4_2
dgMainT4.Visible = True
dgMainT4_1.Visible = True
dgMainT4_2.Visible = True
'..在此編輯EXCEL
Screen.MousePointer = 0
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT5_Click()

'資料排序
Recordset2Excel "LNSL01_DailyShippingReport", rsMainT5
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT6_Click()

'資料排序
Recordset2Excel "LNSL01_DailyGoodsArriveReport", rsMainT6
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT7_Click()

'資料排序
Recordset2Excel "LNSL01_DailyStorageStatusReport", rsMainT7
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
'Dim chc_Orderdate As String, chc_DeliveryDate As String, i As Integer, strSelected As String
Dim strSelected As String
strSelected = ""


''區碼
'For i = 0 To List1.ListCount - 1
'    If List1.Selected(i) Then strSelected = strSelected & "'" & Left(List1.List(i), 2) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t1m.area_code in ( " & strSelected & "'') "
'
''貨運公司
'strSelected = ""
'For i = 0 To List2.ListCount - 1
'    If List2.Selected(i) Then strSelected = strSelected & "'" & mySplit(List2.List(i), "_", 0) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t8m.company_code in ( " & strSelected & "'') "
'
''單別
'strSelected = ""
'For i = 0 To List3.ListCount - 1
'    If List3.Selected(i) Then strSelected = strSelected & "'" & mySplit(List3.List(i), "_", 0) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and isnull(s2.priority,'') in ( " & strSelected & "'') "
'
''維護日期
'chc_Orderdate = ""
'If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
'   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
'ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
'   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateS.Text & "' "
'ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
'   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateE.Text & "' "
'End If
'
''到貨日期
'chc_DeliveryDate = ""
'If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
'   chc_DeliveryDate = "and convert(Char(8),s2.arrive_date,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
'ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
'   chc_DeliveryDate = "and convert(Char(8),arrive_date,112) = '" & txtDeliveryDateS.Text & "' "
'ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
'   chc_DeliveryDate = "and convert(Char(8),arrive_date,112) = '" & txtDeliveryDateE.Text & "' "
'End If
'
''簽單類別
'If optNormal = 0 And optAbnormal = 0 And optNotYet = 0 Then GoTo NextStep
'Dim strStatus As String
'
'strStatus = "and s2.confirm_notes in ("
'
'If optNormal = 1 Then strStatus = strStatus & "'正常訂單',"
'If optAbnormal = 1 Then strStatus = strStatus & "'異常訂單','未出訂單',"
'If optNotYet = 1 Then strStatus = strStatus & "'',"
'
'str_SQL = str_SQL & Left(strStatus, Len(strStatus) - 1) & ")"
'
'NextStep:
'
''貨主
'If Len(RTrim(Combo1.Text)) > 0 Then str_SQL = str_SQL & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' "
'
'If Combo2.Text = "使用者、維護時間" Then
'    str_SQL = str_SQL & "order by s2.confirm_userid,isnull(convert(char(19),s2.confirm_date,121),'') "
'Else
'    str_SQL = str_SQL & "order by isnull(t1m.channel,''),isnull(t1m.short_name,'') "
'End If

If Combo1 = "" Then MsgBox "請選擇配送公司!", 16, Me.Caption: Screen.MousePointer = 0: Exit Sub
If txtDeliveryDateS.Text = "" Then MsgBox "請輸入到貨日期!", 16, Me.Caption: Screen.MousePointer = 0: Exit Sub

'If Combo1 = "信速" Then str_SQL = "exec gs_LABT01SdnList '" & txtDeliveryDateS.Text & "','002-10'"
'If Combo1 = "中區" Then str_SQL = "exec gs_LABT01SdnList '" & txtDeliveryDateS.Text & "','000-31'"
'If Combo1 = "德迅" Then str_SQL = "exec gs_LABT01SdnList '" & txtDeliveryDateS.Text & "','000-70'"

'車號條件edit by Eric 20150112,不使用SP
If Combo1 = "信速" Then strSelected = strSelected & "and s2.vehicle_id_no = '002-10' "
If Combo1 = "中區" Then strSelected = strSelected & "and s2.vehicle_id_no = '000-31' "
If Combo1 = "德迅" Then strSelected = strSelected & "and s2.vehicle_id_no = '000-70' "

'日期條件
strSelected = strSelected & "and rtrim(isnull(s2.arrive_date,'')) = '" & txtDeliveryDateS.Text & "'"

str_SQL = "select " & _
            "客戶名稱 = rtrim(isnull(t1m.short_name,'')) " & _
            ",到貨日 = rtrim(isnull(s2.arrive_date,'')) " & _
            ",訂單類別 = rtrim(isnull(s2.priority,'')) " & _
            ",出貨箱數 = isnull((select sum(otqty) from ort02t where ort02t.receipt_no = s2.receipt_no),0) " & _
            ",訂單號碼 = rtrim(isnull(s2.extern,'')) " & _
            ",驗收單號 = rtrim(isnull(s2.customerorderkey1,'')) " & _
            ",退貨箱數 = isnull(rtrim(o.goodsback),0) " & _
            ",'現金/支票' = o.cash " & _
            ",異常狀況 = case when s2.confirm_notes = '正常訂單' then 'N' when len(rtrim(isnull(s2.confirm_notes,''))) =0 then 'N' else 'Y' end " & _
            "from trp01m t1m right join sdn02t s2 on s2.consigneekey = t1m.consigneekey and s2.storerkey = t1m.storerkey " & _
            "join orders o on o.orderkey = s2.c_receipt_no " & _
            "where s2.storerkey = 'LABT01' " & strSelected & _
            "order by isnull(s2.extern,'')"


Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Call Replication_Recordset(tmp_Rs, rsMain)

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

End With

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT1_Click()

If Len(txtDeliveryDateST1) = 0 Or Len(txtDeliveryDateET1) = 0 Then MsgBox "請輸入起訖日期區間！", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle

Screen.MousePointer = 11
Set dgMainT1.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_DeliveryDate As String

str_SQL = "select * from gv_sdn05tdetail where 貨主 = 'LNSL01' and 到貨日 between '" & txtDeliveryDateST1 & "' and '" & txtDeliveryDateET1 & "' "

Set rsMainT1 = New ADODB.Recordset
rsMainT1.CursorLocation = adUseClient
rsMainT1.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT1.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
rsMainT1.Sort = "到貨日,路線編號,貨主單號"

Set dgMainT1.DataSource = rsMainT1: dgMainT1.Visible = False
rsMainT1.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT1
StatusBar.Panels(2).Text = rsMainT1.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMainT1.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT2_Click()

'If Len(RTrim(txtSdnDateST2)) = 0 Or Len(RTrim(txtSdnDateET2)) = 0 Then MsgBox "請輸入日期區間!", 64, "查詢": Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMainT2.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"

str_SQL = "exec gs_LNSL01Abnormal '" & RTrim(txtSdnDateST2) & "','" & RTrim(txtSdnDateET2) & "','" & RTrim(txtDeliveryDateST2) & "','" & RTrim(txtDeliveryDateET2) & "' "
            
Set rsMainT2 = New ADODB.Recordset
rsMainT2.CursorLocation = adUseClient
rsMainT2.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT2.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT2.DataSource = rsMainT2: dgMainT2.Visible = False
rsMainT2.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT2
StatusBar.Panels(2).Text = rsMainT2.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMainT2.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT3_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST3) = 0 And Len(txtDeliveryDateET3) = 0 And Len(txtOrderDateST3) = 0 And Len(txtOrderDateET3) = 0 Then MsgBox "請輸入起訖日期區間！", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT3.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_DeliveryDate As String

str_SQL = "exec [gs_LABT01receive] '" & txtOrderDateST3 & "','" & txtOrderDateET3 & "','" & txtDeliveryDateST3 & "','" & txtDeliveryDateET3 & "'"

Set rsMainT3 = New ADODB.Recordset
rsMainT3.CursorLocation = adUseClient
rsMainT3.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT3.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT3.DataSource = rsMainT3: dgMainT3.Visible = False
rsMainT3.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT3
StatusBar.Panels(2).Text = rsMainT3.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMainT3.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT4_Click()

On Error GoTo err_Handle
'If Len(txtDeliveryDateST4) = 0 Or Len(txtDeliveryDateET4) = 0 Then MsgBox "請輸入起訖日期區間！", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT4.DataSource = Nothing: Set dgMainT4_1.DataSource = Nothing: Set dgMainT4_2.DataSource = Nothing: StatusBar.Panels(2).Text = ""
Dim strWhere As Integer, intloop As Integer, strTmp As String, chkDeliveryDate As String, chkCar As String, chkAdddateDate As String, chkDelivery As String, chkRoute As String, chkStatus As String, tmp_data() As String

'車號
chkCar = ""
If RTrim(cboCarT4) <> "" Then chkCar = "and o2.vehicle_id_no = '" & cboCarT4 & "' "

'排除車號
If Len(txtNotCarNo) > 0 Then
   '儲位編號：零散，以逗號分割
   tmp_data = Split(txtNotCarNo, ",", -1, vbTextCompare)    '分割輸入的車號
   strTmp = ""
   '將已切割字串加以組合 (非空白字串才加以組合)
   For intloop = LBound(tmp_data) To UBound(tmp_data)
       If Len(tmp_data(intloop)) > 0 Then
          If Len(strTmp) > 0 Then
             strTmp = strTmp & ",'" & tmp_data(intloop) & "'"
          Else
             strTmp = strTmp & "'" & tmp_data(intloop) & "'"
          End If
       End If
   Next intloop
   If Len(strTmp) > 0 Then
      strTmp = " and o2.vehicle_id_no not in (" & strTmp & ") "
      If Len(strWhere) = 0 Then
         strWhere = strTmp
      Else
         chkCar = strTmp
      End If
   End If
End If

'接單日
chkAdddateDate = ""
If Len(RTrim(txtAddDateST4.Text)) > 0 And Len(RTrim(txtAddDateET4.Text)) > 0 Then
   chkAdddateDate = "and convert(char,o.adddate,112) between '" & txtAddDateST4.Text & "' and '" & txtAddDateET4.Text & "' "
ElseIf Len(RTrim(txtAddDateST4.Text)) > 0 And Len(RTrim(txtAddDateET4.Text)) = 0 Then
   chkAdddateDate = "and convert(char,o.adddate,112) = '" & txtAddDateST4.Text & "' "
ElseIf Len(RTrim(txtAddDateST4.Text)) = 0 And Len(RTrim(txtAddDateET4.Text)) > 0 Then
   chkAdddateDate = "and convert(char,o.adddate,112) = '" & txtAddDateET4.Text & "' "
End If

'出車日
chkDelivery = ""
If Len(RTrim(txtDeliveryST4.Text)) > 0 And Len(RTrim(txtDeliveryET4.Text)) > 0 Then
   chkDelivery = "and '20' + substring(o2.route_no,2,6) between '" & txtDeliveryST4.Text & "' and '" & txtDeliveryET4.Text & "' "
ElseIf Len(RTrim(txtDeliveryST4.Text)) > 0 And Len(RTrim(txtDeliveryET4.Text)) = 0 Then
   chkDelivery = "and '20' + substring(o2.route_no,2,6) = '" & txtDeliveryST4.Text & "' "
ElseIf Len(RTrim(txtDeliveryST4.Text)) = 0 And Len(RTrim(txtDeliveryET4.Text)) > 0 Then
   chkDelivery = "and '20' + substring(o2.route_no,2,6) = '" & txtDeliveryET4.Text & "' "
End If

'到貨日
chkDeliveryDate = ""
If Len(RTrim(txtDeliveryDateST4.Text)) > 0 And Len(RTrim(txtDeliveryDateET4.Text)) > 0 Then
   chkDeliveryDate = "and convert(char(8),o2.arrive_Date,112) between '" & txtDeliveryDateST4.Text & "' and '" & txtDeliveryDateET4.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateST4.Text)) > 0 And Len(RTrim(txtDeliveryDateET4.Text)) = 0 Then
   chkDeliveryDate = "and convert(char(8),o2.arrive_Date,112) = '" & txtDeliveryDateST4.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateST4.Text)) = 0 And Len(RTrim(txtDeliveryDateET4.Text)) > 0 Then
   chkDeliveryDate = "and convert(char(8),o2.arrive_Date,112) = '" & txtDeliveryDateET4.Text & "' "
End If

'路編
chkRoute = ""
If Len(RTrim(txtRouteST4.Text)) > 0 And Len(RTrim(txtRouteET4.Text)) > 0 Then
   chkRoute = "and o2.route_no between '" & txtRouteST4.Text & "' and '" & txtRouteET4.Text & "' "
ElseIf Len(RTrim(txtRouteST4.Text)) > 0 And Len(RTrim(txtRouteET4.Text)) = 0 Then
   chkRoute = "and o2.route_no = '" & txtRouteST4.Text & "' "
ElseIf Len(RTrim(txtRouteST4.Text)) = 0 And Len(RTrim(txtRouteET4.Text)) > 0 Then
   chkRoute = "and o2.route_no = '" & txtRouteET4.Text & "' "
End If

'件數狀態
chkStatus = ""
If optNo = True Then chkStatus = "and len(rtrim(isnull(convert(char(20),o2.OTconfirmdate,120),''))) = 0 "
If optYes = True Then chkStatus = "and len(rtrim(isnull(convert(char(20),o2.OTconfirmdate,120),''))) > 0 "

str_SQL = "set nocount on if object_id ('tempdb..#2') is not null drop table #2  " & _
"select  Rtrim(isnull(o2.Extern,''))+'、'+Rtrim(isnull(o.CustomerOrderkey,''))+'、' + Rtrim(isnull(o.InvoiceNo,'')) + '、' + Rtrim(isnull(o.B_Contact2,'')) as 'DN單號' " & _
",Convert(char(10),o2.arrive_date,111) as '到貨日' " & _
",Rtrim(t1m.full_name) as '客戶名稱',isnull(Rtrim(t1m.Address),'') as '客戶地址' " & _
",Rtrim(t1m.Phone) as '客戶電話',Rtrim(convert(char(1000),o.Notes)) as '訂單備註' " & _
",路線編號 = o2.route_no,車號 = rtrim(o2.vehicle_id_no) " & _
", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '箱數' " & _
",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '站所碼' " & _
",代收貨款 = o.Cash,收退貨 = case when isnull(o.GoodsBack,0) = 1 then '順收' else '' end " & _
",件數確認 = isnull(o2.OTconfirmuser,'未確認'),分類 = rtrim(isnull(o.B_City,'')),區域 = left(t1m.area_code,1) into #2 " & _
"from trp02t o2 join trp03t o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
"inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
"inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
"left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
"left join trp02m tm on t1m.zip=tm.zip " & _
"where o.storerkey='LABT01' and o.type<>'刪單' and isnull(o.B_Phone1,'')<>'01' " & chkCar & chkAdddateDate & chkDelivery & chkDeliveryDate & chkRoute & chkStatus & _
"group by left(t1m.area_code,1),o2.OTconfirmuser,o2.route_no,o2.vehicle_id_no,o2.Extern ,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "

str_SQL = str_SQL & "union " & _
"select  Rtrim(isnull(o2.Extern,''))+'、'+Rtrim(isnull(o.CustomerOrderkey,''))+'、' + Rtrim(isnull(o.InvoiceNo,'')) + '、' + Rtrim(isnull(o.B_Contact2,'')) as 'DN單號' " & _
",Convert(char(10),o2.arrive_date,111) as '到貨日' " & _
",Rtrim(t1m.full_name) as '客戶名稱',isnull(Rtrim(t1m.Address),'') as '客戶地址' " & _
",Rtrim(t1m.Phone) as '客戶電話',Rtrim(convert(char(1000),o.Notes)) as '訂單備註' " & _
",路線編號 = o2.route_no,車號 = rtrim(o2.vehicle_id_no) " & _
", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '箱數' " & _
",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '站所碼' " & _
",代收貨款 = o.Cash,收退貨 = case when isnull(o.GoodsBack,0) = 1 then '順收' else '' end " & _
",件數確認 = isnull(o2.OTconfirmuser,'未確認'),分類 = rtrim(isnull(o.B_City,'')),區域 = left(t1m.area_code,1) " & _
"from ort02t o2 join ort03t o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
"inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
"inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
"left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
"left join trp02m tm on t1m.zip=tm.zip " & _
"where o.storerkey='LABT01' and o.type<>'刪單' and isnull(o.B_Phone1,'')<>'01' " & chkCar & chkAdddateDate & chkDelivery & chkDeliveryDate & chkRoute & chkStatus & _
"group by left(t1m.area_code,1),o2.OTconfirmuser,o2.route_no,o2.vehicle_id_no,o2.Extern ,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "


If RTrim(cboCarT4) = "待排車" Or RTrim(cboCarT4) = "" Then
    str_SQL = str_SQL & "union select  Rtrim(isnull(o2.Extern,''))+'、'+Rtrim(isnull(o.CustomerOrderkey,''))+'、' + Rtrim(isnull(o.InvoiceNo,'')) + '、' + Rtrim(isnull(o.B_Contact2,'')) as 'DN單號' " & _
    ",Convert(char(10),o2.arrive_date,111) as '到貨日' " & _
    ",Rtrim(t1m.full_name) as '客戶名稱',isnull(Rtrim(t1m.Address),'') as '客戶地址' " & _
    ",Rtrim(t1m.Phone) as '客戶電話',Rtrim(convert(char(1000),o.Notes)) as '訂單備註' " & _
    ",路線編號 = '待排車',車號 = '待排車' " & _
    ", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '箱數' " & _
    ",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '站所碼' " & _
    ",代收貨款 = o.Cash,收退貨 = case when isnull(o.GoodsBack,0) = 1 then '順收' else '' end " & _
    ",件數確認 = isnull(o2.OTconfirmuser,'未確認'),分類 = rtrim(isnull(o.B_City,'')),區域 = left(t1m.area_code,1) " & _
    "from ort02w o2 join ort03w o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
    "inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
    "inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
    "left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
    "left join trp02m tm on t1m.zip=tm.zip " & _
    "where o.storerkey='LABT01' and o.type<>'刪單' and isnull(o.B_Phone1,'')<>'01' " & chkAdddateDate & chkDeliveryDate & chkRoute & chkStatus & _
    "group by left(t1m.area_code,1),o2.OTconfirmuser,o2.Extern,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2 ,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "

    str_SQL = str_SQL & "union select  Rtrim(isnull(o2.Extern,''))+'、'+Rtrim(isnull(o.CustomerOrderkey,''))+'、' + Rtrim(isnull(o.InvoiceNo,'')) + '、' + Rtrim(isnull(o.B_Contact2,'')) as 'DN單號' " & _
    ",Convert(char(10),o2.arrive_date,111) as '到貨日' " & _
    ",Rtrim(t1m.full_name) as '客戶名稱',isnull(Rtrim(t1m.Address),'') as '客戶地址' " & _
    ",Rtrim(t1m.Phone) as '客戶電話',Rtrim(convert(char(1000),o.Notes)) as '訂單備註' " & _
    ",路線編號 = '待排車',車號 = '待排車' " & _
    ", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '箱數' " & _
    ",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '站所碼' " & _
    ",代收貨款 = o.Cash,收退貨 = case when isnull(o.GoodsBack,0) = 1 then '順收' else '' end " & _
    ",件數確認 = isnull(o2.OTconfirmuser,'未確認'),分類 = rtrim(isnull(o.B_City,'')),區域 = left(t1m.area_code,1) " & _
    "from trp02w o2 join trp03w o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
    "inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
    "inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
    "left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
    "left join trp02m tm on t1m.zip=tm.zip " & _
    "where o.storerkey='LABT01' and o.type<>'刪單' and isnull(o.B_Phone1,'')<>'01' " & chkAdddateDate & chkDeliveryDate & chkRoute & chkStatus & _
    "group by left(t1m.area_code,1),o2.OTconfirmuser,o2.Extern,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2 ,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "

End If

'其他明細表資料
str_SQL = str_SQL & "select 路線編號,DN單號,到貨日,客戶名稱,客戶地址,客戶電話,訂單備註,sum(箱數) as 箱數,Zip,站所碼,代收貨款,收退貨 " & _
",件數=isnull((select sum(isnull(otqty,0)) from trp02t where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1)  and OTconfirmdate is not null),0) " & _
"+isnull((select sum(isnull(otqty,0)) from ort02t where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1)  and OTconfirmdate is not null),0) " & _
"+isnull((select sum(isnull(otqty,0)) from trp02w where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1)  and OTconfirmdate is not null),0) " & _
"+isnull((select sum(isnull(otqty,0)) from ort02w where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1)  and OTconfirmdate is not null),0) " & _
",件數確認,區域 from  #2 where 分類 not like '%杏一%' and 分類 not like '%維康%' group by DN單號,到貨日,客戶名稱,客戶地址,客戶電話,訂單備註,Zip,站所碼 ,代收貨款,收退貨,件數確認,區域,路線編號 order by SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If tmp_Rs.EOF = True Then
    Set rsMainT4 = Nothing
    Screen.MousePointer = 0: MsgBox "查無其他明細表資料！", vbOKOnly + vbInformation, Me.Caption:
Else
    '帶出明細資料
    Call ReDim_Recordset(rsMainT4)
    Call Replication_Recordset(tmp_Rs, rsMainT4)
    tmp_Rs.Close
    rsMainT4.MoveFirst
    Do While Not rsMainT4.EOF
        str_SQL = "select 品名簡稱 = case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end,數量 = isnull(sum(od.originalqty),0) from orders o (nolock) join orderdetail od (nolock)  on o.orderkey = od.orderkey and o.storerkey = 'LABT01' and o.type <> '刪單' " & _
                  "join Exceed_ABT..sku s (nolock)  on s.storerkey = od.storerkey and s.sku = od.sku " & _
                  "where CONVERT(varchar(12),o.deliverydate, 111) = '" & RTrim(rsMainT4.Fields("到貨日")) & "' and o.externorderkey = '" & RTrim(mySplit(rsMainT4.Fields("DN單號"), "、", 0)) & "' group by  case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end "
                  
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                
                If RTrim(rsMainT4.Fields("訂單備註")) = "" Then
                    rsMainT4.Fields("訂單備註") = rsMainT4.Fields("訂單備註") & "出貨產品："
                Else
                    rsMainT4.Fields("訂單備註") = rsMainT4.Fields("訂單備註") & "　出貨產品："
                End If
                
                Do While Not tmp_Rs.EOF
                    rsMainT4.Fields("訂單備註") = rsMainT4.Fields("訂單備註") & " " & RTrim(tmp_Rs.Fields("品名簡稱")) & "*" & RTrim(tmp_Rs.Fields("數量")) & "、"
                    tmp_Rs.MoveNext
                Loop
                rsMainT4.Fields("訂單備註") = Left(rsMainT4.Fields("訂單備註"), Len(rsMainT4.Fields("訂單備註")) - 1)
                tmp_Rs.Close
        rsMainT4.MoveNext
    Loop
    rsMainT4.MoveFirst
    Set dgMainT4.DataSource = rsMainT4
    SSTab1.Tab = 0
    StatusBar.Panels(2).Text = StatusBar.Panels(2).Text & "其他明細:" & rsMainT4.RecordCount & " 筆資料列                   "
    SetDataGridColWidth Me.Caption, dgMainT4
End If

'杏一明細表
str_SQL = "select 路線編號,DN單號,到貨日,客戶名稱,客戶地址,客戶電話,訂單備註,sum(箱數) as 箱數,Zip,站所碼,代收貨款,收退貨,件數=isnull((select sum(isnull(otqty,0)) from trp02t where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1)  and OTconfirmdate is not null),0) " & _
    "+isnull((select sum(isnull(otqty,0)) from ort02t where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1) and OTconfirmdate is not null),0) " & _
    ",件數確認,區域 from  #2 where 分類 like '%杏一%' group by DN單號,到貨日,客戶名稱,客戶地址,客戶電話,訂單備註,Zip,站所碼 ,代收貨款,收退貨,件數確認,區域,路線編號 order by SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = True Then
    Set rsMainT4_1 = Nothing
    Screen.MousePointer = 0: MsgBox "查無杏一明細表資料！", vbOKOnly + vbInformation, Me.Caption:
Else
    '帶出明細資料
    Call ReDim_Recordset(rsMainT4_1)
    Call Replication_Recordset(tmp_Rs, rsMainT4_1)
    tmp_Rs.Close
     rsMainT4_1.MoveFirst
    Do While Not rsMainT4_1.EOF
        str_SQL = "select 品名簡稱 = case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end,數量 = isnull(sum(od.originalqty),0) from orders o (nolock) join orderdetail od (nolock)  on o.orderkey = od.orderkey and o.storerkey = 'LABT01' and o.type <> '刪單' " & _
                  "join Exceed_ABT..sku s (nolock)  on s.storerkey = od.storerkey and s.sku = od.sku " & _
                  "where CONVERT(varchar(12),o.deliverydate, 111) = '" & RTrim(rsMainT4_1.Fields("到貨日")) & "' and o.externorderkey = '" & RTrim(mySplit(rsMainT4_1.Fields("DN單號"), "、", 0)) & "' group by  case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end "
                
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.CursorLocation = 3 '可以修改recordset
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If RTrim(rsMainT4_1.Fields("訂單備註")) = "" Then
                    rsMainT4_1.Fields("訂單備註") = rsMainT4_1.Fields("訂單備註") & "出貨產品："
                Else
                    rsMainT4_1.Fields("訂單備註") = rsMainT4_1.Fields("訂單備註") & "　出貨產品："
                End If
                Do While Not tmp_Rs.EOF
                    rsMainT4_1.Fields("訂單備註") = rsMainT4_1.Fields("訂單備註") & " " & RTrim(tmp_Rs.Fields("品名簡稱")) & "*" & RTrim(tmp_Rs.Fields("數量")) & "、"
                    tmp_Rs.MoveNext
                Loop
                rsMainT4_1.Fields("訂單備註") = Left(rsMainT4_1.Fields("訂單備註"), Len(rsMainT4_1.Fields("訂單備註")) - 1)
                tmp_Rs.Close
        rsMainT4_1.MoveNext
    Loop
        rsMainT4_1.MoveFirst
    Set dgMainT4_1.DataSource = rsMainT4_1
    SetDataGridColWidth Me.Caption, dgMainT4_1
    StatusBar.Panels(2).Text = StatusBar.Panels(2).Text & "杏一明細:" & rsMainT4_1.RecordCount & " 筆資料列                 "
    SSTab1.Tab = 1
End If


'維康明細表
str_SQL = "select 路線編號,DN單號,到貨日,客戶名稱,客戶地址,客戶電話,訂單備註,sum(箱數) as 箱數,Zip,站所碼,代收貨款,收退貨,件數=isnull((select sum(isnull(otqty,0)) from trp02t where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1)  and OTconfirmdate is not null),0) " & _
    "+isnull((select sum(isnull(otqty,0)) from ort02t where extern =  SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1) and OTconfirmdate is not null),0) " & _
    ",件數確認,區域 from  #2 where 分類 like '%維康%' group by DN單號,到貨日,客戶名稱,客戶地址,客戶電話,訂單備註,Zip,站所碼 ,代收貨款,收退貨,件數確認,區域,路線編號 order by SUBSTRING(DN單號 , 1, CHARINDEX('、',  DN單號 )-1) if object_id ('tempdb..#2') is not null drop table #2 set nocount off "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = True Then
    Set rsMainT4_2 = Nothing
    Screen.MousePointer = 0: MsgBox "查無維康明細表資料！", vbOKOnly + vbInformation, Me.Caption:
Else
    '帶出明細資料
    Call ReDim_Recordset(rsMainT4_2)
    Call Replication_Recordset(tmp_Rs, rsMainT4_2)
    tmp_Rs.Close
    rsMainT4_2.MoveFirst
    Do While Not rsMainT4_2.EOF
        str_SQL = "select 品名簡稱 = case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end,數量 = isnull(sum(od.originalqty),0) from orders o (nolock) join orderdetail od (nolock)  on o.orderkey = od.orderkey and o.storerkey = 'LABT01'  and o.type <> '刪單' " & _
                  "join Exceed_ABT..sku s (nolock)  on s.storerkey = od.storerkey and s.sku = od.sku " & _
                  "where CONVERT(varchar(12),o.deliverydate, 111) = '" & RTrim(rsMainT4_2.Fields("到貨日")) & "' and o.externorderkey = '" & RTrim(mySplit(rsMainT4_2.Fields("DN單號"), "、", 0)) & "' group by  case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end "
                  
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If RTrim(rsMainT4_2.Fields("訂單備註")) = "" Then
                    rsMainT4_2.Fields("訂單備註") = rsMainT4_2.Fields("訂單備註") & "出貨產品："
                Else
                    rsMainT4_2.Fields("訂單備註") = rsMainT4_2.Fields("訂單備註") & "　出貨產品："
                End If
                Do While Not tmp_Rs.EOF
                    rsMainT4_2.Fields("訂單備註") = rsMainT4_2.Fields("訂單備註") & " " & RTrim(tmp_Rs.Fields("品名簡稱")) & "*" & RTrim(tmp_Rs.Fields("數量")) & "、"
                    tmp_Rs.MoveNext
                Loop
                rsMainT4_2.Fields("訂單備註") = Left(rsMainT4_2.Fields("訂單備註"), Len(rsMainT4_2.Fields("訂單備註")) - 1)
                tmp_Rs.Close
        rsMainT4_2.MoveNext
    Loop
    rsMainT4_2.MoveFirst
    Set dgMainT4_2.DataSource = rsMainT4_2
    SetDataGridColWidth Me.Caption, dgMainT4_2
    StatusBar.Panels(2).Text = StatusBar.Panels(2).Text & "維康明細:" & rsMainT4_2.RecordCount & " 筆資料列                 "
    SSTab1.Tab = 2
End If

    Screen.MousePointer = 0: dgMainT4.Visible = True: dgMainT4_1.Visible = True: dgMainT4_2.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT5_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST5) = 0 Or Len(txtDeliveryDateET5) = 0 Then MsgBox "請輸入起訖日期區間！", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT5.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_DeliveryDate As String

str_SQL = "exec gs_LNSL01ShippingReport '" & txtDeliveryDateST5 & "','" & txtDeliveryDateET5 & "' "

Set rsMainT5 = New ADODB.Recordset
rsMainT5.CursorLocation = adUseClient
rsMainT5.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT5.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT5.DataSource = rsMainT5: dgMainT5.Visible = False
rsMainT5.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT5
StatusBar.Panels(2).Text = rsMainT5.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMainT5.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT6_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST6) = 0 Or Len(txtDeliveryDateET6) = 0 Then MsgBox "請輸入起訖日期區間！", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT6.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_DeliveryDate As String

If chkT6 = 1 Then
    str_SQL = "exec gs_LNSL01GoodsArriveReport_wild '" & txtDeliveryDateST6 & "','" & txtDeliveryDateET6 & "' "
Else
    str_SQL = "exec gs_LNSL01GoodsArriveReport '" & txtDeliveryDateST6 & "','" & txtDeliveryDateET6 & "' "
End If


Set rsMainT6 = New ADODB.Recordset
rsMainT6.CursorLocation = adUseClient
rsMainT6.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT6.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT6.DataSource = rsMainT6: dgMainT6.Visible = False
rsMainT6.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT6
StatusBar.Panels(2).Text = rsMainT6.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMainT6.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdQueryT7_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST7) = 0 Or Len(txtDeliveryDateET7) = 0 Then MsgBox "請輸入起訖日期區間！", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT7.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_DeliveryDate As String

MsgBox "1.NP、NPP與BULK產品，箱入數不能為0" & vbCrLf & "2.F＆B與空調商品，箱入數不能為0" & vbCrLf & "3.排除212倉別", 64, "注意"

If chkT7 = 1 Then
    str_SQL = "exec gs_LNSL01Storage_Wild '" & txtDeliveryDateST7 & "','" & txtDeliveryDateET7 & "' "
Else
    str_SQL = "exec gs_LNSL01Storage '" & txtDeliveryDateST7 & "','" & txtDeliveryDateET7 & "' "
End If

Set rsMainT7 = New ADODB.Recordset
rsMainT7.CursorLocation = adUseClient
rsMainT7.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT7.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT7.DataSource = rsMainT7: dgMainT7.Visible = False
rsMainT7.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT7
StatusBar.Panels(2).Text = rsMainT7.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMainT7.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdSaveToText_Click()
'資料排序
Recordset2Excel "佰事達物流回單檢核表", rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
'        .Columns("L").Select
'        .Selection.ClearContents
        .Range("B3").Value = Combo1
        .Range("A1").Select
        '備份檔案
        '    If Dir("C:\LTKK01\DelievryTrack", vbDirectory) = "" Then MkDirs "C:\LTKK01\DelievryTrack"
        '    .ActiveWorkbook.SaveAs "C:\LTKK01\DelievryTrack\DelievryTrack" & Format(Now, "yyyymmddhhMMss") & ".xls"
                
    End With
End If
Set MyXlsApp = Nothing
    
End Sub

Private Sub cmdSaveToTextT2_Click()

If rsMainT2 Is Nothing Then Exit Sub
If rsMainT2.EOF Then Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11: cmdSaveToTextT2.Enabled = False: dgMainT2.Enabled = False

Dim i As Integer, j As Integer, strCheck As String, strFileName As String, strFileName1 As String

strFileName = "簽單回檔" & Format(Now, "yyyymmddhhMMss") & ".txt"
strFileName1 = "退貨簽單回檔" & Format(Now, "yyyymmddhhMMss") & ".txt"

'轉文字檔
If Dir("C:\LNSL01\簽單回檔", vbDirectory) = "" Then MkDirs "C:\LNSL01\簽單回檔"
Open "C:\LNSL01\簽單回檔\" & strFileName For Output As #1
Open "C:\LNSL01\簽單回檔\" & strFileName1 For Output As #2

rsMainT2.Sort = "預計到貨日,貨主訂單號碼,項次"

'交易開始
Tran_Level = cn.BeginTrans

rsMainT2.MoveFirst
Do While Not rsMainT2.EOF
    
    If Len(rsMainT2("WMS單號")) > 0 Then
        Print #1, rsMainT2("WMS單號"); rsMainT2("出倉日"); rsMainT2("預計到貨日"); rsMainT2("貨主訂單號碼"); Format(rsMainT2("項次"), "0000000000"); rsMainT2("品號"); Format(rsMainT2("出貨數量"), "00000000"); Format(rsMainT2("簽單數量"), "00000000"); rsMainT2("到期日"); rsMainT2("生產批號"); rsMainT2("倉別"); rsMainT2("備註"); rsMainT2("發票回收"); rsMainT2("客戶編號"); rsMainT2("客戶簡稱"); Format(rsMainT2("此單總筆數"), "00000000")
        i = i + 1
    Else
        Print #2, rsMainT2("WMS單號"); rsMainT2("出倉日"); rsMainT2("預計到貨日"); rsMainT2("貨主訂單號碼"); Format(rsMainT2("項次"), "0000000000"); rsMainT2("品號"); Format(rsMainT2("出貨數量"), "00000000"); Format(rsMainT2("簽單數量"), "00000000"); rsMainT2("到期日"); rsMainT2("生產批號"); rsMainT2("倉別"); rsMainT2("備註"); rsMainT2("發票回收"); rsMainT2("客戶編號"); rsMainT2("客戶簡稱"); Format(rsMainT2("此單總筆數"), "00000000")
        j = j + 1
    End If
    
    '更新為已回傳
    str_SQL = "update sdn02t set sdnfeedback = 1 where receipt_no = '" & RTrim(rsMainT2("TMS單號")) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rsMainT2.MoveNext
Loop

Print #1, "Total Count = " & Format(i, "00000000")
Print #2, "Total Count = " & Format(j, "00000000")

'關閉檔案
Close #1
Close #2

cn.CommitTrans: Tran_Level = 0

Set rsMainT2 = Nothing: Set dgMainT2.DataSource = Nothing
Screen.MousePointer = 0: cmdSaveToTextT2.Enabled = True: dgMainT2.Enabled = True
MsgBox "簽單回傳轉出完成!!" & vbCrLf & "C:\LNSL01\簽單回檔\" & strFileName & vbCrLf & "C:\LNSL01\簽單回檔\" & strFileName1, vbOKOnly, Me.Caption
Exit Sub

err_Handle:
    Screen.MousePointer = 0: cmdSaveToTextT2.Enabled = True: dgMainT2.Enabled = True
    Close
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
    
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT1
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT2
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT3_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT3
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT4_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT4
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT5_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT5
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT6_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT6
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT7_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT7
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
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
Private Sub dgMainT1_HeadClick(ByVal ColIndex As Integer)

If dgMainT1.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT1.Sort = dgMainT1.Columns(ColIndex).Caption & " DESC"
    dgMainT1.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT1.Sort = dgMainT1.Columns(ColIndex).Caption
    dgMainT1.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgMainT2_HeadClick(ByVal ColIndex As Integer)

If dgMainT2.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT2.Sort = dgMainT2.Columns(ColIndex).Caption & " DESC"
    dgMainT2.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT2.Sort = dgMainT2.Columns(ColIndex).Caption
    dgMainT2.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgMainT3_HeadClick(ByVal ColIndex As Integer)

If dgMainT3.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT3.Sort = dgMainT3.Columns(ColIndex).Caption & " DESC"
    dgMainT3.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT3.Sort = dgMainT3.Columns(ColIndex).Caption
    dgMainT3.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgMainT4_HeadClick(ByVal ColIndex As Integer)

If dgMainT4.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT4.Sort = dgMainT4.Columns(ColIndex).Caption & " DESC"
    dgMainT4.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT4.Sort = dgMainT4.Columns(ColIndex).Caption
    dgMainT4.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgMainT4_1_HeadClick(ByVal ColIndex As Integer)

If dgMainT4_1.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT4_1.Sort = dgMainT4_1.Columns(ColIndex).Caption & " DESC"
    dgMainT4_1.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT4_1.Sort = dgMainT4_1.Columns(ColIndex).Caption
    dgMainT4_1.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgMainT4_2_HeadClick(ByVal ColIndex As Integer)

If dgMainT4_2.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT4_2.Sort = dgMainT4_2.Columns(ColIndex).Caption & " DESC"
    dgMainT4_2.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT4_2.Sort = dgMainT4_2.Columns(ColIndex).Caption
    dgMainT4_2.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgMainT5_HeadClick(ByVal ColIndex As Integer)

If dgMainT5.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT5.Sort = dgMainT5.Columns(ColIndex).Caption & " DESC"
    dgMainT5.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT5.Sort = dgMainT5.Columns(ColIndex).Caption
    dgMainT5.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgMainT6_HeadClick(ByVal ColIndex As Integer)

If dgMainT6.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT6.Sort = dgMainT6.Columns(ColIndex).Caption & " DESC"
    dgMainT6.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT6.Sort = dgMainT6.Columns(ColIndex).Caption
    dgMainT6.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgMainT7_HeadClick(ByVal ColIndex As Integer)

If dgMainT7.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMainT7.Sort = dgMainT7.Columns(ColIndex).Caption & " DESC"
    dgMainT7.ClearSelCols
    intColumnIndex = 255

Else
    rsMainT7.Sort = dgMainT7.Columns(ColIndex).Caption
    dgMainT7.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub txtAddDateET4_Click()
Set objMvdateTarget = txtAddDateET4
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtAddDateST4_Click()
Set objMvdateTarget = txtAddDateST4
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryET4_Click()
Set objMvdateTarget = txtDeliveryET4
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryST4_Click()
Set objMvdateTarget = txtDeliveryST4
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtOrderDateS_Click()

Set objMvdateTarget = txtOrderDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateE_Click()

Set objMvdateTarget = txtOrderDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateST1_Click()

Set objMvdateTarget = txtDeliveryDateST1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET1_Click()

Set objMvdateTarget = txtDeliveryDateET1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateST1_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET1_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtOrderDateST3_Click()
Set objMvdateTarget = txtOrderDateST3
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub txtOrderDateET3_Click()
Set objMvdateTarget = txtOrderDateET3
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub txtOrderDateST3_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtOrderDateET3_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtsdnDateST2_Click()

Set objMvdateTarget = txtSdnDateET2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtsdnDateET2_Click()

Set objMvdateTarget = txtSdnDateET2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtsdnDateST2_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtsdnDateET2_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateST2_Click()

Set objMvdateTarget = txtDeliveryDateST2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET2_Click()

Set objMvdateTarget = txtDeliveryDateET2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateST2_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET2_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateST3_Click()

Set objMvdateTarget = txtDeliveryDateST3
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET3_Click()

Set objMvdateTarget = txtDeliveryDateET3
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateST3_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET3_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateST4_Click()

Set objMvdateTarget = txtDeliveryDateST4
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET4_Click()

Set objMvdateTarget = txtDeliveryDateET4
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateST4_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET4_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateST5_Click()

Set objMvdateTarget = txtDeliveryDateST5
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET5_Click()

Set objMvdateTarget = txtDeliveryDateET5
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateST5_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET5_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateST6_Click()

Set objMvdateTarget = txtDeliveryDateST6
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateET6_Click()

Set objMvdateTarget = txtDeliveryDateET6
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateST6_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET6_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateST7_Click()

Set objMvdateTarget = txtDeliveryDateST7
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateET7_Click()

Set objMvdateTarget = txtDeliveryDateET7
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height * 2
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateST7_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET7_KeyPress(KeyAscii As Integer)

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
Sub Recordset2Excel_ABT(str As String, rs As Object)
'**************************************************
'Create by Gemini @20061102 4 Recordset匯出Excel
'使用說明
'1.新增EXCEL範例檔
'2.欲匯入資料的工作表命名為DATA
'3.將欲開始放置資料的儲存格輸入程式抬頭字串(App.Title)，放置位置介於A100-Z100之間
'4.將範例檔套用順序1.程式目錄下"XLT"範例資料夾2.ini檔所指定之路徑
'參數說明
'frm:來源From物件
'rs:來源Recordset
'範例
'    Recordset2Excel Me, rs_Cust
'    '..在此編輯EXCEL
'    Set MyXlsApp = Nothing'終止Excel物件
'宣告於模組
'Public MyXlsApp As Excel.Application
'**************************************************
On Error GoTo err_Handle
If rs Is Nothing Then MsgBox str & "無資料可供轉檔！", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
If rs.RecordCount > 65535 Then MsgBox str & "轉出資料超過Excel限制(65535)！", 16, "Save2Excel終止": Exit Sub
If rs.RecordCount = -1 Then MsgBox str & "無資料可供轉檔！", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String

'MsgBox "系統進行資料轉Excel時，請勿操作其他Excel作業，以免資料轉出錯誤！", vbOKOnly + vbInformation, "Save2Excel" 'add @ 20110402

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .DisplayAlerts = False '工作表新增複製刪除不顯示提示視窗add @ 20110402

    If Dir(App.Path & "\XLT\" & str & ".xlt") = "" Then '找不到本機範例檔
        
        '取範例檔路徑
        Dim objIni As vbIniFile, strXltPath As String
        Set objIni = New vbIniFile
        
        With objIni
        
            .FileName = striniFileName_FullPath
            strXltPath = RTrim(.ReadData("EXCEL", "XLTPATH", "")) '不支援中文資料夾名稱
            
        End With
        Set objIni = Nothing

    End If

    '無指定路徑或範本檔名，不使用範例檔
    If Len(RTrim(strXltPath)) = 0 Then strXltPath = App.Path & "\XLT"
    If Dir(strXltPath, vbDirectory) = "" Or Len(RTrim(str)) = 0 Then GoTo Run
    
    '範例檔
    If Dir(strXltPath & "\" & str & ".xlt") <> "" Then
'        If MsgBox("是否使用範例檔?(" & strXltPath & "\" & str & ".xlt), vbQuestion + vbYesNo, "轉Excel") = vbNo Then GoTo Run
        
        '開啟範例檔
        .Workbooks.Open (strXltPath & "\" & str & ".xlt")
        
        '尋找DATA工作表
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets("Data").Select: Exit For '選定DATA工作表
        Next
        
        '找不到新增DATA工作表
        If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then
            .Sheets.Add
            .ActiveSheet.Name = "DATA"
        Else
            '找到搜尋存放儲存格
            For k = 65 To 90
                For j = 1 To 100
                    If UCase(.Range(Chr(k) & j).Value) = "BESTTRP" Then GoTo NextStep
                Next j
            Next k
            k = 65: j = 2 '沒找到時指定放A1(J=2是因為下面會-1)
        End If
        .ActiveSheet.Name = str
NextStep:
        '寫入標題列
        If j > 1 Then '如果在第一列，則不放欄位名稱
            For i = 1 To rs.Fields.Count - 1
                l = i Mod 26
                .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Name
                '欄位超過26
                If Chr(65 + l) = "Z" Then
                    If strCol = "" Then
                        strCol = "A"
                    Else
                        strCol = Chr(Asc(strCol) + 1)
                    End If
                End If
            Next i
            '寫入recordset資料
            rs.MoveFirst: j = 3
            Do While Not rs.EOF
                For i = 1 To rs.Fields.Count - 1
                    l = i Mod 26
                    .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Value
                    If RTrim(rs.Fields(i).Value) = "未確認" Then .Range(strCol & Chr(k + l - 1 - 1) & j - 1).Value = ""
                    '欄位超過26
                    If Chr(65 + l) = "Z" Then
                        If strCol = "" Then
                            strCol = "A"
                        Else
                            strCol = Chr(Asc(strCol) + 1)
                        End If
                    End If
                Next i
                j = j + 1
                rs.MoveNext
            Loop
        End If

        '資料寫入
        '.Range(Chr(k) & j).CopyFromRecordset rs
        
    Else '不使用範例檔
Run:
        '新增Excel
        .Workbooks.Add: .ActiveSheet.Name = str
              '寫入標題列
        j = 2:  k = 65
        If j > 1 Then '如果在第一列，則不放欄位名稱
            For i = 1 To rs.Fields.Count - 1
                l = i Mod 26
                .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Name
                '欄位超過26
                If Chr(65 + l) = "Z" Then
                    If strCol = "" Then
                        strCol = "A"
                    Else
                        strCol = Chr(Asc(strCol) + 1)
                    End If
                End If
            Next i
            '寫入recordset資料
            rs.MoveFirst: j = 3
            Do While Not rs.EOF
                For i = 1 To rs.Fields.Count - 1
                    l = i Mod 26
                    .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Value
                If RTrim(rs.Fields(i).Value) = "未確認" Then
                    .Range(strCol & Chr(k + l - 1 - 1) & j - 1).Value = ""
                End If
                    '欄位超過26
                    If Chr(65 + l) = "Z" Then
                        If strCol = "" Then
                            strCol = "A"
                        Else
                            strCol = Chr(Asc(strCol) + 1)
                        End If
                    End If
                Next i
                j = j + 1
                rs.MoveNext
            Loop
        End If
    
    End If
    .ActiveWorkbook.SaveAs str & ".xls"
    .ActiveWorkbook.Author = User_id
    .Visible = True
    
End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, "轉EXECL錯誤!!")
End Sub

