VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Report_DivideSku 
   Caption         =   "花王分貨表"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11460
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   2520
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3240
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
      StartOfWeek     =   135593985
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7080
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12488
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm_Report_DivideSku.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fam_Tab0_Header"
      Tab(0).Control(1)=   "dg_Tab0_VLL"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frm_Report_DivideSku.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fam_Tab1_Header"
      Tab(1).Control(1)=   "dg_Tab1_VLLSum"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "排車分貨表"
      TabPicture(2)   =   "frm_Report_DivideSku.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "dg_DivideSku"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   1530
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   11145
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
            Left            =   8775
            Picture         =   "frm_Report_DivideSku.frx":0054
            Style           =   1  '圖片外觀
            TabIndex        =   68
            Top             =   180
            Width           =   1035
         End
         Begin VB.TextBox Text4 
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
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   57
            Top             =   1200
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox Text3 
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
            Left            =   8955
            MaxLength       =   10
            TabIndex        =   56
            Top             =   1200
            Visible         =   0   'False
            Width           =   1365
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
            Left            =   4680
            TabIndex        =   55
            Top             =   1080
            Value           =   1  '核取
            Width           =   1425
         End
         Begin VB.CommandButton Command4 
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
            Left            =   5280
            Picture         =   "frm_Report_DivideSku.frx":17D6
            Style           =   1  '圖片外觀
            TabIndex        =   54
            Top             =   840
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton Command3 
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
            TabIndex        =   53
            Top             =   180
            Width           =   765
         End
         Begin VB.ComboBox cmb_Storerkey 
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
            TabIndex        =   52
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
            Index           =   2
            Left            =   9975
            Picture         =   "frm_Report_DivideSku.frx":1AE0
            Style           =   1  '圖片外觀
            TabIndex        =   51
            Top             =   180
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
            TabIndex        =   50
            Top             =   615
            Width           =   1245
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
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   49
            Top             =   600
            Width           =   1245
         End
         Begin VB.CommandButton Command2 
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
            Left            =   6360
            Picture         =   "frm_Report_DivideSku.frx":1F22
            Style           =   1  '圖片外觀
            TabIndex        =   48
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
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
            Picture         =   "frm_Report_DivideSku.frx":27EC
            Style           =   1  '圖片外觀
            TabIndex        =   47
            Top             =   240
            Width           =   1065
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
            TabIndex        =   67
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label sumlab 
            Caption         =   "總個數:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   1080
            Width           =   1215
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
            Index           =   14
            Left            =   8685
            TabIndex        =   63
            Top             =   1200
            Visible         =   0   'False
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
            Index           =   13
            Left            =   6240
            TabIndex        =   62
            Top             =   1200
            Visible         =   0   'False
            Width           =   960
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
            Left            =   135
            TabIndex        =   61
            Top             =   255
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
            Index           =   11
            Left            =   4200
            TabIndex        =   60
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
            Index           =   10
            Left            =   135
            TabIndex        =   59
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
            Index           =   9
            Left            =   2445
            TabIndex        =   58
            Top             =   690
            Width           =   240
         End
      End
      Begin VB.Frame fam_Tab0_Header 
         Height          =   1665
         Left            =   -74880
         TabIndex        =   35
         Top             =   660
         Width           =   11145
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
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1260
            Width           =   1605
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
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1260
            Width           =   1605
         End
         Begin VB.TextBox txtOrderkeyS 
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
            TabIndex        =   4
            Top             =   900
            Width           =   1605
         End
         Begin VB.TextBox txtOrderkeyE 
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
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   5
            Top             =   900
            Width           =   1605
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
            Left            =   9615
            Picture         =   "frm_Report_DivideSku.frx":33AE
            Style           =   1  '圖片外觀
            TabIndex        =   13
            Top             =   195
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
            Left            =   6105
            Picture         =   "frm_Report_DivideSku.frx":37F0
            Style           =   1  '圖片外觀
            TabIndex        =   10
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_SaveToExcel 
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
            Left            =   7260
            Picture         =   "frm_Report_DivideSku.frx":40BA
            Style           =   1  '圖片外觀
            TabIndex        =   11
            Top             =   195
            Width           =   1065
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
            Left            =   5055
            Style           =   1  '圖片外觀
            TabIndex        =   8
            Top             =   525
            Width           =   630
         End
         Begin VB.CheckBox chk_Tab0_PreView 
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
            Left            =   4785
            TabIndex        =   9
            Top             =   1305
            Value           =   1  '核取
            Width           =   1425
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
            TabIndex        =   1
            Top             =   180
            Width           =   1245
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
            TabIndex        =   0
            Top             =   180
            Width           =   1245
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
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   3
            Top             =   555
            Width           =   1605
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
            TabIndex        =   2
            Top             =   555
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab0_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "出貨單"
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
            Left            =   8445
            Picture         =   "frm_Report_DivideSku.frx":4C7C
            Style           =   1  '圖片外觀
            TabIndex        =   12
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Top             =   1320
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
            Index           =   7
            Left            =   2790
            TabIndex        =   43
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   975
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
            Index           =   5
            Left            =   2790
            TabIndex        =   41
            Top             =   960
            Width           =   240
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
            TabIndex        =   40
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
            Index           =   19
            Left            =   120
            TabIndex        =   39
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
            Index           =   20
            Left            =   2790
            TabIndex        =   38
            Top             =   615
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
            Index           =   21
            Left            =   120
            TabIndex        =   37
            Top             =   615
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
            Index           =   0
            Left            =   4020
            TabIndex        =   36
            Top             =   225
            Width           =   2010
         End
      End
      Begin VB.Frame fam_Tab1_Header 
         Height          =   1530
         Left            =   -74880
         TabIndex        =   28
         Top             =   660
         Width           =   11145
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
            Picture         =   "frm_Report_DivideSku.frx":4F86
            Style           =   1  '圖片外觀
            TabIndex        =   23
            Top             =   195
            Width           =   1065
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
            Picture         =   "frm_Report_DivideSku.frx":5B48
            Style           =   1  '圖片外觀
            TabIndex        =   22
            Top             =   210
            Width           =   1065
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
            TabIndex        =   17
            Top             =   630
            Width           =   1245
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
            TabIndex        =   16
            Top             =   615
            Width           =   1245
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
            Picture         =   "frm_Report_DivideSku.frx":6412
            Style           =   1  '圖片外觀
            TabIndex        =   25
            Top             =   180
            Width           =   1065
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
            TabIndex        =   15
            Top             =   210
            Width           =   3960
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
            TabIndex        =   20
            Top             =   180
            Width           =   765
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
            Picture         =   "frm_Report_DivideSku.frx":6854
            Style           =   1  '圖片外觀
            TabIndex        =   24
            Top             =   195
            Width           =   1065
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
            TabIndex        =   21
            Top             =   1080
            Value           =   1  '核取
            Width           =   1425
         End
         Begin VB.TextBox txt_Tab1_RouteNo_End 
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
            TabIndex        =   19
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab1_RouteNo_Start 
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
            TabIndex        =   18
            Top             =   1080
            Width           =   1365
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
            TabIndex        =   34
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
            Index           =   2
            Left            =   135
            TabIndex        =   33
            Top             =   660
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
            TabIndex        =   32
            Top             =   600
            Width           =   2010
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
            TabIndex        =   31
            Top             =   255
            Width           =   960
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
            TabIndex        =   30
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
            Index           =   30
            Left            =   2565
            TabIndex        =   29
            Top             =   1080
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_VLL 
         Height          =   4545
         Left            =   -74850
         TabIndex        =   14
         Top             =   2445
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8017
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         HeadLines       =   2
         RowHeight       =   20
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
      Begin MSDataGridLib.DataGrid dg_Tab1_VLLSum 
         Height          =   4710
         Left            =   -74850
         TabIndex        =   26
         Top             =   2280
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8308
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
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   120
         Top             =   -480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dg_DivideSku 
         Height          =   4950
         Left            =   150
         TabIndex        =   64
         Top             =   1980
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8731
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
   Begin VB.Label Label3 
      Caption         =   "總個數:"
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
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frm_Report_DivideSku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬

Private strAccessDBFileName_FullPath As String
Private MSAccessAP As access.Application
Private rs_Access As ADODB.Recordset         '報表列印用 >> 轉資料至 Access DB
Private arAreaCode() As String
Private rsMain0 As ADODB.Recordset
Private rsMain1 As ADODB.Recordset
Private rsMain2 As ADODB.Recordset
Private rsMain2_1 As ADODB.Recordset
Private rsMain2_2 As ADODB.Recordset



Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub

Private Sub cmd_Tab0_Print_Click()
'花王分貨表>Barcode列印

On Error GoTo err_Handle
Dim i As Integer, Str_Duck As String, Int_Cube As Integer, Int_Count1 As Long, Int_Count2 As Long
Str_Duck = ""
Int_Cube = 0
Int_Count1 = 0
Int_Count2 = 0
'1. 資料寫出 Access 資料庫 >> 車輛裝載匯總表
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From TIHI_LABEL"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "TIHI_LABEL", cnAccess, adOpenStatic, adLockOptimistic

    rsMain2.MoveFirst
    Do While Not rsMain2.EOF
        If RTrim(rsMain2.Fields("暫存碼頭").Value) <> Str_Duck Then
            Str_Duck = RTrim(rsMain2.Fields("暫存碼頭").Value)
            Int_Cube = Round(rsMain2.Fields("總材積").Value / 40 + 0.5)
            For i = 1 To Int_Cube
                str_SQL = "Insert into TIHI_LABEL (num_1,num_2,DeliveryDate,Full_Name,Duck,Pallet) " & _
                          "Values (" & Int_Count1 & "," & Int_Count2 & ",'" & Trim(rsMain2.Fields("到貨日").Value) & "','" & Trim(rsMain2.Fields("客戶名稱").Value) & "','" & Trim(rsMain2.Fields("暫存碼頭").Value) & "','" & i & " / " & Int_Cube & "')"
            cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Int_Count1 = Int_Count1 + 1 'BarCode排序用
            Int_Count2 = Int_Count2 + 1 'BarCode排序用
            Next
        Else
            rsMain2.MoveNext
        End If
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
       MSAccessAP.DoCmd.OpenReport "TIHI_LABEL", acViewPreview
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

Private Sub cmd_Tab0_PrintReport_Click()
Dim i As Integer, Tran_Level, j As Integer, strTmp As String

'報表列印
If rsMain0 Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub: chk_Tab0_PreView = 0

On Error GoTo err_Handle

'資料寫入 Access 資料庫
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 退貨簽收單"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Dim rs_Access As New ADODB.Recordset
rs_Access.Open "退貨簽收單", cnAccess, adOpenStatic, adLockOptimistic
rsMain0.MoveFirst
Do While Not rsMain0.EOF

   rs_Access.AddNew
   For i = 0 To rsMain0.Fields.Count - 3
   
'    If i = 15 Then'控制每頁15筆資料
'         If strTmp <> rsMain0("退貨單號") Then
'             j = 0: strTmp = rsMain0("退貨單號")
'         End If
'             j = j + 1
'             rs_Access.Fields(i).Value = j
'    Else
    
         
         rs_Access.Fields(i).Value = RTrim(rsMain0.Fields(i).Value)
'    End If
   Next i
   rs_Access.Update
   rsMain0.MoveNext
Loop
rsMain0.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'更新列印次數
'str_SQL = "Update Ort01T Set VLListCount = " & rs_Tab0_VLLSum.Fields("列印次數").Value & ",VLListPrintDate = '" & strPrintDate & "' " & _
'          "Where Route_No = '" & strRouteNo & "' or C_Route_No = '" & strRouteNo & "'"
'cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab0_PreView = 1 Then
   '預覽列印
    MSAccessAP.Visible = True
    MSAccessAP.DoCmd.OpenReport "退貨簽收單", acViewPreview
    MSAccessAP.DoCmd.Maximize
   
Else
   '直接列印至印表機
    MSAccessAP.Visible = False
    MSAccessAP.DoCmd.OpenReport "退貨簽收單", acViewNormal
    MSAccessAP.CloseCurrentDatabase
    MSAccessAP.Quit: Set MSAccessAP = Nothing
End If
'chk_Tab0_PreView = 0
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
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-退貨簽收單-列印", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Query_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dg_Tab0_VLL.DataSource = Nothing
Dim chcOrderby As String, chcDeliveryDate As String, chcRoute As String, chcOrderkey As String, chcExternOrderkey As String
Dim i As Integer

str_SQL = "select 訂單類別 = case o2t.priority when 'RC' then '提貨入庫單' when 'A2B' then '提貨配送單' else case when o2t.storerkey = 'LTKK01' and substring(o2t.extern,3,2) = '12' then '退貨單(換貨)' else '退貨單' end end " & _
        ", 貨主名稱 =  (select rtrim(t16.c_name) from trp16m t16 where t16.storerkey = o2t.storerkey ) " & _
        ", 路線編號 = o2t.route_no , 參考路編 = rtrim(o.ContainerType) " & _
        ", 出車日期 = convert(char(8) , o1t.delivery_date , 112) " & _
        ", 收貨日期 = convert(char(8) , o2t.arrive_date , 112) " & _
        ", 車號 = rtrim(o2t.vehicle_id_no) , 駕駛 = rtrim(t9m.driver) " & _
        ", TMS單號 = o2t.receipt_no , 貨主單號 = rtrim(o2t.extern) " & _
        ", 客戶訂單號碼 = rtrim(o.customerorderkey) " & _
        ", 客戶名稱 = rtrim(t1m.short_name) , 客戶地址 = rtrim(t1m.address) ,電話 = rtrim(t1m.phone), 客戶需求 = t1m.notes " & _
        ", 到貨客戶 = case when o2t.priority in ('R','RC') then '貨送：' + rtrim(o.facility) when len(rtrim(o.b_company)) > 0 then '貨送：' + rtrim(t1ma.short_name) + '-'+ rtrim(t1ma.address) + ' ' + rtrim(t1ma.phone) else '' end " & _
        ", 項次 = rtrim(o3t.seq_no) , 貨號 = Rtrim(o3t.Product_No)  " & _
        ", 品名 = rtrim(sp.descr) " & _
        ", 箱數 =isnull(case when sp.casecnt = 0 then 0 else floor(o3t.order_qty/sp.Casecnt) end ,0) ,大包裝 = isnull(rtrim(sp.busr3),'箱') " & _
        ", 個數 =isnull(case when sp.casecnt = 0 then o3t.order_qty else cast(o3t.order_qty as int)%cast(sp.Casecnt as int) end ,0) , 小包裝 = isnull(rtrim(sp.busr1),'個') " & _
        ", 備註 = case when len(cast(o.notes as varchar(1000))) > 0 or len(cast(od.notes as varchar(1000))) > 0 then cast(o.notes as varchar(1000)) + '_' + cast(od.notes as varchar(1000)) else ' ' end , 總個數= o3t.order_qty " & _
        ", 排車者 = Case When Isnull(o1t.C_Route_No,'') = '' Then Isnull(Rtrim(o1t.AddWho),'') else Rtrim(o1t.AddWho) End , 總材積= o3t.order_qty * sp.stdcube , 總重量= o3t.order_qty * sp.stdgrosswgt " & _
        "from ort01t o1t join ort02t o2t on o1t.route_no = o2t.route_no " & _
        "join ort03t o3t on o3t.receipt_no = o2t.receipt_no " & _
        "join orders o on o.orderkey = o2t.receipt_no " & _
        "left join trp01m t1m on o2t.consigneekey = t1m.consigneekey and t1m.storerkey = o2t.storerkey " & _
        "left join trp01m t1ma on o.b_company = t1ma.consigneekey and t1ma.storerkey = o.storerkey  " & _
        "left join trp09m t9m on t9m.vehicle_id_no = o2t.vehicle_id_no " & _
        "join orderdetail od on od.orderkey = o.orderkey and od.orderlinenumber = o3t.seq_no  " & _
        "join gv_skuxpack sp on sp.sku = od.sku and sp.storerkey = od.storerkey " & _
        "where left(o2t.route_no,1) = 'R' "

chcOrderby = " "

'出車日期
chcDeliveryDate = ""
If Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   chcDeliveryDate = "and convert(char(8) , o1t.delivery_date , 112) between '" & txt_Tab0_DeliveryDate_Start.Text & "' and '" & txt_Tab0_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) = 0 Then
   chcDeliveryDate = "and convert(char(8) , o1t.delivery_date , 112) = '" & txt_Tab0_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) = 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   chcDeliveryDate = "and convert(char(8) , o1t.delivery_date , 112) = '" & txt_Tab0_DeliveryDate_End.Text & "' "
End If

'路線編號
chcRoute = ""
If Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   chcRoute = "and o2t.route_no between '" & txt_Tab0_RouteNo_Start.Text & "' and '" & txt_Tab0_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) = 0 Then
   chcRoute = "and o2t.route_no = '" & txt_Tab0_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) = 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   chcRoute = "and o2t.route_no = '" & txt_Tab0_RouteNo_End.Text & "' "
End If

'訂單號碼
chcOrderkey = ""
If Len(txtOrderkeyS.Text) > 0 And Len(txtOrderkeyE.Text) > 0 Then
   chcOrderkey = "and o2t.receipt_no between '" & txtOrderkeyS.Text & "' and '" & txtOrderkeyE.Text & "' "
ElseIf Len(txtOrderkeyS.Text) > 0 And Len(txtOrderkeyE.Text) = 0 Then
   chcOrderkey = "and o2t.receipt_no = '" & txtOrderkeyS.Text & "' "
ElseIf Len(txtOrderkeyS.Text) = 0 And Len(txtOrderkeyE.Text) > 0 Then
   chcOrderkey = "and o2t.receipt_no = '" & txtOrderkeyE.Text & "' "
End If

'客戶單號
chcExternOrderkey = ""
If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chcExternOrderkey = "and o2t.extern between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
   chcExternOrderkey = "and o2t.extern = '" & txtExternOrderkeyS.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chcExternOrderkey = "and o2t.extern = '" & txtExternOrderkeyE.Text & "' "
End If

'組合字串
str_SQL = str_SQL & chcDeliveryDate & chcRoute & chcOrderkey & chcExternOrderkey & chcOrderby

Set rsMain0 = New ADODB.Recordset
rsMain0.CursorLocation = 3
cn.CommandTimeout = 0
rsMain0.Open str_SQL, cn ', adOpenForwardOnly, adLockPessimistic
If rsMain0.EOF Then MsgBox "查無資料！", 64, Me.Caption: Screen.MousePointer = 0: Exit Sub
Set dg_Tab0_VLL.DataSource = rsMain0
 
With dg_Tab0_VLL
    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
'    .Columns(11).Alignment = dbgRight
'    .Columns(12).Alignment = dbgRight
'    .Columns(13).Alignment = dbgRight
'    .Columns(14).Alignment = dbgRight
'    .Columns(15).Alignment = dbgCenter
'    .Columns(18).Alignment = dbgRight
'    .Columns(19).Alignment = dbgRight

End With

rsMain0.Sort = " 路線編號 , TMS單號 , 項次 "
SetDataGridColWidth Me.Caption, dg_Tab0_VLL
Screen.MousePointer = 0

Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-退貨簽收單-查詢", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Reset_Click()
'VLL上貨表 >> 清除
txt_Tab0_DeliveryDate_Start.Text = "": txt_Tab0_DeliveryDate_End.Text = ""
txt_Tab0_RouteNo_Start.Text = "": txt_Tab0_RouteNo_End.Text = ""
Set dg_Tab0_VLL.DataSource = Nothing
Set rsMain0 = Nothing
End Sub

Private Sub cmd_Tab0_SaveToExcel_Click()
'資料排序
Recordset2Excel "其他排車明細", rsMain0

'..在此編輯EXCEL
With MyXlsApp
   
End With

Set MyXlsApp = Nothing

End Sub

Private Sub cmd_Tab1_PrintReport_Click()
Dim i As Integer, j As Integer, strTmp As String

'報表列印
If rsMain1 Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub

On Error GoTo err_Handle

'資料寫出 Access 資料庫
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From " & "退貨排車一覽表"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Dim rs_Access As New ADODB.Recordset
rs_Access.Open "退貨排車一覽表", cnAccess, adOpenStatic, adLockOptimistic
rsMain1.MoveFirst
Do While Not rsMain1.EOF
   
   rs_Access.AddNew
   For i = 0 To rsMain1.Fields.Count - 1
       rs_Access.Fields(i).Value = rsMain1.Fields(i).Value
   Next i
   rs_Access.Update
   rsMain1.MoveNext
Loop
rsMain1.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

Dim MSAccessAP As New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (App.Path & "\" & App.title & ".mdb")

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab1_PreView = 1 Then
   '預覽列印
    MSAccessAP.Visible = True
    MSAccessAP.DoCmd.OpenReport "退貨排車一覽表", acViewPreview
    MSAccessAP.DoCmd.Maximize
   
Else
   '直接列印至印表機
    MSAccessAP.Visible = False
    MSAccessAP.DoCmd.OpenReport "退貨排車一覽表", acViewNormal
    MSAccessAP.CloseCurrentDatabase
    MSAccessAP.Quit: Set MSAccessAP = Nothing
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
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-退貨排車一覽表-列印", Me.Caption, "cmd_Tab1_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Query_Click()
'排車一覽表 >> 查詢
Set dg_Tab1_VLLSum.DataSource = Nothing
Set rsMain1 = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select 出車日期,區域,運送區域,貨運公司,車牌號碼,車次,一單多車,駕駛人," & _
          "可載重量,可載材積,路線編號,運送點數,運送箱數,運送個數,運送板數,運送重量,運送材積,貨運公司代碼,備註,預計報到日期時間,客戶簡稱 " & _
          "From Report_ORTPlanList "

    
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
If Len(txt_Tab1_RouteNo_Start.Text) > 0 And Len(txt_Tab1_RouteNo_End.Text) > 0 Then
   strTmp = " Rtrim(路線編號) between '" & txt_Tab1_RouteNo_Start.Text & "' and '" & txt_Tab1_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab1_RouteNo_Start.Text) > 0 And Len(txt_Tab1_RouteNo_End.Text) = 0 Then
   strTmp = " Rtrim(路線編號) = '" & txt_Tab1_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab1_RouteNo_Start.Text) = 0 And Len(txt_Tab1_RouteNo_End.Text) > 0 Then
   strTmp = " Rtrim(路線編號) = '" & txt_Tab1_RouteNo_End.Text & "' "
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
   strTmp = " 區域 = '" & arAreaCode(cmb_Tab1_AreaCode.ListIndex) & "' "
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
str_SQL = str_SQL & " order by 出車日期,車牌號碼,車次,區域 "
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之排車資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rsMain1)
tmp_Rs.Close

With dg_Tab1_VLLSum
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
     rsMain1.MoveFirst
     Set .DataSource = rsMain1

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
    .Columns(4).Width = 1500       '貨運公司
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1000       '車牌號碼
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 500        '車次
    .Columns(6).Alignment = dbgCenter
    .Columns(7).Width = 800        '一單多車
    .Columns(7).Alignment = dbgCenter
    .Columns(8).Width = 1000       '駕駛人
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 800        '可載重量
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800       '可載材積
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 1100      '路線編號
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 800       '運送點數
    .Columns(12).Alignment = dbgCenter
    .Columns(13).Width = 800       '運送箱數
    .Columns(13).Alignment = dbgRight
    .Columns(14).Width = 800       '運送個數
    .Columns(14).Alignment = dbgRight
    .Columns(15).Width = 800       '運送板數
    .Columns(15).Alignment = dbgRight
    .Columns(16).Width = 800       '運送重量
    .Columns(16).Alignment = dbgRight
    .Columns(17).Width = 800       '運送材積
    .Columns(17).Alignment = dbgRight
    .Columns(18).Width = 1200       '貨運公司代碼
    .Columns(18).Alignment = dbgLeft
    .Columns(19).Width = 1200       '備註(二次排車路線編號)
    .Columns(19).Alignment = dbgLeft
    .Columns(20).Width = 2000       '預計報到日期時間
    .Columns(20).Alignment = dbgCenter
    .Columns(21).Width = 2000       '客戶簡稱
    .Columns(21).Alignment = dbgLeft
End With
rsMain1.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-退貨排車一覽表-查詢", Me.Caption, "cmd_Tab1_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Reset_Click()
'退貨排車一覽表 >> 清除
cmb_Tab1_AreaCode.ListIndex = -1
txt_Tab1_DeliveryDate_Start.Text = ""
txt_Tab1_DeliveryDate_End.Text = ""
txt_Tab1_RouteNo_Start.Text = ""
txt_Tab1_RouteNo_End.Text = ""
chk_Tab1_PreView = 0
Set dg_Tab1_VLLSum.DataSource = Nothing
Set rsMain1 = Nothing
End Sub

Private Sub cmd_Tab1_SaveToExcel_Click()


'排車一覽表 >> 轉 EXCEL

    If rsMain1 Is Nothing Then Exit Sub
    rsMain1.MoveFirst
    On Error GoTo err_Handle

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
    MyXlsApp.Cells(i, 8).Value = "運送箱數"
    MyXlsApp.Cells(i, 9).Value = "運送個數"
    MyXlsApp.Cells(i, 10).Value = "運送重量"
    MyXlsApp.Cells(i, 11).Value = "運送材積"
    MyXlsApp.Cells(i, 12).Value = "備註"
    MyXlsApp.Cells(i, 13).Value = "時間"
    MyXlsApp.Cells(i, 14).Value = "客戶簡稱"
    MyXlsApp.Cells(i, 15).Value = "追蹤時間"
    MyXlsApp.Cells(i, 16).Value = "確認"
    MyXlsApp.Cells(i, 17).Value = "借出"
    MyXlsApp.Cells(i, 18).Value = "回收"
    MyXlsApp.Cells(i, 19).Value = "隔板"
    i = i + 1
    j = i
    rsMain1.MoveFirst
    '日期,車號,單號,班別,借出,還入
    Do While Not rsMain1.EOF
        If i > 2 Then
            If MyXlsApp.Cells(i - 1, 4).Value <> rsMain1.Fields(5) Then
                '車號不同,隔一行在寫入excel
                MyXlsApp.Cells(i, 8).Value = "=SUM(H" & CStr(j) & ":H" & CStr(i - 1) & ")"  '箱數
                MyXlsApp.Cells(i, 9).Value = "=SUM(I" & CStr(j) & ":I" & CStr(i - 1) & ")"  '個數
                MyXlsApp.Cells(i, 10).Value = "=SUM(J" & CStr(j) & ":J" & CStr(i - 1) & ")" '運送重量
                MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")" '材積
                i = i + 2
                j = i
            End If
        End If
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 2).Value = Trim(rsMain1.Fields(1))
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rsMain1.Fields(2)
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rsMain1.Fields(5)
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rsMain1.Fields(6)
        'MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 6).Value = rsMain1.Fields(8)
        MyXlsApp.Cells(i, 7).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 7).Value = rsMain1.Fields(11)
        MyXlsApp.Cells(i, 8).Value = rsMain1.Fields(13) '運送箱數
        MyXlsApp.Cells(i, 9).Value = rsMain1.Fields(14) '運送個數
        MyXlsApp.Cells(i, 10).Value = rsMain1.Fields(16) '運送重量
        MyXlsApp.Cells(i, 11).Value = rsMain1.Fields(17) '運送材積
        MyXlsApp.Cells(i, 12).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 12).Value = rsMain1.Fields(19)
        MyXlsApp.Cells(i, 13).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 13).Value = Mid(rsMain1.Fields(20), 10, 4)
        MyXlsApp.Cells(i, 14).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 14).Value = rsMain1.Fields(21)
        rsMain1.MoveNext
        i = i + 1
    Loop
                MyXlsApp.Cells(i, 8).Value = "=SUM(H" & CStr(j) & ":H" & CStr(i - 1) & ")"  '箱數
                MyXlsApp.Cells(i, 9).Value = "=SUM(I" & CStr(j) & ":I" & CStr(i - 1) & ")"  '個數
                MyXlsApp.Cells(i, 10).Value = "=SUM(J" & CStr(j) & ":J" & CStr(i - 1) & ")" '運送重量
                MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")" '材積
    i = i + 1
    '最適欄寬
    MyXlsApp.Columns("A:S").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '儲存格格式設定
    MyXlsApp.Columns("H:K").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A1:S" & i - 1).Select
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

Private Sub Command1_Click()
'
''資料排序
'Recordset2Excel "分貨表總表", rsMain2
'
''..在此編輯EXCEL
'With MyXlsApp
'
''    .Range("s:t").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
''    '備份檔案
''    If Dir("C:\LTKK01\配送異常", vbDirectory) = "" Then MkDirs "C:\LTKK01\配送異常"
''    .ActiveWorkbook.SaveAs "C:\LTKK01\配送異常\配送異常_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'
'End With
'
'Set MyXlsApp = Nothing
'Exit Sub

Dim SaveToExcel As Boolean
'總表列印
If rsMain2 Is Nothing Then Exit Sub
    If rsMain2.RecordCount = 0 Then Exit Sub
    Dim ExcelTitle As String
    Call DocStoreDirectory(strDocPath)
    Dim strTranFileName As String           'Excel 檔案名稱
    CmnDialog.DialogTitle = "轉存分貨表總表 Excel 檔"
    CmnDialog.InitDir = "c:\my documents"
    CmnDialog.FileName = cmb_Storerkey.Text & "分貨表總表_" & Format(Now, "YYYYMMDDHHNNSS")
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
    SaveToExcel = False
    On Error GoTo err_Handle
    Screen.MousePointer = vbHourglass
    If SaveTo_ExcelFile_OTHER(strTranFileName, rsMain2, "分貨表總表", 1) = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
    Else
       Screen.MousePointer = vbDefault
    End If
    rsMain2.MoveFirst
    SaveToExcel = True
    
    '分貨表箱數

If rsMain2_1 Is Nothing Then Exit Sub
    If rsMain2_1.RecordCount = 0 Then Exit Sub
    strTranFileName = Replace(CmnDialog.FileName, "總表", "箱數")
    SaveToExcel = False
    Screen.MousePointer = vbHourglass
    If SaveTo_ExcelFile(strTranFileName, rsMain2_1, "分貨表總表_箱", 1) = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
    Else
       Screen.MousePointer = vbDefault
    End If
    rsMain2_1.MoveFirst
    SaveToExcel = True

    '分貨表個數
If rsMain2_2 Is Nothing Then Exit Sub
    If rsMain2_2.RecordCount = 0 Then Exit Sub
   strTranFileName = Replace(CmnDialog.FileName, "總表", "個數")
    SaveToExcel = False
    Screen.MousePointer = vbHourglass
    If SaveTo_ExcelFile(strTranFileName, rsMain2_2, "分貨表總表_個", 1) = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
    Else
       Screen.MousePointer = vbDefault
       If Len(strTranFileName) > 0 Then
          strTranFileName = Replace(CmnDialog.FileName, "總表", "")
          msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
          MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       End If
    End If
    rsMain2_2.MoveFirst
    SaveToExcel = True
    Exit Sub

err_Handle:
   Dim tmpString As String
   SaveToExcel = True
   Screen.MousePointer = vbDefault
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--轉 EXCEL", Me.Caption, "cmd_Tab3SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub Command2_Click()
sumlab(1).Caption = ""
'排車一覽表 >> 查詢
Set dg_DivideSku.DataSource = Nothing
Set rsMain2 = Nothing
Dim Dob_sube As Double, Int_section As Integer, Int_loc As Integer, Str_Custname As String, Int_Cloc As Integer, Str_Loc As String, Dob_Sumqty As Double, N_Count As Long
N_Count = 0
Dob_sube = 0
Int_section = 65
Int_loc = 0
Int_Cloc = 0
Str_Custname = ""
Str_Loc = ""
Dob_Sumqty = 0
Dim Door_Array
Door_Array = Array("", "RA02", "RA03", "RA04", "RA05", "RA06", "RA07", "RA08", "RA09", "RA10", "RA11", "RA12", "RA13", "RA14", _
                         "RB02", "RB03", "RB04", "RB05", "RB06", "RB07", "RB08", "RB09", "RB10", "RB11", "RB12", "RB13", "RB14", _
                         "RC02", "RC03", "RC04", "RC05", "RC06", "RC07", "RC08", "RC09", "RC10", "RC11", "RC12", "RC13", "RC14", _
                         "RD02", "RD03", "RD04", "RD05", "RD06", "RD07", "RD08", "RD09", "RD10", "RD11", "RD12", "RD13", "RD14", _
                         "RE02", "RE03", "RE04", "RE05", "RE06", "RE07", "RE08", "RE09", "RE10", "RE11", "RE12", "RE13", "RE14", _
                         "RF02", "RF03", "RF04", "RF05", "RF06", "RF07", "RF08", "RF09", "RF10", "RF11", "RF12", "RF13", "RF14", _
                        "AB03-1A", "AB04-1A", "AB05-1A", "AB06-1A", "AB07-1A", "AB08-1A", "AB09-1A", "AB10-1A", "AB11-1A", "AB12-1A", "AB13-1A", _
                        "AB14-1A", "AB15-1A", "AB16-1A", "AB17-1A", "AB18-1A", "AB19-1A", "AB20-1A", "AB21-1A", "AB22-1A", "AB23-1A", "AB24-1A", _
                        "AB25-1A", "AB26-1A", "AB27-1A", "AB28-1A", "AB29-1A", "AB30-1A", "AB31-1A", "AB32-1A", "AB33-1A", "AB34-1A", _
                        "AC01-1A", "AC02-1A", "AC03-1A", "AC04-1A", "AC05-1A", "AC06-1A", "AC07-1A", "AC08-1A", "AC09-1A", "AC10-1A", "AC11-1A", "AC12-1A", "AC13-1A", _
                        "AC14-1A", "AC15-1A", "AC16-1A", "AC17-1A", "AC18-1A", "AC19-1A", "AC20-1A", "AC21-1A", "AC22-1A", "AC23-1A", "AC24-1A", _
                        "AC25-1A", "AC26-1A", "AC27-1A", "AC28-1A", "AC29-1A", "AC30-1A", "AC31-1A", "AC32-1A", "AC33-1A", "AC34-1A", _
                        "N1", "N2", "N3", "N4", "N5", "N6", "N7", "N8", "N9", "N10", "N11", "N12", "N13", "N14", "N15", "N16", "N17", "N18", "N19", "N20", _
                        "N21", "N22", "N23", "N24", "N25", "N26", "N27", "N28", "N29", "N30", "N31", "N32", "N33", "N34", "N35", "N36", "N37", "N38", "N39", "N40", _
                        "N41", "N42", "N43", "N44", "N45", "N46", "N47", "N48", "N49", "N50", "N51", "N52", "N53", "N54", "N55", "N56", "N57", "N58", "N59", "N60", _
                        "N61", "N62", "N63", "N64", "N65", "N66", "N67", "N68", "N69", "N70", "N71", "N72", "N73", "N74", "N75", "N76", "N77", "N78", "N79", "N80", _
                        "N81", "N82", "N83", "N84", "N85", "N86", "N87", "N88", "N89", "N90", "N91", "N92", "N93", "N94", "N95", "N96", "N97", "N98", "N99", "N100", _
                        "N101", "N102", "N103", "N104", "N105", "N106", "N107", "N108", "N109", "N110", "N111", "N112", "N113", "N114", "N115", "N116", "N117", "N118", "N119", "N120", _
                        "N121", "N122", "N123", "N124", "N125", "N126", "N127", "N128", "N129", "N130", "N131", "N132", "N133", "N134", "N135", "N136", "N137", "N138", "N139", "N140", _
                        "N141", "N142", "N143", "N144", "N145", "N146", "N147", "N148", "N149", "N150", "N151", "N152", "N153", "N154", "N155", "N156", "N157", "N158", "N159", "N160", _
                        "N161", "N162", "N163", "N164", "N165", "N166", "N167", "N168", "N169", "N170", "N171", "N172", "N173", "N174", "N175", "N176", "N177", "N178", "N179", "N180", _
                        "N181", "N182", "N183", "N184", "N185", "N186", "N187", "N188", "N189", "N190", "N191", "N192", "N193", "N194", "N195", "N196", "N197", "N198", "N199", "N200")


Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "exec Es_DivideSku_new '" & DateS.Text & "','" & DateE.Text & "'"

    
Dim strWhere As String, strTmp As String

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合搜尋條件之分貨表資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rsMain2)
tmp_Rs.Close

str_SQL = "select 貨主,暫存碼頭,箱數,品號,品名,客戶名稱,條碼,箱入數,確認1 = ' ',確認2 = ' ',確認3 = ' ',到貨日 from ##DivideSku2 where 箱數 > 0 order by  到貨日,品號,客戶名稱,暫存碼頭"
'分貨表箱
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF Then
'   Screen.MousePointer = vbDefault
'   msg_text = "查詢結果：無符合搜尋條件之分貨表箱資料"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
'   Exit Sub
'End If
Call Replication_Recordset(tmp_Rs, rsMain2_1)
tmp_Rs.Close

'分貨表個
str_SQL = "select 貨主,品號,條碼,暫存碼頭,個數,品名,客戶名稱,箱入數,確認1=' ',確認2= ' ',確認3= ' ',到貨日 from ##DivideSku2 where 個數 > 0  order by  到貨日,品號,客戶名稱,暫存碼頭"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF Then
'   Screen.MousePointer = vbDefault
'   msg_text = "查詢結果：無符合搜尋條件之分貨表個資料"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
'   Exit Sub
'End If
Call Replication_Recordset(tmp_Rs, rsMain2_2)
tmp_Rs.Close


'str_SQL = "select 總數量 = sum(總數量) from ##DivideSku"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'sumlab(1).Caption = tmp_Rs.Fields("總數量")
'tmp_Rs.Close

'自動補上總表暫存區
rsMain2.MoveFirst
Do While Not rsMain2.EOF
Dob_Sumqty = Dob_Sumqty + Val(rsMain2.Fields("總個數").Value)
'自動以45材為一個暫存區自動編排
    If Str_Custname <> rsMain2.Fields("客戶名稱") Then
        Str_Custname = Trim(rsMain2.Fields("客戶名稱"))
        Dob_sube = Val(rsMain2.Fields("總材積"))
        Int_Cloc = Round(Dob_sube / 40 + 0.5) '換下一個儲位
        Int_loc = Int_loc + Int_Cloc '無條件進位
        
'        If Int_loc >= 109 Then
'            rsMain2.Fields("暫存碼頭").Value = "N"
'            GoTo Exitif
'        End If
'
        If Int_Cloc > 1 Then
            '2個暫存區，有區間
            rsMain2.Fields("暫存碼頭").Value = Door_Array(Int_loc - Int_Cloc + 1) & " ~ " & Door_Array(Int_loc)
            Str_Loc = Door_Array(Int_loc - Int_Cloc + 1) & " ~ " & Door_Array(Int_loc)
        Else
            '一個暫存區
            rsMain2.Fields("暫存碼頭").Value = Door_Array(Int_loc)
            Str_Loc = Door_Array(Int_loc)
        End If

    Else
'        If Int_loc >= 109 Then
'            rsMain2.Fields("暫存碼頭").Value = "N"
'            GoTo Exitif
'        End If

        '同一個客戶延續上一個暫存區
         rsMain2.Fields("暫存碼頭").Value = Str_Loc
    End If
Exitif:
    rsMain2.MoveNext
Loop

'箱數依照總表的來抓
rsMain2.MoveFirst
If rsMain2_1.EOF Then
    '沒有箱數的資料
   msg_text = "查詢結果：無符合搜尋條件之分貨表箱資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
Else
    rsMain2_1.MoveFirst
End If

Do While Not rsMain2_1.EOF
    rsMain2.MoveFirst
    Do While Not rsMain2.EOF
        If rsMain2.Fields("到貨日") = rsMain2_1.Fields("到貨日") And rsMain2.Fields("客戶名稱") = rsMain2_1.Fields("客戶名稱") And rsMain2.Fields("品名") = rsMain2_1.Fields("品名") And rsMain2.Fields("箱數") = rsMain2_1.Fields("箱數") Then rsMain2_1.Fields("暫存碼頭") = rsMain2.Fields("暫存碼頭"): Exit Do
        rsMain2.MoveNext
    Loop
        rsMain2_1.MoveNext
Loop

If rsMain2_1.EOF Then
    '沒有箱數的資料
Else
    rsMain2_1.MoveFirst
End If

'個數依照總表的來抓
rsMain2.MoveFirst
If rsMain2_2.EOF Then
    '沒有個數的資料
    msg_text = "查詢結果：無符合搜尋條件之分貨表個資料"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
Else
    rsMain2_2.MoveFirst
End If
Do While Not rsMain2_2.EOF
    rsMain2.MoveFirst
    Do While Not rsMain2.EOF
        If rsMain2.Fields("到貨日") = rsMain2_2.Fields("到貨日") And rsMain2.Fields("客戶名稱") = rsMain2_2.Fields("客戶名稱") And rsMain2.Fields("品名") = rsMain2_2.Fields("品名") And rsMain2.Fields("個數") = rsMain2_2.Fields("個數") Then rsMain2_2.Fields("暫存碼頭") = rsMain2.Fields("暫存碼頭"): Exit Do
        rsMain2.MoveNext
    Loop
        rsMain2_2.MoveNext
Loop
rsMain2.MoveFirst
If rsMain2_2.EOF Then
    '沒有個數的資料
Else
    rsMain2_2.MoveFirst
End If

sumlab(1).Caption = Dob_Sumqty

With dg_DivideSku
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
     rsMain2.MoveFirst
     Set .DataSource = rsMain2

    .ColumnHeaders = True          '標題行顯示
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1200       '貨主
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000        '出車日
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1000       '品號
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 2000       '品名
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 4000       '條碼
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '客戶
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1600        '暫存碼頭
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 700       '箱入數
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 600        '箱數
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 750        '個數
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 1100        '總數量
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 1100       '確認
    .Columns(12).Alignment = dbgLeft


End With
rsMain2.MoveFirst

DoEvents: DoEvents
Screen.MousePointer = vbDefault
If Int_loc >= 113 Then
            msg_text = "超出暫存區儲位，請轉出時，手動修改"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End If

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-其他排車一覽表-查詢", Me.Caption, "cmd_Tab1_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub DateE_Click()
'退貨排車一覽表 >> 出車日期 >> 起
If Trim(DateE.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(DateE.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(DateE.Text, 4) & "/" & Mid(DateE.Text, 5, 2) & "/" & Right(DateE.Text, 2))
   End If
End If
mvDate.Tag = "分貨表.出車日期.迄"
mvDate.Top = SSTab1.Top + DateE.Top + DateE.Top + DateE.Height
mvDate.Left = SSTab1.Left + Frame1.Left + DateE.Left
mvDate.Visible = True
End Sub


Private Sub DateS_Click()
'退貨排車一覽表 >> 出車日期 >> 起
If Trim(DateS.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(DateS.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(DateS.Text, 4) & "/" & Mid(DateS.Text, 5, 2) & "/" & Right(DateS.Text, 2))
   End If
End If
mvDate.Tag = "分貨表.出車日期.起"
mvDate.Top = SSTab1.Top + DateS.Top + DateS.Top + DateS.Height
mvDate.Left = SSTab1.Left + Frame1.Left + DateS.Left
mvDate.Visible = True

End Sub


Private Sub dg_Tab0_VLL_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dg_Tab0_VLL
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dg_Tab0_VLL_HeadClick(ByVal ColIndex As Integer)
'退貨簽收單
'以滑鼠點選 dg_Tab0_VLL 欄位標題區
Dim OrderFieldName As String
If TypeName(rsMain0) <> "Nothing" Then
   OrderFieldName = "[" & dg_Tab0_VLL.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rsMain0.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rsMain0.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

'Private Sub dg_Tab1_VLLSum_HeadClick(ByVal ColIndex As Integer)
''退貨排車一覽 表
''以滑鼠點選 dg_Tab0_VLL 欄位標題區
'Dim OrderFieldName As String
'If TypeName(rsMain1) <> "Nothing" Then
'   OrderFieldName = "[" & dg_Tab1_VLLSum.Columns(ColIndex).Caption & "]"
'   If strOrder = "ASC" Then
'      strOrder = "DESC"
'      rsMain1.Sort = OrderFieldName & " DESC "
'   Else
'      strOrder = "ASC"
'      rsMain1.Sort = OrderFieldName & " ASC "
'   End If
'End If
'
'End Sub

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
cmb_Tab1_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_Tab1_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If

cmb_Storerkey.AddItem "LKAO01"
cmb_Storerkey.Text = "LKAO01"


cmb_Tab1_AreaCode.ListIndex = -1
tmp_Rs.Close
txt_Tab0_DeliveryDate_Start.Text = Format(Now, "yyyymmdd")
DateS.Text = Format(Now + 1, "yyyymmdd")
DateE.Text = Format(Now + 1, "yyyymmdd")
SSTab1.Tab = 0

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
   dg_DivideSku.Width = dg_DivideSku.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_DivideSku.Height = dg_DivideSku.Height - (dbsrcFormHeight - Me.ScaleHeight)
  
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
   dg_DivideSku.Width = dg_DivideSku.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_DivideSku.Height = dg_DivideSku.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
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
Set rsMain0 = Nothing
Set rsMain1 = Nothing
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'日期選取
Select Case mvDate.Tag
    Case "VLL裝載表.出車日期.起"
         txt_Tab0_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "VLL裝載表.出車日期.迄"
         txt_Tab0_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "退貨排車一覽表.出車日期.起"
         txt_Tab1_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "退貨排車一覽表.出車日期.迄"
         txt_Tab1_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "分貨表.出車日期.起"
        DateS.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "分貨表.出車日期.迄"
        DateE.Text = Format(mvDate.Value, "YYYYMMDD")
    Case Else
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call Form_KeyDown(27, 0)
If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
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
'退貨排車一覽表 >> 出車日期 >> 迄
If Trim(txt_Tab1_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "退貨排車一覽表.出車日期.迄"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_End.Top + txt_Tab1_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DeliveryDate_Start_Click()
'退貨排車一覽表 >> 出車日期 >> 起
If Trim(txt_Tab1_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "退貨排車一覽表.出車日期.起"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_Start.Top + txt_Tab1_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab1_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'退貨排車一覽表 >> 出車日期 >> 起
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab1_DeliveryDate_Start.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
                msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab1_DeliveryDate_Start.SelStart = 0: txt_Tab1_DeliveryDate_Start.SelLength = Len(txt_Tab1_DeliveryDate_Start.Text): txt_Tab1_DeliveryDate_Start.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab1_DeliveryDate_End.SelStart = 0
          txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text)
          txt_Tab1_DeliveryDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab1_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'退貨排車一覽表 >> 出車日期 >> 迄
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab1_DeliveryDate_End.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
                msg_text = "出車日期資料檢核錯誤：" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab1_DeliveryDate_End.SelStart = 0: txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text): txt_Tab1_DeliveryDate_End.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab1_RouteNo_Start.SelStart = 0: txt_Tab1_RouteNo_Start.SelLength = Len(txt_Tab1_RouteNo_Start.Text)
          txt_Tab1_RouteNo_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_RouteNo_End_KeyPress(KeyAscii As Integer)
'退貨排車一覽表 >> 路線編號 >> 迄
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab1_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_RouteNo_Start_KeyPress(KeyAscii As Integer)
'退貨排車一覽表 >> 路線編號 >> 起
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab1_RouteNo_End.SelStart = 0: txt_Tab1_RouteNo_End.SelLength = Len(txt_Tab1_RouteNo_End.Text)
          txt_Tab1_RouteNo_End.SetFocus
   End Select
End Sub

Public Function SaveTo_ExcelFile_OTHER(ByVal strFileName As String, ByRef in_rs As ADODB.Recordset, _
                Optional ByVal title As String, Optional ByVal OrientSelect As Integer) As Integer
'轉存 Excel 檔
Dim excelAP As Excel.Application
Dim tmp_col As Double, tmp_row As Double
Dim tmp_letter As String, tmp_RangNo As String, tmpI As Integer
Dim Dob_CS As Double, Dob_EA As Double, Dob_Total As Double, Str_Sku As String
Dim bl_first As Boolean
bl_first = True
Str_Sku = ""
Dob_CS = 0
Dob_EA = 0
Dob_Total = 0

'SaveTo_ExcelFile = 1
If TypeName(in_rs) = "Nothing" Then
   funRtn_msg = "轉檔錯誤：無轉出資料"
   Exit Function
ElseIf in_rs.RecordCount = 0 Then
   funRtn_msg = "轉檔錯誤：無轉出資料"
   Exit Function
End If

'設定執行狀態 Form 顯示
fgTransferToExcel = True
Load frm_WaitWindows
frm_WaitWindows.Tag = "Transfertoexcel"
frm_WaitWindows.ZOrder
frm_WaitWindows.Refresh
DoEvents: DoEvents

On Error GoTo err_Handle
Set excelAP = New Excel.Application
excelAP.Visible = False
excelAP.Workbooks.Add
DoEvents

'產生第一行：以欄位名稱當標題列
in_rs.MoveFirst
tmp_row = 1
For tmp_col = 0 To in_rs.Fields.Count - 1
    tmp_letter = Chr(65 + tmp_col)      ' A 之 ascii code
    If Asc(tmp_letter) > 90 Then        ' > Z 則變成 AA 起始
       tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
    End If
        tmp_RangNo = tmp_letter & (tmp_row)
        excelAP.Range(tmp_RangNo) = in_rs.Fields(tmp_col).Name

Next tmp_col
excelAP.Range("A1", tmp_RangNo).Select
excelAP.Selection.Font.Name = "新細明體"
excelAP.Selection.Font.FontStyle = "粗體"

'設定：跨頁表格欄位標題列印
With excelAP.ActiveSheet.PageSetup
     .PrintTitleRows = "$1:$1"
End With

'抄寫資料至 Excel File
tmp_row = tmp_row + 1
Do While Not in_rs.EOF
    DoEvents
    '判斷使用者是否取消轉檔作業
    If fgTransferToExcel = False Then
       err.Raise vbObjectError + 513, "Excel 轉檔作業", "使用者要求取消 Excel 轉檔作業，轉檔作業未完成"
    End If
            
         If Trim(in_rs.Fields("暫存碼頭").Value) <> Str_Sku Then
            '多一行加總數量，並將數量重置，重新設定碼頭
            If bl_first = False Then
                For tmp_col = 0 To in_rs.Fields.Count - 1
                    tmp_letter = Chr(65 + tmp_col)      ' A 之 ascii code
                    If Asc(tmp_letter) > 90 Then        ' > Z 則變成 AA 起始
                       tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                    End If
                        tmp_RangNo = tmp_letter & (tmp_row)
                        excelAP.Range(tmp_RangNo) = ""
                
                        With excelAP.Range(tmp_RangNo)
                            .NumberFormatLocal = "@"      '儲存格格式 >> 數字 >> 類別 = 文字
                            '.Font.Name = "新細明體"       '儲存格格式 >> 字型 >> 字型 = Times New Roman
                            '.Font.FontStyle = "標準"      '儲存格格式 >> 字型 >> 外型樣式 = 標準
                            '.Font.Size = 12               '儲存格格式 >> 字型 >> 大小 = 12
                            .Font.Name = "新細明體"
                            .Font.FontStyle = "粗體"
                            .Interior.Color = RGB(173, 255, 47)
                        End With
                        
                        If Left(tmp_RangNo, 1) = "I" Then excelAP.Range(tmp_RangNo) = "總數："
                        If Left(tmp_RangNo, 1) = "J" Then excelAP.Range(tmp_RangNo) = Dob_CS
                        If Left(tmp_RangNo, 1) = "K" Then excelAP.Range(tmp_RangNo) = Dob_EA
                        If Left(tmp_RangNo, 1) = "L" Then excelAP.Range(tmp_RangNo) = Dob_Total
    
                Next tmp_col
            End If
            
            Str_Sku = Trim(in_rs.Fields("暫存碼頭").Value)
            Dob_CS = Trim(in_rs.Fields("箱數").Value): Dob_EA = Trim(in_rs.Fields("個數").Value): Dob_Total = Trim(in_rs.Fields("總個數").Value)
            
            If bl_first = False Then
                tmp_row = tmp_row + 1
            End If
            bl_first = False
        Else
            Dob_CS = Dob_CS + Trim(in_rs.Fields("箱數").Value)
            Dob_EA = Dob_EA + Trim(in_rs.Fields("個數").Value)
            Dob_Total = Dob_Total + Trim(in_rs.Fields("總個數").Value)
        End If
        
    For tmp_col = 0 To in_rs.Fields.Count - 1
        tmp_letter = Chr(65 + tmp_col)      ' A 之 ascii code
        If Asc(tmp_letter) > 90 Then        ' > Z 則變成 AA 起始
           tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
        End If
        tmp_RangNo = tmp_letter & (tmp_row)
        '設定格式
        
        With excelAP.Range(tmp_RangNo)
            .NumberFormatLocal = "@"      '儲存格格式 >> 數字 >> 類別 = 文字
            .Font.Name = "新細明體"       '儲存格格式 >> 字型 >> 字型 = Times New Roman
            .Font.FontStyle = "標準"      '儲存格格式 >> 字型 >> 外型樣式 = 標準
            '.Font.Size = 12               '儲存格格式 >> 字型 >> 大小 = 12
        End With
        excelAP.Range(tmp_RangNo) = Trim(in_rs.Fields(tmp_col).Value)
    Next tmp_col
    in_rs.MoveNext
    tmp_row = tmp_row + 1
Loop

'最後一筆
                For tmp_col = 0 To in_rs.Fields.Count - 1
                    tmp_letter = Chr(65 + tmp_col)      ' A 之 ascii code
                    If Asc(tmp_letter) > 90 Then        ' > Z 則變成 AA 起始
                       tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                    End If
                        tmp_RangNo = tmp_letter & (tmp_row)
                
                        With excelAP.Range(tmp_RangNo)
                            .NumberFormatLocal = "@"      '儲存格格式 >> 數字 >> 類別 = 文字
                            '.Font.Name = "新細明體"       '儲存格格式 >> 字型 >> 字型 = Times New Roman
                            '.Font.FontStyle = "標準"      '儲存格格式 >> 字型 >> 外型樣式 = 標準
                            '.Font.Size = 12               '儲存格格式 >> 字型 >> 大小 = 12
                            .Font.Name = "新細明體"
                            .Font.FontStyle = "粗體"
                            .Interior.Color = RGB(173, 255, 47)
                        End With
                
                        excelAP.Range(tmp_RangNo) = ""
                        If Left(tmp_RangNo, 1) = "I" Then excelAP.Range(tmp_RangNo) = "總數："
                        If Left(tmp_RangNo, 1) = "J" Then excelAP.Range(tmp_RangNo) = Dob_CS
                        If Left(tmp_RangNo, 1) = "K" Then excelAP.Range(tmp_RangNo) = Dob_EA
                        If Left(tmp_RangNo, 1) = "L" Then excelAP.Range(tmp_RangNo) = Dob_Total
    
                Next tmp_col
                tmp_row = tmp_row + 1
            
'劃框線
DoEvents
'判斷使用者是否取消轉檔作業
If fgTransferToExcel = False Then
   err.Raise vbObjectError + 513, "Excel 轉檔作業", "使用者要求取消 Excel 轉檔作業，轉檔作業未完成"
End If
excelAP.Range("A1", tmp_RangNo).Select
With excelAP.Selection
     .Font.Name = "新細明體"
     .Font.Size = 9
     .Borders(xlEdgeLeft).LineStyle = xlContinuous
     .Borders(xlEdgeLeft).Weight = xlThin
     '.Borders(xlEdgeLeft).ColorIndex = xlAutomatic
     .Borders(xlEdgeTop).LineStyle = xlContinuous
     .Borders(xlEdgeTop).Weight = xlThin
     '.Borders(xlEdgeTop).ColorIndex = xlAutomatic
     .Borders(xlEdgeBottom).LineStyle = xlContinuous
     .Borders(xlEdgeBottom).Weight = xlThin
     '.Borders(xlEdgeBottom).ColorIndex = xlAutomatic
     .Borders(xlEdgeRight).LineStyle = xlContinuous
     .Borders(xlEdgeRight).Weight = xlThin
     '.Borders(xlEdgeRight).ColorIndex = xlAutomatic
     .Borders(xlInsideVertical).LineStyle = xlContinuous
     .Borders(xlInsideVertical).Weight = xlThin
     '.Borders(xlInsideVertical).ColorIndex = xlAutomatic
     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
     .Borders(xlInsideHorizontal).Weight = xlThin
     '.Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
End With

'自動調整欄寬
DoEvents
Dim str_cnStart As String, str_cnEnd As String, str_cn As String
Dim int_cn As Integer
int_cn = 1
Do   '取得欄位
  Select Case Mid(tmp_RangNo, int_cn, 1)
         Case 0 To 9
              Exit Do
         Case Else
              int_cn = int_cn + 1
  End Select
Loop
str_cnEnd = Mid(tmp_RangNo, 1, int_cn - 1)
str_cnStart = "A"
DoEvents
Do
   '判斷使用者是否取消轉檔作業
   If fgTransferToExcel = False Then
      err.Raise vbObjectError + 513, "Excel 轉檔作業", "使用者要求取消 Excel 轉檔作業，轉檔作業未完成"
   End If
   str_cn = str_cnStart & ":" & str_cnStart
   excelAP.Columns(str_cn).EntireColumn.AutoFit
   If str_cnStart = str_cnEnd Then
      Exit Do
   End If
   If str_cnStart = "Z" Then
      str_cnStart = "AA"
   Else
      If Len(str_cnStart) > 1 Then
         str_cnStart = "A" & Chr(Asc(Mid(str_cnStart, 2, 1)) + 1)
      Else
         str_cnStart = Chr(Asc(str_cnStart) + 1)
      End If
   End If
   DoEvents
Loop

'記錄轉檔時間
tmp_row = tmp_row + 1
tmp_RangNo = "A" & (tmp_row)
excelAP.Range(tmp_RangNo) = "轉檔人員：" & Get_LoginUserName
tmp_row = tmp_row + 1
tmp_RangNo = "A" & (tmp_row)
excelAP.Range(tmp_RangNo) = "電腦名稱：" & GetComputerName_rtnString
tmp_row = tmp_row + 1
tmp_RangNo = "A" & (tmp_row)
excelAP.Range(tmp_RangNo) = "轉檔時間：" & Format(Now, "yyyy/mm/dd hh:nn:ss")

'自訂頁首
With excelAP.ActiveSheet.PageSetup
     If Len(title) > 0 Then
        .CenterHeader = "&""標楷體,粗體""&18" & title
     End If
     .RightFooter = "共&""Times New Roman,標準"" &N &""細明體,標準""頁，&""新細明體,標準""第&""Times New Roman,標準"" &P &""新細明體,標準""頁"
     If OrientSelect = 1 Then
        .Orientation = xlLandscape    '橫印
        .LeftMargin = excelAP.InchesToPoints(0.1)
        .RightMargin = excelAP.InchesToPoints(0.1)
        .TopMargin = excelAP.InchesToPoints(0.81)
        .BottomMargin = excelAP.InchesToPoints(0.62)
        .HeaderMargin = excelAP.InchesToPoints(0.39)
        .FooterMargin = excelAP.InchesToPoints(0.36)
     End If
End With

DoEvents
If Len(strFileName) > 0 Then
   excelAP.ActiveWorkbook.SaveAs FileName:=strFileName, FileFormat:=xlNormal
   excelAP.ActiveWindow.Close
   excelAP.Visible = False
   Set excelAP = Nothing
Else
   excelAP.Visible = True
End If
in_rs.MoveFirst

'關閉執行狀態 Form
Unload frm_WaitWindows
Set frm_WaitWindows = Nothing
fgTransferToExcel = True

'SaveTo_ExcelFile = 0
Exit Function

err_Handle:
   fgTransferToExcel = False
   Call Release_RunningForm
   Dim tmpString As String
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   funRtn_msg = "轉存 excel 作業程序失敗，錯誤訊息如下：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   If TypeName(excelAP) <> "Nothing" Then
      excelAP.ActiveWorkbook.Close SaveChanges:=False
      Set excelAP = Nothing
   End If
End Function

