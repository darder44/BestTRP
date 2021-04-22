VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Query_Orders 
   Caption         =   "  訂   單   資   料   查   詢"
   ClientHeight    =   7725
   ClientLeft      =   150
   ClientTop       =   960
   ClientWidth     =   12570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   12570
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   5520
      TabIndex        =   1
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
      StartOfWeek     =   98107393
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7320
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   12912
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "查詢條件"
      TabPicture(0)   =   "frm_Query_Orders.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "欄位選取"
      TabPicture(1)   =   "frm_Query_Orders.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "查詢結果"
      TabPicture(2)   =   "frm_Query_Orders.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "CmnDialog"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dg_Result"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmd_Exit(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmd_Tab2SavetoExcel"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame3 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   11415
         Begin VB.Frame Frame15 
            Height          =   885
            Left            =   120
            TabIndex        =   149
            Top             =   5880
            Width           =   4950
            Begin VB.ComboBox cmb_RSC 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1155
               Style           =   2  '單純下拉式
               TabIndex        =   151
               Top             =   180
               Width           =   2250
            End
            Begin VB.ComboBox cmb_RBC 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1155
               Style           =   2  '單純下拉式
               TabIndex        =   150
               Top             =   510
               Width           =   2250
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "異常原因"
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
               Index           =   34
               Left            =   285
               TabIndex        =   153
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "責任歸屬"
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
               Index           =   35
               Left            =   285
               TabIndex        =   152
               Top             =   570
               Width           =   840
            End
         End
         Begin VB.Frame Frame14 
            Height          =   2595
            Left            =   5115
            TabIndex        =   134
            Top             =   4170
            Width           =   6135
            Begin VB.ComboBox cmb_Zip 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '單純下拉式
               TabIndex        =   141
               Top             =   195
               Width           =   2565
            End
            Begin VB.ComboBox cmb_AreaCode 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '單純下拉式
               TabIndex        =   140
               Top             =   525
               Width           =   4455
            End
            Begin VB.ComboBox cmb_ExtraDemand 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '單純下拉式
               TabIndex        =   139
               Top             =   855
               Width           =   4455
            End
            Begin VB.ComboBox cmb_VehicleType 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '單純下拉式
               TabIndex        =   138
               Top             =   1185
               Width           =   4455
            End
            Begin VB.ComboBox cmb_TRPCompany 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '單純下拉式
               TabIndex        =   137
               Top             =   1515
               Width           =   4455
            End
            Begin VB.ComboBox cmb_STRPCompany 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '單純下拉式
               TabIndex        =   136
               Top             =   2175
               Width           =   4455
            End
            Begin VB.ComboBox cmb_SVehicleType 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '單純下拉式
               TabIndex        =   135
               Top             =   1845
               Width           =   4455
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "郵遞區號"
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
               Index           =   14
               Left            =   660
               TabIndex        =   148
               Top             =   255
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "運送區域"
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
               Index           =   9
               Left            =   660
               TabIndex        =   147
               Top             =   585
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "特殊需求"
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
               Index           =   6
               Left            =   660
               TabIndex        =   146
               Top             =   915
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "運送車種"
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
               Index           =   7
               Left            =   660
               TabIndex        =   145
               Top             =   1245
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "貨運公司"
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
               Left            =   660
               TabIndex        =   144
               Top             =   1575
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "二次貨運公司"
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
               Index           =   32
               Left            =   240
               TabIndex        =   143
               Top             =   2235
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "二次運送車種"
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
               Index           =   33
               Left            =   240
               TabIndex        =   142
               Top             =   1905
               Width           =   1260
            End
         End
         Begin VB.Frame Frame13 
            Height          =   1815
            Left            =   120
            TabIndex        =   118
            Top             =   4155
            Width           =   4950
            Begin VB.TextBox txt_SRouteNo_End 
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
               Left            =   3075
               TabIndex        =   125
               Top             =   525
               Width           =   1290
            End
            Begin VB.TextBox txt_SRouteNo_Start 
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
               Left            =   1485
               TabIndex        =   124
               Top             =   525
               Width           =   1290
            End
            Begin VB.TextBox txt_SDeliveryDate_End 
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
               Left            =   3075
               TabIndex        =   123
               Top             =   840
               Width           =   1290
            End
            Begin VB.TextBox txt_SDeliveryDate_Start 
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
               Left            =   1485
               TabIndex        =   122
               Top             =   825
               Width           =   1290
            End
            Begin VB.TextBox txt_SVehicleID 
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
               Left            =   1485
               TabIndex        =   121
               Top             =   1125
               Width           =   1290
            End
            Begin VB.CheckBox chk_SecondPlan 
               Caption         =   "篩選進行二次排車"
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
               Left            =   570
               TabIndex        =   120
               Top             =   225
               Width           =   2355
            End
            Begin VB.TextBox txt_SAddWho 
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
               Left            =   1485
               TabIndex        =   119
               Top             =   1425
               Width           =   1290
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
               Index           =   27
               Left            =   2820
               TabIndex        =   133
               Top             =   555
               Width           =   240
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "二次排車路編"
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
               Index           =   6
               Left            =   180
               TabIndex        =   132
               Top             =   570
               Width           =   1260
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
               Index           =   28
               Left            =   2820
               TabIndex        =   131
               Top             =   855
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "二次出車日期"
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
               Left            =   180
               TabIndex        =   130
               Top             =   885
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "二次車牌號碼"
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
               Index           =   30
               Left            =   180
               TabIndex        =   129
               Top             =   1185
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "二次排車者"
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
               Index           =   31
               Left            =   390
               TabIndex        =   128
               Top             =   1485
               Width           =   1050
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   4
               Left            =   2790
               TabIndex        =   127
               Top             =   1170
               Width           =   240
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   5
               Left            =   2790
               TabIndex        =   126
               Top             =   1470
               Width           =   240
            End
         End
         Begin VB.Frame Frame12 
            Height          =   2370
            Left            =   120
            TabIndex        =   70
            Top             =   1875
            Width           =   8610
            Begin VB.TextBox txt_FReceiptNo 
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
               Left            =   6840
               TabIndex        =   94
               Top             =   1725
               Width           =   1425
            End
            Begin VB.TextBox txt_FAddWho 
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
               Left            =   6840
               TabIndex        =   93
               Top             =   525
               Width           =   1410
            End
            Begin VB.ComboBox cmb_EXEConfirm 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               ItemData        =   "frm_Query_Orders.frx":0054
               Left            =   6840
               List            =   "frm_Query_Orders.frx":0061
               Style           =   2  '單純下拉式
               TabIndex        =   92
               Top             =   195
               Width           =   1440
            End
            Begin VB.TextBox txt_FVehicleID 
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
               Left            =   6840
               TabIndex        =   91
               Top             =   825
               Width           =   1410
            End
            Begin VB.TextBox txt_FDriver 
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
               Left            =   6840
               TabIndex        =   90
               Top             =   1125
               Width           =   1410
            End
            Begin VB.TextBox txt_FDockNo 
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
               Left            =   6840
               TabIndex        =   89
               Top             =   1425
               Width           =   1410
            End
            Begin VB.TextBox txt_FCheckout_End 
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
               Left            =   3240
               TabIndex        =   88
               Top             =   1350
               Width           =   1125
            End
            Begin VB.TextBox txt_FCheckout_Start 
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
               Left            =   1845
               TabIndex        =   87
               Top             =   1350
               Width           =   1125
            End
            Begin VB.TextBox txt_FCheckin_End 
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
               Left            =   3240
               TabIndex        =   86
               Top             =   1050
               Width           =   1125
            End
            Begin VB.TextBox txt_FCheckin_Start 
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
               Left            =   1845
               TabIndex        =   85
               Top             =   1050
               Width           =   1125
            End
            Begin VB.TextBox txt_FPlanCheckin_End 
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
               Left            =   3780
               TabIndex        =   84
               Top             =   1665
               Width           =   1125
            End
            Begin VB.TextBox txt_FPlanCheckin_Start 
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
               Left            =   1845
               TabIndex        =   83
               Top             =   1665
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
               Left            =   3240
               TabIndex        =   82
               Top             =   756
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
               Left            =   1845
               TabIndex        =   81
               Top             =   756
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
               Left            =   3240
               TabIndex        =   80
               Top             =   453
               Width           =   1125
            End
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
               Left            =   1845
               TabIndex        =   79
               Top             =   453
               Width           =   1125
            End
            Begin VB.TextBox txt_FRouteNo_End 
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
               Left            =   3495
               TabIndex        =   78
               Top             =   150
               Width           =   1300
            End
            Begin VB.TextBox txt_FRouteNo_Start 
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
               Left            =   1845
               TabIndex        =   77
               Top             =   150
               Width           =   1300
            End
            Begin VB.TextBox txt_FPlanCheckinTime_Start 
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
               Left            =   2985
               TabIndex        =   76
               Top             =   1665
               Width           =   525
            End
            Begin VB.TextBox txt_FPlanCheckinTime_End 
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
               Left            =   4920
               TabIndex        =   75
               Top             =   1665
               Width           =   525
            End
            Begin VB.TextBox txt_SDNTime_End 
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
               Left            =   4920
               TabIndex        =   74
               Top             =   1950
               Width           =   525
            End
            Begin VB.TextBox txt_SDNTime_Start 
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
               Left            =   2985
               TabIndex        =   73
               Top             =   1950
               Width           =   525
            End
            Begin VB.TextBox txt_SDNDate_Start 
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
               Left            =   1845
               TabIndex        =   72
               Top             =   1950
               Width           =   1125
            End
            Begin VB.TextBox txt_SDNDate_End 
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
               Left            =   3780
               TabIndex        =   71
               Top             =   1950
               Width           =   1125
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "TMS單號"
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
               Index           =   4
               Left            =   5940
               TabIndex        =   117
               Top             =   1785
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "排車者"
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
               Left            =   6150
               TabIndex        =   116
               Top             =   585
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "回傳狀態"
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
               Left            =   5940
               TabIndex        =   115
               Top             =   270
               Width           =   840
            End
            Begin VB.Label Label3 
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
               Height          =   195
               Index           =   18
               Left            =   5940
               TabIndex        =   114
               Top             =   885
               Width           =   840
            End
            Begin VB.Label Label3 
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
               Height          =   195
               Index           =   19
               Left            =   6150
               TabIndex        =   113
               Top             =   1185
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   22
               Left            =   5940
               TabIndex        =   112
               Top             =   1485
               Width           =   840
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   0
               Left            =   8220
               TabIndex        =   111
               Top             =   555
               Width           =   240
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   2
               Left            =   8220
               TabIndex        =   110
               Top             =   885
               Width           =   240
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   3
               Left            =   8220
               TabIndex        =   109
               Top             =   1170
               Width           =   240
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
               Index           =   26
               Left            =   3000
               TabIndex        =   108
               Top             =   1140
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "車輛離倉日期"
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
               Index           =   25
               Left            =   525
               TabIndex        =   107
               Top             =   1455
               Width           =   1260
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
               Index           =   24
               Left            =   3000
               TabIndex        =   106
               Top             =   1440
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "車輛報到日期"
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
               Index           =   23
               Left            =   525
               TabIndex        =   105
               Top             =   1155
               Width           =   1260
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
               Left            =   3525
               TabIndex        =   104
               Top             =   1725
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "預計報到日期時間"
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
               Index           =   20
               Left            =   105
               TabIndex        =   103
               Top             =   1755
               Width           =   1680
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
               Left            =   3000
               TabIndex        =   102
               Top             =   840
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
               Left            =   945
               TabIndex        =   101
               Top             =   840
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
               Left            =   3000
               TabIndex        =   100
               Top             =   510
               Width           =   240
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
               Left            =   945
               TabIndex        =   99
               Top             =   540
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
               Index           =   10
               Left            =   3225
               TabIndex        =   98
               Top             =   210
               Width           =   240
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
               Height          =   195
               Index           =   5
               Left            =   945
               TabIndex        =   97
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "揀貨確認日期時間"
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
               Index           =   36
               Left            =   120
               TabIndex        =   96
               Top             =   2040
               Width           =   1680
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
               Index           =   37
               Left            =   3525
               TabIndex        =   95
               Top             =   2010
               Width           =   360
            End
         End
         Begin VB.CommandButton cmd_Tab0_Reset 
            BackColor       =   &H00C0FFC0&
            Caption         =   "清除查詢條件"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":0081
            Style           =   1  '圖片外觀
            TabIndex        =   69
            Top             =   2220
            Width           =   1935
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
            Height          =   900
            Index           =   3
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":0393
            Style           =   1  '圖片外觀
            TabIndex        =   68
            Top             =   3135
            Width           =   1935
         End
         Begin VB.CommandButton cmd_Tab0_SelectField 
            BackColor       =   &H00C0E0FF&
            Caption         =   "回傳欄位選取"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":07D5
            Style           =   1  '圖片外觀
            TabIndex        =   67
            Top             =   390
            Width           =   1935
         End
         Begin VB.CommandButton cmd_Query 
            BackColor       =   &H008080FF&
            Caption         =   "訂 單 查 詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   0
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":0ADF
            Style           =   1  '圖片外觀
            TabIndex        =   66
            Top             =   1305
            Width           =   1935
         End
         Begin VB.Frame Frame11 
            Height          =   1845
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   8610
            Begin VB.TextBox txt_StorerKey 
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
               TabIndex        =   52
               Top             =   180
               Width           =   1125
            End
            Begin VB.TextBox txt_Extern_Start 
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
               TabIndex        =   51
               Top             =   495
               Width           =   1125
            End
            Begin VB.TextBox txt_Extern_End 
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
               TabIndex        =   50
               Top             =   495
               Width           =   1125
            End
            Begin VB.TextBox txt_OrderDate_Start 
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
               TabIndex        =   49
               Top             =   810
               Width           =   1125
            End
            Begin VB.TextBox txt_OrderDate_End 
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
               TabIndex        =   48
               Top             =   810
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
               TabIndex        =   47
               Top             =   1125
               Width           =   1125
            End
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
               TabIndex        =   46
               Top             =   1125
               Width           =   1125
            End
            Begin VB.TextBox txt_ConsigneeKey 
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
               Left            =   4680
               TabIndex        =   45
               Top             =   180
               Width           =   2640
            End
            Begin VB.TextBox txt_ConsigName 
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
               Left            =   4680
               TabIndex        =   44
               Top             =   495
               Width           =   1890
            End
            Begin VB.TextBox txt_SKU 
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
               TabIndex        =   43
               Top             =   1455
               Width           =   2520
            End
            Begin VB.CheckBox chk_OnlyExpireDate 
               Caption         =   "指定到期日"
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
               Left            =   3825
               TabIndex        =   42
               Top             =   915
               Width           =   1830
            End
            Begin VB.CheckBox txt_OrderNotes 
               Caption         =   "訂單備註"
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
               Left            =   5820
               TabIndex        =   41
               Top             =   915
               Width           =   1455
            End
            Begin VB.CheckBox chk_NotImport 
               Caption         =   "未轉入訂單"
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
               Left            =   3825
               TabIndex        =   40
               Top             =   1215
               Width           =   1455
            End
            Begin VB.CheckBox chk_WaitPlan 
               Caption         =   "已轉入訂單待排車"
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
               Left            =   3825
               TabIndex        =   39
               Top             =   1515
               Width           =   2085
            End
            Begin VB.CheckBox chk_CancelOrder 
               Caption         =   "取消訂單"
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
               Left            =   5820
               TabIndex        =   38
               Top             =   1215
               Width           =   1455
            End
            Begin VB.CheckBox chk_ExpectOrder 
               Caption         =   "簽收異常訂單"
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
               Left            =   5820
               TabIndex        =   37
               Top             =   1500
               Width           =   1665
            End
            Begin VB.CheckBox chk_Ship_qty 
               Caption         =   "未揀貨訂單"
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
               Left            =   7080
               TabIndex        =   36
               Top             =   915
               Width           =   1455
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "貨主"
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
               Index           =   0
               Left            =   570
               TabIndex        =   65
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "訂單編號"
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
               Index           =   0
               Left            =   150
               TabIndex        =   64
               Top             =   555
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
               Index           =   1
               Left            =   2205
               TabIndex        =   63
               Top             =   525
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "訂單日期"
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
               Index           =   2
               Left            =   150
               TabIndex        =   62
               Top             =   870
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
               Index           =   3
               Left            =   2205
               TabIndex        =   61
               Top             =   840
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "送貨日期"
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
               Index           =   4
               Left            =   150
               TabIndex        =   60
               Top             =   1185
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
               Index           =   5
               Left            =   2205
               TabIndex        =   59
               Top             =   1155
               Width           =   240
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "客戶編號"
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
               Left            =   3780
               TabIndex        =   58
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "客戶名稱"
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
               Index           =   2
               Left            =   3780
               TabIndex        =   57
               Top             =   555
               Width           =   840
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "貨號"
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
               Index           =   3
               Left            =   555
               TabIndex        =   56
               Top             =   1515
               Width           =   420
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   1
               Left            =   3570
               TabIndex        =   55
               Top             =   1485
               Width           =   240
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   6
               Left            =   7650
               TabIndex        =   54
               Top             =   225
               Width           =   240
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   7
               Left            =   6915
               TabIndex        =   53
               Top             =   525
               Width           =   240
            End
         End
         Begin VB.Shape Shape11 
            BackColor       =   &H00004080&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000D&
            BorderWidth     =   2
            Height          =   3765
            Left            =   8820
            Top             =   330
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   525
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   9915
         Begin VB.TextBox txt_Tab2_srcTotal_DifSDN 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   8520
            TabIndex        =   32
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_DifPick 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   6700
            TabIndex        =   30
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Case 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1335
            TabIndex        =   26
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_PickCase 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   2940
            TabIndex        =   25
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_SDNCase 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4545
            TabIndex        =   24
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "簽收差異數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   4
            Left            =   7575
            TabIndex        =   33
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "揀貨差異數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   3
            Left            =   5760
            TabIndex        =   31
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "簽收數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   9
            Left            =   3915
            TabIndex        =   29
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "揀貨數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   10
            Left            =   2325
            TabIndex        =   28
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "總計：排車數"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   11
            Left            =   195
            TabIndex        =   27
            Top             =   210
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmd_Tab2SavetoExcel 
         BackColor       =   &H00FFFFC0&
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
         Height          =   1020
         Left            =   10275
         Picture         =   "frm_Query_Orders.frx":0DE9
         Style           =   1  '圖片外觀
         TabIndex        =   22
         Top             =   495
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
         Height          =   1020
         Index           =   0
         Left            =   10275
         Picture         =   "frm_Query_Orders.frx":19AB
         Style           =   1  '圖片外觀
         TabIndex        =   21
         Top             =   2385
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         Height          =   6315
         Left            =   -74685
         TabIndex        =   3
         Top             =   495
         Width           =   10785
         Begin VB.CommandButton cmd_Query 
            BackColor       =   &H008080FF&
            Caption         =   "查  詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Index           =   1
            Left            =   7830
            Picture         =   "frm_Query_Orders.frx":1DED
            Style           =   1  '圖片外觀
            TabIndex        =   17
            Top             =   360
            Width           =   2385
         End
         Begin VB.ListBox lst_AllFields 
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
            Height          =   5580
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   2800
         End
         Begin VB.ListBox lst_SelectedFields 
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
            Height          =   5580
            Left            =   4125
            TabIndex        =   15
            Top             =   600
            Width           =   2800
         End
         Begin VB.CommandButton cmd_Tab1_Add 
            BackColor       =   &H008080FF&
            Caption         =   "＞＞"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3120
            Style           =   1  '圖片外觀
            TabIndex        =   14
            ToolTipText     =   "欄位選取"
            Top             =   3270
            Width           =   855
         End
         Begin VB.CommandButton cmd_Tab1_Remove 
            BackColor       =   &H0080C0FF&
            Caption         =   "＜＜"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3120
            Style           =   1  '圖片外觀
            TabIndex        =   13
            ToolTipText     =   "移除欄位"
            Top             =   3720
            Width           =   855
         End
         Begin VB.CommandButton cmd_Tab1_Down 
            BackColor       =   &H00FF80FF&
            Caption         =   "下移"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   3570
            Style           =   1  '圖片外觀
            TabIndex        =   12
            Top             =   1425
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_Up 
            BackColor       =   &H00FF8080&
            Caption         =   "上移"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   3570
            Style           =   1  '圖片外觀
            TabIndex        =   11
            Top             =   675
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_Reload 
            BackColor       =   &H00C0C0C0&
            Caption         =   "重新載入"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   2940
            Style           =   1  '圖片外觀
            TabIndex        =   10
            ToolTipText     =   "載入查詢設定值"
            Top             =   5010
            Width           =   1185
         End
         Begin VB.ListBox lst_OrderBy 
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
            Height          =   2940
            Left            =   7785
            TabIndex        =   9
            Top             =   2550
            Width           =   2800
         End
         Begin VB.CommandButton cmd_Tab1_OrderRemove 
            BackColor       =   &H0080C0FF&
            Caption         =   "＜＜"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7035
            Style           =   1  '圖片外觀
            TabIndex        =   8
            Top             =   4740
            Width           =   690
         End
         Begin VB.CommandButton cmd_Tab1_OrderAdd 
            BackColor       =   &H008080FF&
            Caption         =   "＞＞"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7035
            Style           =   1  '圖片外觀
            TabIndex        =   7
            Top             =   4290
            Width           =   690
         End
         Begin VB.CommandButton cmd_Tab1_OrderByUp 
            BackColor       =   &H00FF8080&
            Caption         =   "上移"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   7215
            Style           =   1  '圖片外觀
            TabIndex        =   6
            Top             =   2625
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_OrderByDown 
            BackColor       =   &H00FF80FF&
            Caption         =   "下移"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   7215
            Style           =   1  '圖片外觀
            TabIndex        =   5
            Top             =   3375
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_Reset 
            BackColor       =   &H008080FF&
            Caption         =   "清  除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2925
            Style           =   1  '圖片外觀
            TabIndex        =   4
            ToolTipText     =   "清除所有設定值"
            Top             =   5700
            Width           =   1200
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H80000001&
            BackStyle       =   1  '不透明
            BorderColor     =   &H0000C000&
            BorderWidth     =   2
            Height          =   1140
            Left            =   7770
            Top             =   315
            Width           =   2505
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "待 選 欄 位 列 表"
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
            Index           =   0
            Left            =   615
            TabIndex        =   20
            Top             =   285
            Width           =   1905
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "已 選 欄 位 列 表"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   1
            Left            =   4605
            TabIndex        =   19
            Top             =   285
            Width           =   1905
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00C0C000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   0
            Left            =   3510
            Top             =   615
            Width           =   630
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00400040&
            BackStyle       =   1  '不透明
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            Height          =   1005
            Index           =   1
            Left            =   3060
            Top             =   3210
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "排 序 設 定"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   240
            Index           =   2
            Left            =   8610
            TabIndex        =   18
            Top             =   2145
            Width           =   1245
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00400040&
            BackStyle       =   1  '不透明
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            Height          =   1005
            Index           =   2
            Left            =   6975
            Top             =   4230
            Width           =   810
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00C0C000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   3
            Left            =   7155
            Top             =   2565
            Width           =   630
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            Height          =   465
            Index           =   0
            Left            =   495
            Top             =   180
            Width           =   2145
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            Height          =   465
            Index           =   1
            Left            =   4485
            Top             =   180
            Width           =   2145
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404000&
            BackStyle       =   1  '不透明
            Height          =   465
            Index           =   2
            Left            =   8145
            Top             =   2040
            Width           =   2145
         End
      End
      Begin MSDataGridLib.DataGrid dg_Result 
         Height          =   5925
         Left            =   195
         TabIndex        =   2
         Top             =   990
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   10451
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
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   10485
         Top             =   1740
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frm_Query_Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬
Private iLoop As Double

Private arZip() As String            '郵遞區號
Private arAreaCode() As String       '運送區域
Private arExtraDemand() As String    '特殊需求
Private arVehicleType() As String    '車種
Private arTRPCompany() As String     '貨運公司
Private arRSC() As String            '異常原因
Private arRBC() As String            '責任歸屬
Private MyXlsApp As Excel.Application

Private rs_Result As ADODB.Recordset

Private Sub cmd_Tab2SavetoExcel_Click()
'訂單查詢 >> 轉 EXCEL
Recordset2Excel Me.Caption, rs_Result
'..在此編輯EXCEL
Set MyXlsApp = Nothing

'If rs_Result Is Nothing Then Exit Sub
'If rs_Result.RecordCount = 0 Then Exit Sub
'
'Dim ExcelTitle As String
'Call DocStoreDirectory(strDocPath)
'
'Dim strTranFileName As String           'Excel 檔案名稱
'CmnDialog.DialogTitle = "轉存 Excel 檔"
'CmnDialog.InitDir = "c:\my documents"
'CmnDialog.FileName = "訂單查詢_" & Format(Now, "YYYYMMDDHHNNSS")
'CmnDialog.Filter = "Excel檔案(*.xls)|*.xls"
'CmnDialog.FilterIndex = 1
'CmnDialog.CancelError = True
'On Error Resume Next
'CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
'CmnDialog.ShowOpen
'If Err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
'   msg_text = "選擇 [取消] 按鈕，必須於 Excel 中自行存檔"
'   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
'   strTranFileName = ""
'Else
'   strTranFileName = CmnDialog.FileName
'   If Dir(strTranFileName) <> "" Then
'      Kill strTranFileName
'   End If
'End If
'
'On Error GoTo err_Handle
'Screen.MousePointer = vbHourglass
'If SaveTo_ExcelFile(strTranFileName, rs_Result) = 1 Then
'   Screen.MousePointer = vbDefault
'   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
'Else
'   Screen.MousePointer = vbDefault
'   If Len(strTranFileName) > 0 Then
'      msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
'      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   End If
'End If
'If rs_Result Is Nothing Then Exit Sub
'rs_Result.MoveFirst
'Exit Sub

'err_Handle:
'   Dim tmpString As String
'   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'   CreateErrorLog Me.Name & "-轉 EXCEL", Me.Caption, "cmd_Tab2SavetoExcel_Click", tmpString
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Query_Click(Index As Integer)
'檢查排序欄位有的則不可出現在待排序中
Dim i As Integer: Dim j As Integer
For i = 0 To lst_OrderBy.ListCount
    For j = 0 To lst_AllFields.ListCount
         If lst_OrderBy.List(i) = lst_AllFields.List(j) And lst_OrderBy.List(i) <> "" Then
               msg_text = "請檢查'排序欄位'不可出現在'待選欄位'中!" & vbCrLf & "錯誤欄位:" & lst_OrderBy.List(i)
               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
               GoTo err:
         End If
    Next
Next

' 查詢
If lst_SelectedFields.ListCount = 0 Then
   msg_text = "作業程序錯誤：並未選取查詢結果回傳的欄位"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Screen.MousePointer = vbHourglass
DoEvents: DoEvents
On Error GoTo err_Handle

'使用者查詢環境設定值存檔
Call SaveQueryEnv
Set dg_Result.DataSource = Nothing
Set rs_Result = Nothing

'組合選取欄位
Dim strSelect As String, strOrderBy As String
strSelect = ""
For iLoop = 0 To lst_SelectedFields.ListCount - 1
    If strSelect = "" Then
       strSelect = lst_SelectedFields.List(iLoop)
    Else
       strSelect = strSelect & "," & lst_SelectedFields.List(iLoop)
    End If
Next iLoop
strSelect = "Select Distinct " & strSelect & " From Query_OrdersData "
strOrderBy = ""
For iLoop = 0 To lst_OrderBy.ListCount - 1
    If strOrderBy = "" Then
       strOrderBy = lst_OrderBy.List(iLoop)
    Else
       strOrderBy = strOrderBy & "," & lst_OrderBy.List(iLoop)
    End If
Next iLoop

'組合查詢條件
Dim str_Where As String, strSubwhere As String, intloop As Integer, tmp_data() As String
str_Where = ""
'Storer
txt_StorerKey.Text = Trim(txt_StorerKey.Text)
strSubwhere = ""
If txt_StorerKey.Text <> "" Then
   strSubwhere = " 貨主 = '" & txt_StorerKey.Text & "' "
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'貨主單號
txt_Extern_Start.Text = Trim(txt_Extern_Start.Text)
txt_Extern_End.Text = Trim(txt_Extern_End.Text)
strSubwhere = ""
If Len(txt_Extern_Start.Text) > 0 And Len(txt_Extern_End.Text) > 0 Then
   strSubwhere = " 貨主單號 Between '" & txt_Extern_Start.Text & "' and '" & txt_Extern_End.Text & "' "
ElseIf Len(txt_Extern_Start.Text) > 0 And Len(txt_Extern_End.Text) = 0 Then
   strSubwhere = " 貨主單號 = '" & txt_Extern_Start.Text & "' "
ElseIf Len(txt_Extern_Start.Text) = 0 And Len(txt_Extern_End.Text) > 0 Then
   strSubwhere = " 貨主單號 = '" & txt_Extern_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'訂單日期
txt_OrderDate_Start.Text = Trim(txt_OrderDate_Start.Text)
txt_OrderDate_End.Text = Trim(txt_OrderDate_End.Text)
strSubwhere = ""
If Len(txt_OrderDate_Start.Text) > 0 And Len(txt_OrderDate_End.Text) > 0 Then
   strSubwhere = " 訂單日期 Between '" & txt_OrderDate_Start.Text & "' and '" & txt_OrderDate_End.Text & "' "
ElseIf Len(txt_OrderDate_Start.Text) > 0 And Len(txt_OrderDate_End.Text) = 0 Then
   strSubwhere = " 訂單日期 = '" & txt_Extern_Start.Text & "' "
ElseIf Len(txt_OrderDate_Start.Text) = 0 And Len(txt_OrderDate_End.Text) > 0 Then
   strSubwhere = " 訂單日期 = '" & txt_Extern_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'送貨日期
txt_DeliveryDate_Start.Text = Trim(txt_DeliveryDate_Start.Text)
txt_DeliveryDate_End.Text = Trim(txt_DeliveryDate_End.Text)
strSubwhere = ""
If Len(txt_DeliveryDate_Start.Text) > 0 And Len(txt_DeliveryDate_End.Text) > 0 Then
   strSubwhere = " 送貨日期 Between '" & txt_DeliveryDate_Start.Text & "' and '" & txt_DeliveryDate_End.Text & "' "
ElseIf Len(txt_DeliveryDate_Start.Text) > 0 And Len(txt_DeliveryDate_End.Text) = 0 Then
   strSubwhere = " 送貨日期 = '" & txt_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_DeliveryDate_Start.Text) = 0 And Len(txt_DeliveryDate_End.Text) > 0 Then
   strSubwhere = " 送貨日期 = '" & txt_DeliveryDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'貨號
txt_SKU.Text = Trim(txt_SKU.Text)
strSubwhere = ""
If Len(txt_SKU.Text) > 0 Then
   If InStr(txt_SKU.Text, ",") > 0 Then
      tmp_data = Split(txt_SKU.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 貨號 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 貨號 in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " 貨號 like '%" & txt_SKU.Text & "%' "
      Else
         str_Where = str_Where & " and 貨號 like '%" & txt_SKU.Text & "%' "
      End If
   End If
End If
'客戶編號
txt_ConsigneeKey.Text = Trim(txt_ConsigneeKey.Text)
strSubwhere = ""
If Len(txt_ConsigneeKey.Text) > 0 Then
   If InStr(txt_ConsigneeKey.Text, ",") > 0 Then
      tmp_data = Split(txt_ConsigneeKey.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 客戶編號 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 客戶編號 in (" & strSubwhere & ") "
      End If
  Else
      If Len(str_Where) = 0 Then
         str_Where = " 客戶編號 like '%" & txt_ConsigneeKey.Text & "%' "
      Else
         str_Where = str_Where & " and 客戶編號 like '%" & txt_ConsigneeKey.Text & "%' "
      End If
  End If
End If
'客戶名稱
txt_ConsigName.Text = Trim(txt_ConsigName.Text)
If txt_ConsigName.Text <> "" Then
   If Len(str_Where) = 0 Then
      str_Where = " 客戶名稱 like '%" & txt_ConsigName.Text & "%' "
   Else
      str_Where = str_Where & " and 客戶名稱 like '%" & strSubwhere & "%' "
   End If
End If
'指定到期日
If chk_OnlyExpireDate.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " 註記 <> '' "
   Else
      str_Where = str_Where & " and 註記 <> ''"
   End If
End If
'訂單備註
If txt_OrderNotes.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " 訂單備註 <> '' "
   Else
      str_Where = str_Where & " and 訂單備註 <> '' "
   End If
End If
'取消訂單
If chk_CancelOrder.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " 簽單類別 = '取消訂單' "
   Else
      str_Where = str_Where & " and 簽單類別 = '未出訂單' "
   End If
End If
'簽收異常訂單
If chk_ExpectOrder.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " 簽單類別 = '異常訂單' "
   Else
      str_Where = str_Where & " and 簽單類別 = '異常訂單' "
   End If
End If
'未揀貨訂單chk_Ship_qty
If chk_Ship_qty.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " 揀貨量 = '0' "
   Else
      str_Where = str_Where & " and 揀貨量 = '0' "
   End If
End If
'轉入排車系統識別欄位：Orders.B_PHONE2 >> 00 已轉入
If chk_NotImport.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " 轉入識別 = '' "
   Else
      str_Where = str_Where & " and  轉入識別 = '' "
   End If
End If
'已轉入，尚未排車(尚未產生路線編號)
If chk_WaitPlan.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " (轉入識別 = 'V' and 路線編號 = '') "
   Else
      str_Where = str_Where & " and (轉入識別 = 'V' and 路線編號 = '') "
   End If
End If
'一次排車路線編號
txt_FRouteNo_Start.Text = Trim(txt_FRouteNo_Start.Text)
txt_FRouteNo_End.Text = Trim(txt_FRouteNo_End.Text)
strSubwhere = ""
If Len(txt_FRouteNo_Start.Text) > 0 And Len(txt_FRouteNo_End.Text) > 0 Then
   strSubwhere = " 路線編號 Between '" & txt_FRouteNo_Start.Text & "' and '" & txt_FRouteNo_End.Text & "' "
ElseIf Len(txt_FRouteNo_Start.Text) > 0 And Len(txt_FRouteNo_End.Text) = 0 Then
   strSubwhere = " 路線編號 = '" & txt_FRouteNo_Start.Text & "' "
ElseIf Len(txt_FRouteNo_Start.Text) = 0 And Len(txt_FRouteNo_End.Text) > 0 Then
   strSubwhere = " 路線編號 = '" & txt_FRouteNo_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'一次排車日期
txt_FPlanDate_Start.Text = Trim(txt_FPlanDate_Start.Text)
txt_FPlanDate_End.Text = Trim(txt_FPlanDate_End.Text)
strSubwhere = ""
If Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   strSubwhere = " 排車日期 Between '" & txt_FPlanDate_Start.Text & "' and '" & txt_FPlanDate_End.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) = 0 Then
   strSubwhere = " 排車日期 = '" & txt_FPlanDate_Start.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) = 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   strSubwhere = " 排車日期 = '" & txt_FPlanDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'一次出車日期
txt_FDeliveryDate_Start.Text = Trim(txt_FDeliveryDate_Start.Text)
txt_FDeliveryDate_End.Text = Trim(txt_FDeliveryDate_End.Text)
strSubwhere = ""
If Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   strSubwhere = " 出車日期 Between '" & txt_FDeliveryDate_Start.Text & "' and '" & txt_FDeliveryDate_End.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) = 0 Then
   strSubwhere = " 出車日期 = '" & txt_FDeliveryDate_Start.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) = 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   strSubwhere = " 出車日期 = '" & txt_FDeliveryDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'一次預計報到日期<daniel 20041005>
txt_FPlanCheckin_Start.Text = Trim(txt_FPlanCheckin_Start.Text)
txt_FPlanCheckin_End.Text = Trim(txt_FPlanCheckin_End.Text)
txt_FPlanCheckinTime_Start.Text = Trim(txt_FPlanCheckinTime_Start.Text)
txt_FPlanCheckinTime_End.Text = Trim(txt_FPlanCheckinTime_End.Text)
strSubwhere = ""
If Len(txt_FPlanCheckin_Start.Text) > 0 And Len(txt_FPlanCheckin_End.Text) > 0 And Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
   strSubwhere = " 預計報到日期 Between '" & txt_FPlanCheckin_Start.Text & "' and '" & txt_FPlanCheckin_End.Text & "' "
ElseIf Len(txt_FPlanCheckin_Start.Text) > 0 And Len(txt_FPlanCheckin_End.Text) = 0 And Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
   strSubwhere = " 預計報到日期 = '" & txt_FPlanCheckin_Start.Text & "' "
ElseIf Len(txt_FPlanCheckin_Start.Text) = 0 And Len(txt_FPlanCheckin_End.Text) > 0 And Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
   strSubwhere = " 預計報到日期 = '" & txt_FPlanCheckin_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'一次排車：預計報到時間<daniel 20041005>
If Len(Trim(txt_FPlanCheckinTime_Start.Text)) <> 0 Then
    If Len(txt_FPlanCheckinTime_Start.Text) <> 4 Then
        msg_text = "預計報到時間：資料格式 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_FPlanCheckinTime_Start.Text, 2)
        Case "00" To "23"
        Case Else
             msg_text = "預計報到時間：資料格式 hhss "
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             Screen.MousePointer = vbDefault
             txt_FPlanCheckinTime_Start.SetFocus
             Exit Sub
     End Select
     Select Case Right(txt_FPlanCheckinTime_Start.Text, 2)
        Case "00" To "59"
        Case Else
             msg_text = "預計報到時間：資料格式 hhss "
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_FPlanCheckinTime_Start.SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
     End Select
End If
If Len(Trim(txt_FPlanCheckinTime_End.Text)) <> 0 Then
    If Len(txt_FPlanCheckinTime_End.Text) <> 4 Then
        msg_text = "預計報到時間：資料格式 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_FPlanCheckinTime_End.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "預計報到時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_FPlanCheckinTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
    Select Case Right(txt_FPlanCheckinTime_End.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "預計報到時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_FPlanCheckinTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
End If
strSubwhere = ""
If Len(txt_FPlanCheckinTime_Start.Text) > 0 And Len(txt_FPlanCheckinTime_End.Text) > 0 Then
    If Len(txt_FPlanCheckin_Start.Text) = 0 And Len(txt_FPlanCheckin_End.Text) = 0 Then
        msg_text = "請輸入預計報到日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " 預計報到日期+預計報到時間 Between '" & txt_FPlanCheckin_Start.Text & txt_FPlanCheckinTime_Start.Text & "' and '" & txt_FPlanCheckin_End.Text & txt_FPlanCheckinTime_End.Text & "' "
ElseIf Len(txt_FPlanCheckinTime_Start.Text) > 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
    If Len(txt_FPlanCheckin_Start.Text) = 0 Then
        msg_text = "請輸入預計報到日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " 預計報到日期+預計報到時間 = '" & txt_FPlanCheckin_Start.Text & txt_FPlanCheckinTime_Start.Text & "' "
ElseIf Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) > 0 Then
    If Len(txt_FPlanCheckin_End.Text) = 0 Then
        msg_text = "請輸入預計報到日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " 預計報到日期+預計報到時間 = '" & txt_FPlanCheckin_End.Text & txt_FPlanCheckinTime_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If

'揀貨確認日期<daniel 20041005>
txt_SDNDate_Start.Text = Trim(txt_SDNDate_Start.Text)
txt_SDNDate_End.Text = Trim(txt_SDNDate_End.Text)
strSubwhere = ""
If Len(txt_SDNDate_Start.Text) > 0 And Len(txt_SDNDate_End.Text) > 0 And Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) = 0 Then
    strSubwhere = " left(揀貨確認時間,8) Between '" & txt_SDNDate_Start.Text & "' and '" & txt_SDNDate_End.Text & "' "
ElseIf Len(txt_SDNDate_Start.Text) > 0 And Len(txt_SDNDate_End.Text) = 0 And Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) = 0 Then
    strSubwhere = " left(揀貨確認時間,8) = '" & txt_SDNDate_Start.Text & "' "
ElseIf Len(txt_SDNDate_Start.Text) = 0 And Len(txt_SDNDate_End.Text) > 0 And Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) = 0 Then
    strSubwhere = " left(揀貨確認時間,8) = '" & txt_SDNDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'揀貨確認時間<daniel 20041005>
txt_SDNTime_Start.Text = Trim(txt_SDNTime_Start.Text)
txt_SDNTime_End.Text = Trim(txt_SDNTime_End.Text)
If Len(Trim(txt_SDNTime_Start.Text)) <> 0 Then
    If Len(txt_SDNTime_Start.Text) <> 4 Then
        msg_text = "揀貨確認時間：資料格式 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_SDNTime_Start.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "揀貨確認時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_Start.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
    Select Case Right(txt_SDNTime_Start.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "揀貨確認時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_Start.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
End If
If Len(Trim(txt_SDNTime_End.Text)) <> 0 Then
    If Len(txt_SDNTime_End.Text) <> 4 Then
        msg_text = "揀貨確認時間：資料格式 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_SDNTime_End.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "揀貨確認時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
    Select Case Right(txt_SDNTime_End.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "揀貨確認時間：資料格式 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
End If
strSubwhere = ""
If Len(txt_SDNTime_Start.Text) > 0 And Len(txt_SDNTime_End.Text) > 0 Then
    If Len(txt_SDNDate_Start.Text) = 0 Or Len(txt_SDNDate_End.Text) = 0 Then
        msg_text = "請輸入揀貨確認日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " 揀貨確認時間  Between '" & txt_SDNDate_Start.Text & txt_SDNTime_Start.Text & "' and '" & txt_SDNDate_End.Text & txt_SDNTime_End.Text & "' "
ElseIf Len(txt_SDNTime_Start.Text) > 0 And Len(txt_SDNTime_End.Text) = 0 Then
    If Len(txt_SDNDate_Start.Text) = 0 Then
        msg_text = "請輸入揀貨確認日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " 揀貨確認時間 = '" & txt_SDNDate_Start.Text & txt_SDNTime_Start.Text & "' "
ElseIf Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) > 0 Then
    If Len(txt_SDNDate_End.Text) = 0 Then
        msg_text = "請輸入揀貨確認日期"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " 揀貨確認時間 = '" & txt_SDNDate_End.Text & txt_SDNTime_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If


'一次排車：車輛報到日期
txt_FCheckin_Start.Text = Trim(txt_FCheckin_Start.Text)
txt_FCheckin_End.Text = Trim(txt_FCheckin_End.Text)
strSubwhere = ""
If Len(txt_FCheckin_Start.Text) > 0 And Len(txt_FCheckin_End.Text) > 0 Then
   strSubwhere = " 報到日期 Between '" & txt_FCheckin_Start.Text & "' and '" & txt_FCheckin_End.Text & "' "
ElseIf Len(txt_FCheckin_Start.Text) > 0 And Len(txt_FCheckin_End.Text) = 0 Then
   strSubwhere = " 報到日期 = '" & txt_FCheckin_Start.Text & "' "
ElseIf Len(txt_FCheckin_Start.Text) = 0 And Len(txt_FCheckin_End.Text) > 0 Then
   strSubwhere = " 報到日期 = '" & txt_FCheckin_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If

'一次排車：車輛離倉日期
txt_FCheckout_Start.Text = Trim(txt_FCheckout_Start.Text)
txt_FCheckout_End.Text = Trim(txt_FCheckout_End.Text)
strSubwhere = ""
If Len(txt_FCheckout_Start.Text) > 0 And Len(txt_FCheckout_End.Text) > 0 Then
   strSubwhere = " 離倉日期 Between '" & txt_FCheckout_Start.Text & "' and '" & txt_FCheckout_End.Text & "' "
ElseIf Len(txt_FCheckout_Start.Text) > 0 And Len(txt_FCheckout_End.Text) = 0 Then
   strSubwhere = " 離倉日期 = '" & txt_FCheckout_Start.Text & "' "
ElseIf Len(txt_FCheckout_Start.Text) = 0 And Len(txt_FCheckout_End.Text) > 0 Then
   strSubwhere = " 離倉日期 = '" & txt_FCheckout_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'一次排車：exe回傳狀態
If cmb_EXEConfirm.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " 回傳狀態 = '" & cmb_EXEConfirm.List(cmb_EXEConfirm.ListIndex) & "' "
   Else
      str_Where = str_Where & " 回傳狀態 = '" & cmb_EXEConfirm.List(cmb_EXEConfirm.ListIndex) & "' "
   End If
End If
'一次排車：排車人員
txt_FAddWho.Text = Trim(txt_FAddWho.Text)
strSubwhere = ""
If Len(txt_FAddWho.Text) > 0 Then
   If InStr(txt_FAddWho.Text, ",") > 0 Then
      tmp_data = Split(txt_FAddWho.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 排車者 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 排車者 in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " 排車者 like '%" & txt_FAddWho.Text & "%' "
      Else
         str_Where = str_Where & " and 排車者 like '%" & txt_FAddWho.Text & "%' "
      End If
   End If
End If
'一次排車：車牌號碼
txt_FVehicleID.Text = Trim(txt_FVehicleID.Text)
strSubwhere = ""
If Len(txt_FVehicleID.Text) > 0 Then
   If InStr(txt_FVehicleID.Text, ",") > 0 Then
      tmp_data = Split(txt_FVehicleID.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 車牌號碼 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 車牌號碼 in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " 車牌號碼 like '%" & txt_FVehicleID.Text & "' "
      Else
         str_Where = str_Where & " and 車牌號碼 like '%" & txt_FVehicleID.Text & "' "
      End If
   End If
End If
'一次排車：駕駛人
txt_FDriver.Text = Trim(txt_FDriver.Text)
strSubwhere = ""
If Len(txt_FDriver.Text) > 0 Then
   If InStr(txt_FDriver.Text, ",") > 0 Then
      tmp_data = Split(txt_FDriver.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 駕駛人 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 駕駛人 in (" & strSubwhere & ") "
      End If
   Else   '沒輸入逗點間隔，用 Like 進行查詢
      If Len(str_Where) = 0 Then
         str_Where = " 駕駛人 like '%" & txt_FDriver.Text & "%' "
      Else
         str_Where = str_Where & " and 駕駛人 like '%" & txt_FDriver.Text & "%' "
      End If
   End If
End If
'一次排車：碼頭暫存
txt_FDockNo.Text = Trim(txt_FDockNo.Text)
strSubwhere = ""
If Len(txt_FDockNo.Text) > 0 Then
   If InStr(txt_FDockNo.Text, ",") > 0 Then
      tmp_data = Split(txt_FDockNo.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 碼頭暫存 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 碼頭暫存 in (" & strSubwhere & ") "
      End If
   Else   '沒輸入逗點間隔，用 Like 進行查詢
      If Len(str_Where) = 0 Then
         str_Where = " 碼頭暫存 like '%" & txt_FDockNo.Text & "%' "
      Else
         str_Where = str_Where & " and 碼頭暫存 like '%" & txt_FDockNo.Text & "%' "
      End If
   End If
End If
'一次排車：TMS單號
txt_FReceiptNo.Text = Trim(txt_FReceiptNo.Text)
strSubwhere = ""
If Len(txt_FReceiptNo.Text) > 0 Then
   If InStr(txt_FReceiptNo.Text, ",") > 0 Then
      tmp_data = Split(txt_FReceiptNo.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " TMS單號 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and TMS單號 in (" & strSubwhere & ") "
      End If
   Else   '沒輸入逗點間隔，用 Like 進行查詢
      If Len(str_Where) = 0 Then
         str_Where = " TMS單號 like '%" & txt_FReceiptNo.Text & "%' "
      Else
         str_Where = str_Where & " and TMS單號 like '%" & txt_FReceiptNo.Text & "%' "
      End If
   End If
End If
'篩選進行二次排車
If chk_SecondPlan.Value = vbChecked Then
      If Len(str_Where) = 0 Then
         str_Where = " 二次路線編號 <> '' "
      Else
         str_Where = str_Where & " and 二次路線編號 <> '' "
      End If
End If
'二次排車：路線編號
txt_SRouteNo_Start.Text = Trim(txt_SRouteNo_Start.Text)
txt_SRouteNo_End.Text = Trim(txt_SRouteNo_End.Text)
strSubwhere = ""
If Len(txt_SRouteNo_Start.Text) > 0 And Len(txt_SRouteNo_End.Text) > 0 Then
   strSubwhere = "  二次路線編號 Between '" & txt_SRouteNo_Start.Text & "' and '" & txt_SRouteNo_End.Text & "' "
ElseIf Len(txt_SRouteNo_Start.Text) > 0 And Len(txt_SRouteNo_End.Text) = 0 Then
   strSubwhere = "  二次路線編號 = '" & txt_SRouteNo_Start.Text & "' "
ElseIf Len(txt_SRouteNo_Start.Text) = 0 And Len(txt_SRouteNo_End.Text) > 0 Then
   strSubwhere = "  二次路線編號 = '" & txt_SRouteNo_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If

'二次排車：出車日期
txt_SDeliveryDate_Start.Text = Trim(txt_SDeliveryDate_Start.Text)
txt_SDeliveryDate_End.Text = Trim(txt_SDeliveryDate_End.Text)
strSubwhere = ""
If Len(txt_SDeliveryDate_Start.Text) > 0 And Len(txt_SDeliveryDate_End.Text) > 0 Then
   strSubwhere = "  二次出車日期 Between '" & txt_SDeliveryDate_Start.Text & "' and '" & txt_SDeliveryDate_End.Text & "' "
ElseIf Len(txt_SDeliveryDate_Start.Text) > 0 And Len(txt_SDeliveryDate_End.Text) = 0 Then
   strSubwhere = "  二次出車日期 = '" & txt_SDeliveryDate_Start.Text & "' "
ElseIf Len(txt_SDeliveryDate_Start.Text) = 0 And Len(txt_SDeliveryDate_End.Text) > 0 Then
   strSubwhere = "  二次出車日期 = '" & txt_SDeliveryDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'二次排車：車牌號碼
txt_SVehicleID.Text = Trim(txt_SVehicleID.Text)
strSubwhere = ""
If Len(txt_SVehicleID.Text) > 0 Then
   If InStr(txt_SVehicleID.Text, ",") > 0 Then
      tmp_data = Split(txt_SVehicleID.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 二次車牌號碼 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 二次車牌號碼 in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " 二次車牌號碼 like '%" & txt_SVehicleID.Text & "' "
      Else
         str_Where = str_Where & " and 二次車牌號碼 like '%" & txt_SVehicleID.Text & "' "
      End If
   End If
End If
'二次排車：排車人員
txt_SAddWho.Text = Trim(txt_SAddWho.Text)
strSubwhere = ""
If Len(txt_SAddWho.Text) > 0 Then
   If InStr(txt_SAddWho.Text, ",") > 0 Then
      tmp_data = Split(txt_SAddWho.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " 二次排車者 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and 二次排車者 in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " 二次排車者 like '%" & txt_SAddWho.Text & "%' "
      Else
         str_Where = str_Where & " and 二次排車者 like '%" & txt_SAddWho.Text & "%' "
      End If
   End If
End If
'郵遞區號
If cmb_ZIP.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " 郵遞區號 = '" & arZip(cmb_ZIP.ListIndex) & "' "
   Else
      str_Where = str_Where & " and 郵遞區號 = '" & arZip(cmb_ZIP.ListIndex) & "' "
   End If
End If
'運送區域
If cmb_AreaCode.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " Area = '" & arAreaCode(cmb_AreaCode.ListIndex) & "' "
   Else
      str_Where = str_Where & " and Area = '" & arAreaCode(cmb_AreaCode.ListIndex) & "' "
   End If
End If
'特殊需求
If cmb_ExtraDemand.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " (特殊需求碼1 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "' OR 特殊需求碼2 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "') "
   Else
      str_Where = str_Where & " and (特殊需求碼1 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "' OR 特殊需求碼2 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "') "
   End If
End If
'一次排車：運送車種
If cmb_VehicleType.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " 車種代碼 = '" & arVehicleType(cmb_VehicleType.ListIndex) & "' "
   Else
      str_Where = str_Where & " and 車種代碼 = '" & arVehicleType(cmb_VehicleType.ListIndex) & "' "
   End If
End If
'一次排車：貨運公司
If cmb_TRPCompany.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " 貨運公司代碼 = '" & arTRPCompany(cmb_TRPCompany.ListIndex) & "' "
   Else
      str_Where = str_Where & " and 貨運公司代碼 = '" & arTRPCompany(cmb_TRPCompany.ListIndex) & "' "
   End If
End If
'二次排車：運送車種
If cmb_SVehicleType.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " 二次車種代碼 = '" & arVehicleType(cmb_SVehicleType.ListIndex) & "' "
   Else
      str_Where = str_Where & " and 二次車種代碼 = '" & arVehicleType(cmb_SVehicleType.ListIndex) & "' "
   End If
End If
'二次排車：貨運公司
If cmb_STRPCompany.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " 二次貨運公司代碼 = '" & arTRPCompany(cmb_STRPCompany.ListIndex) & "' "
   Else
      str_Where = str_Where & " and 二次貨運公司代碼 = '" & arTRPCompany(cmb_STRPCompany.ListIndex) & "' "
   End If
End If

If Len(str_Where) = 0 Then
   Call Unload_RunLogForm
   Screen.MousePointer = vbDefault
   msg_text = "注意：資料量太大，請輸入查詢條件以減少資料量"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
str_SQL = strSelect & " Where " & str_Where
If strOrderBy <> "" Then
   str_SQL = str_SQL & " Order by " & strOrderBy
End If

SSTab1.Tab = 2
DoEvents: DoEvents

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   Call Unload_RunLogForm
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合設定條件之訂單資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab2_srcTotal_Case.Text = "0"
   txt_Tab2_srcTotal_PickCase.Text = "0"
   txt_Tab2_srcTotal_SDNCase.Text = "0"
   txt_Tab2_srcTotal_DifPick = "0"
   txt_Tab2_srcTotal_DifSDN = "0"
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Result)
tmp_Rs.Close

With dg_Result
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 2                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 300                '設定DataGrid 控制項中所有資料列的高
End With
rs_Result.MoveFirst
Set dg_Result.DataSource = rs_Result
With dg_Result
    .RowHeight = 250
    For iLoop = 0 To .Columns.Count - 1
'        .Columns(iLoop).Width = GetFieldWidth(rs_Result.Fields(iLoop).Name)
        .Columns(iLoop).Alignment = GetFieldAlignment(rs_Result.Fields(iLoop).Name)
    Next iLoop
End With

'設定欄寬
SetDataGridColWidth "訂單資料查詢結果", dg_Result

'統計<daniel 20041005>
str_SQL = "select sum(排車量),sum(揀貨量),sum(簽收量) from Query_OrdersData Where " & str_Where
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   txt_Tab2_srcTotal_Case.Text = "0"
   txt_Tab2_srcTotal_PickCase.Text = "0"
   txt_Tab2_srcTotal_SDNCase.Text = "0"
   txt_Tab2_srcTotal_DifPick = "0"
   txt_Tab2_srcTotal_DifSDN = "0"
   tmp_Rs.Close
   Exit Sub
End If
txt_Tab2_srcTotal_Case.Text = tmp_Rs.Fields(0)
txt_Tab2_srcTotal_PickCase.Text = tmp_Rs.Fields(1)
txt_Tab2_srcTotal_SDNCase.Text = tmp_Rs.Fields(2)
txt_Tab2_srcTotal_DifPick = Val(tmp_Rs.Fields(0)) - Val(tmp_Rs.Fields(1))
txt_Tab2_srcTotal_DifSDN = tmp_Rs.Fields(0) - Val(tmp_Rs.Fields(2))
tmp_Rs.Close

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單查詢", Me.Caption, "cmd_Query", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
      
err:
End Sub

Private Sub cmd_Tab0_Reset_Click()
'清除查詢條件
Set dg_Result.DataSource = Nothing
Set rs_Result = Nothing

Call ClearForm_AllField(Me)

End Sub

Private Sub cmd_Tab1_OrderAdd_Click()
'欄位選取 >> 排序方式-加入
If lst_SelectedFields.SelCount > 0 Then
   lst_OrderBy.AddItem lst_SelectedFields.List(lst_SelectedFields.ListIndex)
End If

End Sub

Private Sub cmd_Tab1_OrderByDown_Click()
'欄位選取 >> 排序欄位順序下移
If lst_OrderBy.ListCount = 0 Then Exit Sub
If lst_OrderBy.SelCount > 0 Then
   Dim afItem As String, selItem As String
   Dim afIndex As Double, selIndex As Double
   selItem = lst_OrderBy.List(lst_OrderBy.ListIndex)
   selIndex = lst_OrderBy.ListIndex
   If (lst_OrderBy.ListIndex + 1) > (lst_OrderBy.ListCount - 1) Then
      afItem = lst_OrderBy.List(0)
      afIndex = 0
   Else
      afItem = lst_OrderBy.List(lst_OrderBy.ListIndex + 1)
      afIndex = (lst_OrderBy.ListIndex + 1)
   End If
   lst_OrderBy.List(afIndex) = selItem
   lst_OrderBy.List(selIndex) = afItem
   lst_OrderBy.Selected(afIndex) = True
End If

End Sub

Private Sub cmd_Tab1_OrderByUp_Click()
'欄位選取 >> 排序欄位順序上移
If lst_OrderBy.ListCount = 0 Then Exit Sub
If lst_OrderBy.SelCount > 0 Then
   Dim preItem As String, selItem As String
   Dim preIndex As Double, selIndex As Double
   selItem = lst_OrderBy.List(lst_OrderBy.ListIndex)
   selIndex = lst_OrderBy.ListIndex
   If (lst_OrderBy.ListIndex - 1) < 0 Then
      preItem = lst_OrderBy.List(lst_OrderBy.ListCount - 1)
      preIndex = (lst_OrderBy.ListCount - 1)
   Else
      preItem = lst_OrderBy.List(lst_OrderBy.ListIndex - 1)
      preIndex = (lst_OrderBy.ListIndex - 1)
   End If
   lst_OrderBy.List(preIndex) = selItem
   lst_OrderBy.List(selIndex) = preItem
   lst_OrderBy.Selected(preIndex) = True
End If

End Sub

Private Sub cmd_Tab1_OrderRemove_Click()
'欄位選取 >> DoubleClick 移除選取排序方式欄位
If lst_OrderBy.SelCount > 0 Then
   lst_OrderBy.RemoveItem lst_OrderBy.ListIndex
End If
End Sub

Private Sub cmd_Tab0_SelectField_Click()
'查詢條件 >> 欄位選取
SSTab1.Tab = 1
End Sub

Private Sub cmd_Tab1_Add_Click()
'欄位選取 >> 加入
If lst_AllFields.SelCount > 0 Then
   lst_SelectedFields.AddItem lst_AllFields.List(lst_AllFields.ListIndex)
   lst_AllFields.RemoveItem lst_AllFields.ListIndex
End If
End Sub

Private Sub cmd_Tab1_Down_Click()
'欄位選取 >> 欄位順序下移
If lst_SelectedFields.ListCount = 0 Then Exit Sub
If lst_SelectedFields.SelCount > 0 Then
   Dim afItem As String, selItem As String
   Dim afIndex As Double, selIndex As Double
   selItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   selIndex = lst_SelectedFields.ListIndex
   If (lst_SelectedFields.ListIndex + 1) > (lst_SelectedFields.ListCount - 1) Then
      afItem = lst_SelectedFields.List(0)
      afIndex = 0
   Else
      afItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex + 1)
      afIndex = (lst_SelectedFields.ListIndex + 1)
   End If
   lst_SelectedFields.List(afIndex) = selItem
   lst_SelectedFields.List(selIndex) = afItem
   lst_SelectedFields.Selected(afIndex) = True
End If

End Sub

Private Sub cmd_Tab1_Reload_Click()
'欄位選取>>重新載入
'取得所有可用的查詢欄位
Call GetAllFields
End Sub

Private Sub cmd_Tab1_Remove_Click()
'欄位選取 >> 移除
If lst_SelectedFields.SelCount > 0 Then
   lst_AllFields.AddItem lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   lst_SelectedFields.RemoveItem lst_SelectedFields.ListIndex

End If

End Sub

Private Sub cmd_Tab1_Reset_Click()
'欄位選取 >> 清除
On Error GoTo err_Handle
Tran_Level = 0
Tran_Level = cn.BeginTrans
str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYFIELDS' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYORDER' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans
Tran_Level = 0
'取回所有欄位
Call GetAllFields
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單查詢-設定值清除", Me.Caption, "cmd_Tab1_Reset", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Up_Click()
'欄位選取 >> 欄位順序上移
If lst_SelectedFields.ListCount = 0 Then Exit Sub
If lst_SelectedFields.SelCount > 0 Then
   Dim preItem As String, selItem As String
   Dim preIndex As Double, selIndex As Double
   selItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   selIndex = lst_SelectedFields.ListIndex
   If (lst_SelectedFields.ListIndex - 1) < 0 Then
      preItem = lst_SelectedFields.List(lst_SelectedFields.ListCount - 1)
      preIndex = (lst_SelectedFields.ListCount - 1)
   Else
      preItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex - 1)
      preIndex = (lst_SelectedFields.ListIndex - 1)
   End If
   lst_SelectedFields.List(preIndex) = selItem
   lst_SelectedFields.List(selIndex) = preItem
   lst_SelectedFields.Selected(preIndex) = True
End If
End Sub

Private Sub dg_Result_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

Dim objDataGrid As Object: Set objDataGrid = dg_Result
If Len(objDataGrid.Columns(ColIndex).DataField) = 0 Or objDataGrid.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, "訂單資料查詢結果" & objDataGrid.Name, objDataGrid.Columns(ColIndex).DataField, objDataGrid.Columns(ColIndex).Width

End Sub

Private Sub lst_AllFields_DblClick()
'欄位選取 >> DoubleClick 加入選取
If lst_AllFields.SelCount > 0 Then
   lst_SelectedFields.AddItem lst_AllFields.List(lst_AllFields.ListIndex)
   lst_AllFields.RemoveItem lst_AllFields.ListIndex
End If
End Sub

Private Sub lst_OrderBy_DblClick()
'欄位選取 >> DoubleClick 移除選取排序方式欄位
If lst_OrderBy.SelCount > 0 Then
   lst_OrderBy.RemoveItem lst_OrderBy.ListIndex
End If
End Sub

Private Sub lst_SelectedFields_DblClick()
'欄位選取 >> DoubleClick 移除選取
If lst_SelectedFields.SelCount > 0 Then
   lst_AllFields.AddItem lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   lst_SelectedFields.RemoveItem lst_SelectedFields.ListIndex
End If
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "訂單查詢作業"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
Me.Height = 7650: Me.Width = 11595
SSTab1.Tab = 0
If SSTab1.Tab = 0 Then cmd_Tab2SavetoExcel.Visible = False: cmd_Exit(0).Visible = False
'取得所有可用的查詢欄位
Call GetAllFields

'查詢條件待選清單建立
Dim dbZip As Double, dbAreaCode As Double, dbExtraDemand As Double, dbVehicleType As Double, dbTRPCompany As Double
Dim dbRSC As Double, dbRBC As Double
cmb_ZIP.Clear: cmb_AreaCode.Clear: cmb_ExtraDemand.Clear
cmb_VehicleType.Clear: cmb_TRPCompany.Clear
cmb_RSC.Clear: cmb_RBC.Clear
str_SQL = "Select 區分,代碼,說明 From Query_OrdersBaseData Order by 區分,代碼"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arZip(1) As String
ReDim arAreaCode(1) As String
ReDim arExtraDemand(1) As String
ReDim arVehicleType(1) As String
ReDim arTRPCompany(1) As String
ReDim arRSC(1) As String
ReDim arRBC(1) As String
If Not tmp_Rs.EOF Then
   dbZip = 0: dbAreaCode = 0: dbExtraDemand = 0: dbVehicleType = 0: dbTRPCompany = 0
   Do While Not tmp_Rs.EOF
      Select Case tmp_Rs.Fields("區分").Value
         Case "郵遞區號"
              arZip(dbZip) = tmp_Rs.Fields("代碼").Value
              cmb_ZIP.AddItem tmp_Rs.Fields("代碼").Value & Space(6 - Len(Trim(tmp_Rs.Fields("代碼").Value))) & tmp_Rs.Fields("說明").Value
              dbZip = dbZip + 1
              If dbZip = UBound(arZip) Then
                 ReDim Preserve arZip(UBound(arZip) + 2) As String
              End If
         Case "運送區域"
              arAreaCode(dbAreaCode) = tmp_Rs.Fields("代碼").Value
              cmb_AreaCode.AddItem tmp_Rs.Fields("代碼").Value & Space(6 - Len(Trim(tmp_Rs.Fields("代碼").Value))) & tmp_Rs.Fields("說明").Value
              dbAreaCode = dbAreaCode + 1
              If dbAreaCode = UBound(arAreaCode) Then
                 ReDim Preserve arAreaCode(UBound(arAreaCode) + 2) As String
              End If
         Case "特殊需求"
              arExtraDemand(dbExtraDemand) = tmp_Rs.Fields("代碼").Value
              cmb_ExtraDemand.AddItem tmp_Rs.Fields("說明").Value
              dbExtraDemand = dbExtraDemand + 1
              If dbExtraDemand = UBound(arExtraDemand) Then
                 ReDim Preserve arExtraDemand(UBound(arExtraDemand) + 2) As String
              End If
         Case "車種"
              arVehicleType(dbVehicleType) = tmp_Rs.Fields("代碼").Value
              cmb_VehicleType.AddItem tmp_Rs.Fields("代碼").Value & Space(6 - Len(Trim(tmp_Rs.Fields("代碼").Value))) & tmp_Rs.Fields("說明").Value
              cmb_SVehicleType.AddItem tmp_Rs.Fields("代碼").Value & Space(6 - Len(Trim(tmp_Rs.Fields("代碼").Value))) & tmp_Rs.Fields("說明").Value
              dbVehicleType = dbVehicleType + 1
              If dbVehicleType = UBound(arVehicleType) Then
                 ReDim Preserve arVehicleType(UBound(arVehicleType) + 2) As String
              End If
         Case "貨運公司"
              arTRPCompany(dbTRPCompany) = tmp_Rs.Fields("代碼").Value
              cmb_TRPCompany.AddItem tmp_Rs.Fields("說明").Value
              cmb_STRPCompany.AddItem tmp_Rs.Fields("說明").Value
              dbTRPCompany = dbTRPCompany + 1
              If dbTRPCompany = UBound(arTRPCompany) Then
                 ReDim Preserve arTRPCompany(UBound(arTRPCompany) + 2) As String
              End If
         Case "異常原因"
              arRSC(dbRSC) = tmp_Rs.Fields("代碼").Value
              cmb_RSC.AddItem tmp_Rs.Fields("代碼").Value & "  " & tmp_Rs.Fields("說明").Value
              dbRSC = dbRSC + 1
              If dbRSC = UBound(arRSC) Then
                 ReDim Preserve arRSC(UBound(arRSC) + 2) As String
              End If
         Case "責任歸屬"
              arRBC(dbRBC) = tmp_Rs.Fields("代碼").Value
              cmb_RBC.AddItem tmp_Rs.Fields("代碼").Value & "  " & tmp_Rs.Fields("說明").Value
              dbRBC = dbRBC + 1
              If dbRBC = UBound(arRBC) Then
                 ReDim Preserve arRBC(UBound(arRBC) + 2) As String
              End If
              
      End Select
      tmp_Rs.MoveNext
   Loop
End If
tmp_Rs.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'攔截整個表單鍵盤按鍵事件
'用途：使用者按下 Esc 則不傳回任何資料，且關閉日期選取視窗
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
End If
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    
    Frame2.Top = Me.Top + (Me.ScaleHeight - Frame2.Height) / 2 + 240
    Frame3.Top = Me.Top + (Me.ScaleHeight - Frame3.Height) / 2 + 240
    Frame2.Left = Me.Left + (Me.ScaleWidth - Frame2.Width) / 2 + 240
    Frame3.Left = Me.Left + (Me.ScaleWidth - Frame3.Width) / 2 + 240
    
    If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
        SSTab1.Height = Me.ScaleHeight
        dg_Result.Height = Me.ScaleHeight - cmd_Tab2SavetoExcel.Width - 360

    End If
    
    If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
        SSTab1.Width = Me.ScaleWidth
        dg_Result.Width = Me.ScaleWidth - cmd_Tab2SavetoExcel.Width - 360
        cmd_Tab2SavetoExcel.Left = dg_Result.Width + 240
        cmd_Tab2SavetoExcel.Top = dg_Result.Top
        cmd_Exit(0).Left = dg_Result.Width + 240
    End If

End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_Query_Orders = Nothing
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub

Private Sub GetAllFields()
'取得所有可使用欄位
On Error GoTo err_Handle
lst_AllFields.Clear
lst_SelectedFields.Clear
lst_OrderBy.Clear

'使用者選取之欄位暫存
Dim rs_UserSelectedFields As ADODB.Recordset
Call ReDim_Recordset(rs_UserSelectedFields)
With rs_UserSelectedFields
     .Fields.Append "編號", adDouble
     .Fields.Append "欄位名稱", adVarChar, 40
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
'查詢結果之欄位排序暫存
Dim rs_OrderByFields As ADODB.Recordset
Call ReDim_Recordset(rs_OrderByFields)
With rs_OrderByFields
     .Fields.Append "編號", adDouble
     .Fields.Append "欄位名稱", adVarChar, 40
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
'取回訂單查詢所有欄位
str_SQL = "Select ac.FieldName as 欄位名稱 , Isnull(Cast( cd1.Description as integer),0) as SeqNo,Isnull(Cast( cd2.Description as integer),0) as OrderNo  " & _
          "From Query_UserSelectedField ac " & _
          "Left outer join CodeLKUP cd1 on cd1.ListName = 'ORDERSQUERYFIELDS' and cd1.Code = ac.ViewName and cd1.Long = ac.FieldName and cd1.Short = '" & User_id & "' " & _
          "Left outer join CodeLKUP cd2 on cd2.ListName = 'ORDERSQUERYORDER' and cd2.Code = ac.ViewName and cd2.Long = ac.FieldName and cd2.Short = '" & User_id & "' " & _
          "Where ac.ViewName = 'Query_OrdersData' Order by ac.ColIndex"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
Do While Not tmp_Rs.EOF
   If tmp_Rs.Fields("SeqNo").Value = 0 Then
      lst_AllFields.AddItem tmp_Rs.Fields("欄位名稱").Value
   Else
      rs_UserSelectedFields.AddNew
      rs_UserSelectedFields.Fields("編號").Value = tmp_Rs.Fields("SeqNo").Value
      rs_UserSelectedFields.Fields("欄位名稱").Value = tmp_Rs.Fields("欄位名稱").Value
      rs_UserSelectedFields.Update
   End If
   If tmp_Rs.Fields("OrderNo").Value <> 0 Then
      rs_OrderByFields.AddNew
      rs_OrderByFields.Fields("編號").Value = tmp_Rs.Fields("OrderNo").Value
      rs_OrderByFields.Fields("欄位名稱").Value = tmp_Rs.Fields("欄位名稱").Value
      rs_OrderByFields.Update
   End If
   tmp_Rs.MoveNext
Loop
Set tmp_Rs = Nothing

'查詢結果欄位
If rs_UserSelectedFields.EOF Then
   Set rs_UserSelectedFields = Nothing
   Exit Sub
Else
   rs_UserSelectedFields.Sort = " 編號 "
   rs_UserSelectedFields.MoveFirst
   Do While Not rs_UserSelectedFields.EOF
      lst_SelectedFields.AddItem rs_UserSelectedFields.Fields("欄位名稱").Value
      rs_UserSelectedFields.MoveNext
   Loop
   Set rs_UserSelectedFields = Nothing
End If

'排序依據
If rs_OrderByFields.EOF Then
   Set rs_OrderByFields = Nothing
   Exit Sub
Else
   rs_OrderByFields.Sort = " 編號 "
   rs_OrderByFields.MoveFirst
   Do While Not rs_OrderByFields.EOF
      lst_OrderBy.AddItem rs_OrderByFields.Fields("欄位名稱").Value
      rs_OrderByFields.MoveNext
   Loop
   Set rs_OrderByFields = Nothing
End If

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單查詢-載入欄位", Me.Caption, "From 內部 Subprogram GetAllFields", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub SaveQueryEnv()
'儲存使用者查詢設定值
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

On Error GoTo err_Handle
Tran_Level = 0
Tran_Level = cn.BeginTrans
'查詢結果欄位選取值存檔
If lst_SelectedFields.ListCount <> 0 Then
   str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYFIELDS' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   For iLoop = 0 To lst_SelectedFields.ListCount - 1
       str_SQL = "Insert into Codelkup (ListName,Code,Long,Short,Description,AddWho,EditWho) Values ('ORDERSQUERYFIELDS','Query_OrdersData','" & _
                 lst_SelectedFields.List(iLoop) & "','" & User_id & "'," & iLoop + 1 & ",'" & User_id & "','" & User_id & "')"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   Next iLoop
End If
'查詢結果排序依據存檔
If lst_OrderBy.ListCount <> 0 Then
   str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYORDER' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   For iLoop = 0 To lst_OrderBy.ListCount - 1
       str_SQL = "Insert into Codelkup (ListName,Code,Long,Short,Description,AddWho,EditWho) Values ('ORDERSQUERYORDER','Query_OrdersData','" & _
                 lst_OrderBy.List(iLoop) & "','" & User_id & "'," & iLoop + 1 & ",'" & User_id & "','" & User_id & "')"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   Next iLoop
End If
cn.CommitTrans
Tran_Level = 0
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單查詢-設定值存檔", Me.Caption, "From 內部 Subprogram [SaveQueryEnv]", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'日期選取
Select Case mvDate.Tag
       Case "訂單日期.起"
            txt_OrderDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "訂單日期.迄"
            txt_OrderDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "送貨日期.起"
            txt_DeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "送貨日期.迄"
            txt_DeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.排車日期.起"
            txt_FPlanDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.排車日期.迄"
            txt_FPlanDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.出車日期.起"
            txt_FDeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.出車日期.迄"
            txt_FDeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.預計報到日期.起"
            txt_FPlanCheckin_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.預計報到日期.迄"
            txt_FPlanCheckin_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.車輛報到日期.起"
            txt_FCheckin_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.車輛報到日期.迄"
            txt_FCheckin_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.車輛離倉日期.起"
            txt_FCheckout_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "一次排車.車輛離倉日期.迄"
            txt_FCheckout_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "二次排車.車輛出車日期.起"
            txt_SDeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "二次排車.車輛出車日期.迄"
            txt_SDeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "揀貨確認日期.起"
            txt_SDNDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "揀貨確認日期.迄"
            txt_SDNDate_End.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then cmd_Tab2SavetoExcel.Visible = False: cmd_Exit(0).Visible = False: Frame3.Visible = True
If SSTab1.Tab = 1 Then cmd_Tab2SavetoExcel.Visible = False: cmd_Exit(0).Visible = False: Frame2.Visible = True: Frame3.Visible = False
If SSTab1.Tab = 2 Then cmd_Tab2SavetoExcel.Visible = True: cmd_Exit(0).Visible = True: Frame3.Visible = False: Frame2.Visible = False
End Sub

Private Sub txt_SDNDate_End_Click()
'揀貨確認日期：迄
If Trim(txt_SDNDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDNDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDNDate_End.Text, 4) & "/" & Mid(txt_SDNDate_End.Text, 5, 2) & "/" & Right(txt_SDNDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_SDNDate_End.Left
mvDate.Top = Frame12.Top + txt_SDNDate_End.Top + txt_SDNDate_End.Height
mvDate.Tag = "揀貨確認日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_SDNDate_Start_Click()
'揀貨確認日期：起
If Trim(txt_SDNDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDNDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDNDate_Start.Text, 4) & "/" & Mid(txt_SDNDate_Start.Text, 5, 2) & "/" & Right(txt_SDNDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_SDNDate_Start.Left
mvDate.Top = Frame12.Top + txt_SDNDate_Start.Top + txt_SDNDate_Start.Height
mvDate.Tag = "揀貨確認日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_StorerKey_KeyPress(KeyAscii As Integer)
'貨主
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_Extern_Start.SelStart = 0: txt_Extern_Start.SelLength = Len(txt_Extern_Start.Text)
         txt_Extern_Start.SetFocus
End Select
End Sub

Private Sub txt_Extern_Start_KeyPress(KeyAscii As Integer)
'貨主單號：起
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_Extern_End.SelStart = 0: txt_Extern_End.SelLength = Len(txt_Extern_End.Text)
         txt_Extern_End.SetFocus
End Select
End Sub

Private Sub txt_Extern_End_KeyPress(KeyAscii As Integer)
'貨主單號：迄
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_OrderDate_Start.SelStart = 0: txt_OrderDate_Start.SelLength = Len(txt_OrderDate_Start.Text)
         txt_OrderDate_Start.SetFocus
End Select
End Sub

Private Sub txt_OrderDate_Start_Click()
'訂單日期：起
If Trim(txt_OrderDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_OrderDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_OrderDate_Start.Text, 4) & "/" & Mid(txt_OrderDate_Start.Text, 5, 2) & "/" & Right(txt_OrderDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_OrderDate_Start.Left
mvDate.Top = Frame11.Top + txt_OrderDate_Start.Top + txt_OrderDate_Start.Height
mvDate.Tag = "訂單日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_OrderDate_Start_KeyPress(KeyAscii As Integer)
'訂單日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_OrderDate_End.SelStart = 0: txt_OrderDate_End.SelLength = Len(txt_OrderDate_End.Text)
         txt_OrderDate_End.SetFocus
End Select
End Sub

Private Sub txt_OrderDate_End_Click()
'訂單日期：迄
If Trim(txt_OrderDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_OrderDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_OrderDate_End.Text, 4) & "/" & Mid(txt_OrderDate_End.Text, 5, 2) & "/" & Right(txt_OrderDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_OrderDate_End.Left
mvDate.Top = Frame11.Top + txt_OrderDate_End.Top + txt_OrderDate_End.Height
mvDate.Tag = "訂單日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_OrderDate_End_KeyPress(KeyAscii As Integer)
'訂單日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_DeliveryDate_Start.SelStart = 0: txt_DeliveryDate_Start.SelLength = Len(txt_DeliveryDate_Start.Text)
         txt_DeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_DeliveryDate_Start_Click()
'送貨日期：起
If Trim(txt_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_DeliveryDate_Start.Left
mvDate.Top = Frame11.Top + txt_DeliveryDate_Start.Top + txt_DeliveryDate_Start.Height
mvDate.Tag = "送貨日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'送貨日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_DeliveryDate_End.SelStart = 0: txt_DeliveryDate_End.SelLength = Len(txt_DeliveryDate_End.Text)
         txt_DeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_DeliveryDate_End_Click()
'送貨日期：迄
If Trim(txt_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate_End.Text, 4) & "/" & Mid(txt_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_DeliveryDate_End.Left
mvDate.Top = Frame11.Top + txt_DeliveryDate_End.Top + txt_DeliveryDate_End.Height
mvDate.Tag = "送貨日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'送貨日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_SKU.SelStart = 0: txt_SKU.SelLength = Len(txt_SKU.Text): txt_SKU.SetFocus
End Select
End Sub

Private Sub txt_SKU_KeyPress(KeyAscii As Integer)
'貨號
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_ConsigneeKey.SelStart = 0: txt_ConsigneeKey.SelLength = Len(txt_ConsigneeKey.Text): txt_ConsigneeKey.SetFocus
End Select
End Sub

Private Sub txt_ConsigneeKey_KeyPress(KeyAscii As Integer)
'客戶編號
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_ConsigName.SelStart = 0: txt_ConsigName.SelLength = Len(txt_ConsigName.Text): txt_ConsigName.SetFocus
End Select
End Sub

Private Sub txt_ConsigName_KeyPress(KeyAscii As Integer)
'客戶名稱
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FRouteNo_Start.SelStart = 0: txt_FRouteNo_Start.SelLength = Len(txt_FRouteNo_Start.Text)
         txt_FRouteNo_Start.SetFocus
End Select
End Sub

Private Sub txt_FRouteNo_Start_KeyPress(KeyAscii As Integer)
'一次排車：路線編號：起
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FRouteNo_End.SelStart = 0: txt_FRouteNo_End.SelLength = Len(txt_FRouteNo_End.Text)
         txt_FRouteNo_End.SetFocus
End Select
End Sub

Private Sub txt_FRouteNo_End_KeyPress(KeyAscii As Integer)
'一次排車：路線編號：迄
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FPlanDate_Start.SelStart = 0: txt_FPlanDate_Start.SelLength = Len(txt_FPlanDate_Start.Text)
         txt_FPlanDate_Start.SetFocus
End Select
End Sub

Private Sub txt_FPlanDate_Start_Click()
'一次排車：排車日期：起
If Trim(txt_FPlanDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_Start.Text, 4) & "/" & Mid(txt_FPlanDate_Start.Text, 5, 2) & "/" & Right(txt_FPlanDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanDate_Start.Left
mvDate.Top = Frame12.Top + txt_FPlanDate_Start.Top + txt_FPlanDate_Start.Height
mvDate.Tag = "一次排車.排車日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_Start_KeyPress(KeyAscii As Integer)
'一次排車：排車日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanDate_End.SelStart = 0: txt_FPlanDate_End.SelLength = Len(txt_FPlanDate_End.Text)
         txt_FPlanDate_End.SetFocus
End Select
End Sub

Private Sub txt_FPlanDate_End_Click()
'一次排車：排車日期：迄
If Trim(txt_FPlanDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_End.Text, 4) & "/" & Mid(txt_FPlanDate_End.Text, 5, 2) & "/" & Right(txt_FPlanDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanDate_End.Left
mvDate.Top = Frame12.Top + txt_FPlanDate_End.Top + txt_FPlanDate_End.Height
mvDate.Tag = "一次排車.排車日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_End_KeyPress(KeyAscii As Integer)
'一次排車：排車日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_Start.SelStart = 0: txt_FDeliveryDate_Start.SelLength = Len(txt_FDeliveryDate_Start.Text)
         txt_FDeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_Start_Click()
'一次排車：出車日期：起
If Trim(txt_FDeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_Start.Text, 4) & "/" & Mid(txt_FDeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FDeliveryDate_Start.Left
mvDate.Top = Frame12.Top + txt_FDeliveryDate_Start.Top + txt_FDeliveryDate_Start.Height
mvDate.Tag = "一次排車.出車日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_Start_KeyPress(KeyAscii As Integer)
'一次排車：出車日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_End.SelStart = 0: txt_FDeliveryDate_End.SelLength = Len(txt_FDeliveryDate_End.Text)
         txt_FDeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_End_Click()
'一次排車：出車日期：迄
If Trim(txt_FDeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_End.Text, 4) & "/" & Mid(txt_FDeliveryDate_End.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FDeliveryDate_End.Left
mvDate.Top = Frame12.Top + txt_FDeliveryDate_End.Top + txt_FDeliveryDate_End.Height
mvDate.Tag = "一次排車.出車日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_End_KeyPress(KeyAscii As Integer)
'一次排車：出車日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanCheckin_Start.SelStart = 0: txt_FPlanCheckin_Start.SelLength = Len(txt_FPlanCheckin_Start.Text)
         txt_FPlanCheckin_Start.SetFocus
End Select
End Sub

Private Sub txt_FPlanCheckin_Start_Click()
'一次排車：預計報到日期：起
If Trim(txt_FPlanCheckin_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanCheckin_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanCheckin_Start.Text, 4) & "/" & Mid(txt_FPlanCheckin_Start.Text, 5, 2) & "/" & Right(txt_FPlanCheckin_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanCheckin_Start.Left
mvDate.Top = Frame12.Top + txt_FPlanCheckin_Start.Top + txt_FPlanCheckin_Start.Height
mvDate.Tag = "一次排車.預計報到日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanCheckin_Start_KeyPress(KeyAscii As Integer)
'一次排車：預計報到日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanCheckin_End.SelStart = 0: txt_FPlanCheckin_End.SelLength = Len(txt_FPlanCheckin_End.Text)
         txt_FPlanCheckin_End.SetFocus
End Select
End Sub

Private Sub txt_FPlanCheckin_End_Click()
'一次排車：預計報到日期：迄
If Trim(txt_FPlanCheckin_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanCheckin_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanCheckin_End.Text, 4) & "/" & Mid(txt_FPlanCheckin_End.Text, 5, 2) & "/" & Right(txt_FPlanCheckin_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanCheckin_End.Left
mvDate.Top = Frame12.Top + txt_FPlanCheckin_End.Top + txt_FPlanCheckin_End.Height
mvDate.Tag = "一次排車.預計報到日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanCheckin_End_KeyPress(KeyAscii As Integer)
'一次排車：預計報到日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckin_Start.SelStart = 0: txt_FCheckin_Start.SelLength = Len(txt_FCheckin_Start.Text)
         txt_FCheckin_Start.SetFocus
End Select
End Sub

Private Sub txt_FCheckin_Start_Click()
'一次排車：車輛報到日期：起
If Trim(txt_FCheckin_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckin_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckin_Start.Text, 4) & "/" & Mid(txt_FCheckin_Start.Text, 5, 2) & "/" & Right(txt_FCheckin_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckin_Start.Left
mvDate.Top = Frame12.Top + txt_FCheckin_Start.Top + txt_FCheckin_Start.Height
mvDate.Tag = "一次排車.車輛報到日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckin_Start_KeyPress(KeyAscii As Integer)
'一次排車：車輛報到日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckin_End.SelStart = 0: txt_FCheckin_End.SelLength = Len(txt_FCheckin_End.Text)
         txt_FCheckin_End.SetFocus
End Select
End Sub

Private Sub txt_FCheckin_End_Click()
'一次排車：車輛報到日期：迄
If Trim(txt_FCheckin_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckin_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckin_End.Text, 4) & "/" & Mid(txt_FCheckin_End.Text, 5, 2) & "/" & Right(txt_FCheckin_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckin_End.Left
mvDate.Top = Frame12.Top + txt_FCheckin_End.Top + txt_FCheckin_End.Height
mvDate.Tag = "一次排車.車輛報到日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckin_End_KeyPress(KeyAscii As Integer)
'一次排車：車輛報到日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckout_Start.SelStart = 0: txt_FCheckout_Start.SelLength = Len(txt_FCheckout_Start.Text)
         txt_FCheckout_Start.SetFocus
End Select
End Sub

Private Sub txt_FCheckout_Start_Click()
'一次排車：車輛離倉日期：起
If Trim(txt_FCheckout_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckout_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckout_Start.Text, 4) & "/" & Mid(txt_FCheckout_Start.Text, 5, 2) & "/" & Right(txt_FCheckout_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckout_Start.Left
mvDate.Top = Frame12.Top + txt_FCheckout_Start.Top + txt_FCheckout_Start.Height
mvDate.Tag = "一次排車.車輛離倉日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckout_Start_KeyPress(KeyAscii As Integer)
'一次排車：車輛離倉日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckout_End.SelStart = 0: txt_FCheckout_End.SelLength = Len(txt_FCheckout_End.Text)
         txt_FCheckout_End.SetFocus
End Select
End Sub

Private Sub txt_FCheckout_End_Click()
'一次排車：車輛離倉日期：迄
If Trim(txt_FCheckout_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckout_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckout_End.Text, 4) & "/" & Mid(txt_FCheckout_End.Text, 5, 2) & "/" & Right(txt_FCheckout_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckout_End.Left
mvDate.Top = Frame12.Top + txt_FCheckout_End.Top + txt_FCheckout_End.Height
mvDate.Tag = "一次排車.車輛離倉日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckout_End_KeyPress(KeyAscii As Integer)
'一次排車：車輛離倉日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
End Select
End Sub

Private Sub txt_FVehicleID_KeyPress(KeyAscii As Integer)
'一次排車：車牌號碼
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FDriver.SetFocus
End Select
End Sub

Private Sub txt_FReceiptNo_KeyPress(KeyAscii As Integer)
'一次排車：排車訂單編號
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
End Select
End Sub

Private Sub txt_SRouteNo_Start_KeyPress(KeyAscii As Integer)
'二次排車：路線編號：起
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_SRouteNo_End.SelStart = 0: txt_SRouteNo_End.SelLength = Len(txt_SRouteNo_End.Text)
         txt_SRouteNo_End.SetFocus
End Select
End Sub

Private Sub txt_SRouteNo_End_KeyPress(KeyAscii As Integer)
'二次排車：路線編號：迄
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_SDeliveryDate_Start.SelStart = 0: txt_SDeliveryDate_Start.SelLength = Len(txt_SDeliveryDate_Start.Text)
         txt_SDeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_SDeliveryDate_Start_Click()
'二次排車：車輛出車日期：起
If Trim(txt_SDeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDeliveryDate_Start.Text, 4) & "/" & Mid(txt_SDeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_SDeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame13.Left + txt_SDeliveryDate_Start.Left
mvDate.Top = Frame13.Top + (txt_SDeliveryDate_Start.Top + txt_SDeliveryDate_Start.Height) - (mvDate.Height + txt_SDeliveryDate_Start.Height)
mvDate.Tag = "二次排車.車輛出車日期.起"
mvDate.Visible = True
End Sub

Private Sub txt_SDeliveryDate_Start_KeyPress(KeyAscii As Integer)
'二次排車：車輛出車日期：起
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_SDeliveryDate_End.SelStart = 0: txt_SDeliveryDate_End.SelLength = Len(txt_SDeliveryDate_End.Text)
         txt_SDeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_SDeliveryDate_End_Click()
'二次排車：車輛出車日期：迄
If Trim(txt_SDeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDeliveryDate_End.Text, 4) & "/" & Mid(txt_SDeliveryDate_End.Text, 5, 2) & "/" & Right(txt_SDeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame13.Left + txt_SDeliveryDate_End.Left
mvDate.Top = Frame13.Top + (txt_SDeliveryDate_End.Top + txt_SDeliveryDate_End.Height) - (mvDate.Height + txt_SDeliveryDate_End.Height)
mvDate.Tag = "二次排車.車輛出車日期.迄"
mvDate.Visible = True
End Sub

Private Sub txt_SDeliveryDate_End_KeyPress(KeyAscii As Integer)
'二次排車：車輛出車日期：迄
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '不允許輸入字元
         KeyAscii = 0
    Case vbKeyReturn
         txt_SVehicleID.SelStart = 0: txt_SVehicleID.SelLength = Len(txt_SVehicleID.Text)
         txt_SVehicleID.SetFocus
End Select
End Sub

Private Sub txt_SVehicleID_KeyPress(KeyAscii As Integer)
'二次排車：車牌號碼
Select Case KeyAscii
    Case 97 To 122     '小寫字元改為大寫字元
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_SAddWho.SelStart = 0: txt_SAddWho.SelLength = Len(txt_SAddWho.Text)
         txt_SAddWho.SetFocus
End Select
End Sub

Private Function GetFieldWidth(ByVal strFieldName) As Double
'取得查詢結果欄位寬度
Select Case strFieldName
       Case "編號", "貨主", "項次", "Area", "ZIP", "註記", "車次"
            GetFieldWidth = 500
       Case "貨號", "出車數", "矩陣碼", "訂單量", "揀貨量", "單箱重", "PalletTI", "PalletHI", "駕駛人", "裝載量", _
            "簽收量", "異常碼", "責屬碼"
            GetFieldWidth = 800
       Case "OrderKey", "貨主單號", "訂單日期", "送貨日期", "聯絡人", "客戶電話", "揀貨板數", "揀貨重量", "揀貨材積", _
            "車牌號碼", "電話", "碼頭暫存", "裝載板數", "裝載重量", "裝載材積", "報到日期", "報到時間", "離倉日期 ", _
            "離倉時間", "車種代碼", "簽單狀態", "簽收日期", "簽單類別"
            GetFieldWidth = 1000
       Case "單箱材積", "轉入識別", "排車量", "排車板數", "排車重量", "排車材積", "排車日期", "排車時間", "排車者", _
            "回傳日期", "回傳時間", "回傳狀態", "出車日期", "二次排車者", "二次車次", "二次電話", "簽收板數", "簽收重量", _
            "簽收材積"
            GetFieldWidth = 1000
       Case "客戶編號", "排車訂單編號", "路線編號", "預計報到日期", "預計報到時間", "二次路線編號", "二次出車日期", _
            "二次車牌號碼", "二次駕駛人", "二次配送車種", "二次碼頭暫存", "二次報到日期", "二次報到時間", "二次離倉日期", _
            "二次離倉時間", "特殊需求碼1", "特殊需求碼2", "貨運公司代碼", "二次車種代碼"
            GetFieldWidth = 1200
       Case "郵遞區號", "訂單備註", "註記內容", "二次預計報到日期", "二次預期報到時間", "二次貨運公司代碼", _
            "異常原因", "責任歸屬", "簽單輸入時間", "簽單輸入人員"
            GetFieldWidth = 1600
       Case "客戶簡稱", "二次運輸公司"
            GetFieldWidth = 2000
       Case "客戶名稱", "運送區域", "地址", "客戶配送車種", "特殊需求1", "特殊需求2", "品名", "貨運公司", "配送車種"
            GetFieldWidth = 2500
       Case Else
            GetFieldWidth = 1000
End Select
End Function

Private Function GetFieldAlignment(ByVal strFieldName) As Double
'取得查詢結果欄位寬度
Select Case strFieldName
       Case "註記", "貨號", "駕駛人", "聯絡人", "客戶電話", "電話", "碼頭暫存", "離倉時間", "車種代碼", "排車者", "回傳狀態", _
            "二次排車者", "二次車次", "二次電話", "二次車牌號碼", "二次駕駛人", "二次配送車種", "二次碼頭暫存", _
            "特殊需求碼1", "特殊需求碼2", "貨運公司代碼", "二次車種代碼", "郵遞區號", "訂單備註", "註記內容", _
            "二次預計報到日期", "二次預期報到時間", "二次貨運公司代碼", "客戶簡稱", "二次運輸公司", "客戶名稱", _
            "運送區域", "地址", "客戶配送車種", "特殊需求1", "特殊需求2", "品名", "貨運公司", "配送車種", _
            "異常原因", "責任歸屬", "簽單輸入人員", "簽單類別"
            GetFieldAlignment = dbgLeft
       Case "項次", "出車數", "訂單量", "揀貨量", "單箱重", "PalletTI", "PalletHI", "裝載板數", "裝載重量", "裝載材積", _
            "排車量", "排車板數", "排車重量", "排車材積", "裝載量", "揀貨板數", "揀貨重量", "揀貨材積", "單箱材積", _
            "簽收量", "簽收板數", "簽收重量", "簽收材積"
            GetFieldAlignment = dbgRight
       Case "編號", "貨主", "Area", "ZIP", "車次", "矩陣碼", "OrderKey", "貨主單號", "訂單日期", "送貨日期", "車牌號碼", _
            "報到日期", "報到時間", "離倉日期 ", "轉入識別", "排車日期", "排車時間", "回傳日期", "回傳時間", "出車日期", _
            "客戶編號", "排車訂單編號", "路線編號", "預計報到日期", "預計報到時間", "二次路線編號", "二次出車日期", _
            "二次報到日期", "二次報到時間", "二次離倉日期", "二次離倉時間", "異常碼", "責屬碼", "簽單狀態", "簽收日期", _
            "簽單輸入時間"
            GetFieldAlignment = dbgCenter
       Case Else
            GetFieldAlignment = dbgGeneral
End Select
End Function


