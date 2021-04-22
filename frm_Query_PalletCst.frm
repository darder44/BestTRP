VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Query_PalletCst 
   Caption         =   "統計結餘"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   11385
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "依客戶統計結餘"
      TabPicture(0)   =   "frm_Query_PalletCst.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmnDialog"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd_Text"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd_DetailExcel"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd_Exit(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_Query"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_Cust"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_HeadExcel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_DateS"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "依車號統計結餘"
      TabPicture(1)   =   "frm_Query_PalletCst.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboUserType"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt_CarDate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmd_TextCar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmd_CarExcelDetail"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd_Exit(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmd_QueryCar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt_Car"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmd_CarExcel"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label15"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label11"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label9"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Shape1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label5"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.ComboBox cboUserType 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         TabIndex        =   38
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txt_CarDate 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74640
         TabIndex        =   9
         Top             =   1080
         Width           =   1395
      End
      Begin VB.TextBox txt_DateS 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C000&
         ForeColor       =   &H80000008&
         Height          =   4800
         Left            =   -74880
         TabIndex        =   28
         Top             =   2040
         Width           =   11085
         Begin VB.TextBox txt_CarIn 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   840
            TabIndex        =   31
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txt_CarOut 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   2400
            TabIndex        =   30
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txt_CarSum 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3960
            TabIndex        =   29
            Top             =   240
            Width           =   795
         End
         Begin MSDataGridLib.DataGrid dg_PalletCar 
            Height          =   3960
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   6985
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   16
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
         Begin MSDataGridLib.DataGrid dg_CarPalletDetail 
            Height          =   3960
            Left            =   4320
            TabIndex        =   17
            Top             =   720
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   6985
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   16
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
         Begin VB.Label Label8 
            BackStyle       =   0  '透明
            Caption         =   "借出"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '透明
            Caption         =   "還入"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '透明
            Caption         =   "結餘"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   32
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.CommandButton cmd_TextCar 
         BackColor       =   &H00FFC0C0&
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
         Height          =   990
         Left            =   -68760
         Picture         =   "frm_Query_PalletCst.frx":0038
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   720
         Width           =   1050
      End
      Begin VB.CommandButton cmd_CarExcelDetail 
         BackColor       =   &H00C0FFFF&
         Caption         =   "明細轉Excel"
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
         Left            =   -66360
         Picture         =   "frm_Query_PalletCst.frx":0342
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   720
         Width           =   1050
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
         Height          =   990
         Index           =   1
         Left            =   -65160
         Picture         =   "frm_Query_PalletCst.frx":0C0C
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   720
         Width           =   1050
      End
      Begin VB.CommandButton cmd_QueryCar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "查  詢"
         DownPicture     =   "frm_Query_PalletCst.frx":104E
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
         Left            =   -69960
         Picture         =   "frm_Query_PalletCst.frx":27D0
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txt_Car 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -73200
         TabIndex        =   10
         Top             =   1080
         Width           =   1395
      End
      Begin VB.CommandButton cmd_CarExcel 
         BackColor       =   &H00C0FFFF&
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
         Height          =   990
         Left            =   -67560
         Picture         =   "frm_Query_PalletCst.frx":2C12
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   720
         Width           =   1050
      End
      Begin VB.CommandButton cmd_HeadExcel 
         BackColor       =   &H00C0FFFF&
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
         Height          =   990
         Left            =   7380
         Picture         =   "frm_Query_PalletCst.frx":34DC
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   660
         Width           =   1050
      End
      Begin VB.TextBox txt_Cust 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C000&
         ForeColor       =   &H80000008&
         Height          =   4800
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   11085
         Begin VB.TextBox txt_Sum 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3960
            TabIndex        =   23
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txt_Out 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   2400
            TabIndex        =   22
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txt_In 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   840
            TabIndex        =   21
            Top             =   240
            Width           =   795
         End
         Begin MSDataGridLib.DataGrid dg_PalletCst 
            Height          =   3960
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   6985
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   16
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
         Begin MSDataGridLib.DataGrid dg_PalletDetail 
            Height          =   3960
            Left            =   4320
            TabIndex        =   8
            Top             =   720
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   6985
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   16
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
         Begin VB.Label Label4 
            BackStyle       =   0  '透明
            Caption         =   "結餘"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   26
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '透明
            Caption         =   "還入"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   25
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            Caption         =   "借出"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.CommandButton cmd_Query 
         BackColor       =   &H00C0FFC0&
         Caption         =   "查  詢"
         DownPicture     =   "frm_Query_PalletCst.frx":3DA6
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
         Left            =   5040
         Picture         =   "frm_Query_PalletCst.frx":5528
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   660
         Width           =   1050
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
         Height          =   990
         Index           =   0
         Left            =   9720
         Picture         =   "frm_Query_PalletCst.frx":596A
         Style           =   1  '圖片外觀
         TabIndex        =   6
         Top             =   660
         Width           =   1050
      End
      Begin VB.CommandButton cmd_DetailExcel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "明細轉Excel"
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
         Left            =   8550
         Picture         =   "frm_Query_PalletCst.frx":5DAC
         Style           =   1  '圖片外觀
         TabIndex        =   5
         Top             =   660
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Text 
         BackColor       =   &H00FFC0C0&
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
         Height          =   990
         Left            =   6210
         Picture         =   "frm_Query_PalletCst.frx":6676
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   660
         Width           =   1050
      End
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   5880
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '透明
         Caption         =   "棧板類別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71640
         TabIndex        =   39
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '透明
         Caption         =   "指定日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   37
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '透明
         Caption         =   "指定日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '透明
         Caption         =   "車號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73080
         TabIndex        =   35
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00008080&
         BorderWidth     =   2
         Height          =   1260
         Index           =   0
         Left            =   -74760
         Top             =   600
         Width           =   10725
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "客戶名稱"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72480
         TabIndex        =   27
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "客戶名稱"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00008080&
         BorderWidth     =   2
         Height          =   1260
         Index           =   4
         Left            =   240
         Top             =   540
         Width           =   10725
      End
   End
End
Attribute VB_Name = "frm_Query_PalletCst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private disp_rs As ADODB.Recordset
Private disp_rsd As ADODB.Recordset
Private Cardisp_rs As ADODB.Recordset
Private Cardisp_rsd As ADODB.Recordset
Private PalletChang As Boolean

Private Sub cmd_CarExcel_Click()
    If Cardisp_rs Is Nothing Then Exit Sub
    If Cardisp_rs.RecordCount = 0 Then Exit Sub
    PalletChang = False
    Dim ExcelTitle As String
    Call DocStoreDirectory(strDocPath)
    
    Dim strTranFileName As String           'Excel 檔案名稱
    CmnDialog.DialogTitle = "轉存 Excel 檔"
    CmnDialog.InitDir = "c:\my documents"
    CmnDialog.FileName = "棧板結餘_" & Format(Now, "YYYYMMDDHHNNSS")
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
    On Error GoTo err_handle
    Screen.MousePointer = vbHourglass
    If SaveTo_ExcelFile(strTranFileName, Cardisp_rs) = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
    Else
       Screen.MousePointer = vbDefault
       If Len(strTranFileName) > 0 Then
          msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
          MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       End If
    End If
    Cardisp_rs.MoveFirst
    SaveToExcel = True
    PalletChang = True
    Exit Sub
err_handle:
   Dim tmpString As String
   SaveToExcel = True
   Screen.MousePointer = vbDefault
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--棧板結餘", Me.Caption, "cmd_Tab3SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   PalletChang = True
End Sub

Private Sub cmd_CarExcelDetail_Click()
    If Cardisp_rsd Is Nothing Then Exit Sub
    If Cardisp_rsd.RecordCount = 0 Then Exit Sub
    PalletChang = False
    
    
    Dim MyXlsApp As Excel.Application   '開啟excel檔
    Dim objFld As Field
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '新增Wookbooks
    MyXlsApp.Workbooks.Add
    '新增Sheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "明細表"
    MyXlsApp.ActiveSheet.Name = "明細表"
    i = 3
    'Convert(Varchar,adddate,111) as 日期,Customer as 客戶,CarNo as 車號,CheckNo as 單號,UserType as 班別,QtyIn as '借出',QtyOut as '還入',Notes as 備註
    MyXlsApp.Cells(i, 1).Value = Trim(Cardisp_rsd.Fields(1).Name)
    MyXlsApp.Cells(i, 2).Value = Trim(Cardisp_rsd.Fields(2).Name)
    MyXlsApp.Cells(i, 3).Value = Trim(Cardisp_rsd.Fields(3).Name)
    MyXlsApp.Cells(i, 4).Value = Trim(Cardisp_rsd.Fields(4).Name)
    MyXlsApp.Cells(i, 5).Value = Trim(Cardisp_rsd.Fields(5).Name)
    MyXlsApp.Cells(i, 6).Value = Trim(Cardisp_rsd.Fields(6).Name)
    MyXlsApp.Cells(i, 7).Value = Trim(Cardisp_rsd.Fields(7).Name)
    MyXlsApp.Cells(i, 8).Value = Trim(Cardisp_rsd.Fields(8).Name)
    i = i + 1
    Cardisp_rsd.MoveFirst
    '日期,客戶,車號, 單號,班別,'借出', '還入',備註
    Do While Not Cardisp_rsd.EOF
        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 1).Value = Trim(Cardisp_rsd.Fields(1))
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 2).Value = Trim(Cardisp_rsd.Fields(2))
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = Trim(Cardisp_rsd.Fields(3))
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = Trim(Cardisp_rsd.Fields(4))
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = Trim(Cardisp_rsd.Fields(5))
        MyXlsApp.Cells(i, 6).Value = Trim(Cardisp_rsd.Fields(6))
        MyXlsApp.Cells(i, 7).Value = Trim(Cardisp_rsd.Fields(7))
        MyXlsApp.Cells(i, 8).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 8).Value = Trim(Cardisp_rsd.Fields(8))
        Cardisp_rsd.MoveNext
        i = i + 1
    Loop
    '合併儲存格
    MyXlsApp.Range("A1:H1").Select
    With MyXlsApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    MyXlsApp.Selection.Merge
    MyXlsApp.Selection.Font.Bold = True
    With Selection.Font
        .Name = "新細明體"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Cells(1, 1).Value = "結餘統計表"
   
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A3:H" & i - 1).Select
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
    Dim ExcelTitle As String
    Call DocStoreDirectory(strDocPath)
    
    Cardisp_rsd.MoveFirst
    SaveToExcel = True
    PalletChang = True
    Exit Sub

err_handle:
   Dim tmpString As String
   SaveToExcel = True
   Screen.MousePointer = vbDefault
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--棧板結餘", Me.Caption, "cmd_Tab3SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   PalletChang = True
End Sub

Private Sub cmd_DetailExcel_Click()
    If disp_rsd Is Nothing Then Exit Sub
    If disp_rsd.RecordCount = 0 Then Exit Sub
    PalletChang = False
    
    Dim MyXlsApp As Excel.Application   '開啟excel檔
    Dim objFld As Field
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '新增Wookbooks
    MyXlsApp.Workbooks.Add
    '新增Sheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "明細表"
    MyXlsApp.ActiveSheet.Name = "明細表"
    i = 3
    'Convert(Varchar,adddate,111) as 日期,Customer as 客戶,CarNo as 車號,CheckNo as 單號,UserType as 班別,QtyIn as '借出',QtyOut as '還入',Notes as 備註
    MyXlsApp.Cells(i, 1).Value = Trim(disp_rsd.Fields(1).Name)
    MyXlsApp.Cells(i, 2).Value = Trim(disp_rsd.Fields(2).Name)
    MyXlsApp.Cells(i, 3).Value = Trim(disp_rsd.Fields(3).Name)
    MyXlsApp.Cells(i, 4).Value = Trim(disp_rsd.Fields(4).Name)
    MyXlsApp.Cells(i, 5).Value = Trim(disp_rsd.Fields(5).Name)
    MyXlsApp.Cells(i, 6).Value = Trim(disp_rsd.Fields(6).Name)
    MyXlsApp.Cells(i, 7).Value = Trim(disp_rsd.Fields(7).Name)
    MyXlsApp.Cells(i, 8).Value = Trim(disp_rsd.Fields(8).Name)
    i = i + 1
    disp_rsd.MoveFirst
    '日期,客戶,車號, 單號,班別,'借出', '還入',備註
    Do While Not disp_rsd.EOF
        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 1).Value = Trim(disp_rsd.Fields(1))
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 2).Value = Trim(disp_rsd.Fields(2))
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = Trim(disp_rsd.Fields(3))
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = Trim(disp_rsd.Fields(4))
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = Trim(disp_rsd.Fields(5))
        MyXlsApp.Cells(i, 6).Value = Trim(disp_rsd.Fields(6))
        MyXlsApp.Cells(i, 7).Value = Trim(disp_rsd.Fields(7))
        MyXlsApp.Cells(i, 8).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 8).Value = Trim(disp_rsd.Fields(8))
        disp_rsd.MoveNext
        i = i + 1
    Loop
    '合併儲存格
    MyXlsApp.Range("A1:H1").Select
    With MyXlsApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    MyXlsApp.Selection.Merge
    MyXlsApp.Selection.Font.Bold = True
    With MyXlsApp.Selection.Font
        .Name = "新細明體"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Cells(1, 1).Value = "結餘統計表"
   
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A3:H" & i - 1).Select
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

    disp_rsd.MoveFirst
    SaveToExcel = True
    PalletChang = True
    Exit Sub

err_handle:
   Dim tmpString As String
   SaveToExcel = True
   Screen.MousePointer = vbDefault
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--棧板結餘", Me.Caption, "cmd_Tab3SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   PalletChang = True
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmd_HeadExcel_Click()
    If disp_rs Is Nothing Then Exit Sub
    If disp_rs.RecordCount = 0 Then Exit Sub
    On Error Resume Next
    PalletChang = False
    Dim MyXlsApp As Excel.Application   '開啟excel檔
    Dim objFld As Field
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '新增Wookbooks
    MyXlsApp.Workbooks.Add
    '新增Sheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "結餘統計表"
    MyXlsApp.ActiveSheet.Name = "結餘統計表"
    i = 3
    'select Customer as 客戶,sum(QtyIn) as 借出,sum(QtyOut) as 還入,sum(QtyIn)-sum(QtyOut) as 結餘
    MyXlsApp.Cells(i, 1).Value = Trim(disp_rs.Fields(1).Name)
    MyXlsApp.Cells(i, 2).Value = Trim(disp_rs.Fields(2).Name)
    MyXlsApp.Cells(i, 3).Value = Trim(disp_rs.Fields(3).Name)
    MyXlsApp.Cells(i, 4).Value = Trim(disp_rs.Fields(4).Name)
    i = i + 1
    disp_rs.MoveFirst
    '日期,客戶,車號,單號,班別,借出,還入,備註
    Do While Not disp_rs.EOF
        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '儲存格格式 >> 數字 >> 類別 = 文字
        MyXlsApp.Cells(i, 1).Value = Trim(disp_rs.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(disp_rs.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(disp_rs.Fields(3))
        MyXlsApp.Cells(i, 4).Value = Trim(disp_rs.Fields(4))
        disp_rs.MoveNext
        i = i + 1
    Loop
    '合併儲存格
    MyXlsApp.Range("A1:D1").Select
    With MyXlsApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    MyXlsApp.Selection.Merge
    MyXlsApp.Selection.Font.Bold = True
    With Selection.Font
        .Name = "新細明體"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Cells(1, 1).Value = "結餘統計表"
   
    '全部反白
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '畫線h
    MyXlsApp.Range("A3:D" & i - 1).Select
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
    PalletChang = True
    
    

    Exit Sub
err_handle:
   Dim tmpString As String
   'SaveToExcel = True
   Screen.MousePointer = vbDefault
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--棧板結餘", Me.Caption, "cmd_Tab3SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   PalletChang = True
End Sub

Private Sub cmd_QueryCar_Click()
    On Error Resume Next
    Dim strWhere As String
    Dim strTmp As String
    If Len(Trim(Me.txt_CarDate.Text)) > 0 Then
        strWhere = "and Convert(Varchar(8),adddate,112) <= '" & Trim(Me.txt_CarDate.Text) & "'"
    End If
    'Convert(Varchar,adddate,111)
    '車號
    If Len(Trim(Me.txt_Car.Text)) > 0 Then strWhere = strWhere & " and CarNo = '" & Me.txt_Car.Text & "'"

    '棧板類別
    If Len(RTrim(cboUserType.Text)) > 0 Then strWhere = strWhere & " and usertype = '" & cboUserType.Text & "'"
    
'    '客戶
'    If Len(Trim(strwhere)) > 0 Then
'        strwhere = "where " & strwhere
'    End If

'    If Len(Trim(Me.txt_Car.Text)) = 0 Then
'        str_SQL = "select CarNo as 車號,棧板類別 = usertype,sum(QtyIn) as 借出,sum(QtyOut) as 還入,sum(QtyIn)-sum(QtyOut) as 結餘 from Pallet_Cst where 1 = 1 " & strwhere & " group by usertype "
'    Else
        str_SQL = "select rtrim(CarNo) as 車號,棧板類別 = rtrim(usertype),sum(QtyIn) as 借出,sum(QtyOut) as 還入,sum(QtyIn)-sum(QtyOut) as 結餘 from Pallet_Cst where 1 = 1 " & strWhere & " group by CarNo , usertype "
'    End If
    
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無符合查詢條件之資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Exit Sub
    End If
    
    '複製 recordset
    Call Replication_Recordset(tmp_Rs, Cardisp_rs)
    tmp_Rs.Close
    disp_rs.MoveFirst
    With dg_PalletCst
         .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
         .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
         .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
         .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
    End With
    Set dg_PalletCar.DataSource = Cardisp_rs
    With dg_PalletCar
        .Columns(0).Width = 500      '序號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1200
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 600
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 600
        .Columns(3).Alignment = dbgRight
        .Columns(4).Width = 600
        .Columns(4).Alignment = dbgRight
    End With
    Cardisp_rs.MoveFirst
    Screen.MousePointer = vbDefault
    SSTab1.SetFocus
    str_SQL = "select sum(QtyIn) as 借出,sum(QtyOut) as 還入,sum(QtyIn)-sum(QtyOut) as 結餘 from Pallet_Cst where 1 = 1 " & strWhere & ""
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    Me.txt_CarIn.Text = Trim(tmp_Rs.Fields(0))
    Me.txt_CarOut.Text = Trim(tmp_Rs.Fields(1))
    Me.txt_CarSum.Text = Trim(tmp_Rs.Fields(2))
    tmp_Rs.Close
    Screen.MousePointer = vbDefault
    Exit Sub

err_handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "查詢", Me.Caption, "cmd_Tab1Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title

End Sub

Private Sub cmd_Text_Click()
        '查詢結果 >> 轉文字檔
    PalletChang = False
    If disp_rsd Is Nothing Then Exit Sub
    If disp_rsd.RecordCount = 0 Then Exit Sub
    Call DocStoreDirectory(strDocPath)
    
    Dim strTranFileName As String           '文字檔檔案名稱
    CmnDialog.DialogTitle = "轉存文字檔"
    CmnDialog.InitDir = "c:\my documents"
    CmnDialog.FileName = "棧板結餘_" & Format(Now, "YYYYMMDDHHNNSS")
    CmnDialog.Filter = "純文字檔(*.txt)|*.txt"
    CmnDialog.FilterIndex = 1
    CmnDialog.CancelError = True
    On Error Resume Next
    CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
    CmnDialog.ShowOpen
    If err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
       msg_text = "選擇 [取消] 按鈕，必須於文字檔中自行存檔"
       MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
       strTranFileName = ""
    Else
       strTranFileName = CmnDialog.FileName
       If Dir(strTranFileName) <> "" Then
          Kill strTranFileName
       End If
    End If
    
    On Error GoTo err_handle
    Screen.MousePointer = vbHourglass: DoEvents
    If SaveTo_TextFile(strTranFileName, disp_rs, Me.Name & "棧板結餘") = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
       If disp_rs Is Nothing Then Exit Sub
       disp_rs.MoveFirst
    Else
       Screen.MousePointer = vbDefault
       If Len(strTranFileName) > 0 Then
          msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
          MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       End If
    End If
    PalletChang = True
Exit Sub

err_handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--空儲位查詢結果轉文字檔", Me.Caption, "cmd_SaveToText_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   If disp_rs Is Nothing Then Exit Sub
   PalletChang = True
   disp_rs.MoveFirst
End Sub



Private Sub cmd_TextCar_Click()
    PalletChang = False
    If Cardisp_rsd Is Nothing Then Exit Sub
    If Cardisp_rsd.RecordCount = 0 Then Exit Sub
    Call DocStoreDirectory(strDocPath)
    
    Dim strTranFileName As String           '文字檔檔案名稱
    CmnDialog.DialogTitle = "轉存文字檔"
    CmnDialog.InitDir = "c:\my documents"
    CmnDialog.FileName = "棧板結餘_" & Format(Now, "YYYYMMDDHHNNSS")
    CmnDialog.Filter = "純文字檔(*.txt)|*.txt"
    CmnDialog.FilterIndex = 1
    CmnDialog.CancelError = True
    On Error Resume Next
    CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
    CmnDialog.ShowOpen
    If err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
       msg_text = "選擇 [取消] 按鈕，必須於文字檔中自行存檔"
       MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
       strTranFileName = ""
    Else
       strTranFileName = CmnDialog.FileName
       If Dir(strTranFileName) <> "" Then
          Kill strTranFileName
       End If
    End If
    
    On Error GoTo err_handle
    Screen.MousePointer = vbHourglass: DoEvents
    If SaveTo_TextFile(strTranFileName, Cardisp_rs, Me.Name & "棧板結餘") = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
       If disp_rs Is Nothing Then Exit Sub
       disp_rs.MoveFirst
    Else
       Screen.MousePointer = vbDefault
       If Len(strTranFileName) > 0 Then
          msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
          MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       End If
    End If
    PalletChang = True
Exit Sub

err_handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--空儲位查詢結果轉文字檔", Me.Caption, "cmd_SaveToText_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   If disp_rs Is Nothing Then Exit Sub
   PalletChang = True
   Cardisp_rs.MoveFirst
End Sub

Private Sub dg_PalletCar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If PalletChang = False Then Exit Sub
    If Len(Trim(Me.txt_DateS.Text)) > 0 Then
        str_SQL = "select Convert(Varchar(8),adddate,112) as 日期,Customer as 客戶,CarNo as 車號,CheckNo as 單號,adduser as Keyin,QtyIn as '借出',QtyOut as '還入'," & _
              "Notes as 備註 from Pallet_Cst where CarNo= '" & Trim(Cardisp_rs.Fields(1)) & "' and Convert(Varchar(8),adddate,112) <= '" & Trim(Me.txt_CarDate.Text) & "' order by adddate"
    Else
        str_SQL = "select Convert(Varchar(8),adddate,112) as 日期,Customer as 客戶,CarNo as 車號,CheckNo as 單號,adduser as Keyin,QtyIn as '借出',QtyOut as '還入'," & _
              "Notes as 備註 from Pallet_Cst where CarNo= '" & Trim(Cardisp_rs.Fields(1)) & "' order by adddate"
    End If
    
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       Set dg_PalletDetail.DataSource = Nothing
       Exit Sub
    End If
    Call ReDim_Recordset(Cardisp_rsd)
    Call Replication_Recordset(tmp_Rs, Cardisp_rsd)
    tmp_Rs.Close
    Cardisp_rsd.MoveFirst
    Set dg_CarPalletDetail.DataSource = Cardisp_rsd
    With dg_CarPalletDetail
         .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
         .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
         .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
         .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
    End With
    With dg_CarPalletDetail
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000       'sku
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 800
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 600
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 600
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 2000
        .Columns(8).Alignment = dbgLeft
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub dg_PalletCst_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If PalletChang = False Then Exit Sub
    If Len(Trim(Me.txt_DateS.Text)) > 0 Then
        str_SQL = "select Convert(Varchar(8),adddate,112) as 日期,Customer as 客戶,CarNo as 車號,CheckNo as 單號,adduser as Keyin,QtyIn as '借出',QtyOut as '還入'," & _
              "Notes as 備註 from Pallet_Cst where Customer= '" & disp_rs.Fields(1) & "' and Convert(Varchar(8),adddate,112)<= '" & Trim(Me.txt_DateS.Text) & "' order by adddate"
    Else
        str_SQL = "select Convert(Varchar(8),adddate,112) as 日期,Customer as 客戶,CarNo as 車號,CheckNo as 單號,adduser as Keyin,QtyIn as '借出',QtyOut as '還入'," & _
              "Notes as 備註 from Pallet_Cst where Customer= '" & disp_rs.Fields(1) & "' order by adddate"
    End If
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       Set dg_PalletDetail.DataSource = Nothing
       Exit Sub
    End If
    Call ReDim_Recordset(disp_rsd)
    Call Replication_Recordset(tmp_Rs, disp_rsd)
    tmp_Rs.Close
    disp_rsd.MoveFirst
    Set dg_PalletDetail.DataSource = disp_rsd
    With dg_PalletDetail
         .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
         .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
         .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
         .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
    End With
    With dg_PalletDetail
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000       'sku
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 800
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 600
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 600
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 2000
        .Columns(8).Alignment = dbgLeft
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
    Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
    msg_title = "統計結餘"
End Sub

Private Sub Form_Load()
    Me.Height = 7600: Me.Width = 11500
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200
    PalletChang = True
    
    '倉庫別
    '取參數
    Dim objIni As vbIniFile, arrTmp, i As Integer
    Set objIni = New vbIniFile
    objIni.FileName = striniFileName_FullPath
    
    arrTmp = Split(objIni.ReadData("OPTION", "WAREHOUSE", "0"), ";")
    
    For i = 0 To UBound(arrTmp)
        cboUserType.AddItem arrTmp(i)
    Next
    
End Sub
Private Sub cmd_Query_Click()
    On Error Resume Next
    Dim strWhere As String
    Dim strTmp As String
    If Len(Trim(Me.txt_DateS.Text)) > 0 Then
        strWhere = "Convert(Varchar(8),adddate,112) <= '" & Trim(Me.txt_DateS.Text) & "'"
    End If
    'Convert(Varchar,adddate,111)
    
    If Len(Trim(Me.txt_Cust.Text)) > 0 Then
        If Len(Trim(strWhere)) > 0 Then
            strWhere = strWhere & " and Customer like '" & Me.txt_Cust.Text & "%'"
        Else
            strWhere = strWhere & " Customer like '" & Me.txt_Cust.Text & "%'"
        End If
    End If
    '客戶
    If Len(Trim(strWhere)) > 0 Then
        strWhere = "where " & strWhere
    End If
    str_SQL = "select Customer as 客戶,sum(QtyIn) as 借出,sum(QtyOut) as 還入,sum(QtyIn)-sum(QtyOut) as 結餘 from Pallet_Cst  " & strWhere & " group by Customer"
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無符合查詢條件之資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Exit Sub
    End If
    
    '複製 recordset
    Call Replication_Recordset(tmp_Rs, disp_rs)
    tmp_Rs.Close
    disp_rs.MoveFirst
    With dg_PalletCst
         .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
         .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
         .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
         .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
    End With
    Set dg_PalletCst.DataSource = disp_rs
    With dg_PalletCst
        .Columns(0).Width = 500      '序號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1200
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 600
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 600
        .Columns(3).Alignment = dbgRight
        .Columns(4).Width = 600
        .Columns(4).Alignment = dbgRight
    End With
    Screen.MousePointer = vbDefault
    SSTab1.SetFocus
    '統計總結餘
    str_SQL = "select sum(QtyIn) as 借出,sum(QtyOut) as 還入,sum(QtyIn)-sum(QtyOut) as 結餘 from Pallet_Cst  " & strWhere & ""
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    Me.txt_In.Text = Trim(tmp_Rs.Fields(0))
    Me.txt_Out.Text = Trim(tmp_Rs.Fields(1))
    Me.txt_Sum.Text = Trim(tmp_Rs.Fields(2))
    tmp_Rs.Close
    Screen.MousePointer = vbDefault
    Exit Sub

err_handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "查詢", Me.Caption, "cmd_Tab1Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

