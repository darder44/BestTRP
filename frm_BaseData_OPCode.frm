VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_BaseData_OPCode 
   Caption         =   "   作   業   代  碼   資   料   維   護"
   ClientHeight    =   6405
   ClientLeft      =   810
   ClientTop       =   1530
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   10665
   Begin TabDlg.SSTab SSTab1 
      Height          =   6360
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   11218
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "基本代碼1"
      TabPicture(0)   =   "frm_BaseData_OPCode.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd_Exit(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "車輛種類"
      TabPicture(1)   =   "frm_BaseData_OPCode.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(1)"
      Tab(1).Control(1)=   "dg_Tab1CarType"
      Tab(1).Control(2)=   "cmd_Tab1CarType_Show"
      Tab(1).Control(3)=   "cmd_Exit(2)"
      Tab(1).Control(4)=   "cmd_Tab1CarType_Save"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "特殊需求"
      TabPicture(2)   =   "frm_BaseData_OPCode.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Label1(0)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "運送區域"
      TabPicture(3)   =   "frm_BaseData_OPCode.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(2)"
      Tab(3).Control(1)=   "dg_Tab3Area"
      Tab(3).Control(2)=   "cmd_Tab3Area_Show"
      Tab(3).Control(3)=   "cmd_Exit(3)"
      Tab(3).Control(4)=   "cmd_Tab3Area_Save"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "郵遞區號"
      TabPicture(4)   =   "frm_BaseData_OPCode.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmd_Tab4Zip_Show"
      Tab(4).Control(1)=   "cmd_Exit(4)"
      Tab(4).Control(2)=   "cmd_Tab4Zip_Save"
      Tab(4).Control(3)=   "dg_Tab4Zip"
      Tab(4).Control(4)=   "Label1(3)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "矩陣圖碼"
      TabPicture(5)   =   "frm_BaseData_OPCode.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(4)"
      Tab(5).Control(1)=   "dg_Tab5GridCode"
      Tab(5).Control(2)=   "cmd_Tab5GridCode_Show"
      Tab(5).Control(3)=   "cmd_Exit(5)"
      Tab(5).Control(4)=   "cmd_Tab5GridCodeSave"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "基本代碼2"
      TabPicture(6)   =   "frm_BaseData_OPCode.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1(4)"
      Tab(6).Control(1)=   "Frame1(5)"
      Tab(6).Control(2)=   "cmd_Exit(6)"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   " "
      TabPicture(7)   =   "frm_BaseData_OPCode.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(5)"
      Tab(7).Control(1)=   "dg_Tab7_TRP17M"
      Tab(7).Control(2)=   "cmd_Tab7_DisPlay"
      Tab(7).Control(3)=   "cmd_Tab7_Delete"
      Tab(7).Control(4)=   "cmd_Tab7_Save"
      Tab(7).Control(5)=   "cmd_Exit(7)"
      Tab(7).ControlCount=   6
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
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
         Height          =   945
         Index           =   7
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":00E0
         Style           =   1  '圖片外觀
         TabIndex        =   62
         Top             =   4800
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab7_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "存檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":0522
         Style           =   1  '圖片外觀
         TabIndex        =   61
         Top             =   2640
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab7_Delete 
         BackColor       =   &H00C0FFC0&
         Caption         =   "刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":082C
         Style           =   1  '圖片外觀
         TabIndex        =   60
         Top             =   3720
         Width           =   1035
      End
      Begin VB.Frame Frame3 
         Caption         =   "特殊需求細項"
         Height          =   5535
         Left            =   -69600
         TabIndex        =   56
         Top             =   720
         Width           =   5055
         Begin VB.CommandButton cmd_Tab2_Delete 
            BackColor       =   &H00C0FFC0&
            Caption         =   "刪  除"
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
            Left            =   1440
            Style           =   1  '圖片外觀
            TabIndex        =   63
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab2SpecDemandDetail_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   2895
            Picture         =   "frm_BaseData_OPCode.frx":0B36
            TabIndex        =   58
            Top             =   240
            Width           =   1830
         End
         Begin VB.CommandButton cmd_Tab2SpecDemandDetail_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   255
            Style           =   1  '圖片外觀
            TabIndex        =   57
            Top             =   240
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab2SpecDemandDetail 
            Height          =   4650
            Left            =   120
            TabIndex        =   59
            Top             =   780
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   8202
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.CommandButton cmd_Tab7_DisPlay 
         BackColor       =   &H00FFC0C0&
         Caption         =   "顯示所有資料"
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
         Left            =   -68310
         Picture         =   "frm_BaseData_OPCode.frx":0E40
         TabIndex        =   49
         Top             =   480
         Width           =   1830
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
         Height          =   480
         Index           =   6
         Left            =   -65625
         Style           =   1  '圖片外觀
         TabIndex        =   48
         Top             =   435
         Width           =   1050
      End
      Begin VB.Frame Frame1 
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
         ForeColor       =   &H00C00000&
         Height          =   2565
         Index           =   5
         Left            =   -74805
         TabIndex        =   44
         Top             =   930
         Width           =   5100
         Begin VB.CommandButton cmd_Tab6RSC_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":114A
            TabIndex        =   46
            Top             =   285
            Width           =   1500
         End
         Begin VB.CommandButton cmd_Tab6RSC_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":1454
            TabIndex        =   45
            Top             =   285
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab6RSC 
            Height          =   1785
            Left            =   90
            TabIndex        =   47
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "異常責屬"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2565
         Index           =   4
         Left            =   -69660
         TabIndex        =   40
         Top             =   930
         Width           =   5100
         Begin VB.CommandButton cmd_Tab6RBC_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":175E
            TabIndex        =   42
            Top             =   285
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab6RBC_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":1A68
            TabIndex        =   41
            Top             =   285
            Width           =   1500
         End
         Begin MSDataGridLib.DataGrid dg_Tab6RBC 
            Height          =   1785
            Left            =   90
            TabIndex        =   43
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.CommandButton cmd_Tab5GridCodeSave 
         BackColor       =   &H00FFC0C0&
         Caption         =   "存  檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -65985
         Picture         =   "frm_BaseData_OPCode.frx":1D72
         Style           =   1  '圖片外觀
         TabIndex        =   39
         Top             =   3840
         Width           =   1125
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
         Height          =   1050
         Index           =   5
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":207C
         Style           =   1  '圖片外觀
         TabIndex        =   38
         Top             =   5055
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Tab5GridCode_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "顯示所有資料"
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
         Left            =   -68355
         Picture         =   "frm_BaseData_OPCode.frx":24BE
         TabIndex        =   35
         Top             =   600
         Width           =   1830
      End
      Begin VB.CommandButton cmd_Tab4Zip_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "顯示所有資料"
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
         Left            =   -68040
         Picture         =   "frm_BaseData_OPCode.frx":27C8
         TabIndex        =   32
         Top             =   510
         Width           =   1830
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
         Height          =   1050
         Index           =   4
         Left            =   -65865
         Picture         =   "frm_BaseData_OPCode.frx":2AD2
         Style           =   1  '圖片外觀
         TabIndex        =   31
         Top             =   4710
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Tab4Zip_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "存  檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -65865
         Picture         =   "frm_BaseData_OPCode.frx":2F14
         Style           =   1  '圖片外觀
         TabIndex        =   30
         Top             =   3495
         Width           =   1125
      End
      Begin VB.CommandButton cmd_Tab3Area_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "存  檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -65880
         Picture         =   "frm_BaseData_OPCode.frx":321E
         Style           =   1  '圖片外觀
         TabIndex        =   27
         Top             =   3525
         Width           =   1125
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
         Height          =   1050
         Index           =   3
         Left            =   -65895
         Picture         =   "frm_BaseData_OPCode.frx":3528
         Style           =   1  '圖片外觀
         TabIndex        =   26
         Top             =   4740
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Tab3Area_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "顯示所有資料"
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
         Left            =   -68070
         Picture         =   "frm_BaseData_OPCode.frx":396A
         TabIndex        =   25
         Top             =   540
         Width           =   1830
      End
      Begin VB.CommandButton cmd_Tab1CarType_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "存  檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -66105
         Picture         =   "frm_BaseData_OPCode.frx":3C74
         Style           =   1  '圖片外觀
         TabIndex        =   21
         Top             =   3630
         Width           =   1050
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
         Height          =   945
         Index           =   2
         Left            =   -66120
         Picture         =   "frm_BaseData_OPCode.frx":3F7E
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   4830
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab1CarType_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "顯示所有資料"
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
         Left            =   -68430
         Picture         =   "frm_BaseData_OPCode.frx":43C0
         TabIndex        =   19
         Top             =   540
         Width           =   1830
      End
      Begin VB.Frame Frame1 
         Caption         =   "搬運工具"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2565
         Index           =   3
         Left            =   5325
         TabIndex        =   15
         Top             =   3615
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Move_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":46CA
            TabIndex        =   17
            Top             =   285
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab0Move_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":49D4
            TabIndex        =   16
            Top             =   285
            Width           =   1500
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Move 
            Height          =   1770
            Left            =   90
            TabIndex        =   18
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3122
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "僱用方式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2565
         Index           =   2
         Left            =   5340
         TabIndex        =   11
         Top             =   855
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Employ_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":4CDE
            TabIndex        =   13
            Top             =   285
            Width           =   1500
         End
         Begin VB.CommandButton cmd_Tab0Employ_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":4FE8
            TabIndex        =   12
            Top             =   285
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Employ 
            Height          =   1785
            Left            =   90
            TabIndex        =   14
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "裝卸方式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2565
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   3615
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Load_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":52F2
            TabIndex        =   9
            Top             =   285
            Width           =   1500
         End
         Begin VB.CommandButton cmd_Tab0Load_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":55FC
            TabIndex        =   8
            Top             =   285
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Load 
            Height          =   1785
            Left            =   90
            TabIndex        =   10
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "車廂形式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2565
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   855
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Box_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":5906
            TabIndex        =   5
            Top             =   285
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab0Box_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":5C10
            TabIndex        =   3
            Top             =   285
            Width           =   1500
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Box 
            Height          =   1785
            Left            =   90
            TabIndex        =   4
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
         Height          =   480
         Index           =   0
         Left            =   9375
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   360
         Width           =   1050
      End
      Begin MSDataGridLib.DataGrid dg_Tab1CarType 
         Height          =   5130
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab3Area 
         Height          =   5130
         Left            =   -74745
         TabIndex        =   28
         Top             =   960
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab4Zip 
         Height          =   5130
         Left            =   -74715
         TabIndex        =   33
         Top             =   930
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab5GridCode 
         Height          =   5130
         Left            =   -74790
         TabIndex        =   36
         Top             =   1020
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab7_TRP17M 
         Height          =   5130
         Left            =   -74640
         TabIndex        =   50
         Top             =   900
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame2 
         Caption         =   "客戶特殊需求"
         Height          =   5535
         Left            =   -74880
         TabIndex        =   52
         Top             =   720
         Width           =   5055
         Begin VB.CommandButton cmd_Tab2SpecDemand_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "存  檔"
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
            Left            =   255
            Style           =   1  '圖片外觀
            TabIndex        =   54
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab2SpecDemand_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "顯示所有資料"
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
            Left            =   2895
            Picture         =   "frm_BaseData_OPCode.frx":5F1A
            TabIndex        =   53
            Top             =   240
            Width           =   1830
         End
         Begin MSDataGridLib.DataGrid dg_Tab2SpecDemand 
            Height          =   4650
            Left            =   120
            TabIndex        =   55
            Top             =   780
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   8202
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制作業：僅允許新增、修改，不允許刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   0
         Left            =   -74595
         TabIndex        =   24
         Top             =   375
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制作業：僅允許新增、修改，不允許刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   5
         Left            =   -74580
         TabIndex        =   51
         Top             =   495
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制作業：僅允許新增、修改，不允許刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   4
         Left            =   -74745
         TabIndex        =   37
         Top             =   570
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制作業：僅允許新增、修改，不允許刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   3
         Left            =   -74670
         TabIndex        =   34
         Top             =   480
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制作業：僅允許新增、修改，不允許刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   2
         Left            =   -74700
         TabIndex        =   29
         Top             =   510
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制作業：僅允許新增、修改，不允許刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   -74700
         TabIndex        =   23
         Top             =   555
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制作業：僅允許新增、修改，不允許刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   13
         Left            =   1815
         TabIndex        =   6
         Top             =   465
         Width           =   4845
      End
   End
End
Attribute VB_Name = "frm_BaseData_OPCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬
Private intloop As Integer
'基本代碼維護 1
Private rs_Tab0CarBox As ADODB.Recordset       '車廂形式
Private rs_Tab0Load As ADODB.Recordset         '裝卸方式
Private rs_Tab0Employ As ADODB.Recordset       '僱用方式
Private rs_Tab0Move As ADODB.Recordset         '搬運工具

Private rs_Tab1CarType As ADODB.Recordset      '車種代碼
Private rs_Tab2SpecDemand As ADODB.Recordset   '特殊需求
Private rs_Tab3Area As ADODB.Recordset         '運送區域
Private rs_Tab4Zip As ADODB.Recordset          '郵遞區號
Private rs_Tab5GridCode As ADODB.Recordset     '矩陣圖碼
Private rs_Tab7_TRP17M  As ADODB.Recordset     '計費代碼
Private rs_Tab2SpecDemandDetail As ADODB.Recordset   '特殊需求細項
'基本代碼維護 2
Private rs_Tab6RSC As ADODB.Recordset          '異常原因
Private rs_Tab6RBC As ADODB.Recordset          '異常責屬


Private Sub cmd_Tab0Box_Save_Click()
'基本代碼1 >> 車廂形式 >> 存檔
If rs_Tab0CarBox Is Nothing Then
   msg_text = "請先查詢所有 車廂形式 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0CarBox.RecordCount = 0 Then
   msg_text = "請先查詢所有 車廂形式 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0CarBox.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0CarBox.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0CarBox.Fields("車廂形式").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'CARBOXTYPE' And Code = '" & rs_Tab0CarBox.Fields("代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'CARBOXTYPE','" & rs_Tab0CarBox.Fields("代碼").Value & "','" & rs_Tab0CarBox.Fields("車廂形式").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0CarBox.MoveNext
Loop
rs_Tab0CarBox.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-車廂形式-存檔", Me.Caption, "cmd_Tab0Box_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Box_Show_Click()
'基本代碼1 >> 車廂形式 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Box.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0CarBox)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 車廂形式 " & _
          "From CodeLKUP Where ListName = 'CARBOXTYPE'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0CarBox)
tmp_Rs.Close

With dg_Tab0Box
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab0CarBox.MoveFirst
Set dg_Tab0Box.DataSource = rs_Tab0CarBox
With dg_Tab0Box
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '車廂形式
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車廂形式-顯示所有資料", Me.Caption, "cmd_Tab0Box_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Employ_Save_Click()
'基本代碼1 >> 僱用方式 >> 存檔
If rs_Tab0Employ Is Nothing Then
   msg_text = "請先查詢所有 僱用方式 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0Employ.RecordCount = 0 Then
   msg_text = "請先查詢所有 僱用方式 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0Employ.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0Employ.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0Employ.Fields("僱用方式").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'EMPLOYTYPE' And Code = '" & rs_Tab0Employ.Fields("代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'EMPLOYTYPE','" & rs_Tab0Employ.Fields("代碼").Value & "','" & rs_Tab0Employ.Fields("僱用方式").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0Employ.MoveNext
Loop
rs_Tab0Employ.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-僱用方式-存檔", Me.Caption, "cmd_Tab0Employ_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Employ_Show_Click()
'基本代碼1 >> 僱用方式 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Employ.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0Employ)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 僱用方式 " & _
          "From CodeLKUP Where ListName = 'EMPLOYTYPE'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0Employ)
tmp_Rs.Close

With dg_Tab0Employ
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab0Employ.MoveFirst
Set dg_Tab0Employ.DataSource = rs_Tab0Employ
With dg_Tab0Employ
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '僱用方式
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-僱用方式-顯示所有資料", Me.Caption, "cmd_Tab0Box_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Load_Save_Click()
'基本代碼1 >> 裝卸方式 >> 存檔
If rs_Tab0Load Is Nothing Then
   msg_text = "請先查詢所有 裝卸方式 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0Load.RecordCount = 0 Then
   msg_text = "請先查詢所有 裝卸方式 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0Load.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0Load.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0Load.Fields("裝卸方式").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'LOADUNLOADTYPE' And Code = '" & rs_Tab0Load.Fields("代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'LOADUNLOADTYPE','" & rs_Tab0Load.Fields("代碼").Value & "','" & rs_Tab0Load.Fields("裝卸方式").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0Load.MoveNext
Loop
rs_Tab0Load.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-裝卸方式-存檔", Me.Caption, "cmd_Tab0Load_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Load_Show_Click()
'基本代碼1 >> 裝卸方式 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Load.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0Load)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 裝卸方式 " & _
          "From CodeLKUP Where ListName = 'LOADUNLOADTYPE'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0Load)
tmp_Rs.Close

With dg_Tab0Load
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab0Load.MoveFirst
Set dg_Tab0Load.DataSource = rs_Tab0Load
With dg_Tab0Load
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '裝卸方式
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-裝卸方式-顯示所有資料", Me.Caption, "cmd_Tab0Load_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Move_Save_Click()
'基本代碼1 >> 搬運工具 >> 存檔
If rs_Tab0Move Is Nothing Then
   msg_text = "請先查詢所有 搬運工具 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0Move.RecordCount = 0 Then
   msg_text = "請先查詢所有 搬運工具 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0Move.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0Move.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0Move.Fields("搬運工具").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'MOVETOOL' And Code = '" & rs_Tab0Move.Fields("代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'MOVETOOL','" & rs_Tab0Move.Fields("代碼").Value & "','" & rs_Tab0Move.Fields("搬運工具").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0Move.MoveNext
Loop
rs_Tab0Move.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-裝卸方式-存檔", Me.Caption, "cmd_Tab0Move_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Move_Show_Click()
'基本代碼1 >> 搬運方式 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Move.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0Move)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 搬運工具 " & _
          "From CodeLKUP Where ListName = 'MOVETOOL'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0Move)
tmp_Rs.Close

With dg_Tab0Move
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab0Move.MoveFirst
Set dg_Tab0Move.DataSource = rs_Tab0Move
With dg_Tab0Move
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '搬運工具
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-搬運工具-顯示所有資料", Me.Caption, "cmd_Tab0Move_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1CarType_Save_Click()
'車輛種類 >> 存檔
If rs_Tab1CarType Is Nothing Then
   msg_text = "請先查詢所有車輛種類資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab1CarType.RecordCount = 0 Then
   msg_text = "請先查詢所有車輛種類資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab1CarType.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab1CarType.EOF
   str_SQL = "Update TRP15M " & _
             "Set Description = '" & rs_Tab1CarType.Fields("車輛種類").Value & "' ,Car_Type = '" & rs_Tab1CarType("計費類別") & "' " & _
             "Where Vehicle_Type = '" & rs_Tab1CarType.Fields("代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP15M(Vehicle_Type,Description,Car_Type) Values (" & _
                "'" & rs_Tab1CarType.Fields("代碼").Value & "','" & rs_Tab1CarType("車輛種類") & "','" & rs_Tab1CarType("計費類別") & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab1CarType.MoveNext
Loop
rs_Tab1CarType.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-車輛種類-存檔", Me.Caption, "cmd_Tab1CarType_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1CarType_Show_Click()
'車種代碼 >> 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab1CarType.DataSource = Nothing
Call ReDim_Recordset(rs_Tab1CarType)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Vehicle_Type) AS 代碼, RTRIM(Description) AS 車輛種類 ,計費類別 = car_type " & _
          "From TRP15M Order by Vehicle_Type"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab1CarType)
tmp_Rs.Close

With dg_Tab1CarType
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With

rs_Tab1CarType.MoveFirst
Set dg_Tab1CarType.DataSource = rs_Tab1CarType
With dg_Tab1CarType
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 500       '代碼
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 5000      '車輛種類
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000      '計費類別
    .Columns(3).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車輛種類-顯示所有資料", Me.Caption, "cmd_Tab1CarType_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Delete_Click()

If rs_Tab2SpecDemandDetail Is Nothing Then Exit Sub

Tran_Level = cn.BeginTrans

If MsgBox("確認刪除貨主：" & rs_Tab2SpecDemandDetail.Fields("貨主") & "；客戶群組：" & rs_Tab2SpecDemandDetail("客戶群組") & "；品號：" & rs_Tab2SpecDemandDetail("品號") & " !!", vbOKCancel, "刪除") <> vbOK Then Exit Sub

cn.Execute "delete trp18m where storerkey = '" & rs_Tab2SpecDemandDetail.Fields("貨主") & "' and consigneekey = '" & rs_Tab2SpecDemandDetail.Fields("客戶群組") & "' and code = '" & rs_Tab2SpecDemandDetail.Fields("品號") & "' ", RowsAffect, adExecuteNoRecords

If RowsAffect = 1 Then
    cn.CommitTrans: Tran_Level = 0
    rs_Tab2SpecDemandDetail.Delete
Else
    cn.RollbackTrans
End If

End Sub

Private Sub cmd_Tab2SpecDemand_Save_Click()
'車輛種類 >> 存檔
If rs_Tab2SpecDemand Is Nothing Then
   msg_text = "請先查詢所有特殊需求資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab2SpecDemand.RecordCount = 0 Then
   msg_text = "請先查詢所有特殊需求資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab2SpecDemand.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab2SpecDemand.EOF
   str_SQL = "Update TRP04M " & _
             "Set Description = '" & rs_Tab2SpecDemand.Fields("特殊需求").Value & "' " & _
             "Where Extra_Demand_Code = '" & rs_Tab2SpecDemand.Fields("代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP04M(Extra_Demand_Code,Description) Values (" & _
                "'" & rs_Tab2SpecDemand.Fields("代碼").Value & "','" & rs_Tab2SpecDemand.Fields("特殊需求").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab2SpecDemand.MoveNext
Loop
rs_Tab2SpecDemand.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-特殊需求-存檔", Me.Caption, "cmd_Tab2SpecDemand_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2SpecDemand_Show_Click()
'特殊需求 >> 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab2SpecDemand.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2SpecDemand)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Extra_Demand_Code) AS 代碼, RTRIM(Description) AS 特殊需求 " & _
          "From TRP04M Order by Extra_Demand_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab2SpecDemand)
tmp_Rs.Close

With dg_Tab2SpecDemand
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab2SpecDemand.MoveFirst
Set dg_Tab2SpecDemand.DataSource = rs_Tab2SpecDemand
With dg_Tab2SpecDemand
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 6000       '特殊需求說明
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-特殊需求-顯示所有資料", Me.Caption, "cmd_Tab2SpecDemand_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2SpecDemandDetail_Save_Click()
'特殊需求-->特殊需求細項>> 存檔
If rs_Tab2SpecDemandDetail Is Nothing Then
   msg_text = "請先查詢所有特殊需求資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab2SpecDemandDetail.RecordCount = 0 Then
   msg_text = "請先查詢所有特殊需求資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab2SpecDemandDetail.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab2SpecDemandDetail.EOF
   str_SQL = "Update TRP18M " & _
             "Set Description = '" & rs_Tab2SpecDemandDetail.Fields("需求說明").Value & "'," & _
             " Consigneekey = '" & rs_Tab2SpecDemandDetail.Fields("客戶群組").Value & "'," & _
             " Code = '" & rs_Tab2SpecDemandDetail.Fields("品號").Value & "'," & _
             " Storerkey = '" & rs_Tab2SpecDemandDetail.Fields("貨主").Value & "'" & _
             " Where Code = '" & rs_Tab2SpecDemandDetail("品號") & "' and Storerkey = '" & rs_Tab2SpecDemandDetail("貨主") & "' and consigneekey = '" & rs_Tab2SpecDemandDetail("客戶群組") & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP18M(Storerkey,Consigneekey,Code,Description) Values (" & _
                "'" & UCase(rs_Tab2SpecDemandDetail("貨主")) & "','" & UCase(rs_Tab2SpecDemandDetail.Fields("客戶群組").Value) & "'," & _
                "'" & UCase(rs_Tab2SpecDemandDetail("品號")) & "','" & rs_Tab2SpecDemandDetail.Fields("需求說明").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab2SpecDemandDetail.MoveNext
Loop

rs_Tab2SpecDemandDetail.MoveFirst
cn.CommitTrans: Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-特殊需求細項-存檔", Me.Caption, "cmd_Tab2SpecDemandDetail_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2SpecDemandDetail_Show_Click()
'特殊需求細項 >> 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Set dg_Tab2SpecDemandDetail.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2SpecDemandDetail)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT " & _
        "RTRIM(Storerkey) AS 貨主, " & _
        "RTRIM(consigneekey) AS 客戶群組 , " & _
        "RTRIM(code) as 品號 , " & _
        " isnull(RTRIM(Description),'') AS 需求說明  " & _
        "From TRP18M "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If tmp_rs.EOF Then
'   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
'   Exit Sub
'End If

Call Replication_Recordset(tmp_Rs, rs_Tab2SpecDemandDetail)
tmp_Rs.Close

With dg_Tab2SpecDemandDetail
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With

If Not rs_Tab2SpecDemandDetail.EOF Then rs_Tab2SpecDemandDetail.MoveFirst

Set dg_Tab2SpecDemandDetail.DataSource = rs_Tab2SpecDemandDetail
With dg_Tab2SpecDemandDetail
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 800       '貨主
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '客戶群組
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '品號
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000       '備註
    .Columns(4).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-特殊需求-顯示所有資料", Me.Caption, "cmd_Tab2SpecDemand_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3Area_Save_Click()
'運送區域 >> 存檔
If rs_Tab3Area Is Nothing Then
   msg_text = "請先查詢所有[運送區域]資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab3Area.RecordCount = 0 Then
   msg_text = "請先查詢所有[運送區域]資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab3Area.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab3Area.EOF
   str_SQL = "Update TRP03M " & _
             "Set Description = '" & rs_Tab3Area.Fields("運送區域").Value & "', "
   If Trim(rs_Tab3Area.Fields("最大車種限制").Value) = "" Then
      str_SQL = str_SQL & "Max_Size_Limit = null,"
   Else
      str_SQL = str_SQL & "Max_Size_Limit = " & Val(rs_Tab3Area.Fields("最大車種限制").Value) & ","
   End If
   If Trim(rs_Tab3Area.Fields("最小車種限制").Value) = "" Then
      str_SQL = str_SQL & "Min_Size_Limit = null "
   Else
      str_SQL = str_SQL & "Min_Size_Limit = " & Val(rs_Tab3Area.Fields("最小車種限制").Value) & " "
   End If
   str_SQL = str_SQL & _
             "Where Area_Code = '" & rs_Tab3Area.Fields("代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP03M(Area_Code,Max_Size_Limit,Min_Size_Limit,Description) Values (" & _
                "'" & rs_Tab3Area.Fields("代碼").Value & "',"
      If Trim(rs_Tab3Area.Fields("最大車種限制").Value) = "" Then
         str_SQL = str_SQL & "null,"
      Else
         str_SQL = str_SQL & Val(rs_Tab3Area.Fields("最大車種限制").Value) & ", "
      End If
      If Trim(rs_Tab3Area.Fields("最小車種限制").Value) = "," Then
         str_SQL = str_SQL & "null,'"
      Else
         str_SQL = str_SQL & Val(rs_Tab3Area.Fields("最小車種限制").Value) & ",'"
      End If
      str_SQL = str_SQL & rs_Tab3Area.Fields("運送區域").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab3Area.MoveNext
Loop
rs_Tab3Area.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-運送區域-存檔", Me.Caption, "cmd_Tab3Area_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3Area_Show_Click()
'特殊需求 >> 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab3Area.DataSource = Nothing
Call ReDim_Recordset(rs_Tab3Area)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Area_Code) AS 代碼, RTRIM(Isnull(Cast(MAX_SIZE_LIMIT as varchar(300)),'')) AS 最大車種限制,RTRIM(Isnull(Cast(MIN_SIZE_LIMIT as varchar(300)),'')) AS 最小車種限制, " & _
          "RTRIM(Isnull(Description,'')) AS 運送區域 " & _
          "From TRP03M Order by Area_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab3Area)
tmp_Rs.Close

With dg_Tab3Area
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab3Area.MoveFirst
Set dg_Tab3Area.DataSource = rs_Tab3Area
With dg_Tab3Area
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '最大車種限制
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1300       '最小車種限制
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 4000       '運送區域
    .Columns(4).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-運送區域-顯示所有資料", Me.Caption, "cmd_Tab3Area_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4Zip_Save_Click()
'郵遞區號 >> 存檔
If rs_Tab4Zip Is Nothing Then
   msg_text = "請先查詢所有 [郵遞區號] 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab4Zip.RecordCount = 0 Then
   msg_text = "請先查詢所有 [郵遞區號] 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab4Zip.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab4Zip.EOF
   str_SQL = "Update TRP02M " & _
             "Set Description = '" & rs_Tab4Zip.Fields("說明").Value & "',city = '" & rs_Tab4Zip.Fields("城市").Value & "',dcode = '" & rs_Tab4Zip("信速到著碼") & "',E_Abb = '" & rs_Tab4Zip("縮寫") & "', "
   If Trim(rs_Tab4Zip.Fields("運送區域代碼").Value) = "" Then
      str_SQL = str_SQL & "Area_Code = null "
   Else
      str_SQL = str_SQL & "Area_Code = '" & Trim(rs_Tab4Zip.Fields("運送區域代碼").Value) & "' "
   End If
   str_SQL = str_SQL & _
             "Where ZIP = '" & rs_Tab4Zip.Fields("郵遞區號").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP02M(ZIP,city,dcode,Area_Code,Description,E_Abb) Values (" & _
                "'" & rs_Tab4Zip.Fields("郵遞區號").Value & "', '" & rs_Tab4Zip.Fields("城市").Value & "','" & rs_Tab4Zip.Fields("信速到著碼").Value & "',"
      If Trim(rs_Tab4Zip.Fields("運送區域代碼").Value) = "" Then
         str_SQL = str_SQL & "null,'"
      Else
         str_SQL = str_SQL & "'" & Trim(rs_Tab4Zip.Fields("運送區域代碼").Value) & "', '"
      End If
      str_SQL = str_SQL & rs_Tab4Zip.Fields("說明").Value & "','" & rs_Tab4Zip.Fields("縮寫") & "') "
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab4Zip.MoveNext
Loop
rs_Tab4Zip.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-運送區域-存檔", Me.Caption, "cmd_Tab3Area_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4Zip_Show_Click()
'郵遞區號 >> 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab4Zip.DataSource = Nothing
Call ReDim_Recordset(rs_Tab4Zip)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(ZIP) AS 郵遞區號,RTRIM(Area_Code) AS 運送區域代碼,RTRIM(city) AS 城市,RTRIM(Isnull(Description,'')) AS 說明,信速到著碼=rtrim(isnull(dcode,'')),縮寫 = isnull(E_Abb,'') " & _
          "From TRP02M Order by ZIP "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab4Zip)
tmp_Rs.Close

With dg_Tab4Zip
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab4Zip.MoveFirst
Set dg_Tab4Zip.DataSource = rs_Tab4Zip
With dg_Tab4Zip
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '郵遞區號
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '運送區域代碼
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '說明
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1500       '說明
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1000       '信速到著碼
    .Columns(5).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-郵遞區號-顯示所有資料", Me.Caption, "cmd_Tab4ZIP_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5GridCode_Show_Click()
'矩陣圖碼 >> 顯示所有資料
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab5GridCode.DataSource = Nothing
Call ReDim_Recordset(rs_Tab5GridCode)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select Rtrim(Grid_Code) as 矩陣圖碼 , Rtrim(Isnull(Grid_Type ,'')) as 類別,Rtrim(Isnull(X_Coordinate,'')) as X座標," & _
          "   Rtrim(Isnull(Y_Coordinate,'')) as Y座標 , Rtrim(Isnull(Description,'')) as 說明 " & _
          "From TRP14M order by GRID_CODE"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab5GridCode)
tmp_Rs.Close

With dg_Tab5GridCode
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab5GridCode.MoveFirst
Set dg_Tab5GridCode.DataSource = rs_Tab5GridCode
With dg_Tab5GridCode
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '矩陣圖碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '矩陣圖類別
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       'X座標
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000       'Y座標
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 3000       '說明
    .Columns(5).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-矩陣圖碼-顯示所有資料", Me.Caption, "cmd_Tab5GridCode_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5GridCodeSave_Click()
'矩陣圖碼 >> 存檔
If rs_Tab5GridCode Is Nothing Then
   msg_text = "請先查詢所有 [矩陣圖碼] 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab5GridCode.RecordCount = 0 Then
   msg_text = "請先查詢所有 [矩陣圖碼] 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab5GridCode.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab5GridCode.EOF
   str_SQL = "Update TRP14M " & _
             "Set Grid_Type = '" & rs_Tab5GridCode.Fields("類別").Value & "',X_Coordinate = '" & rs_Tab5GridCode.Fields("X座標").Value & "'," & _
             " Y_Coordinate = '" & rs_Tab5GridCode.Fields("Y座標").Value & "',Description = '" & rs_Tab5GridCode.Fields("說明").Value & "' " & _
             "Where Grid_Code = '" & rs_Tab5GridCode.Fields("矩陣圖碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP14M(Grid_Code,Grid_Type,X_Coordinate,Y_Coordinate,Description) Values (" & _
                "'" & rs_Tab4Zip.Fields("矩陣圖碼").Value & "','" & rs_Tab5GridCode.Fields("類別").Value & "','" & rs_Tab5GridCode.Fields("X座標").Value & "','" & _
                rs_Tab5GridCode.Fields("Y座標").Value & "','" & rs_Tab5GridCode.Fields("說明").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab5GridCode.MoveNext
Loop
rs_Tab5GridCode.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-矩陣圖碼-存檔", Me.Caption, "cmd_Tab5GridCode_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6RBC_Save_Click()
'基本代碼2 >> 異常責屬 >> 存檔
If rs_Tab6RBC Is Nothing Then
   msg_text = "請先查詢所有 異常責屬 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab6RBC.RecordCount = 0 Then
   msg_text = "請先查詢所有 車廂責屬 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab6RBC.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab6RBC.EOF
   str_SQL = "Update TRP06M " & _
             "Set Description = '" & rs_Tab6RBC.Fields("異常責屬").Value & "' " & _
             "Where RBC_Code = '" & rs_Tab6RBC.Fields("責屬代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP06M(RBC_Code,Description) Values (" & _
                "'" & rs_Tab6RBC.Fields("責屬代碼").Value & "','" & rs_Tab6RBC.Fields("異常責屬").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab6RBC.MoveNext
Loop
rs_Tab6RBC.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-異常責屬-存檔", Me.Caption, "cmd_Tab6RBC_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6RBC_Show_Click()
'基本代碼2 >> 異常責屬
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab6RBC.DataSource = Nothing
Call ReDim_Recordset(rs_Tab6RBC)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(RBC_Code) AS 責屬代碼, RTRIM(Description) AS 異常責屬 " & _
          "From TRP06M Order by RBC_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab6RBC)
tmp_Rs.Close

With dg_Tab6RBC
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab6RBC.MoveFirst
Set dg_Tab6RBC.DataSource = rs_Tab6RBC
With dg_Tab6RBC
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '責屬代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2500       '異常責屬
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-異常責屬-顯示所有資料", Me.Caption, "cmd_Tab6RBC_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmd_Tab6RSC_Save_Click()
'基本代碼2 >> 異常原因 >> 存檔
If rs_Tab6RSC Is Nothing Then
   msg_text = "請先查詢所有 異常原因 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab6RSC.RecordCount = 0 Then
   msg_text = "請先查詢所有 車廂形式 資料，確認後再執行 [存檔] 作業"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab6RSC.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab6RSC.EOF
   str_SQL = "Update TRP05M " & _
             "Set Description = '" & rs_Tab6RSC.Fields("異常原因").Value & "' " & _
             "Where RSC_Code = '" & rs_Tab6RSC.Fields("異常代碼").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '找不到可更新的資料列 >> 新增此筆資料
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP05M(RSC_Code,Description) Values (" & _
                "'" & rs_Tab6RSC.Fields("異常代碼").Value & "','" & rs_Tab6RSC.Fields("異常原因").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab6RSC.MoveNext
Loop
rs_Tab6RSC.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "存檔作業完成"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-異常原因-存檔", Me.Caption, "cmd_Tab6RSC_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6RSC_Show_Click()
'基本代碼2 >> 異常原因
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab6RSC.DataSource = Nothing
Call ReDim_Recordset(rs_Tab6RSC)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(RSC_Code) AS 異常代碼, RTRIM(Description) AS 異常原因 " & _
          "From TRP05M Order by RSC_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab6RSC)
tmp_Rs.Close

With dg_Tab6RSC
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab6RSC.MoveFirst
Set dg_Tab6RSC.DataSource = rs_Tab6RSC
With dg_Tab6RSC
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '異常代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2500       '異常原因
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-異常原因-顯示所有資料", Me.Caption, "cmd_Tab6RSC_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab7_Delete_Click()

If rs_Tab7_TRP17M Is Nothing Then Exit Sub

Call ReDim_Recordset(tmp_Rs)

str_SQL = "select * from sdn05t s5 join sdn02t s2 on s5.sdn_no = s2.receipt_no and s2.storerkey = '" & rs_Tab7_TRP17M.Fields("貨主") & "' and s5.costcode = '" & rs_Tab7_TRP17M.Fields("代碼") & "' "
tmp_Rs.Open str_SQL, cn
If Not tmp_Rs.EOF Then MsgBox "使用中計費代碼無法刪除!!", 64, "刪除": Exit Sub

If MsgBox("確認刪除計費代碼 " & rs_Tab7_TRP17M.Fields("貨主") & "-" & rs_Tab7_TRP17M("代碼") & " !!", vbOKCancel, "刪除") <> vbOK Then Exit Sub

cn.Execute "delete trp17m where storerkey = '" & rs_Tab7_TRP17M.Fields("貨主") & "' and costcode = '" & rs_Tab7_TRP17M.Fields("代碼") & "' ", RowsAffect, adExecuteNoRecords

rs_Tab7_TRP17M.Delete

tmp_Rs.Close

End Sub

Private Sub cmd_Tab7_DisPlay_Click()
    '車種代碼 >> 顯示所有資料
    On Error GoTo err_Handle
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab1CarType.DataSource = Nothing
    Call ReDim_Recordset(rs_Tab7_TRP17M)
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "SELECT 貨主 = rtrim(storerkey) ,RTRIM(CostCode) AS 代碼,rtrim(CostKind) as 請款類別,單位 = rtrim(UOM) ,Receivable as 應收單價,Payable as 應付單價,rtrim(AreaStart) as 起點,rtrim(AreaEnd) as 迄點,rtrim(CostName) as 計費名稱,rtrim(CostNote) as 說明 " & _
              "From TRP17M Order by storerkey,CostCode"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       msg_text = "資料錯誤：查詢結果傳回 0 列資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
'    Call Replication_Recordset(tmp_rs, rs_Tab7_TRP17M)
    Call OffLineRecordset(tmp_Rs, rs_Tab7_TRP17M)
    tmp_Rs.Close
    
    With dg_Tab1CarType
         .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
         .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
         .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
         .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
    End With
    rs_Tab7_TRP17M.MoveFirst
    Set dg_Tab7_TRP17M.DataSource = rs_Tab7_TRP17M
    With dg_Tab7_TRP17M
        .RowHeight = 250
        .Columns(0).Width = 800        '序號
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 800      '說明
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000       '代碼
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800       '代碼
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1000      '客戶名稱
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 800      '應收單價
        .Columns(5).Alignment = dbgRight
        .Columns(6).Width = 800      '應付單價
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1000      '起點
        .Columns(7).Alignment = dbgLeft
        .Columns(8).Width = 1000      '迄點
        .Columns(8).Alignment = dbgLeft
        .Columns(9).Width = 3000      '說明
        .Columns(9).Alignment = dbgLeft
    End With
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車輛種類-顯示所有資料", Me.Caption, "cmd_Tab1CarType_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab7_Save_Click()
    '車輛種類 >> 存檔
    If rs_Tab7_TRP17M Is Nothing Then
        msg_text = "請先查詢所有資料，確認後再執行 [存檔] 作業"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
        If rs_Tab7_TRP17M.RecordCount = 0 Then
        msg_text = "請先查詢所有資料，確認後再執行 [存檔] 作業"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    On Error GoTo err_Handle
    Screen.MousePointer = vbHourglass
    dg_Tab7_TRP17M.Enabled = False
    DoEvents: DoEvents
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    rs_Tab7_TRP17M.MoveFirst
    Call DB_CheckConnectStatus
    Do While Not rs_Tab7_TRP17M.EOF
    
    If Len(Trim(rs_Tab7_TRP17M.Fields("貨主"))) = 0 Then MsgBox "請輸入貨主資料", 64, "存檔": Tran_Level = 0: cn.RollbackTrans: Screen.MousePointer = 0: dg_Tab7_TRP17M.Enabled = True: Exit Sub
    If Len(Trim(rs_Tab7_TRP17M.Fields("代碼"))) = 0 Then MsgBox "請輸入代碼資料", 64, "存檔": Tran_Level = 0: cn.RollbackTrans: Screen.MousePointer = 0: dg_Tab7_TRP17M.Enabled = True: Exit Sub
    
       str_SQL = "Update TRP17M " & _
                  "Set Storerkey = '" & Trim(rs_Tab7_TRP17M.Fields("貨主")) & "' ,CostName = '" & rs_Tab7_TRP17M.Fields("計費名稱").Value & "',Receivable = '" & rs_Tab7_TRP17M.Fields("應收單價").Value & "', " & _
                  "Payable = '" & rs_Tab7_TRP17M.Fields("應付單價").Value & "',AreaStart = '" & rs_Tab7_TRP17M.Fields("起點").Value & "'," & _
                  "AreaEnd = '" & rs_Tab7_TRP17M.Fields("迄點").Value & "',CostNote = '" & rs_Tab7_TRP17M.Fields("說明").Value & "'," & _
                  "CostKind = '" & rs_Tab7_TRP17M.Fields("請款類別").Value & "' ,UOM = '" & rs_Tab7_TRP17M("單位") & "' " & _
                  "Where Storerkey = '" & rs_Tab7_TRP17M.Fields("貨主") & "' and CostCode = '" & Trim(rs_Tab7_TRP17M.Fields("代碼").Value) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '找不到可更新的資料列 >> 新增此筆資料
        If RowsAffect = 0 Then
            str_SQL = "Insert into TRP17M (Storerkey,CostCode,CostName,Receivable,Payable,AreaStart,AreaEnd,CostNote,CostKind,adduser,UOM) Values (" & _
                      "'" & Trim(rs_Tab7_TRP17M.Fields("貨主").Value) & "','" & Trim(rs_Tab7_TRP17M.Fields("代碼").Value) & "','" & rs_Tab7_TRP17M.Fields("計費名稱").Value & "', '" & rs_Tab7_TRP17M.Fields("應收單價").Value & "', " & _
                      "'" & rs_Tab7_TRP17M.Fields("應付單價").Value & "','" & rs_Tab7_TRP17M.Fields("起點").Value & "'," & _
                      "'" & rs_Tab7_TRP17M.Fields("迄點").Value & "','" & rs_Tab7_TRP17M.Fields("說明").Value & "','" & rs_Tab7_TRP17M.Fields("請款類別").Value & "','" & User_id & "' , '" & rs_Tab7_TRP17M("單位") & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
        rs_Tab7_TRP17M.MoveNext
    Loop
    rs_Tab7_TRP17M.MoveFirst
    cn.CommitTrans: Tran_Level = 0
    
    dg_Tab7_TRP17M.Enabled = True
    msg_text = "存檔作業完成"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    
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
    CreateErrorLog Me.Name & "-計費代碼-存檔", Me.Caption, "cmd_Tab1CarType_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "作業代碼資料維護"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 6405
dbsrcFormWidth = 10665
Me.Height = 6915: Me.Width = 10800: Me.Left = 0
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

SSTab1.Tab = 0

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
Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_BaseData_OPCode = Nothing
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub
