VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_NSL 
   Caption         =   "NSL�ݨD����"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "�ө���"
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
   ScaleWidth      =   10755
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   4080
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   4320
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
         Name            =   "�s�ө���"
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
   Begin TabDlg.SSTab SSTab 
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm_Report_NSL.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " �ХI�ڸ�Ʃ���"
      TabPicture(1)   =   "frm_Report_NSL.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " �t�e���`��"
      TabPicture(2)   =   "frm_Report_NSL.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frm_Report_NSL.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame7"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frm_Report_NSL.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).Control(1)=   "Frame9"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frm_Report_NSL.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame11"
      Tab(5).Control(1)=   "Frame12"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   " "
      TabPicture(6)   =   "frm_Report_NSL.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame14"
      Tab(6).Control(1)=   "Frame13"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   " "
      TabPicture(7)   =   "frm_Report_NSL.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame16"
      Tab(7).Control(1)=   "Frame15"
      Tab(7).ControlCount=   2
      Begin VB.Frame Frame15 
         BackColor       =   &H80000004&
         Caption         =   "Daily Storage Status Report"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   93
         Top             =   720
         Width           =   8295
         Begin VB.CheckBox chkT7 
            Caption         =   "��WH-Y�a�q"
            BeginProperty Font 
               Name            =   "�ө���"
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
            TabIndex        =   108
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtDeliveryDateST7 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateET7 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToTextT7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":00E0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   98
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":03EA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   97
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":16E4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   96
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Picture         =   "frm_Report_NSL.frx":19EE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   95
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdResetT7 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":2B600
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   94
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   102
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�w�s���"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   101
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame16 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   91
         Top             =   2880
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT7 
            Height          =   2295
            Left            =   120
            TabIndex        =   92
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   89
         Top             =   2880
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT6 
            Height          =   2295
            Left            =   120
            TabIndex        =   90
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
         Caption         =   "Goods Arrive Report"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   79
         Top             =   720
         Width           =   8295
         Begin VB.CheckBox chkT6 
            Caption         =   "��WH-Y�a�q"
            BeginProperty Font 
               Name            =   "�ө���"
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
            TabIndex        =   110
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdResetT6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":2B912
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   86
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Picture         =   "frm_Report_NSL.frx":2BC24
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   85
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":55836
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   84
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":55B40
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   83
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdSaveToTextT6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":56E3A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   82
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtDeliveryDateET6 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   81
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST6 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   80
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ư��~���Ĥ@�X�^��}�Y���ӫ~"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   112
            Top             =   720
            Width           =   3360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��f���"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   88
            Top             =   1005
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   87
            Top             =   1020
            Width           =   360
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H80000004&
         Caption         =   "Daily Shipping Report"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   69
         Top             =   660
         Width           =   8295
         Begin VB.CheckBox chkT5 
            Caption         =   "��WH-Y�a�q"
            BeginProperty Font 
               Name            =   "�ө���"
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
            TabIndex        =   111
            Top             =   1320
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtDeliveryDateST5 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   76
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateET5 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   75
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToTextT5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":57144
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   74
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":5744E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   73
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":58748
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   72
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Picture         =   "frm_Report_NSL.frx":58A52
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   71
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdResetT5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":82664
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   70
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   78
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��f���"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   77
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame12 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   67
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT5 
            Height          =   2295
            Left            =   120
            TabIndex        =   68
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
            Name            =   "�s�ө���"
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
         Begin MSDataGridLib.DataGrid dgMainT4 
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
         Caption         =   "Daily WH Picking Report"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Begin VB.CheckBox chkT4 
            Caption         =   "��WH-Y�a�q"
            BeginProperty Font 
               Name            =   "�ө���"
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
            TabIndex        =   109
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdResetT4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":82976
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   63
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Picture         =   "frm_Report_NSL.frx":82C88
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   62
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":AC89A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   61
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":ACBA4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   60
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdSaveToTextT4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":ADE9E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   59
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtDeliveryDateET4 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   58
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST4 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   57
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�t�f���"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   65
            Top             =   1005
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   64
            Top             =   1020
            Width           =   360
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   49
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT3 
            Height          =   2295
            Left            =   120
            TabIndex        =   50
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
         Caption         =   "������Ӫ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   39
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtDeliveryDateST3 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   46
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateET3 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
         Begin VB.CommandButton cmdSaveToTextT3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":AE1A8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":AE4B2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   43
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":AF7AC
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   42
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Picture         =   "frm_Report_NSL.frx":AFAB6
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   41
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdResetT3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":D96C8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   40
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   48
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��f���"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   47
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   34
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT2 
            Height          =   2295
            Left            =   120
            TabIndex        =   35
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   25
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtDeliveryDateET2 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   104
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST2 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   103
            Top             =   1320
            Width           =   1485
         End
         Begin VB.CommandButton cmdResetT2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":D99DA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   36
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtSdnDateST2 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   31
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtSdnDateET2 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   30
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToTextT2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":D9CEC
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   29
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":D9FF6
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   28
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":DB2F0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   27
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Picture         =   "frm_Report_NSL.frx":DB5FA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   26
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��f���"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   106
            Top             =   1365
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   105
            Top             =   1380
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   33
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ñ����"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   32
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   14
         Top             =   660
         Width           =   8295
         Begin VB.CheckBox optRepack 
            Caption         =   "�[�u�p�O"
            Height          =   255
            Left            =   3360
            TabIndex        =   107
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox optTMS 
            Caption         =   "�B��д�"
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox optWMS 
            Caption         =   "���x�д�"
            Height          =   255
            Left            =   2280
            TabIndex        =   52
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton cmdResetT1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":10520C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   37
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Picture         =   "frm_Report_NSL.frx":10551E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   23
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":12F130
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   19
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd2ExcelT1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":12F43A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   18
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdSaveToTextT1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":130734
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   17
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtDeliveryDateET1 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   16
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateST1 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   15
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p�O�϶�"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   21
            Top             =   1005
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   20
            Top             =   1020
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Top             =   2820
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT1 
            Height          =   2295
            Left            =   120
            TabIndex        =   13
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
         Caption         =   "�w�s��Ʀ^��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   4
         Top             =   660
         Width           =   8295
         Begin VB.CheckBox chkLocation 
            Caption         =   "�t Location"
            BeginProperty Font 
               Name            =   "�ө���"
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
            TabIndex        =   51
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton cmdSaveToText 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���r��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":130A3E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "���}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Index           =   0
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":130D48
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   22
            Top             =   1200
            Width           =   1065
         End
         Begin VB.TextBox txtOrderDateE 
            Alignment       =   2  '�m�����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   9
            Top             =   960
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateS 
            Alignment       =   2  '�m�����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   8
            Top             =   960
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CommandButton cmd2Excel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   5880
            Picture         =   "frm_Report_NSL.frx":15A95A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   7
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdReset 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7080
            Picture         =   "frm_Report_NSL.frx":15BC54
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   6
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   4680
            Picture         =   "frm_Report_NSL.frx":15BF66
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�C��08:00�t�Φ۰ʦ^��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   2490
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   11
            Top             =   1005
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            TabIndex        =   10
            Top             =   1020
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "�s�ө���"
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
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7605
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "���A"
            TextSave        =   "���A"
            Object.ToolTipText     =   "���A"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   12330
            MinWidth        =   2646
            Object.ToolTipText     =   "��Ƶ���"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.ToolTipText     =   "�ϥΪ�"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Report_NSL"
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
Unload Me '�������{��
'End �������ε{��
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
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

SSTab.Tab = 1

txtDeliveryDateST3 = Format(Now + 1, "yyyymmdd")
txtDeliveryDateET3 = Format(Now + 7, "yyyymmdd")
txtDeliveryDateST4 = Format(Now, "yyyymm") + "01"
txtDeliveryDateET4 = Format(Now + 1, "yyyymmdd")
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

StatusBar.Panels(2).Text = "0 ����ƦC"
If SSTab.Tab = 0 And (rsMain Is Nothing) = False Then StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
If SSTab.Tab = 1 And (rsMainT1 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT1.RecordCount & " ����ƦC"
If SSTab.Tab = 2 And (rsMainT2 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT2.RecordCount & " ����ƦC"
If SSTab.Tab = 3 And (rsMainT3 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT3.RecordCount & " ����ƦC"
If SSTab.Tab = 4 And (rsMainT4 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT4.RecordCount & " ����ƦC"
If SSTab.Tab = 5 And (rsMainT5 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT5.RecordCount & " ����ƦC"
If SSTab.Tab = 6 And (rsMainT6 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT6.RecordCount & " ����ƦC"
If SSTab.Tab = 7 And (rsMainT7 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT7.RecordCount & " ����ƦC"
    
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    SSTab.Height = Me.ScaleHeight - StatusBar.Height
    Frame2.Height = SSTab.Height - Frame1.Height - Frame1.Top - 120: dgMain.Height = Frame2.Height - 360
    Frame4.Height = SSTab.Height - Frame3.Height - Frame1.Top - 120: dgMainT1.Height = Frame4.Height - 360
    Frame6.Height = SSTab.Height - Frame5.Height - Frame1.Top - 120: dgMainT2.Height = Frame6.Height - 360
    Frame8.Height = SSTab.Height - Frame7.Height - Frame1.Top - 120: dgMainT3.Height = Frame8.Height - 360
    Frame10.Height = SSTab.Height - Frame9.Height - Frame1.Top - 120: dgMainT4.Height = Frame10.Height - 360
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
    Frame12.Width = SSTab.Width - 360: dgMainT5.Width = Frame12.Width - 240
    Frame14.Width = SSTab.Width - 360: dgMainT6.Width = Frame14.Width - 240
    Frame16.Width = SSTab.Width - 360: dgMainT7.Width = Frame16.Width - 240
End If

End Sub

Private Sub cmdReset_Click()

'���]
txtOrderDateS.Text = "": txtOrderDateE.Text = ""

End Sub

Private Sub cmdResetT1_Click()
'���]
txtDeliveryDateST1 = "": txtDeliveryDateET1 = ""
End Sub

Private Sub cmdResetT2_Click()
'���]
txtSdnDateST2 = "": txtSdnDateET2 = ""
txtDeliveryDateST2 = "": txtDeliveryDateET2 = ""
End Sub

Private Sub cmdResetT3_Click()
'���]
txtDeliveryDateST3 = "": txtDeliveryDateET3 = ""
End Sub
Private Sub cmdResetT4_Click()
'���]
txtDeliveryDateST4 = "": txtDeliveryDateET4 = ""
End Sub
Private Sub cmdResetT5_Click()
'���]
txtDeliveryDateST5 = "": txtDeliveryDateET5 = ""
End Sub
Private Sub cmdResetT6_Click()
'���]
txtDeliveryDateST6 = "": txtDeliveryDateET6 = ""
End Sub
Private Sub cmdResetT7_Click()
'���]
txtDeliveryDateST7 = "": txtDeliveryDateET7 = ""
End Sub
Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel "LNSL01�w�s��Ʀ^��", rsMain
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT1_Click()
If optTMS + optWMS + optRepack = 0 Then MsgBox "�п�ܽдڳ������O�I", vbOKOnly, Me.Caption: Exit Sub
If rsMainT1 Is Nothing Then MsgBox "�L��ƥi�����ɡI", vbOKOnly + vbInformation, "Save2Excel": Exit Sub

MsgBox "�t�ζi��j�q�����Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel"

On Error GoTo err_Handle
Screen.MousePointer = 11
Dim rsTmp As New ADODB.Recordset
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    
    If Dir(App.Path & "\XLT\���_�ХI�ک���.xlt") = "" Then '�䤣�쥻���d����
        
        '���d���ɸ��|
        Dim objIni As vbIniFile, strXltPath As String
        Set objIni = New vbIniFile
        
        With objIni
        
            .FileName = striniFileName_FullPath
            strXltPath = RTrim(.ReadData("EXCEL", "XLTPATH", "")) '���䴩�����Ƨ��W��
            
        End With
        Set objIni = Nothing

    End If

    '�L���w���|�ϥΥ������|
    If Len(RTrim(strXltPath)) = 0 Then strXltPath = App.Path & "\XLT"
    
    '�M�䥻���d����
    If Dir(strXltPath & "\���_�ХI�ک���.xlt") <> "" Then
        
        '�}�ҽd����
        .Workbooks.Open (strXltPath & "\���_�ХI�ک���.xlt")
    Else
        '�s�WExcel
        .Workbooks.Add
    End If
    
.ActiveWorkbook.Author = User_id

'TMS�дڳ���
If optTMS = 1 Then

    '���_�p�O���Ӹ��
    '�M��u�@��
    strSheet = "�B�O���Ӹ��"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then .Sheets.Add: .ActiveSheet.Name = strSheet

    Call WriteOut_RunLog("�B��дڡG1/5.�B�O���Ӹ��..")
    rsMainT1.MoveFirst
'    Call OffLineRecordset(rsMainT1, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1
    For i = 0 To rsMainT1.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsMainT1.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsMainT1
    
'    rsTmp.Close

    '�����
    '�M��u�@��
    strSheet = "�����"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
'    str_SQL = "select * from gv_Charge where �f�D = 'LNSL01' and ���f��� between '" & txtDeliveryDateST1 & "' and '" & txtDeliveryDateET1 & "' order by ���f���,����,�д����O "
    str_SQL = "exec gs_Charge 'LNSL01' , '" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    Call WriteOut_RunLog("�B��дڡG2/5.��X�����..")
    tmp_Rs.CursorLocation = adUseClient
    tmp_Rs.Open str_SQL, cn
    tmp_Rs.Sort = "���f���,����,�д����O"
    Call OffLineRecordset(tmp_Rs, rsTmp)
    tmp_Rs.Sort = ""
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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
    
'�q��t�e
    Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�q��t�e"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01ShipCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("�B��дڡG3/5.��X�q��t�e�O...")
    tmp_Rs.Open str_SQL, cn
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

'���f
    Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "���f"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01RCCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("�B��дڡG4/5.��X���f�O....")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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
    
'�h�f�t�e
    Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�h�f"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01returnCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("�B��дڡG5/5.��X�h�f�O...")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

'WMS�дڳ���
If optWMS = 1 Then
    '�i�f
    Screen.MousePointer = 11
        '�M��u�@��
        strSheet = "�i�f"
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
        Next
    
        '�䤣��s�W�u�@��
        If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
        
        str_SQL = "exec gs_LNSL01ReceiptDetailCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
    
        Call WriteOut_RunLog("���x�дڡG1/5.��X�i�f���")
        tmp_Rs.Open str_SQL, cn
        
        Call Replication_Recordset(tmp_Rs, rsTmp)
    
        '�g�J���D�C
        k = 65: j = 1
        For i = 0 To rsTmp.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
            '���W�L26
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

'�X�f�z�f
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�X�f�z�f"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01PickingCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("���x�дڡG2/5.��X�X�f�z�f�O���")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

'���f����
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "���f����"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01RCDetailCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("���x�дڡG3/5.��X���f���Ӹ��")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

'�h�f
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�h�f����"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
    str_SQL = "exec gs_LNSL01ReturnReceiptDetailCost '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("���x�дڡG4/5.��X�h�f���Ӹ��")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

'���f����-�Ѧ�
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "���f����-�Ѧ�"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
    str_SQL = "exec gs_LNSL01ReceiptDetail '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("���x�дڡG5/5.��X�i�f���ӰѦҸ��")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

'�[�u�p�O����
If optRepack = 1 Then
    Screen.MousePointer = 11
        '�M��u�@��
        strSheet = "NPP�[�u�p�O"
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
        Next
    
        '�䤣��s�W�u�@��
        If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
        
        str_SQL = "exec gs_LNSL01repackcharge01 '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
    
        Call WriteOut_RunLog("�[�u�p�O�G1/5.��XNPP�[�u�p�O")
        tmp_Rs.Open str_SQL, cn
        
        Call Replication_Recordset(tmp_Rs, rsTmp)
    
        '�g�J���D�C
        k = 65: j = 1
        For i = 0 To rsTmp.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
            '���W�L26
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

'�DNPP�[�u�p�O���
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�DNPP�[�u�p�O"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_LNSL01repackcharge02 '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("�[�u�p�O�G2/5.��X�DNPP�[�u�p�O")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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
    
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�Q�Z�[�u�U�[�O"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
    str_SQL = "exec gs_LNSL01RepackingPickCharge '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)

    Call WriteOut_RunLog("�[�u�p�O�G3/5.��X�Q�Z�[�u�U�[�O")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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
    
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�Q�Z�[�u�W�[�O"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
    
    str_SQL = "exec gs_LNSL01RepackingPWChage '" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)

    Call WriteOut_RunLog("�[�u�p�O�G4/5.��X�Q�Z�[�u�W�[�O")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�[�u�p�O����"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec gs_Repackcharge 'LNSL01','" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Call WriteOut_RunLog("�[�u�p�O�G5/5.��X�[�u�p�O����")
    tmp_Rs.Open str_SQL, cn
    
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
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

'��ƱƧ�
Recordset2Excel "LNSL01ñ��^��", rsMainT2
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT3_Click()

'��ƱƧ�
Recordset2Excel "LNSL01������Ӫ�", rsMainT3
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT4_Click()

'��ƱƧ�
Recordset2Excel "LNSL01_DailyWHPickingReport", rsMainT4
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT5_Click()

'��ƱƧ�
Recordset2Excel "LNSL01_DailyShippingReport", rsMainT5
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT6_Click()

'��ƱƧ�
Recordset2Excel "LNSL01_DailyGoodsArriveReport", rsMainT6
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT7_Click()

'��ƱƧ�
Recordset2Excel "LNSL01_DailyStorageStatusReport", rsMainT7
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String

If chkLocation = 0 Then
    str_SQL = "select * from gv_INV2LNSL02 b order by PROD,WHCODE,LOTNO "
    MsgBox "�`�N!" & vbCrLf & "1.���t�~���}�Y""190TB""�P�~���^��}�Y���~���w�s�C" & vbCrLf & "2.���t""212,999,D50,D55,D65,R19,R20,R21,R45,R44,Z999""�ܧO���w�s�C" & vbCrLf & "3.���t�f���]�t""TC""�r��w�s�C" & vbCrLf & "4.�f��""4""�}�Y�A�B�ݩ�Packaging���O�ӫ~�A��ܤp���w�s�ƶq�C" & vbCrLf & "5.�t�εL�]�w�j���ƶq�ɡA��ܤp���w�s�ƶq�C" & vbCrLf & "6.�ư�N-Packaging���O�ӫ~�C" & vbCrLf & "7.�S�w�~���w�s�q�|�ഫ���(�Ԭ���F)�C", 64, "�w�s��Ʀ^��"
Else
    str_SQL = "select * ,Location = (select count(distinct l.loc) from " & strWMSDB & "..lotxloc l join " & strWMSDB & "..lotattribute la on l.lot = la.lot and l.qty > 0 and l.storerkey = la.storerkey and l.storerkey = 'LNSL01' and i.WHCODE = la.lottable06 and i.PROD = l.sku and i.EXPDAT = isnull(convert(char(8),la.lottable05,112),'') and i.LOTNO = rtrim(la.lottable03)) from gv_INV2LNSL01 i order by PROD,WHCODE,LOTNO "
End If

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT1_Click()

If Len(txtDeliveryDateST1) = 0 Or Len(txtDeliveryDateET1) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle

Screen.MousePointer = 11
Set dgMainT1.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

str_SQL = "select * from gv_sdn05tdetail where �f�D = 'LNSL01' and ��f�� between '" & txtDeliveryDateST1 & "' and '" & txtDeliveryDateET1 & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: tmp_Rs.Close: tmp_Rs.Sort = "": Exit Sub
tmp_Rs.Sort = "��f��,���u�s��,�f�D�渹"

Set rsMainT1 = New ADODB.Recordset
rsMainT1.CursorLocation = adUseClient
Call Replication_Recordset(tmp_Rs, rsMainT1)
tmp_Rs.Close: tmp_Rs.Sort = ""

Set dgMainT1.DataSource = rsMainT1: dgMainT1.Visible = False
rsMainT1.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT1
StatusBar.Panels(2).Text = rsMainT1.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT1.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT2_Click()

'If Len(RTrim(txtSdnDateST2)) = 0 Or Len(RTrim(txtSdnDateET2)) = 0 Then MsgBox "�п�J����϶�!", 64, "�d��": Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMainT2.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"

str_SQL = "exec gs_LNSL01Abnormal '" & RTrim(txtSdnDateST2) & "','" & RTrim(txtSdnDateET2) & "','" & RTrim(txtDeliveryDateST2) & "','" & RTrim(txtDeliveryDateET2) & "' "
            
Set rsMainT2 = New ADODB.Recordset
rsMainT2.CursorLocation = adUseClient
rsMainT2.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT2.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT2.DataSource = rsMainT2: dgMainT2.Visible = False
rsMainT2.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT2
StatusBar.Panels(2).Text = rsMainT2.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT2.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT3_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST3) = 0 Or Len(txtDeliveryDateET3) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT3.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

str_SQL = "exec gs_LNSL01OrderStatus '" & txtDeliveryDateST3 & "','" & txtDeliveryDateET3 & "' "

Set rsMainT3 = New ADODB.Recordset
rsMainT3.CursorLocation = adUseClient
rsMainT3.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT3.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT3.DataSource = rsMainT3: dgMainT3.Visible = False
rsMainT3.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT3
StatusBar.Panels(2).Text = rsMainT3.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT3.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT4_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST4) = 0 Or Len(txtDeliveryDateET4) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT4.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

If chkT4 = 1 Then
    str_SQL = "exec gs_LNSL01WHPickingReport_Wild '" & txtDeliveryDateST4 & "','" & txtDeliveryDateET4 & "' "
Else
    str_SQL = "exec gs_LNSL01WHPickingReport '" & txtDeliveryDateST4 & "','" & txtDeliveryDateET4 & "' "
End If

Set rsMainT4 = New ADODB.Recordset
rsMainT4.CursorLocation = adUseClient
rsMainT4.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT4.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT4.DataSource = rsMainT4: dgMainT4.Visible = False
rsMainT4.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT4
StatusBar.Panels(2).Text = rsMainT4.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT4.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT5_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST5) = 0 Or Len(txtDeliveryDateET5) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT5.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

str_SQL = "exec gs_LNSL01ShippingReport '" & txtDeliveryDateST5 & "','" & txtDeliveryDateET5 & "' "

Set rsMainT5 = New ADODB.Recordset
rsMainT5.CursorLocation = adUseClient
rsMainT5.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT5.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT5.DataSource = rsMainT5: dgMainT5.Visible = False
rsMainT5.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT5
StatusBar.Panels(2).Text = rsMainT5.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT5.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryT6_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST6) = 0 Or Len(txtDeliveryDateET6) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT6.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

If chkT6 = 1 Then
    str_SQL = "exec gs_LNSL01GoodsArriveReport_wild '" & txtDeliveryDateST6 & "','" & txtDeliveryDateET6 & "' "
Else
    str_SQL = "exec gs_LNSL01GoodsArriveReport '" & txtDeliveryDateST6 & "','" & txtDeliveryDateET6 & "' "
End If


Set rsMainT6 = New ADODB.Recordset
rsMainT6.CursorLocation = adUseClient
rsMainT6.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT6.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT6.DataSource = rsMainT6: dgMainT6.Visible = False
rsMainT6.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT6
StatusBar.Panels(2).Text = rsMainT6.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT6.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdQueryT7_Click()

On Error GoTo err_Handle
If Len(txtDeliveryDateST7) = 0 Or Len(txtDeliveryDateET7) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT7.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

MsgBox "1.NP�BNPP�PBULK���~�A�c�J�Ƥ��ର0" & vbCrLf & "2.F��B�P�Žհӫ~�A�c�J�Ƥ��ର0" & vbCrLf & "3.�ư�212�ܧO", 64, "�`�N"

If chkT7 = 1 Then
    str_SQL = "exec gs_LNSL01Storage_Wild '" & txtDeliveryDateST7 & "','" & txtDeliveryDateET7 & "' "
Else
    str_SQL = "exec gs_LNSL01Storage '" & txtDeliveryDateST7 & "','" & txtDeliveryDateET7 & "' "
End If

Set rsMainT7 = New ADODB.Recordset
rsMainT7.CursorLocation = adUseClient
rsMainT7.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT7.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMainT7.DataSource = rsMainT7: dgMainT7.Visible = False
rsMainT7.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT7
StatusBar.Panels(2).Text = rsMainT7.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT7.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdSaveToText_Click()

If rsMain Is Nothing Then Exit Sub: If rsMain.EOF Then Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11: cmdSaveToText.Enabled = False: dgMain.Enabled = False

Dim i As Integer, strCheck As String, strFileName As String

strFileName = "�X�f�^��" & Format(Now, "yyyymmddhhMMss") & ".txt"

'���r��
If Dir("C:\LNSL01\�X�f�^��", vbDirectory) = "" Then MkDirs "C:\LNSL01\�X�f�^��"
Open "C:\LNSL01\�X�f�^��\" & strFileName For Output As #1

rsMain.Sort = "WMS�渹"

'����}�l
Tran_Level = cn.BeginTrans

rsMain.MoveFirst
Do While Not rsMain.EOF
    Print #1, rsMain("WMS�渹"); rsMain("�X�ܤ�"); rsMain("�w�p��f��"); rsMain("�f�D�q�渹�X"); Format(rsMain("����"), "0000000000"); rsMain("�~��"); Format(rsMain("�ƶq"), "00000000"); rsMain("���"); rsMain("�����"); rsMain("�Ͳ��帹"); rsMain("�ܧO"); rsMain("�Ȥ�s��"); rsMain("�Ȥ�²��"); Format(rsMain("�����`����"), "00000000")
   
    '��s���w�^��
    str_SQL = "update " & strWMSDB & "..orders set yfystatus = '2' ,TrafficCop = null where orderkey = '" & RTrim(rsMain("WMS�渹")) & "' and status = 9 "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rsMain.MoveNext
Loop

Print #1, "Total Count = " & Format(rsMain.RecordCount, "00000000")

'�����ɮ�
Close

cn.CommitTrans: Tran_Level = 0

Set rsMain = Nothing: Set dgMain.DataSource = Nothing
Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMain.Enabled = True
MsgBox "�X�f�����X����!!" & vbCrLf & "C:\LNSL01\�X�f�^��\" & strFileName, vbOKOnly, Me.Caption
Exit Sub

err_Handle:
    Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMain.Enabled = True
    Close
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
    
End Sub

Private Sub cmdSaveToTextT2_Click()

If rsMainT2 Is Nothing Then Exit Sub
If rsMainT2.EOF Then Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11: cmdSaveToTextT2.Enabled = False: dgMainT2.Enabled = False

Dim i As Integer, j As Integer, strCheck As String, strFileName As String, strFileName1 As String

strFileName = "ñ��^��" & Format(Now, "yyyymmddhhMMss") & ".txt"
strFileName1 = "�h�fñ��^��" & Format(Now, "yyyymmddhhMMss") & ".txt"

'���r��
If Dir("C:\LNSL01\ñ��^��", vbDirectory) = "" Then MkDirs "C:\LNSL01\ñ��^��"
Open "C:\LNSL01\ñ��^��\" & strFileName For Output As #1
Open "C:\LNSL01\ñ��^��\" & strFileName1 For Output As #2

rsMainT2.Sort = "�w�p��f��,�f�D�q�渹�X,����"

'����}�l
Tran_Level = cn.BeginTrans

rsMainT2.MoveFirst
Do While Not rsMainT2.EOF
    
    If Len(rsMainT2("WMS�渹")) > 0 Then
        Print #1, rsMainT2("WMS�渹"); rsMainT2("�X�ܤ�"); rsMainT2("�w�p��f��"); rsMainT2("�f�D�q�渹�X"); Format(rsMainT2("����"), "0000000000"); rsMainT2("�~��"); Format(rsMainT2("�X�f�ƶq"), "00000000"); Format(rsMainT2("ñ��ƶq"), "00000000"); rsMainT2("�����"); rsMainT2("�Ͳ��帹"); rsMainT2("�ܧO"); rsMainT2("�Ƶ�"); rsMainT2("�o���^��"); rsMainT2("�Ȥ�s��"); rsMainT2("�Ȥ�²��"); Format(rsMainT2("�����`����"), "00000000")
        i = i + 1
    Else
        Print #2, rsMainT2("WMS�渹"); rsMainT2("�X�ܤ�"); rsMainT2("�w�p��f��"); rsMainT2("�f�D�q�渹�X"); Format(rsMainT2("����"), "0000000000"); rsMainT2("�~��"); Format(rsMainT2("�X�f�ƶq"), "00000000"); Format(rsMainT2("ñ��ƶq"), "00000000"); rsMainT2("�����"); rsMainT2("�Ͳ��帹"); rsMainT2("�ܧO"); rsMainT2("�Ƶ�"); rsMainT2("�o���^��"); rsMainT2("�Ȥ�s��"); rsMainT2("�Ȥ�²��"); Format(rsMainT2("�����`����"), "00000000")
        j = j + 1
    End If
    
    '��s���w�^��
    str_SQL = "update sdn02t set sdnfeedback = 1 where receipt_no = '" & RTrim(rsMainT2("TMS�渹")) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rsMainT2.MoveNext
Loop

Print #1, "Total Count = " & Format(i, "00000000")
Print #2, "Total Count = " & Format(j, "00000000")

'�����ɮ�
Close #1
Close #2

cn.CommitTrans: Tran_Level = 0

Set rsMainT2 = Nothing: Set dgMainT2.DataSource = Nothing
Screen.MousePointer = 0: cmdSaveToTextT2.Enabled = True: dgMainT2.Enabled = True
MsgBox "ñ��^����X����!!" & vbCrLf & "C:\LNSL01\ñ��^��\" & strFileName & vbCrLf & "C:\LNSL01\ñ��^��\" & strFileName1, vbOKOnly, Me.Caption
Exit Sub

err_Handle:
    Screen.MousePointer = 0: cmdSaveToTextT2.Enabled = True: dgMainT2.Enabled = True
    Close
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
    
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT1
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT2
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT3_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT3
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT4_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT4
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT5_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT5
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT6_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT6
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT7_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT7
'�L��Ʃ���e�Ӥp�A���s�e��
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

Private Sub txtsdnDateST2_Click()

Set objMvdateTarget = txtSdnDateST2
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
