VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Report_VTL 
   Caption         =   "VTL�ݨD����"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
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
   ScaleHeight     =   8010
   ScaleWidth      =   11595
   Begin TabDlg.SSTab SSTab 
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   13573
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm_Report_VTL.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "mvDate"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frm_Report_VTL.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "�̪O��Ʀ^��"
      TabPicture(2)   =   "frm_Report_VTL.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin MSComCtl2.MonthView mvDate 
         Height          =   2220
         Left            =   -69120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3600
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
         StartOfWeek     =   136314881
         TitleBackColor  =   -2147483646
         TitleForeColor  =   16777215
         TrailingForeColor=   -2147483643
         CurrentDate     =   38233
         MaxDate         =   2958455
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
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT2 
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
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   8295
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
            TabIndex        =   35
            Top             =   960
            Width           =   1485
         End
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
            TabIndex        =   34
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
            Picture         =   "frm_Report_VTL.frx":0054
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   33
            ToolTipText     =   "����XLVTL����"
            Top             =   1200
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
            Picture         =   "frm_Report_VTL.frx":035E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   32
            Top             =   240
            Width           =   1065
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
            Picture         =   "frm_Report_VTL.frx":1658
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   31
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
            Picture         =   "frm_Report_VTL.frx":196A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   30
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
            Picture         =   "frm_Report_VTL.frx":1C74
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   29
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���r�ɶȭ�""VTL""����"
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
            Index           =   8
            Left            =   1060
            TabIndex        =   41
            Top             =   1680
            Width           =   2625
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�`�N�G�̪O���O�}�Y��""���s�B�հ�""�P""VTL"""
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
            Index           =   7
            Left            =   -225
            TabIndex        =   40
            Top             =   1440
            Width           =   5835
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
            TabIndex        =   37
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�̪O���"
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
            TabIndex        =   36
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   8295
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
            Picture         =   "frm_Report_VTL.frx":2B886
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   26
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
            Picture         =   "frm_Report_VTL.frx":55498
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   22
            Top             =   240
            Width           =   1065
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
            Picture         =   "frm_Report_VTL.frx":557A2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   21
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
            Picture         =   "frm_Report_VTL.frx":55AB4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   20
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
            Picture         =   "frm_Report_VTL.frx":56DAE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   960
            Width           =   1485
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
            Index           =   3
            Left            =   120
            TabIndex        =   24
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
            TabIndex        =   23
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
         Left            =   -74880
         TabIndex        =   14
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT1 
            Height          =   2295
            Left            =   120
            TabIndex        =   15
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
         Top             =   360
         Width           =   8295
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
            Picture         =   "frm_Report_VTL.frx":570B8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   27
            Top             =   1200
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
            Picture         =   "frm_Report_VTL.frx":573C2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   25
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
            Top             =   600
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
            Top             =   600
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
            Picture         =   "frm_Report_VTL.frx":80FD4
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
            Picture         =   "frm_Report_VTL.frx":822CE
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
            Picture         =   "frm_Report_VTL.frx":825E0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ܮw���b��C��1400�e�^��"
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
            Left            =   1440
            TabIndex        =   12
            Top             =   960
            Width           =   2880
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
            Top             =   645
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
            Top             =   660
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
         Top             =   2520
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
      Top             =   7740
      Width           =   11595
      _ExtentX        =   20452
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
            Object.Width           =   13838
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
Attribute VB_Name = "frm_Report_VTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private rsMainT1 As ADODB.Recordset
Private rsMainT2 As ADODB.Recordset
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

End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

SSTab.Tab = 2

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
    
End Sub
Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    SSTab.Height = Me.ScaleHeight - StatusBar.Height
    Frame2.Height = SSTab.Height - Frame1.Height - Frame1.Top - 120: dgMain.Height = Frame2.Height - 360
    Frame4.Height = SSTab.Height - Frame3.Height - Frame1.Top - 120: dgMainT1.Height = Frame4.Height - 360
    Frame6.Height = SSTab.Height - Frame5.Height - Frame1.Top - 120: dgMainT2.Height = Frame4.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab.Width = Me.ScaleWidth
    Frame2.Width = SSTab.Width - 240: dgMain.Width = Frame2.Width - 240
    Frame4.Width = SSTab.Width - 240: dgMainT1.Width = Frame4.Width - 240
    Frame6.Width = SSTab.Width - 240: dgMainT2.Width = Frame4.Width - 240
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
txtDeliveryDateST2 = "": txtDeliveryDateET2 = ""
End Sub
Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel SSTab.TabCaption(0), rsMain
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub
Private Sub cmd2ExcelT1_Click()
On Error GoTo err_Handle
If rsMainT1 Is Nothing Then MsgBox "�L��ƥi�����ɡI", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
Screen.MousePointer = 11
Call WriteOut_RunLog("1/6.��X�p�O���Ӹ��")
Recordset2Excel "LVTL01�����b�ک��Ӫ�", rsMainT1

'..�b���s��EXCEL
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String
With MyXlsApp: .Visible = False
    
Dim rsTmp As New ADODB.Recordset

'�|�p�ХI�ڸ��
'�M��u�@��
strSheet = "�����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

'.Sheets.Add: .ActiveSheet.Name = "�|�p�ХI�ڸ��"
str_SQL = "exec gs_Charge 'LVTL01','" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("2/6.��X�������")
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

'�F���t�e
'�M��u�@��
strSheet = "���y�B�O-28"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
       
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open "exec [gs_LVTL01AR2_VK] '" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "','A280'", cn
Call WriteOut_RunLog("3/6.��X���y�B�O-28���")
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

'�賡�t�e
'�M��u�@��
strSheet = "���y�B�O-14.5"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "exec [gs_LVTL01AR2_VK] '" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "','A145'", cn
Call WriteOut_RunLog("4/6.��X���y�B�O-14.5���")
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

'���s�t�e
'�M��u�@��
strSheet = "����s�t�e"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open "exec [gs_LVTL01AR2_VR] '" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "'", cn
Call WriteOut_RunLog("5/6.��X����s�t�e���")
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

'�����I
'�M��u�@��
strSheet = "�����I"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open "exec es_OtherARP 'LVTL01','" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "'", cn
Call WriteOut_RunLog("6/6.��X�����I���")

'�g�J���D�C
k = 65: j = 1: strCol = ""
For i = 0 To tmp_Rs.Fields.Count - 1
    l = i Mod 26
    .Range(strCol & Chr(k + l) & j).Value = tmp_Rs.Fields(i).Name
    '���W�L26
    If Chr(65 + l) = "Z" Then
        If strCol = "" Then
            strCol = "A"
        Else
            strCol = Chr(Asc(strCol) + 1)
        End If
    End If
Next i

.Range("A2").CopyFromRecordset tmp_Rs

tmp_Rs.Close

.Visible = True: End With

Set MyXlsApp = Nothing
Screen.MousePointer = 0
Call Unload_RunLogForm

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmd2ExcelT2_Click()

'��ƱƧ�
Recordset2Excel SSTab.TabCaption(2), rsMainT2
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub
Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String

str_SQL = "select * from gv_ship2vtl "

'���
'chc_Orderdate = ""
'If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
'   chc_Orderdate = "and �w�s���b�ɶ� between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
'ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
'   chc_Orderdate = "and �w�s���b�ɶ� = '" & txtOrderDateS.Text & "' "
'ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
'   chc_Orderdate = "and �w�s���b�ɶ� = '" & txtOrderDateE.Text & "' "
'End If
'
''�զX�r��
'str_SQL = str_SQL & chc_Orderdate & " and �f�D�s�� ='" & Combo1.Text & "'order by �q�渹�X , ���� , �s�y��"


Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "�X�f�渹,����"

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

On Error GoTo err_Handle
If Len(txtDeliveryDateST1) = 0 Or Len(txtDeliveryDateET1) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT1.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

str_SQL = "select * from gv_sdn05tdetail where �f�D = 'LVTL01' and ��f�� between '" & txtDeliveryDateST1 & "' and '" & txtDeliveryDateET1 & "' "

Set rsMainT1 = New ADODB.Recordset
rsMainT1.CursorLocation = adUseClient
rsMainT1.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT1.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
rsMainT1.Sort = "��f��,���u�s��,�f�D�渹"

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

If Len(RTrim(txtDeliveryDateST2)) = 0 Or Len(RTrim(txtDeliveryDateET2)) = 0 Then MsgBox "�п�J����϶�!", 16, "�d��": Exit Sub

On Error GoTo err_Handle

Screen.MousePointer = 11
Set dgMainT2.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"

'str_SQL = "select ��� ,�w�O,�渹,����,�Ȥ�,�Ȥ�渹=rtrim(�Ȥ�渹),�̪O�s�� = case when ���O = '���s����' then 'PTD1W110110' " & _
'                                                                                    "when ���O = '���s�콦�O' then 'PTK1P100120' " & _
'                                                                                    "when ���O = '���s�콦��' then 'PTK1P100120' " & _
'                                                                                    "when ���O = '���s�콦��' then 'PTK1P101120' " & _
'                                                                                    "when ���O = '���s���ج�' then 'PTK1P102120' " & _
'                                                                                    "when ���O = 'VTL-��O' then 'PTA1W110140' " & _
'                                                                                    "when ���O = 'VTL-�n���h' then 'PTA1W120100' " & _
'                                                                                    "when ���O = 'VTL-���Ťp' then 'PTA3P110110' " & _
'                                                                                    "when ���O = 'VTL-���Ťj' then 'PTA3P140110' " & _
'                                                                                    "when ���O = 'VTL-�콦��' then 'PTB1P110110' " & _
'                                                                                    "when ���O = 'VTL-�[�z�O' then 'PTB1W' " & _
'                                                                                    "when ���O = 'VTL-�����O' then 'PTD1W110110' " & _
'                                                                                    "when ���O = 'VTL-���O' then 'PTE1W110110' " & _
'                                                                                    "when ���O = 'VTL-����O' then 'PTG1W100120' " & _
'                                                                                    "else ���O end,�ɥX,�٤J " & _
'            "from gv_palletdetail pd where (left(���O,2) = '���s' or left(���O,3) = 'VTL') "
'
'Dim chc_DeliveryDate As String
''���
'chc_DeliveryDate = ""
'If Len(txtDeliveryDateST2.Text) > 0 And Len(txtDeliveryDateET2.Text) > 0 Then
'   chc_DeliveryDate = "and ��� between '" & txtDeliveryDateST2.Text & "' and '" & txtDeliveryDateET2.Text & "' "
'ElseIf Len(txtDeliveryDateST2.Text) > 0 And Len(txtDeliveryDateET2.Text) = 0 Then
'   chc_DeliveryDate = "and ��� = '" & txtDeliveryDateST2.Text & "' "
'ElseIf Len(txtDeliveryDateST2.Text) = 0 And Len(txtDeliveryDateET2.Text) > 0 Then
'   chc_DeliveryDate = "and ��� = '" & txtDeliveryDateET2.Text & "' "
'End If
'
'str_SQL = str_SQL & chc_DeliveryDate
           
           
           
str_SQL = "Select ��� = Convert(Char(8), p.adddate, 112),�w�O = isnull(rtrim(o.stop),''),�渹 = rtrim(p.checkno),���� = Rtrim(p.carno),�Ȥ� = isnull(rtrim(pd.customer),''),�Ȥ�渹 = isnull(rtrim(pd.customersheetno),''),�̪O�s�� = case pd.usertype when '���s����' then 'PTD1W110110' " & _
                        "when '���s�콦�O' then 'PTK1P100120' " & _
                        "when '���s�콦��' then 'PTK1P100120' " & _
                        "when '���s�콦��' then 'PTK1P101120' " & _
                        "when '���s���ج�' then 'PTK1P102120' " & _
                        "when 'VTL-��O' then 'PTA1W110140' " & _
                        "when 'VTL-�n���h' then 'PTA1W120100' " & _
                        "when 'VTL-���Ťp' then 'PTA3P110110' " & _
                        "when 'VTL-���Ťj' then 'PTA3P140110' " & _
                        "when 'VTL-�콦��' then 'PTB1P110110' " & _
                        "when 'VTL-�[�z�O' then 'PTB1W' " & _
                        "when 'VTL��O' then 'PTA1W110140' " & _
                        "when 'VTL�j�a' then 'PTA1W120100' " & _
                        "when 'VTL����p' then 'PTA3P110110' " & _
                        "when 'VTL����j' then 'PTA3P140110' " & _
                        "when 'VTL�[�z��' then 'PTB1P110110' " & _
                        "when 'VTL����' then 'PTD1W110110' " & _
                        "when 'VTL���' then 'PTE1W110110' " & _
                        "when 'VTL����' then 'PTG1W100120' when '�հ��콦��' then 'PTK1P100120' when '�հ����ج�' then 'PTK1P100120-1' " & _
                        "when 'VTL-�����O' then 'PTD1W110110' " & _
                        "when 'VTL-���O' then 'PTE1W110110' " & _
                        "when 'VTL-����O' then 'PTG1W100120' else pd.usertype end " & _
",�ɥX = isnull(pd.qtyin,0),�٤J = isnull(pd.qtyout,0) " & _
"From pallet_cds p join pallet_cst pd on p.checkno = pd.checkno and (left(pd.usertype,2) in ('���s','�հ�') or left(pd.usertype,3) = 'VTL') and Convert(Char(8), p.adddate, 112) between '" & txtDeliveryDateST2.Text & "' and '" & txtDeliveryDateET2.Text & "' " & _
"left join orders o on o.externorderkey = pd.customersheetno and o.type <> '�R��' order by Convert(char(8),p.adddate,112),isnull(rtrim(pd.customersheetno),'') "
          
Set rsMainT2 = New ADODB.Recordset
rsMainT2.CursorLocation = adUseClient
rsMainT2.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMainT2.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
rsMainT2.Sort = "���,�Ȥ�渹"

Set dgMainT2.DataSource = rsMainT2: dgMainT2.Visible = False
rsMainT2.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT2
StatusBar.Panels(2).Text = rsMainT2.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT2.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdSaveToText_Click()

If rsMain Is Nothing Then Exit Sub: If rsMain.EOF Then Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11: cmdSaveToText.Enabled = False: dgMain.Enabled = False

Dim i As Integer, strCheck As String, strFileName As String

strFileName = "BestTransaction_" & Format(Now, "yyyymmddhhMMss") & ".csv"

'���r��
If Dir("C:\LVTL01\�X�f��Ʀ^��", vbDirectory) = "" Then MkDirs "C:\LVTL01\�X�f��Ʀ^��"
Open "C:\LVTL01\�X�f��Ʀ^��\" & strFileName For Output As #1

'����}�l
Tran_Level = cn.BeginTrans

rsMain.MoveFirst
Do While Not rsMain.EOF
    Print #1, rsMain("�X�f�渹"); ","; rsMain("��X���"); ","; rsMain("�b�ګȤ�"); ","; rsMain("�e�f�Ȥ�"); ","; rsMain("����"); ","; rsMain("���~�s��"); ","; rsMain("���"); ","; Trim(rsMain("�q��ƶq")); ","; Trim(rsMain("�O������")); ","; rsMain("�帹"); ","; Trim(rsMain("�O��")); ","; Trim(rsMain("�X�f�ƶq")); ","; rsMain("�ϥδ̪OA")
   
    '��s���w�^��
    str_SQL = "update " & strWMSDB & "..orders " & _
                "set yfystatus = '2' ,TrafficCop = null where externorderkey = '" & RTrim(rsMain("�X�f�渹")) & "' and status = 9 and storerkey = 'LVTL01' "
'                "where orderkey in (select od.orderkey from " & strWMSDB & "..orderdetail od where od.externorderkey = '" & strD & "' and od.externlineno = '" & RTrim(strE) & RTrim(strA) & "')"

    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rsMain.MoveNext
Loop

'�����ɮ�
Close

cn.CommitTrans: Tran_Level = 0

Set rsMain = Nothing: Set dgMain.DataSource = Nothing
Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMain.Enabled = True
MsgBox "�X�f�����X����!!" & vbCrLf & "C:\LVTL01\�X�f��Ʀ^��\" & strFileName, vbOKOnly, Me.Caption
Exit Sub

err_Handle:
    Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMain.Enabled = True
    Close
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
    
End Sub

Private Sub cmdSaveToTextT2_Click()

If rsMainT2 Is Nothing Then Exit Sub: If rsMain.EOF Then Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11: cmdSaveToText.Enabled = False: dgMain.Enabled = False

Dim i As Integer, strCheck As String, strFileName As String, strWH As String, strFileName1 As String

strFileName = "���s�̪O" & Format(Now, "yyyymmddhhMMss") & ".csv"
strFileName1 = "���ƴ̪O" & Format(Now, "yyyymmddhhMMss") & ".csv"

'���r��
If Dir("C:\Best\\LVTL01\�̪O��Ʀ^��", vbDirectory) = "" Then MkDirs "C:\Best\LVTL01\�̪O��Ʀ^��"
Open "C:\Best\LVTL01\�̪O��Ʀ^��\" & strFileName For Output As #1
Open "C:\Best\LVTL01\�̪O��Ʀ^��\" & strFileName1 For Output As #2

'����}�l
'Tran_Level = cn.BeginTrans

rsMainT2.MoveFirst
Do While Not rsMainT2.EOF

If UCase(rsMainT2("����")) = "LVTL" Then

    strWH = UCase(rsMainT2("�w�O"))

    If UCase(Left(rsMainT2("�w�O"), 1)) = "P" Then
        Print #2, rsMainT2("���"); ","; strWH; ","; mySplit(rsMainT2("�Ȥ�"), "-", 0); ","; rsMainT2("�Ȥ�渹"); ","; rsMainT2("�̪O�s��"); ","; rsMainT2("�٤J"); ","; rsMainT2("�̪O�s��"); ","; rsMainT2("�ɥX")
    Else
        Print #1, rsMainT2("���"); ","; strWH; ","; mySplit(rsMainT2("�Ȥ�"), "-", 0); ","; rsMainT2("�Ȥ�渹"); ","; rsMainT2("�̪O�s��"); ","; rsMainT2("�٤J"); ","; rsMainT2("�̪O�s��"); ","; rsMainT2("�ɥX")
    End If
    
End If
    rsMainT2.MoveNext
Loop

'�����ɮ�
Close

'cn.CommitTrans: Tran_Level = 0

Set rsMainT2 = Nothing: Set dgMainT2.DataSource = Nothing
Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMainT2.Enabled = True
MsgBox "�̪O�����X����!!" & vbCrLf & "C:\Best\LVTL01\�̪O��Ʀ^��\" & strFileName & vbCrLf & "C:\Best\LVTL01\�̪O��Ʀ^��\" & strFileName1, vbOKOnly, Me.Caption
Exit Sub

err_Handle:
    Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMainT2.Enabled = True
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
Private Sub txtOrderDateS_Click()

Set objMvdateTarget = txtOrderDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateE_Click()

Set objMvdateTarget = txtOrderDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateST1_Click()

Set objMvdateTarget = txtDeliveryDateST1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET1_Click()

Set objMvdateTarget = txtDeliveryDateET1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateST1_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET1_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateST2_Click()

Set objMvdateTarget = txtDeliveryDateST2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET2_Click()

Set objMvdateTarget = txtDeliveryDateET2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateST2_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateET2_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub
