VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_TRPTrack 
   Caption         =   "��f�l�ܪ�"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14055
      Begin VB.CommandButton cmdImportExternShipKey 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�U�B�渹�פJ"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12840
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDeliveryE 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         Caption         =   "��ܤC�Ѥ��˸��I"
         BeginProperty Font 
            Name            =   "�ө���"
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
         ToolTipText     =   "��ܸ˸��I�A�d�߻ݸ��[���ɶ�"
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3480
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   28
         ToolTipText     =   "��f���A"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   4560
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   27
         ToolTipText     =   "�q�檬�A"
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox List4 
         Columns         =   2
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   120
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   4
         ToolTipText     =   "��O"
         Top             =   1320
         Width           =   3405
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   6360
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   5
         ToolTipText     =   "�f�D"
         Top             =   240
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Columns         =   3
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   8880
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   6
         ToolTipText     =   "�ϽX"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPreView 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�w���C�L"
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
         Left            =   6720
         Picture         =   "frm_Report_TRPTrack.frx":0342
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FF8080&
         Caption         =   "����C�L"
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
         Left            =   5640
         Picture         =   "frm_Report_TRPTrack.frx":064C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         Left            =   10440
         Picture         =   "frm_Report_TRPTrack.frx":0956
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   8
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
         Left            =   11640
         Picture         =   "frm_Report_TRPTrack.frx":1C50
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   1200
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
         Left            =   11640
         Picture         =   "frm_Report_TRPTrack.frx":2B862
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
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
         Left            =   10440
         Picture         =   "frm_Report_TRPTrack.frx":2BB74
         Style           =   1  '�Ϥ��~�[
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
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�X�����"
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
         Left            =   120
         TabIndex        =   34
         Top             =   285
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
         Index           =   2
         Left            =   2640
         TabIndex        =   33
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "A2B.���f�t�e"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "RC.���f�J�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "R.�h�f�q��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "A.�Q�����"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "I.���`�q��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   19
         Top             =   660
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
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���u�s��"
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
         TabIndex        =   17
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
         Index           =   1
         Left            =   2655
         TabIndex        =   16
         Top             =   1020
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
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
            Text            =   "���A"
            TextSave        =   "���A"
            Object.ToolTipText     =   "���A"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   18177
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
msg_title = "���B�渹�פJ"

On Error Resume Next
Dim strFileName As String, strFieldName As String, k As Integer, j As Integer, i As Integer, arrTmp

With dlgCommonDialog
    .DialogTitle = "���B�渹�פJ"
    .CancelError = True
    .InitDir = App.Path
    'ToDo: �]�w�q�ι�ܤ��������X�Ф��ݩ�
    .Filter = "*.csv|*.csv"
    .ShowOpen
    strFileName = .FileName
    
    If err.Number = cdlCancel Then strFileName = "": Exit Sub
    
    If Len(strFileName) = 0 Then Exit Sub

End With

On Error GoTo err_Handle
If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "���B�渹�פJ": Exit Sub '�䤣���ɮ�

Screen.MousePointer = 11

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

'    '�M����w�u�@��
'    For i = 1 To .Sheets.Count
'        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
'        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '���Ĥ@�Ӥu�@��W��
'        .Sheets(.Sheets(i).Name).Select
'    Next
'
'    '�䤣����w�u�@��A��βĤ@��
'    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(1).Select
.Sheets(1).Select
    
'    For i = 1 To .Sheets.Count
'        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
'
'    Next
'
'    If RTrim(.Sheets(i).Name) <> strSheetName Then
'        '�䤣��βĤ@��
'
''        MsgBox "�䤣�� " & strSheetName & "�u�@��I", vbOKOnly + vbInformation, "Excel2Recordset"
''        .Quit: Set MyXlsApp = Nothing
''        Exit Sub
'    End If
    
    k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
'    If strFieldName = "" Then
'        '�����W��
        For i = 1 To 255
            If Len(RTrim(.Cells(1, i) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(RTrim(.Cells(1, i))) & Chr(9)
        Next i
        k = 2 '�ѲĤG�C�}�l�פJ
'    End If
    
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '�g�JRecordset
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
    str_SQL = "UPDATE Orders set ExternShipKey='" & RTrim(rsTmp("�Q�X�f��")) & "' " & _
         "where EXTERNOrderkey='" & RTrim(rsTmp("�q�渹�X")) & "' and type <> '�R��' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

   If RowsAffect > 0 Then intLine = intLine + 1
rsTmp.MoveNext
Loop

rsTmp.Close: Set rsTmp = Nothing
Screen.MousePointer = 0

MsgBox "��s " & intLine & "���U�B�渹!", 64, msg_title

Exit Sub
err_Handle:
Dim str As String
If MyXlsApp Is Nothing = False Then MyXlsApp.Quit: Set MyXlsApp = Nothing

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & Chr(64 + j) & k & ")�A��ƬO�_���~�I"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdPreView_Click()

Dim i As Integer, j As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "�L��ƥi�ѦC�L�I", vbOKOnly + vbInformation, "����C�L": Exit Sub
Screen.MousePointer = 11

'��Ƽg�J Access ��Ʈw
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From ��f�l�ܪ�", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "��f�l�ܪ�", cnAccess, adOpenStatic, adLockOptimistic

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
    
    '�g�JUSER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    .Reports(Me.Caption).[User_id].Caption = User_id
    .DoCmd.Close

    .DoCmd.OpenReport "��f�l�ܪ�", acViewPreview
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

If rsMain Is Nothing Then MsgBox "�L��ƥi�ѦC�L�I", vbOKOnly + vbInformation, "����C�L": Exit Sub
Screen.MousePointer = 11

'��Ƽg�J Access ��Ʈw
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From ��f�l�ܪ�", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "��f�l�ܪ�", cnAccess, adOpenStatic, adLockOptimistic

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
    
    '�g�JUSER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    .Reports(Me.Caption).[User_id].Caption = User_id
    .DoCmd.Close
    
    '�����C�L�ܦL���
    .Visible = False
    .DoCmd.OpenReport "��f�l�ܪ�", acViewNormal
    .CloseCurrentDatabase
    .Quit: Set MSAccessAP = Nothing

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd2Excel_Click()
'��ƱƧ�
Recordset2Excel Me.Caption, rsMain

'..�b���s��EXCEL
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
'If Len(RTrim(txtDeliveryS)) = 0 And Len(RTrim(txtDeliveryE)) = 0 Then MsgBox "�п�J����϶�!", 64, Me.Caption: Exit Sub
If Len(RTrim(txtDeliveryDateS)) = 0 And Len(RTrim(txtDeliveryDateS)) = 0 Then MsgBox "�п�J��f����϶�!", 64, Me.Caption: Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"

cn.Execute "if object_id ('tempdb..##" & strViewName & "') is not null drop table ##" & strViewName, RowsAffect, adExecuteNoRecords

'����J
str_SQL = "select ��O = RTrim(o.Priority),���A = '����J                ',�ϽX = rtrim(t1m.area_code),�G������ = '                  ',�G���r�p�H = '                  ',�G�����s = '          ',�@������ = '                  ',�@���r�p�H = '                  ',�@�����s = '          ',�q���� = o.orderdate,��f��� = o.DeliveryDate " & _
",ñ���� = '                    ',POD�Ѽ� = '          ',�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS�渹 = rtrim(o.orderkey),�q�渹�X = rtrim(o.externorderkey) " & _
",�Ȥ�W�� = rtrim(t1m.short_name),���� = rtrim(t2m.description),�q��O�� = round(sum(case when s.pallet = 0 then 0 else od.originalqty/s.pallet end),3),�q��c�� = sum(case when s.casecnt = 0 then 0 else floor(od.originalqty/s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then od.originalqty else cast(od.originalqty as int) % cast(s.casecnt as int) end),�`�Ӽ� = sum(od.originalqty),�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(od.originalqty /s.casecnt) end) " & _
",�J�w���� = '     ',�q�歫�q = sum(od.originalqty*s.stdgrosswgt),�q����n = sum(od.originalqty*s.stdcube),�q��Ƶ� = cast(o.notes as varchar(1000)),�w����f = '                    ' " & _
",�F�� = '     ',��� = ' ',�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) into ##" & strViewName & _
" from orders o (nolock) join orderdetail od (nolock) on o.orderkey = od.orderkey and o.b_phone2 is null and isnull(o.type,'') <> '�R��' " & _
"join gv_skuxpack s(nolock) on s.sku = od.sku and s.storerkey = o.storerkey join trp16m t16 on t16.storerkey = o.storerkey " & _
"left join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end and t1m.storerkey = o.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),o.deliverydate,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),o.priority,o.orderkey,o.externorderkey,o.DeliveryDate,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,cast(o.notes as varchar(1000)),t2m.description,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'����
str_SQL = "insert into ##" & strViewName & " select ��O = RTrim(w2.Priority),���A = '����' " & _
",�ϽX = rtrim(t1m.area_code),�G������ = '',�G���r�p�H = '',�G�����s = '          ',�@������ = '',�@���r�p�H = '',�@�����s = '          ',�q���� = o.orderdate,��f�� = w2.arrive_date " & _
",ñ���� = '',POD�Ѽ� = '',�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS�渹 = rtrim(w2.receipt_no) " & _
",�q�渹�X = rtrim(w2.extern),�Ȥ�W�� = rtrim(t1m.short_name),���� = rtrim(t2m.description),�q��O�� = round(sum(case when s.pallet = 0 then 0 else w3.order_qty/s.pallet end),3) " & _
",�q��c�� = sum(case when s.casecnt = 0 then 0 else floor(w3.order_qty/s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then w3.order_qty else cast(w3.order_qty as int) % cast(s.casecnt as int) end) " & _
",�`�Ӽ� = sum(w3.order_qty),�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(w3.order_qty /s.casecnt) end),�J�w���� = '' " & _
",�q�歫�q = sum(w3.order_qty*s.stdgrosswgt),�q����n = sum(w3.order_qty*s.stdcube),�q��Ƶ� = rtrim(w2.description),�w����f = '' " & _
",�F�� = ' ',��� = ' ',�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from trp03w w3 (nolock) join trp02w w2 (nolock) on w2.receipt_no = w3.receipt_no join orders o (nolock) on o.orderkey = w2.c_receipt_no " & _
"join gv_skuxpack s(nolock) on s.sku = w3.product_no and s.storerkey = w2.storerkey join trp16m t16 on t16.storerkey = w2.storerkey " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else w2.consigneekey end and t1m.storerkey = w2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),w2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),w2.priority,w2.receipt_no,w2.extern ,w2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,w2.DESCRIPTION,o.adddate,o.addwho"

'For i = 0 To List5.ListCount - 1
'    If List5.Selected(i) Then If List5.List(i) = "����" Then cn.Execute str_SQL, RowsAffect, adExecuteNoRecords: Exit For
'Next

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'����
str_SQL = "insert into ##" & strViewName & " select ��O = RTrim(w2.Priority),���A = '����',�ϽX = rtrim(t1m.area_code),�G������ = '',�G���r�p�H = '',�G�����s = '          ',�@������ = '',�@���r�p�H = '',�@�����s = '          ' " & _
",�q���� = o.orderdate,��f�� = w2.arrive_date,ñ���� = '',POD�Ѽ� = '',�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS�渹 = rtrim(w2.receipt_no) " & _
",�q�渹�X = rtrim(w2.extern),�Ȥ�W�� = rtrim(t1m.short_name),���� = rtrim(t2m.description),�q��O�� = round(sum(case when s.pallet = 0 then 0 else w3.order_qty/s.pallet end),3) " & _
",�q��c�� = sum(case when s.casecnt = 0 then 0 else floor(w3.order_qty/s.casecnt) end),�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then w3.order_qty else cast(w3.order_qty as int) % cast(s.casecnt as int) end) " & _
",�`�Ӽ� = sum(w3.order_qty),�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(w3.order_qty /s.casecnt) end) " & _
",�J�w���� = '',�q�歫�q = sum(w3.order_qty*s.stdgrosswgt),�q����n = sum(w3.order_qty*s.stdcube),�q��Ƶ� = rtrim(w2.description) " & _
",�w����f = '',�F�� = ' ',��� = ' ',�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from ort03w w3 (nolock) join ort02w w2 (nolock) on w2.receipt_no = w3.receipt_no join orders o (nolock) on o.orderkey = w2.c_receipt_no " & _
"join gv_skuxpack s(nolock) on s.sku = w3.product_no and s.storerkey = w2.storerkey join trp16m t16 on t16.storerkey = w2.storerkey " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else w2.consigneekey end and t1m.storerkey = w2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),w2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),w2.priority,w2.receipt_no,w2.extern ,w2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,w2.DESCRIPTION,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�O�d
str_SQL = "insert into ##" & strViewName & " select ��O = RTrim(t2.Priority),���A = '�O�d',�ϽX = rtrim(t1m.area_code),�G������ = '',�G���r�p�H = '',�G�����s = '          ',�@������ = '',�@���r�p�H = '',�@�����s = '          ' " & _
",�q���� = o.orderdate,��f�� = t2.arrive_date,ñ���� = '',POD�Ѽ� = '',�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS�渹 = rtrim(t2.receipt_no),�q�渹�X = rtrim(t2.extern),�Ȥ�W�� = rtrim(t1m.short_name),���� = rtrim(t2m.description) " & _
",�q��O�� = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3),�q��c�� = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),�`�Ӽ� = sum(t3.order_qty) " & _
",�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end),�J�w���� = '',�q�歫�q = sum(t3.order_qty*s.stdgrosswgt) " & _
",�q����n = sum(t3.order_qty*s.stdcube),�q��Ƶ� = rtrim(t2.description),�w����f = '',�F�� = ' ',��� = ' ' " & _
",�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from trp03t t3 (nolock) join trp02t t2 (nolock) on t2.receipt_no = t3.receipt_no and t2.route_no = 'D' " & _
"join orders o (nolock) on o.orderkey = t2.c_receipt_no join gv_skuxpack s on s.sku = t3.product_no and s.storerkey = t2.storerkey " & _
"join trp16m t16(nolock) on t16.storerkey = t2.storerkey join trp01m t1m on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end and t1m.storerkey = t2.storerkey " & _
"left join trp02m t2m(nolock) on t2m.zip = t1m.zip " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),t2.priority,t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,o.adddate,o.addwho"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�O�d
str_SQL = "insert into ##" & strViewName & " select ��O = RTrim(t2.Priority),���A = '�O�d',�ϽX = rtrim(t1m.area_code),�G������ = '',�G���r�p�H = '',�G�����s = '          ',�@������ = '',�@���r�p�H = '',�@�����s = '          ' " & _
",�q���� = o.orderdate,��f�� = t2.arrive_date,ñ���� = '',POD�Ѽ� = '',�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS�渹 = rtrim(t2.receipt_no),�q�渹�X = rtrim(t2.extern),�Ȥ�W�� = rtrim(t1m.short_name),���� = rtrim(t2m.description) " & _
",�q��O�� = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3),�q��c�� = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),�`�Ӽ� = sum(t3.order_qty) " & _
",�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end),�J�w���� = '',�q�歫�q = sum(t3.order_qty*s.stdgrosswgt) " & _
",�q����n = sum(t3.order_qty*s.stdcube),�q��Ƶ� = rtrim(t2.description),�w����f = '',�F�� = ' ',��� = ' ',�q��ӷ� = isnull(o.updatesource,'') " & _
",�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"from ort03t t3 (nolock) join ort02t t2 (nolock) on t2.receipt_no = t3.receipt_no and t2.route_no = 'D' join orders o (nolock) on o.orderkey = t2.c_receipt_no " & _
"join gv_skuxpack s(nolock) on s.sku = t3.product_no and s.storerkey = t2.storerkey join trp16m t16 on t16.storerkey = t2.storerkey " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end and t1m.storerkey = t2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),t2.priority,t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�w��
str_SQL = "insert into ##" & strViewName & " Select ��O = RTrim(t2.Priority),���A = '�w��',�ϽX = rtrim(t1m.area_code) " & _
",�G������ = isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No) ,�G���r�p�H = Rtrim(Isnull(t09m.Driver,'')),�G�����s = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),�@������ = rtrim(isnull((select top 1 t9.VEHICLE_ID_NO  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),�r�p�H = RTRIM(ISNULL((select top 1 t9.DRIVER  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),�@�����s = RTRIM(ISNULL(t2.route_no,'')),�q���� = o.orderdate " & _
",��f��� = t2.arrive_date,ñ���� = '',POD�Ѽ� = '',�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS�渹 = rtrim(t2.receipt_no),�q�渹�X = rtrim(t2.extern),�Ȥ�W�� = rtrim(t1m.short_name),���� = rtrim(t2m.description)" & _
",�q��O�� = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3),�q��c�� = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),�`�Ӽ� = sum(t3.order_qty) " & _
",�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end),�J�w���� = '',�q�歫�q = sum(t3.order_qty*s.stdgrosswgt) " & _
",�q����n = sum(t3.order_qty*s.stdcube),�q��Ƶ� = rtrim(t2.description),�w����f = isnull(convert(char(20),t2.scheduledate,120),''),�F�� = ' ',��� = ' ' " & _
",�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From TRP01T t01t  (nolock) join trp02t t2 (nolock) on t2.Route_No = t01t.Route_No and left(t01t.route_no,1) = 'F' " & _
"join trp03t t3 (nolock) on t3.receipt_no = t2.receipt_no join orders o (nolock) on o.orderkey = t2.c_receipt_no " & _
"join gv_skuxpack s (nolock) on s.storerkey = t2.storerkey and s.sku = t3.product_no " & _
"join trp01m t1m (nolock) on t1m.storerkey = t2.storerkey and t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end " & _
"join trp16m t16 (nolock) on t16.storerkey = t2.storerkey join TRP09M t09m on t09m.Vehicle_ID_No = isnull(t01t.C_Vehicle_ID_No,t2.Vehicle_ID_No) " & _
"join trp02m t2m (nolock) on t2m.zip = t1m.zip join TRP05T t05t (nolock) on t05t.Route_No = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO) and isnull(t05t.sdnstatus,'0') = '0' " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),isnull(convert(char(20),t2.scheduledate,120),''),t2.priority,isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No),t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,t09m.DRIVER,o.adddate,o.addwho,t2.VEHICLE_ID_NO,t2.ROUTE_NO "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�w��
str_SQL = "insert into ##" & strViewName & " select ��O = RTrim(t2.Priority),���A = '�w��',�ϽX = rtrim(t1m.area_code) " & _
",�G������ = isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No) ,�G���r�p�H = Rtrim(Isnull(t09m.Driver,'')),�G�����s = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),�@������ = rtrim(isnull((select top 1 t9.VEHICLE_ID_NO  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),�r�p�H = RTRIM(ISNULL((select top 1 t9.DRIVER  from trp09m t9 where t9.VEHICLE_ID_NO = t2.VEHICLE_ID_NO),'')),�@�����s = RTRIM(ISNULL(t2.route_no,'')),�q���� = o.orderdate,��f��� = t2.arrive_date,ñ���� = '',POD�Ѽ� = '' " & _
",�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS�渹 = rtrim(t2.receipt_no),�q�渹�X = rtrim(t2.extern),�Ȥ�W�� = rtrim(t1m.short_name) " & _
",���� = rtrim(t2m.description),�q��O�� = round(sum(case when s.pallet = 0 then 0 else t3.order_qty/s.pallet end),3) " & _
",�q��c�� = sum(case when s.casecnt = 0 then 0 else floor(t3.order_qty/s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then t3.order_qty else cast(t3.order_qty as int) % cast(s.casecnt as int) end),�`�Ӽ� = sum(t3.order_qty) " & _
",�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(t3.order_qty /s.casecnt) end) " & _
",�J�w���� = '',�q�歫�q = sum(t3.order_qty*s.stdgrosswgt),�q����n = sum(t3.order_qty*s.stdcube),�q��Ƶ� = rtrim(t2.description),�w����f = '',�F�� = ' ' " & _
",��� = ' ',�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From ort01T t01t (nolock) join ort02t t2 (nolock) on t2.Route_No = t01t.Route_No and left(t01t.route_no,1) = 'R' join ort03t t3 (nolock) on t3.receipt_no = t2.receipt_no " & _
"join orders o (nolock) on o.orderkey = t2.c_receipt_no join gv_skuxpack s on s.storerkey = t2.storerkey and s.sku = t3.product_no " & _
"join trp01m t1m (nolock) on t1m.storerkey = t2.storerkey and t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else t2.consigneekey end join trp16m t16 on t16.storerkey = t2.storerkey " & _
"join TRP09M t09m (nolock) on t09m.Vehicle_ID_No = isnull(t01t.C_Vehicle_ID_No,t2.Vehicle_ID_No) join trp02m t2m on t2m.zip = t1m.zip " & _
"join ort05T t05t (nolock) on t05t.Route_No = isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO) and isnull(t05t.sdnstatus,'0') = '0' " & _
"where convert(char(8),t2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),t2.priority,isnull(t01t.c_ROUTE_NO ,t01t.ROUTE_NO),isnull(t01t.c_Vehicle_ID_No,t2.Vehicle_ID_No),t2.receipt_no,t2.extern ,t2.arrive_date,t1m.area_code,t1m.short_name,t16.storerkey,t16.short_name,t2m.DESCRIPTION,t2.DESCRIPTION,t09m.DRIVER,o.adddate,o.addwho,t2.VEHICLE_ID_NO,t2.ROUTE_NO "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�ݭ���
str_SQL = "insert into ##" & strViewName & " select ��O = RTrim(o.Priority),���A = '�ݭ���',�ϽX = rtrim(t1m.area_code),�G������ = '',�G���r�p�H = '',�G�����s = '          ',�@������ = '',�@���r�p�H = '',�@�����s = '          ' " & _
",�q���� = o.orderdate,��f��� = cast(s2.arrive_date as datetime),ñ���� = '',POD�Ѽ� = '',�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name) " & _
",TMS�渹 = rtrim(s2.receipt_no),�q�渹�X = rtrim(s2.extern),�Ȥ�W�� = rtrim(t1m.short_name),���� = rtrim(t2m.description) " & _
",�q��O�� = round( sum(case when isnull(s.pallet,0) = 0 then 0 else s3.order_qty /s.pallet end) ,3),�q��c�� = sum(case when isnull(s.casecnt,0) = 0 then 0 else floor(s3.order_qty /s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then s3.order_qty else cast(s3.order_qty as int) % cast(s.casecnt as int) end),�`�Ӽ� = sum(s3.order_qty) " & _
",�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(s3.order_qty /s.casecnt) end),�J�w���� = '' " & _
",�q�歫�q = round( sum(s3.order_qty * s.stdgrosswgt),3),�q����n = round( sum( s3.order_qty * s.stdcube),3),�q��Ƶ� = cast(o.notes as varchar(1000)) " & _
",�w����f = '',�F�� = ' ',��� = ' ',�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From sdn02w s2 (nolock) join sdn03w s3 (nolock) on s3.receipt_no = s2.receipt_no join gv_skuxpack s on s.storerkey = s3.storerkey and s.sku = s3.product_no " & _
"join orders o (nolock) on o.orderkey = s2.c_receipt_no " & _
"join trp01m t1m(nolock) on t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end and t1m.storerkey = o.storerkey join trp16m t16 on t16.storerkey = o.storerkey " & _
"left join trp02m t2m(nolock) on t2m.zip = t1m.zip " & _
"where convert(char(8),s2.arrive_date,112) between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),o.orderdate,isnull(o.updatesource,''),o.priority,cast(o.notes as varchar(1000)),s2.receipt_no,t16.storerkey,t16.short_name,t1m.area_code,s2.arrive_date,s2.extern,t1m.short_name,t2m.DESCRIPTION,o.adddate,o.addwho "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�X��
str_SQL = "set nocount on insert into ##" & strViewName & " Select ��O = RTrim(s2.Priority),���A = '�X��' + '-' + case when s2.sdnback = 0 then 'ñ�楼�^' else rtrim(s2.confirm_notes) end " & _
",�ϽX = rtrim(t1m.area_code),�G������ = rtrim(s1.c_Vehicle_ID_No),�G���r�p�H = Rtrim(Isnull(s1.Driver,'')),�G�����s = rtrim(s1.c_route_No) ,�@������ = rtrim(isnull(s2.VEHICLE_ID_NO,'')),�r�p�H = RTRIM(ISNULL(t9.driver,'')),�@�����s = RTRIM(ISNULL(s2.route_no,'')) " & _
",�q���� = o.orderdate,��f��� = cast(s2.arrive_date as datetime),ñ���� = isnull(convert(char(10),s2.sdnsenddate,20),'') " & _
",POD�Ѽ� = case when rtrim(s2.confirm_notes) = '' then '' when s2.sdnsenddate is null then '' else cast(datediff(dd,cast(s2.arrive_date as datetime),s2.sdnsenddate+1) as varchar(4)) end " & _
",�f�D = rtrim(t16.storerkey) + '_' + rtrim(t16.short_name),TMS�渹 = rtrim(s2.receipt_no),�q�渹�X = rtrim(s2.extern),�Ȥ�W�� = rtrim(t1m.short_name) " & _
",���� = rtrim(t2m.description),�q��O�� = round( sum(case when isnull(s.pallet,0) = 0 then 0 else s3.order_qty /s.pallet end) ,3),�q��c�� = sum(case when isnull(s.casecnt,0) = 0 then 0 else floor(s3.order_qty /s.casecnt) end) " & _
",�q��Ӽ� = sum(case when isnull(s.casecnt,0) = 0 then s3.order_qty else cast(s3.order_qty as int) % cast(s.casecnt as int) end),�`�Ӽ� = sum(s3.order_qty),�w����� = sum(case when s.casecnt = 0 then 1 else ceiling(s3.order_qty /s.casecnt) end) " & _
",�J�w���� = isnull(s2.invback,''),�q�歫�q = round( sum(s3.order_qty * s.stdgrosswgt),3),�q����n = round( sum(s3.order_qty * s.stdcube),3),�q��Ƶ� = rtrim(s2.description) " & _
",�w����f = isnull(convert(char(20),isnull(s2.scheduledate,s2.custsigndate),120),''),�F�� = case when s2.ontimedelivery = 9 then 'V' else ' ' end,��� = case when s2.ontimedelivery = 5 then 'V' else ' ' end " & _
",�q��ӷ� = isnull(o.updatesource,''),�q��s�W�ɶ� = o.adddate,�q��s�W�H�� = rtrim(o.addwho),�U�B�渹 = rtrim(isnull(o.ExternShipkey,'')) " & _
"From sdn02t s2 (nolock) join sdn01T s1(nolock) on s1.c_route_no = s2.c_route_no " & _
"join sdn03t s3 (nolock) on s3.receipt_no = s2.receipt_no  join orders o (nolock) on o.orderkey = s2.c_receipt_no join gv_skuxpack s on s.storerkey = s3.storerkey and s.sku = s3.product_no join trp01m t1m on t1m.consigneekey =case when o.priority ='A2B' then o.b_company else s2.consigneekey end and t1m.storerkey = s2.storerkey " & _
"join trp09m t9 (nolock) on t9.Vehicle_ID_No = s2.Vehicle_ID_No left join trp16m t16 on t16.storerkey = s2.storerkey left join trp02m t2m on t2m.zip = t1m.zip " & _
"where s2.arrive_date between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' " & _
"group by isnull(o.ExternShipkey,''),s2.sdnback,o.orderdate,isnull(o.updatesource,''),isnull(convert(char(20),isnull(s2.scheduledate,s2.custsigndate),120),''),s2.confirm_notes,s2.OnTimeDelivery,s2.PRIORITY,t2m.DESCRIPTION,s2.receipt_no,t16.storerkey,t16.short_name,t1m.area_code,s1.c_Vehicle_ID_No,s1.Driver, s1.c_route_No,s2.arrive_date,s2.extern,t1m.short_name,s2.description,isnull(s2.invback,''),s2.sdnsenddate,o.adddate,o.addwho,rtrim(isnull(s2.VEHICLE_ID_NO,'')),RTRIM(ISNULL(t9.driver,'')),RTRIM(ISNULL(s2.route_no,'')) set nocount off"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "select * from ##" & strViewName & " Where 1 = 1 "
'str_SQL = "select * from gv_TRPTrack Where 1 = 1 "

'���f�D
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) Then strSelected = strSelected & "'" & List2.List(i) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and �f�D in ( " & strSelected & "'') "

'���ϽX
strSelected = ""
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then strSelected = strSelected & "'" & Trim(List1.List(i)) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and �ϽX in ( " & strSelected & "'') "

'���F��
strSelected = ""
For i = 0 To List3.ListCount - 1
    If List3.Selected(i) And Trim(List3.List(i)) = "���F" Then strSelected = strSelected & "(rtrim(���) = '' and rtrim(�F��) = '') or "
    If List3.Selected(i) And Trim(List3.List(i)) = "���" Then strSelected = strSelected & "rtrim(���) = 'V' or "
    If List3.Selected(i) And Trim(List3.List(i)) = "�F��" Then strSelected = strSelected & "rtrim(�F��) = 'V' or "
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and (" & strSelected & " 1 = 0) "

'����O
strSelected = ""
For i = 0 To List4.ListCount - 1
    If List4.Selected(i) Then strSelected = strSelected & "'" & mySplit(Trim(List4.List(i)), "_", 0) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and ��O in ( " & strSelected & "'') "

'�����A
strSelected = ""
For i = 0 To List5.ListCount - 1
    If List5.Selected(i) Then strSelected = strSelected & "'" & Trim(List5.List(i)) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & "and ���A in ( " & strSelected & "'') "

'���u�s��
chc_Route = ""
If Len(txtRouteS.Text) > 0 And Len(txtRouteE.Text) > 0 Then
   chc_Route = "and �G�����s between '" & txtRouteS.Text & "' and '" & txtRouteE.Text & "' "
ElseIf Len(txtRouteS.Text) > 0 And Len(txtRouteE.Text) = 0 Then
   chc_Route = "and �G�����s = '" & txtRouteS.Text & "' "
ElseIf Len(txtRouteS.Text) = 0 And Len(txtRouteE.Text) > 0 Then
   chc_Route = "and �G�����s = '" & txtRouteE.Text & "' "
End If

'��f���
chc_DeliveryDate = ""
chc_DeliveryDate1 = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(char(8),��f���,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
   chc_DeliveryDate1 = "and convert(char(8),o.deliverydate,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_DeliveryDate = "and convert(char(8),��f���,112) = '" & txtDeliveryDateS.Text & "' "
   chc_DeliveryDate1 = "and convert(char(8),o.deliverydate,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(char(8),��f���,112) = '" & txtDeliveryDateE.Text & "' "
   chc_DeliveryDate1 = "and convert(char(8),o.deliverydate,112) = '" & txtDeliveryDateE.Text & "' "
End If

'�X�����
chc_DeliveryDate2 = ""
If Len(txtDeliveryS.Text) > 0 And Len(txtDeliveryE.Text) > 0 Then
   chc_DeliveryDate2 = "and '20' + substring(�G�����s,2,6) between '" & txtDeliveryS.Text & "' and '" & txtDeliveryE.Text & "' "
ElseIf Len(txtDeliveryS.Text) > 0 And Len(txtDeliveryE.Text) = 0 Then
   chc_DeliveryDate2 = "and '20' + substring(�G�����s,2,6) = '" & txtDeliveryS.Text & "' "
ElseIf Len(txtDeliveryS.Text) = 0 And Len(txtDeliveryE.Text) > 0 Then
   chc_DeliveryDate2 = "and '20' + substring(�G�����s,2,6) = '" & txtDeliveryE.Text & "' "
End If

'�զX�r��
str_SQL = str_SQL & chc_Route & chc_DeliveryDate & chc_DeliveryDate2 & "order by ��f���,�ϽX,�G������,�G�����s,�f�D,�q�渹�X "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
cn.CommandTimeout = 600
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close

rsMain.MoveFirst

If chkShowWH = 1 Then

    '���t�m���
    Dim rsTmp As New ADODB.Recordset
    '    str_SQL = "select distinct sectionkey ,o.updatesource from " & strWMSDB & "..orders o join " & strWMSDB & "..pickdetail p on p.orderkey = o.orderkey and convert(char(8),o.deliverydate,112) > convert(char(8),getdate()-7,112) join " & strWMSDB & "..loc l on l.loc = p.loc order by sectionkey "
    str_SQL = "select distinct sectionkey ,o.updatesource from " & strWMSDB & "..orders o join " & strWMSDB & "..pickdetail p on p.orderkey = o.orderkey " & chc_DeliveryDate1 & " join " & strWMSDB & "..loc l on l.loc = p.loc order by sectionkey "
    tmp_Rs.Open str_SQL, cn
    Call Replication_Recordset(tmp_Rs, rsTmp)
    tmp_Rs.Close
    
    Do While Not rsMain.EOF
    
        If rsMain("��f���") > Format(Now() - 8, "yyyymmdd") Then
    
            rsTmp.Filter = "(updatesource = '" & rsMain("TMS�渹") & "')"
        
            strSectionKey = ""
        
            If rsTmp.EOF Then
                rsMain("�˸��I") = "���t�m"
            Else
                Do While Not rsTmp.EOF
                    If UCase(RTrim(rsTmp("sectionkey"))) <> "FACILITY" Then strSectionKey = strSectionKey & RTrim(rsTmp("sectionkey")) & ";"
                    rsTmp.MoveNext
                Loop
                rsMain("�˸��I") = strSectionKey
            End If
        
            If rsMain("��O") = "R" Or rsMain("��O") = "RC" Or rsMain("��O") = "A2B" Then rsMain("�˸��I") = ""
        
            rsTmp.Filter = ""
        End If
        rsMain.MoveNext
    Loop
    rsTmp.Close
    
End If

Set dgMain.DataSource = rsMain

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True
cn.Execute "if object_id ('tempdb..##" & strViewName & "') is not null drop table ##" & strViewName, RowsAffect, adExecuteNoRecords

Exit Sub
err_Handle:
cn.Execute "if object_id ('tempdb..##" & strViewName & "') is not null drop table ##" & strViewName, RowsAffect, adExecuteNoRecords
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub



Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
If dg.Col > 0 Then If rsMain.Fields(dg.Col).Name = "�w����f" Then dg.Columns(ColIndex).Width = dtpDeliveryTime.Width

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
dtpDeliveryTime.Visible = False
If dgMain.DataSource Is Nothing Then Exit Sub
If rsMain.RecordCount = 0 Then Exit Sub
If rsMain.EOF Then Exit Sub
If dgMain.Col = -1 Then Exit Sub
If Left(rsMain("���A"), 2) <> "�X��" Then Exit Sub

With dgMain

'��f�ɶ�
If rsMain.Fields(.Col).Name = "�w����f" Then

    dtpDeliveryTime.Visible = True
    dtpDeliveryTime.Move .Left + .Columns(.Col).Left + Frame2.Left + 15, .Top + .RowTop(.Row) + Frame2.Top, .Columns(.Col).Width
    
    If dtpDeliveryTime.Left + dtpDeliveryTime.Width - Frame2.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
        dtpDeliveryTime.Width = dtpDeliveryTime.Width + .Left + .Width - dtpDeliveryTime.Left - dtpDeliveryTime.Width
    End If
    dtpDeliveryTime.Value = IIf(RTrim(rsMain("�w����f")) = "", Now, rsMain("�w����f"))

Else
    dtpDeliveryTime.Visible = False
End If

'�F��
If rsMain.Fields(.Col).Name = "�F��" Then
    If Trim(rsMain("�F��")) = "" And rsMain("���") = "V" Then
        If MsgBox("�F��T�{?", vbOKCancel, "���A�ܧ�") = vbOK Then rsMain("�F��") = "V": rsMain("���") = " ": cn.Execute "update sdn02t set ontimedelivery = 9 where receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
    ElseIf Trim(rsMain("�F��")) = "V" Then
        If MsgBox("�F������T�{?", vbOKCancel, "���A�ܧ�") = vbOK Then rsMain("�F��") = " ": cn.Execute "update sdn02t set ontimedelivery = 0 where receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
    Else
        If MsgBox("�F��T�{?", vbOKCancel, "���A�ܧ�") = vbOK Then rsMain("�F��") = "V": cn.Execute "update sdn02t set ontimedelivery = 9 where receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
    End If
    .Col = 18
End If

'���
If rsMain.Fields(.Col).Name = "���" Then
    If Trim(rsMain("���")) = "" And rsMain("�F��") = "V" Then
        If MsgBox("���T�{?", vbOKCancel, "���A�ܧ�") = vbOK Then rsMain("���") = "V": rsMain("�F��") = " ": cn.Execute "update sdn02t set ontimedelivery = 5 where receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
    ElseIf Trim(rsMain("���")) = "V" Then
        If MsgBox("�������T�{?", vbOKCancel, "���A�ܧ�") = vbOK Then rsMain("���") = " ": cn.Execute "update sdn02t set ontimedelivery = 0 where receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
    Else
        If MsgBox("���T�{?", vbOKCancel, "���A�ܧ�") = vbOK Then rsMain("���") = "V": cn.Execute "update sdn02t set ontimedelivery = 5 where receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
    End If
    .Col = 18
End If

End With

End Sub

Private Sub dtpDeliveryTime_LostFocus()
     If MsgBox("�w����f�ɶ��ܧ�?", vbOKCancel, "�T�{") = vbOK Then
        rsMain("�w����f") = Format(dtpDeliveryTime, "yyyy-mm-dd HH:MM")
        cn.Execute "update sdn02t set scheduledate = '" & rsMain("�w����f") & "' where receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
        dtpDeliveryTime.Visible = False
     End If
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

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

'���]
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
Unload Me '�������{��
'End �������ε{��
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

'�f�D
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
'tmp_rs.Open "select distinct(�f�D)  from gv_TRPTrack order by �f�D ", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.Open "select distinct rtrim(storerkey) + '_' + rtrim(short_name) as �f�D from trp16m order by rtrim(storerkey) + '_' + rtrim(short_name)", cn, adOpenKeyset, adLockPessimistic


If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        List2.AddItem RTrim(tmp_Rs("�f�D"))
        tmp_Rs.MoveNext
    Next
    tmp_Rs.Close: Set tmp_Rs = Nothing
End If
    
'�ϰ�
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

'�F��
List3.AddItem "���F"
List3.AddItem "�F��"
List3.AddItem "���"

'��O
List4.AddItem "C_�V�w�q��"
List4.AddItem "I_���`�q��"
List4.AddItem "A_���"
List4.AddItem "R_�h�f�q��"
List4.AddItem "RC_���f�J�w"
List4.AddItem "A2B_���f�t�e"

'���A
List5.AddItem "����J" ': List5.Selected(0) = True
List5.AddItem "����" ': List5.Selected(1) = True
List5.AddItem "�O�d" ': List5.Selected(2) = True
List5.AddItem "�w��" ': List5.Selected(3) = True
List5.AddItem "�ݭ���" ': List5.Selected(4) = True
List5.AddItem "�X��-ñ�楼�^" ': List5.Selected(5) = True
List5.AddItem "�X��-���`�q��"
List5.AddItem "�X��-���`�q��"
List5.AddItem "�X��-���X�q��"

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
