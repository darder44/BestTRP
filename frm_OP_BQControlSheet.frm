VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_OP_BQControlSheet 
   Caption         =   "BQ�ި��"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15555
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
   ScaleHeight     =   9420
   ScaleWidth      =   15555
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   4410
      _ExtentX        =   7779
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
      StartOfWeek     =   61276161
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.TextBox txtFlash1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      MaxLength       =   8
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtFlash 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.ComboBox cboCustomer1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_BQControlSheet.frx":0000
      Left            =   2640
      List            =   "frm_OP_BQControlSheet.frx":0002
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "cboCustomer"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboUserType1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_BQControlSheet.frx":0004
      Left            =   3840
      List            =   "frm_OP_BQControlSheet.frx":0006
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "cboUserType2"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboCustomer 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_BQControlSheet.frx":0008
      Left            =   2640
      List            =   "frm_OP_BQControlSheet.frx":000A
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "cboCustomer"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboUserType 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_BQControlSheet.frx":000C
      Left            =   3840
      List            =   "frm_OP_BQControlSheet.frx":000E
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "cboUserType"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   860
      Left            =   360
      Picture         =   "frm_OP_BQControlSheet.frx":0010
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   1060
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   9150
      Width           =   15555
      _ExtentX        =   27437
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
            Object.Width           =   20823
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
   Begin TabDlg.SSTab SSTab 
      Height          =   9135
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16113
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "�ި��"
      TabPicture(0)   =   "frm_OP_BQControlSheet.frx":0322
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   " �ХI�ڸ�Ʃ���"
      TabPicture(1)   =   "frm_OP_BQControlSheet.frx":033E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frm_OP_BQControlSheet.frx":035A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
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
         Height          =   1215
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtOrderDateS 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            MaxLength       =   8
            TabIndex        =   61
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtOrderDateE 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   60
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            MaxLength       =   20
            TabIndex        =   59
            Top             =   600
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            MaxLength       =   20
            TabIndex        =   58
            Top             =   600
            Width           =   1365
         End
         Begin VB.CommandButton cmdPickSave 
            BackColor       =   &H00FFFF80&
            Caption         =   "�s��"
            Enabled         =   0   'False
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
            Picture         =   "frm_OP_BQControlSheet.frx":0376
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   4080
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
            Height          =   880
            Left            =   3840
            Picture         =   "frm_OP_BQControlSheet.frx":0680
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   56
            Top             =   240
            Width           =   1060
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
            Left            =   2055
            TabIndex        =   65
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���"
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
            TabIndex        =   64
            Top             =   285
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�渹"
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
            TabIndex        =   63
            Top             =   645
            Width           =   480
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
            Index           =   5
            Left            =   2055
            TabIndex        =   62
            Top             =   660
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "�̪O��"
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
         TabIndex        =   53
         Top             =   1680
         Width           =   5055
         Begin MSDataGridLib.DataGrid dgRoute 
            Height          =   2295
            Left            =   120
            TabIndex        =   54
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
      Begin VB.Frame Frame4 
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
         Height          =   3375
         Left            =   5280
         TabIndex        =   44
         Top             =   360
         Width           =   9855
         Begin VB.TextBox Text22 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3705
            MaxLength       =   8
            TabIndex        =   97
            Top             =   2880
            Width           =   765
         End
         Begin VB.TextBox Text21 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4545
            MaxLength       =   8
            TabIndex        =   96
            Top             =   2880
            Width           =   765
         End
         Begin VB.TextBox Text20 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3705
            MaxLength       =   8
            TabIndex        =   94
            Top             =   2520
            Width           =   765
         End
         Begin VB.TextBox Text19 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4545
            MaxLength       =   8
            TabIndex        =   93
            Top             =   2520
            Width           =   765
         End
         Begin VB.TextBox Text18 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6465
            MaxLength       =   8
            TabIndex        =   91
            Top             =   2520
            Width           =   765
         End
         Begin VB.TextBox Text17 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7305
            MaxLength       =   8
            TabIndex        =   90
            Top             =   2520
            Width           =   765
         End
         Begin VB.TextBox Text16 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   825
            MaxLength       =   8
            TabIndex        =   88
            Top             =   2880
            Width           =   765
         End
         Begin VB.TextBox Text15 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1665
            MaxLength       =   8
            TabIndex        =   87
            Top             =   2880
            Width           =   765
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   825
            MaxLength       =   8
            TabIndex        =   85
            Top             =   2520
            Width           =   765
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1665
            MaxLength       =   8
            TabIndex        =   84
            Top             =   2520
            Width           =   765
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   79
            Top             =   1680
            Width           =   765
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   78
            Top             =   1320
            Width           =   765
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   77
            Top             =   960
            Width           =   765
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   76
            Top             =   600
            Width           =   765
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   74
            Top             =   1680
            Width           =   1245
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   73
            Top             =   1320
            Width           =   1245
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   72
            Top             =   960
            Width           =   1245
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   71
            Top             =   600
            Width           =   1245
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
            Height          =   880
            Index           =   0
            Left            =   8640
            Picture         =   "frm_OP_BQControlSheet.frx":098A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   69
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00FFC0FF&
            Caption         =   "�R��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   880
            Left            =   7440
            Picture         =   "frm_OP_BQControlSheet.frx":2A59C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   68
            Top             =   1200
            Width           =   1060
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�ק�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   880
            Left            =   8640
            Picture         =   "frm_OP_BQControlSheet.frx":2B5DE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   67
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdAddNew 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�s�W"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   880
            Left            =   7440
            Picture         =   "frm_OP_BQControlSheet.frx":31E30
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   66
            Top             =   240
            Width           =   1060
         End
         Begin VB.TextBox txtPalletKey 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            MaxLength       =   20
            TabIndex        =   48
            Top             =   960
            Width           =   1725
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            MaxLength       =   8
            TabIndex        =   47
            Top             =   600
            Width           =   1725
         End
         Begin VB.TextBox txtDriver 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            MaxLength       =   8
            TabIndex        =   46
            Top             =   1680
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.ComboBox cboCarno 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   720
            TabIndex        =   45
            Top             =   1320
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ɥX"
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
            Index           =   29
            Left            =   6480
            TabIndex        =   104
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�٤J"
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
            Index           =   28
            Left            =   7320
            TabIndex        =   103
            Top             =   2280
            Width           =   480
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   5760
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�٤J"
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
            Index           =   27
            Left            =   4560
            TabIndex        =   102
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ɥX"
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
            Index           =   26
            Left            =   3720
            TabIndex        =   101
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�٤J"
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
            Index           =   25
            Left            =   1680
            TabIndex        =   100
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ɥX"
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
            Index           =   24
            Left            =   840
            TabIndex        =   99
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���y�c"
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
            Index           =   23
            Left            =   2925
            TabIndex        =   98
            Top             =   2940
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�j�̪O"
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
            Index           =   22
            Left            =   2925
            TabIndex        =   95
            Top             =   2580
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "Ţ��"
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
            Index           =   21
            Left            =   5805
            TabIndex        =   92
            Top             =   2580
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ŪO"
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
            Left            =   165
            TabIndex        =   89
            Top             =   2940
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���O"
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
            Index           =   19
            Left            =   165
            TabIndex        =   86
            Top             =   2580
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "XD��"
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
            Index           =   18
            Left            =   2895
            TabIndex        =   83
            Top             =   1740
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "901��"
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
            Left            =   2880
            TabIndex        =   82
            Top             =   1020
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "81��"
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
            Left            =   3000
            TabIndex        =   81
            Top             =   660
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "1001��"
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
            Left            =   2760
            TabIndex        =   80
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�O��"
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
            Left            =   4920
            TabIndex        =   75
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n��"
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
            Left            =   3600
            TabIndex        =   70
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�渹"
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
            Left            =   120
            TabIndex        =   52
            Top             =   1005
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����"
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
            Left            =   120
            TabIndex        =   51
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���"
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
            Left            =   120
            TabIndex        =   50
            Top             =   645
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�r�p"
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
            TabIndex        =   49
            Top             =   1725
            Visible         =   0   'False
            Width           =   480
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
         Height          =   3405
         Left            =   5280
         TabIndex        =   40
         Top             =   6840
         Width           =   4095
         Begin VB.CommandButton cmdAddSortingCost 
            BackColor       =   &H0000FFFF&
            Caption         =   "�s�W"
            Height          =   375
            Left            =   1320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdDeleteSortingCost 
            BackColor       =   &H00FFC0FF&
            Caption         =   "�R��"
            Height          =   375
            Left            =   2280
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid dgSortingCost 
            Height          =   2175
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            TabAction       =   1
            AllowDelete     =   -1  'True
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
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���`����"
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
            Index           =   31
            Left            =   120
            TabIndex        =   106
            Top             =   300
            Width           =   960
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
         Height          =   3135
         Left            =   5280
         TabIndex        =   36
         Top             =   3720
         Width           =   4095
         Begin VB.CommandButton cmdAddPalletDetail 
            BackColor       =   &H0000FFFF&
            Caption         =   "�s�W"
            Height          =   375
            Left            =   1440
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   38
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdDeletePalletDetail 
            BackColor       =   &H00FFC0FF&
            Caption         =   "�R��"
            Height          =   375
            Left            =   2400
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   37
            Top             =   240
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid dgPalletDetail 
            Height          =   1935
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            TabAction       =   1
            AllowDelete     =   -1  'True
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
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�h�f�P�ռ�"
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
            Index           =   30
            Left            =   120
            TabIndex        =   105
            Top             =   300
            Width           =   1200
         End
      End
      Begin VB.Frame Frame9 
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
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMain 
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
      Begin VB.Frame Frame8 
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
         TabIndex        =   23
         Top             =   360
         Width           =   8295
         Begin VB.CommandButton Command2 
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
            Picture         =   "frm_OP_BQControlSheet.frx":33CA2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   30
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
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
            Picture         =   "frm_OP_BQControlSheet.frx":33FAC
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   29
            Top             =   240
            Width           =   1065
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
            Picture         =   "frm_OP_BQControlSheet.frx":342BE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   28
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox Text4 
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
            TabIndex        =   27
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            TabIndex        =   26
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
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
            Picture         =   "frm_OP_BQControlSheet.frx":355B8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   25
            Top             =   1200
            Width           =   1065
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
            Picture         =   "frm_OP_BQControlSheet.frx":5F1CA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
            Top             =   1200
            Width           =   1065
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
            Index           =   12
            Left            =   2655
            TabIndex        =   33
            Top             =   660
            Visible         =   0   'False
            Width           =   360
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
            Index           =   11
            Left            =   120
            TabIndex        =   32
            Top             =   645
            Visible         =   0   'False
            Width           =   960
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
            Index           =   8
            Left            =   1440
            TabIndex        =   31
            Top             =   960
            Width           =   2880
         End
      End
      Begin VB.Frame Frame7 
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
         TabIndex        =   21
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT1 
            Height          =   2295
            Left            =   120
            TabIndex        =   22
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
         TabIndex        =   11
         Top             =   360
         Width           =   8295
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
            TabIndex        =   18
            Top             =   960
            Width           =   1485
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
            TabIndex        =   17
            Top             =   960
            Width           =   1485
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
            Picture         =   "frm_OP_BQControlSheet.frx":5F4D4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
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
            Picture         =   "frm_OP_BQControlSheet.frx":5F7DE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   15
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
            Picture         =   "frm_OP_BQControlSheet.frx":60AD8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   14
            Top             =   240
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
            Picture         =   "frm_OP_BQControlSheet.frx":60DEA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   13
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
            Picture         =   "frm_OP_BQControlSheet.frx":610F4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   12
            Top             =   1200
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
            Index           =   6
            Left            =   2640
            TabIndex        =   20
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
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   1005
            Width           =   960
         End
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2220
         Left            =   -67680
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5040
         Visible         =   0   'False
         Width           =   4410
         _ExtentX        =   7779
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
         StartOfWeek     =   61276161
         TitleBackColor  =   -2147483646
         TitleForeColor  =   16777215
         TrailingForeColor=   -2147483643
         CurrentDate     =   38233
         MaxDate         =   2958455
      End
   End
End
Attribute VB_Name = "frm_OP_BQControlSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsRoute As ADODB.Recordset
Private rsPalletDetail As ADODB.Recordset
Private rsSortingCost As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
'Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cboCustomer_Change()
Call cboCustomer_Click
End Sub
Private Sub cboCustomer1_Change()
Call cboCustomer1_Click
End Sub

Private Sub cboUserType_Change()
Call cboUserType_Click
End Sub
'
'Private Sub cboUserType1_Change()
'Call cboUserType1_Click
'End Sub

Private Sub cmdAddnew_Click()

'����ˬd
If Not myIsDate(txtDate) Then Call txtdate_Click: Exit Sub
If Len(RTrim(txtPalletKey)) = 0 Then MsgBox "�п�J�渹!!", vbOKOnly, Me.Caption: txtPalletKey.SetFocus: Exit Sub
If Len(RTrim(cboCarno)) = 0 Then MsgBox "�п�J����!!", vbOKOnly, Me.Caption: cboCarno.SetFocus: Exit Sub
'If rsPalletDetail.RecordCount + rsSortingCost.RecordCount = 0 Then MsgBox "�п�J���Ӹ��!!", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle
dgPalletDetail.Col = 0: dgSortingCost.Col = 0
Dim rsTmp As New ADODB.Recordset

'�渹�ˬd
rsTmp.Open "select checkno from pallet_cds where checkno = '" & RTrim(txtPalletKey) & "' ", cn
If Not rsTmp.EOF Then MsgBox "�t�γ渹����!(" & RTrim(txtPalletKey) & ")", 64, "�s�W����!": rsTmp.Close: Exit Sub
rsTmp.Close

'�����ˬd
rsTmp.Open "select driver = isnull(driver,'') from trp09m where vehicle_id_no = '" & RTrim(cboCarno) & "' ", cn
If rsTmp.EOF Then MsgBox "�t�εL������!(" & RTrim(cboCarno) & ")", 64, "�s�W����!": rsTmp.Close: Exit Sub
'�Ȧs���
Dim strDriver As String: strDriver = rsTmp("driver")
rsTmp.Close

cn.BeginTrans: Tran_Level = 1

'�ˬd�X���T�{�ᨮ���O�_�ۦP
rsTmp.Open "select carno = rtrim(c_vehicle_id_no) from sdn01t where c_route_no = '" & RTrim(txtPalletKey) & "' ", cn

If Not rsTmp.EOF Then '�������s
    If rsTmp("carno") <> RTrim(cboCarno) Then '��������
        If MsgBox("�̪O�渹�P���u�s�� (" & txtPalletKey & ") �A�X���T�{��������!" & vbCrLf & "�O�_�P�B��s�X���T�{�����H", vbOKCancel, "�̪O��s�W") = vbOK Then cn.Execute "update sdn01t set c_vehicle_id_no = '" & RTrim(cboCarno) & "',driver = '" & strDriver & "',editdate = getdate() , edituser = '" & User_id & "' where c_route_no = '" & RTrim(txtPalletKey) & "' ", RowsAffect, adExecuteNoRecords
    End If
End If

'�g�J���Y���
str_SQL = "insert into pallet_cds(checkno,storer,carno,usertype,adddate,adduser,edituser,keyindate,editdate) " & _
    "values('" & RTrim(txtPalletKey) & "','BEST','" & UCase(RTrim(cboCarno)) & "','','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & RTrim(txtDate) & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�R����
str_SQL = "delete pallet_cst where checkno = '" & RTrim(txtPalletKey) & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�g�J�̪O���
If rsPalletDetail.RecordCount > 0 Then
    rsPalletDetail.MoveFirst
   
    Do While Not rsPalletDetail.EOF
        If Len(RTrim(rsPalletDetail("���O"))) = 0 Or Len(RTrim(rsPalletDetail("�Ȥ�"))) = 0 Then MsgBox "�п�J�̪O���O�ΫȤ�?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
        If Val(rsPalletDetail("�ɥX")) = 0 And Val(rsPalletDetail("�٤J")) = 0 Then MsgBox "�ɥX�P�٤J�ƶq���o���� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub

        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & RTrim(txtPalletKey) & "','" & RTrim(rsPalletDetail("����")) & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsPalletDetail("���O") & "','" & rsPalletDetail("�Ȥ�") & "','" & rsPalletDetail("�Ȥ�渹") & "','" & RTrim(txtDate) & "','" & Val(rsPalletDetail("�ɥX")) & "','" & Val(rsPalletDetail("�٤J")) & "',0,'" & rsPalletDetail("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rsPalletDetail.MoveNext
    Loop
End If
    
 '�g�J�z�f���
If rsSortingCost.RecordCount > 0 Then
    rsSortingCost.MoveFirst
    
    Do While Not rsSortingCost.EOF
        If Len(RTrim(rsSortingCost("���O"))) = 0 Then MsgBox "�п�J���O?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
        If Val(rsSortingCost("�p�O�ƶq")) = 0 Then MsgBox "�p�O�ƶq���o�� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
           
        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & RTrim(txtPalletKey) & "','" & rsSortingCost("����") & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsSortingCost("���O") & "','" & rsSortingCost("�Ȥ�") & "','" & rsSortingCost("�Ȥ�渹") & "','" & RTrim(txtDate) & "',0,0,'" & Val(rsSortingCost("�p�O�ƶq")) & "','" & rsSortingCost("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rsSortingCost.MoveNext
    Loop
End If

cn.CommitTrans: Tran_Level = 0
MsgBox "�s�W����!", 0, RTrim(txtPalletKey)

'�Ȧs���
Dim strPalletKey As String, strDate As String, strCarno As String
strPalletKey = RTrim(txtPalletKey)
strDate = RTrim(txtDate)
strCarno = RTrim(cboCarno)

rsRoute.Find "�渹 = '" & RTrim(strPalletKey) & "'"
If rsRoute.EOF Then rsRoute.AddNew

rsRoute("���") = RTrim(strDate)
rsRoute("���@") = "V"
rsRoute("�渹") = RTrim(strPalletKey)
rsRoute("����") = RTrim(strCarno)
rsRoute("����") = User_id
rsRoute("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")

If rsPalletDetail.RecordCount = 0 Then rsRoute("���@") = "X"
    
Call dgRoute_RowColChange(dgRoute.Row, dgRoute.Col)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub cmdAddPalletDetail_Click()
If rsPalletDetail Is Nothing Then Exit Sub

'������
Dim i As Integer, j As Integer
If rsPalletDetail.RecordCount > 0 Then rsPalletDetail.MoveLast: i = rsPalletDetail("����")
If rsSortingCost.RecordCount > 0 Then rsSortingCost.MoveLast: j = rsSortingCost("����")

'�s�W
rsPalletDetail.AddNew

If i > j Then
    rsPalletDetail("����") = i + 1
Else
    rsPalletDetail("����") = j + 1
End If

rsPalletDetail("����") = User_id
rsPalletDetail("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")

End Sub

Private Sub cmdAddSortingCost_Click()
If rsSortingCost Is Nothing Then Exit Sub

'������
Dim i As Integer, j As Integer
If rsPalletDetail.RecordCount > 0 Then rsPalletDetail.MoveLast: i = rsPalletDetail("����")
If rsSortingCost.RecordCount > 0 Then rsSortingCost.MoveLast: j = rsSortingCost("����")

'�s�W
rsSortingCost.AddNew

If i > j Then
    rsSortingCost("����") = i + 1
Else
    rsSortingCost("����") = j + 1
End If

rsSortingCost("����") = User_id
rsSortingCost("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")

End Sub

Private Sub cmdDelete_Click()
On Error GoTo err_Handle

If rsRoute Is Nothing Then Exit Sub
If rsRoute.RecordCount = 0 Then Exit Sub
If Len(Trim(rsRoute("���@"))) = 0 Then Exit Sub
If MsgBox("�渹�G" & Trim(txtPalletKey) & " �T�w�R���H", vbOKCancel, Me.Caption) <> vbOK Then Exit Sub

cn.BeginTrans: Tran_Level = 1

    '�R�����Y
    str_SQL = "delete pallet_cds where checkno = '" & Trim(txtPalletKey) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�R����
    str_SQL = "delete pallet_cst where checkno = '" & Trim(txtPalletKey) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

Call cmdQueryDetail_Click

rsRoute("���@") = ""
rsRoute("����") = ""
rsRoute("���ʤ��") = ""

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdDeletePalletDetail_Click()

If dgPalletDetail.DataSource Is Nothing Then Exit Sub
If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub
rsPalletDetail.Delete

End Sub

Private Sub cmdDeleteSortingCost_Click()

If dgSortingCost.DataSource Is Nothing Then Exit Sub
If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
rsSortingCost.Delete

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgRoute.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Set dgPalletDetail.DataSource = Nothing: Set dgSortingCost.DataSource = Nothing
txtDate = "": cboCarno = "": txtDriver = "": txtPalletKey = ""
Dim chc_PalletNo As String, chc_DeliveryDate As String, chc_Storerkey As String

str_SQL = "Select ��� = IsNull(Convert(Char(8), p.adddate, 112), C.YMD) " & _
            ",���@ = case when p.checkno is not null then 'V' Else '' end " & _
            ",�Ȥ� = isnull(isnull(rtrim(p.storer),t1m.consigneekey),'') " & _
            ",�渹 = rtrim(isnull(p.checkno,'')) " & _
            ",���� = Rtrim(isnull(p.carno,'')) " & _
            ",���� = isnull(p.edituser,'') " & _
            ",���ʤ�� = isnull(convert(char(20),p.editdate,20),'') " & _
            "from trp01m t1m join calender c on t1m.storerkey = 'LTRI02' and c.ymd > = '20080801' " & _
            "full join pallet_cds p on p.storer = t1m.consigneekey and Convert(char(8),p.adddate,112) = c.ymd where 1 = 1 "

'���
chc_DeliveryDate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_DeliveryDate = "and IsNull(Convert(Char(8), p.adddate, 112), C.YMD) between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_DeliveryDate = "and IsNull(Convert(Char(8), p.adddate, 112), C.YMD) = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_DeliveryDate = "and IsNull(Convert(Char(8), p.adddate, 112), C.YMD) = '" & txtOrderDateE.Text & "' "
End If

'�渹
chc_PalletNo = ""
If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 Then
   chc_PalletNo = "and rtrim(isnull(p.checkno,'')) between '" & Text1.Text & "' and '" & Text2.Text & "' "
ElseIf Len(Text1.Text) > 0 And Len(Text2.Text) = 0 Then
   chc_PalletNo = "and rtrim(isnull(p.checkno,'')) = '" & Text1.Text & "' "
ElseIf Len(Text1.Text) = 0 And Len(Text2.Text) > 0 Then
   chc_PalletNo = "and rtrim(isnull(p.checkno,'')) = '" & Text2.Text & "' "
End If

str_SQL = str_SQL & chc_DeliveryDate & chc_PalletNo & "order by isnull(Convert(char(8),p.adddate,112),c.YMD) "

Dim rsTmp As New ADODB.Recordset
rsTmp.Open str_SQL, cn, adOpenStatic, adLockPessimistic
If rsTmp.EOF = True Then MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Call cmdQueryDetail_Click

Set rsRoute = New ADODB.Recordset
rsRoute.CursorLocation = adUseClient

Call OffLineRecordset(rsTmp, rsRoute)
rsTmp.Close: Set rsTmp = Nothing

Set dgRoute.DataSource = rsRoute: dgRoute.Visible = False
If rsRoute.EOF = False Then rsRoute.MoveFirst

Set dgRoute.DataSource = rsRoute

SetDataGridColWidth Me.Caption, dgRoute
StatusBar.Panels(2).Text = rsRoute.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgRoute.Visible = True

Call dgRoute_RowColChange(1, 1)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdQueryDetail_Click()
On Error GoTo err_Handle
dgPalletDetail.Visible = False: dgSortingCost.Visible = False
Screen.MousePointer = 11

'�̪O����
str_SQL = "select ���� =linenumber " & _
            ",���O = usertype " & _
            ",�ɥX = qtyin " & _
            ",�٤J = qtyout " & _
            ",�Ȥ�渹 = isnull(customersheetno,'') " & _
            ",���ӳƵ� = isnull(notes,'') " & _
            ",���� = edituser " & _
            ",���ʤ�� = editdate " & _
            "From Pallet_cst where linenumber > 0 and checkno = '" & RTrim(txtPalletKey) & "' order by linenumber "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic

Set rsPalletDetail = New ADODB.Recordset: rsPalletDetail.CursorLocation = 3

Call Replication_Recordset(tmp_Rs, rsPalletDetail)
tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgPalletDetail.DataSource = rsPalletDetail
SetDataGridColWidth Me.Caption, dgPalletDetail

'�z�f����
str_SQL = "select " & _
            "���� " & _
            ",���O " & _
            ",�Ȥ� " & _
            ",�p�O�ƶq " & _
            ",�Ȥ�渹 " & _
            ",���ӳƵ� " & _
            ",���� = ���Ӳ��� " & _
            ",���ʤ�� = ���Ӳ��ʤ�� " & _
            "From gv_PalletDetail where ���� > 0 and �渹 = '" & RTrim(txtPalletKey) & "' and ���O in ('½�O��','�z�f��','�K��','�\��') order by ���� "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic

Set rsSortingCost = New ADODB.Recordset: rsSortingCost.CursorLocation = 3

Call Replication_Recordset(tmp_Rs, rsSortingCost)
tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgSortingCost.DataSource = rsSortingCost
SetDataGridColWidth Me.Caption, dgSortingCost
dgPalletDetail.Columns.item(0).Visible = False: dgSortingCost.Columns.item(0).Visible = False
cboCustomer.Visible = False: cboCustomer1.Visible = False
cboUserType.Visible = False: cboUserType1.Visible = False
txtFlash.Visible = False: txtFlash1.Visible = False
Screen.MousePointer = 0: dgPalletDetail.Visible = True: dgSortingCost.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdExit_Click(Index As Integer)
Unload Me '�������{��
'End �������ε{��
End Sub

Private Sub ShowUserType()

With dgPalletDetail
    .RowHeight = cboUserType.Height - 10
    If .Col = 2 Then
        If .Columns(.Col).Left > 0 Then
                cboUserType.Visible = True
                cboUserType.Move .Left + .Columns(.Col).Left + Frame2.Left + 15, .Top + .RowTop(.Row) + Frame2.Top, .Columns(.Col).Width
                If cboUserType.Left + cboUserType.Width - Frame2.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboUserType.Width = cboUserType.Width + .Left + .Width - cboUserType.Left - cboUserType.Width
                End If
                cboUserType.Text = rsPalletDetail("���O")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboUserType.Visible = False
        End If
    Else
        cboUserType.Visible = False
    End If
    
End With
End Sub

Private Sub cboUserType_Click()
If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub
rsPalletDetail("���O") = cboUserType.Text

End Sub
Private Sub ShowUserType1()

With dgSortingCost
    .RowHeight = cboUserType.Height - 10
    If .Col = 2 Then
        If .Columns(.Col).Left > 0 Then
                cboUserType1.Visible = True
                cboUserType1.Move .Left + .Columns(.Col).Left + Frame6.Left + 15, .Top + .RowTop(.Row) + Frame6.Top, .Columns(.Col).Width
                '�p�G���W�XDataGrid����ܽd�򪺳B�z
                If cboUserType1.Left + cboUserType1.Width - Frame2.Left > .Left + .Width Then
                    cboUserType1.Width = cboUserType1.Width + .Left + .Width - cboUserType1.Left - cboUserType1.Width
                End If
                cboUserType1.Text = rsSortingCost("���O")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboUserType1.Visible = False
        End If
    Else
        cboUserType1.Visible = False
    End If
End With
End Sub
Private Sub cboUserType1_Click()
If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
rsSortingCost("���O") = cboUserType1.Text

End Sub

Private Sub ShowCustomer()

With dgPalletDetail
    .RowHeight = cboUserType.Height - 10
    If .Col = 3 Then
        If .Columns(.Col).Left > 0 Then
                cboCustomer.Visible = True
                cboCustomer.Move .Left + .Columns(.Col).Left + Frame2.Left + 15, .Top + .RowTop(.Row) + Frame2.Top, .Columns(.Col).Width
                If cboCustomer.Left + cboCustomer.Width - Frame2.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboCustomer.Width = cboCustomer.Width + .Left + .Width - cboCustomer.Left - cboCustomer.Width
                End If
                cboCustomer.Text = rsPalletDetail("�Ȥ�")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboCustomer.Visible = False
        End If
    Else
        cboCustomer.Visible = False
    End If
End With
End Sub

Private Sub cboCustomer_Click()

If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub
rsPalletDetail("�Ȥ�") = cboCustomer.Text

End Sub

Private Sub ShowCustomer1()

With dgSortingCost
    .RowHeight = cboUserType.Height - 10
    If .Col = 3 Then
        If .Columns(.Col).Left > 0 Then
                cboCustomer1.Visible = True
                cboCustomer1.Move .Left + .Columns(.Col).Left + Frame6.Left + 15, .Top + .RowTop(.Row) + Frame6.Top, .Columns(.Col).Width
                If cboCustomer1.Left + cboCustomer1.Width - Frame6.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboCustomer1.Width = cboCustomer1.Width + .Left + .Width - cboCustomer1.Left - cboCustomer1.Width
                End If
                cboCustomer1.Text = rsSortingCost("�Ȥ�")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboCustomer1.Visible = False
        End If
    Else
        cboCustomer1.Visible = False
    End If
End With
End Sub

Private Sub cboCustomer1_Click()

If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
rsSortingCost("�Ȥ�") = cboCustomer1.Text

End Sub

Private Sub ShowText1()

With dgSortingCost
.RowHeight = txtFlash1.Height - 10
    If .Columns(.Col).Left > 0 Then
            txtFlash1.Visible = True
            txtFlash1.Move .Left + .Columns(.Col).Left + Frame6.Left + 15, .Top + .RowTop(.Row) + Frame6.Top, .Columns(.Col).Width
            If txtFlash1.Left + txtFlash1.Width - Frame6.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                txtFlash1.Width = txtFlash1.Width + .Left + .Width - txtFlash1.Left - txtFlash.Width
            End If
            txtFlash1.Text = rsSortingCost.Fields(.Col)  '��stxt����
            txtFlash1.SetFocus
    Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
        txtFlash1.Visible = False
    End If

End With
End Sub

Private Sub cmdEdit_Click()

'����ˬd
If Not myIsDate(txtDate) Then Call txtdate_Click: Exit Sub
If Len(RTrim(txtPalletKey)) = 0 Then MsgBox "�п�J�渹!!", vbOKOnly, Me.Caption: txtPalletKey.SetFocus: Exit Sub
If Len(RTrim(cboCarno)) = 0 Then MsgBox "�п�J����!!", vbOKOnly, Me.Caption: cboCarno.SetFocus: Exit Sub
'If rsPalletDetail.RecordCount + rsSortingCost.RecordCount = 0 Then MsgBox "�п�J���Ӹ��!!", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle
dgPalletDetail.Col = 0: dgSortingCost.Col = 0
Dim rsTmp As New ADODB.Recordset

'�渹�ˬd
rsTmp.Open "select checkno from pallet_cds where checkno = '" & RTrim(txtPalletKey) & "' ", cn
If rsTmp.EOF Then MsgBox "�t�εL���渹!(" & RTrim(txtPalletKey) & ")", 64, "��s����!": rsTmp.Close: Exit Sub

'�����ˬd
rsTmp.Close
rsTmp.Open "select driver = isnull(driver,'') from trp09m where vehicle_id_no = '" & RTrim(cboCarno) & "' ", cn
If rsTmp.EOF Then MsgBox "�t�εL������!(" & RTrim(cboCarno) & ")", 64, "��s����!": rsTmp.Close: Exit Sub

'�Ȧs���
Dim strDriver As String: strDriver = rsTmp("driver")

cn.BeginTrans: Tran_Level = 1

'�ˬd�X���T�{�ᨮ���O�_�ۦP
rsTmp.Close
rsTmp.Open "select carno = rtrim(c_vehicle_id_no) from sdn01t where c_route_no = '" & RTrim(txtPalletKey) & "' ", cn

If Not rsTmp.EOF Then '�������s
    If rsTmp("carno") <> RTrim(cboCarno) Then '��������
        If MsgBox("�̪O�渹�P���u�s�� (" & txtPalletKey & ") �A�X���T�{��������!" & vbCrLf & "�O�_�P�B��s�X���T�{�����H", vbOKCancel, "�̪O���s") = vbOK Then cn.Execute "update sdn01t set c_vehicle_id_no = '" & RTrim(cboCarno) & "',driver = '" & strDriver & "',editdate = getdate() , edituser = '" & User_id & "' where c_route_no = '" & RTrim(txtPalletKey) & "' ", RowsAffect, adExecuteNoRecords
    End If
End If

'��s���Y
    str_SQL = "update pallet_cds set " & _
              "carno = '" & UCase(RTrim(cboCarno)) & "' " & _
              ",adddate = '" & RTrim(txtDate) & "' " & _
              ",edituser = '" & User_id & "' " & _
              ",editdate = '" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "' " & _
              "where checkno = '" & RTrim(txtPalletKey) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�R����
str_SQL = "delete pallet_cst where checkno = '" & RTrim(txtPalletKey) & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�g�J�̪O���
If rsPalletDetail.RecordCount > 0 Then
    rsPalletDetail.MoveFirst
   
    Do While Not rsPalletDetail.EOF
        If Len(RTrim(rsPalletDetail("���O"))) = 0 Or Len(RTrim(rsPalletDetail("�Ȥ�"))) = 0 Then MsgBox "�п�J�̪O���O�ΫȤ�?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
        If Val(rsPalletDetail("�ɥX")) = 0 And Val(rsPalletDetail("�٤J")) = 0 Then MsgBox "�ɥX�P�٤J�ƶq���o���� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub

        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & RTrim(txtPalletKey) & "','" & RTrim(rsPalletDetail("����")) & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsPalletDetail("���O") & "','" & rsPalletDetail("�Ȥ�") & "','" & rsPalletDetail("�Ȥ�渹") & "','" & RTrim(txtDate) & "','" & Val(rsPalletDetail("�ɥX")) & "','" & Val(rsPalletDetail("�٤J")) & "',0,'" & rsPalletDetail("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rsPalletDetail.MoveNext
    Loop
End If
    
 '�g�J�z�f���
If rsSortingCost.RecordCount > 0 Then
    rsSortingCost.MoveFirst
    
    Do While Not rsSortingCost.EOF
        If Len(RTrim(rsSortingCost("���O"))) = 0 Then MsgBox "�п�J���O?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
        If Val(rsSortingCost("�p�O�ƶq")) = 0 Then MsgBox "�p�O�ƶq���o�� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
           
        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & RTrim(txtPalletKey) & "','" & rsSortingCost("����") & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsSortingCost("���O") & "','" & rsSortingCost("�Ȥ�") & "','" & rsSortingCost("�Ȥ�渹") & "','" & RTrim(txtDate) & "',0,0,'" & Val(rsSortingCost("�p�O�ƶq")) & "','" & rsSortingCost("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rsSortingCost.MoveNext
    Loop
End If

cn.CommitTrans: Tran_Level = 0
MsgBox "��s����!", 0, RTrim(txtPalletKey)

    '��s
    rsRoute("���") = RTrim(txtDate)
    rsRoute("���@") = "V"
    rsRoute("�渹") = RTrim(txtPalletKey)
    rsRoute("����") = RTrim(cboCarno)
    rsRoute("����") = User_id
    rsRoute("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")
    
If rsPalletDetail.RecordCount = 0 Then rsRoute("���@") = "X"

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgPalletDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub

With dgPalletDetail
    '�����\���ܯS�w���
    If .Col < 2 Or .Col > 7 Then .Col = Abs(LastCol): Exit Sub
    cboCustomer.Visible = False: cboCustomer1.Visible = False
    cboUserType.Visible = False: cboUserType1.Visible = False
    txtFlash.Visible = False: txtFlash1.Visible = False
    
    '���O
    If .Col = 2 Then
        ShowUserType
    '�Ȥ�
    ElseIf .Col = 3 Then
        ShowCustomer
    '��L
    Else
'        ShowText
    End If

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgPalletDetail_Scroll(Cancel As Integer)
If cboUserType.Visible = True Then ShowUserType
If cboCustomer.Visible = True Then ShowCustomer
End Sub

Private Sub dgSortingCost_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
cboCustomer.Visible = False: cboCustomer1.Visible = False
cboUserType.Visible = False: cboUserType1.Visible = False
txtFlash.Visible = False: txtFlash1.Visible = False
        
With dgSortingCost
    '�����\���ܯS�w���
    If .Col < 2 Or .Col > 6 Then .Col = Abs(LastCol): Exit Sub

    '���O
    If .Col = 2 Then
        ShowUserType1
    '�Ȥ�
    ElseIf .Col = 3 Then
        ShowCustomer1
    '��L
    Else
'        ShowText1
'        txtFlash1.SelStart = 0: txtFlash1.SelLength = Len(txtFlash1.Text)
'        txtFlash1.SetFocus
'        DoEvents: DoEvents
    End If

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgroute_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgRoute
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgPalletDetail_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgPalletDetail
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgSortingCost_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgSortingCost
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgRoute_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'�P�@����
If LastRow = Empty Then Exit Sub

'�O�_�����
If rsRoute Is Nothing Then Exit Sub
If rsRoute.RecordCount = 0 Then Exit Sub
If rsRoute.EOF Then Exit Sub

txtDate = rsRoute("���")
txtPalletKey = rsRoute("�渹"): Frame4.Caption = rsRoute("�渹")
cboCarno = rsRoute("����")
'txtDriver = rsRoute("�r�p")
Call cmdQueryDetail_Click

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgSortingCost_Scroll(Cancel As Integer)
If cboUserType1.Visible = True Then ShowUserType1
If cboCustomer1.Visible = True Then ShowCustomer1
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    SSTab.Height = Me.ScaleHeight - StatusBar.Height
    Frame6.Width = Frame2.Width
    Frame3.Height = SSTab.Height - Frame1.Height - Frame1.Top - StatusBar.Height + 60
    dgRoute.Height = Frame3.Height - 360
    dgPalletDetail.Height = Frame2.Height - dgPalletDetail.Top - 120
    dgSortingCost.Height = Frame6.Height - dgSortingCost.Top - 120
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab.Width = Me.ScaleWidth
    Frame2.Width = SSTab.Width - Frame3.Width - 360: Frame6.Width = Frame2.Width
    dgPalletDetail.Width = Frame2.Width - 240: dgSortingCost.Width = Frame6.Width - 240
    dgRoute.Width = Frame3.Width - 240
    
End If

End Sub

Private Sub cmdReset_Click()

'���]
Call ClearForm_AllField(Me)
Call cmdQueryDetail_Click
'Combo1.ListIndex = 0

End Sub

Private Sub dgroute_HeadClick(ByVal ColIndex As Integer)

If dgRoute.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsRoute.Sort = dgRoute.Columns(ColIndex).Caption & " DESC"
    dgRoute.ClearSelCols
    intColumnIndex = 255

Else
    rsRoute.Sort = dgRoute.Columns(ColIndex).Caption
    dgRoute.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgPalletDetail_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

'�̪O���O
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(���O) from gv_palletdetail order by ���O", cn, adOpenKeyset, adLockPessimistic
If Not tmp_Rs.EOF Then
tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboUserType.AddItem tmp_Rs("���O")
        tmp_Rs.MoveNext
    Next
    cboUserType.ListIndex = 0
End If
tmp_Rs.Close

''�z�f���O
'For i = 1 To 4
'cboUserType1.AddItem Choose(i, "½�O��", "�z�f��", "�K��", "�\��")
'Next
'cboUserType1.ListIndex = 0

'�Ȥ�
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(�Ȥ�) from gv_palletdetail order by �Ȥ� ", cn, adOpenKeyset, adLockPessimistic
If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboCustomer.AddItem tmp_Rs("�Ȥ�")
        cboCustomer1.AddItem tmp_Rs("�Ȥ�")
        tmp_Rs.MoveNext
    Next
    cboCustomer.ListIndex = 0: cboCustomer1.ListIndex = 0
End If
tmp_Rs.Close

'����
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(����) from gv_palletdetail order by ���� ", cn, adOpenKeyset, adLockPessimistic
If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboCarno.AddItem tmp_Rs("����")
        tmp_Rs.MoveNext
    Next
    cboCarno.ListIndex = -1
End If
tmp_Rs.Close

txtOrderDateS = Format(Now - 3, "YYYYMMDD")
txtOrderDateE = Format(Now, "YYYYMMDD")
Set tmp_Rs = Nothing

Call cmdQuery_Click
Call cmdQueryDetail_Click
    
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsRoute = Nothing
Set rsPalletDetail = Nothing
Set rsSortingCost = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtdate_Click()

Set objMvdateTarget = txtDate
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width + Frame4.Left, objMvdateTarget.Top + objMvdateTarget.Height + Frame4.Top
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtFlash1_Change()

If rsSortingCost Is Nothing Then Exit Sub
rsSortingCost.Fields(dgSortingCost.Col) = txtFlash1.Text

End Sub

Private Sub txtOrderDateE_Click()

Set objMvdateTarget = txtOrderDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateS_Click()

Set objMvdateTarget = txtOrderDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
