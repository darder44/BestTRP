VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Report_ABT 
   Caption         =   "ABT�ݨD����"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13665
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
         Name            =   "�s�ө���"
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
      TabCaption(0)   =   " �^���ˮ֪�"
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
      TabCaption(3)   =   "�N���f��"
      TabPicture(3)   =   "frm_Report_ABT.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "�q����Ӫ�"
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
         TabIndex        =   121
         Top             =   660
         Width           =   13695
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
            Left            =   10080
            Picture         =   "frm_Report_ABT.frx":00E0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   139
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
            Left            =   12480
            Picture         =   "frm_Report_ABT.frx":03EA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   138
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
            Index           =   0
            Left            =   12480
            Picture         =   "frm_Report_ABT.frx":06FC
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   137
            Top             =   1200
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
            Left            =   11280
            Picture         =   "frm_Report_ABT.frx":2A30E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   136
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtOrderDateS 
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
            TabIndex        =   135
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateE 
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
            TabIndex        =   134
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   133
            Top             =   240
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
            TabIndex        =   132
            Top             =   960
            Visible         =   0   'False
            Width           =   1485
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
            TabIndex        =   131
            Top             =   960
            Width           =   1485
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
            Left            =   10080
            Picture         =   "frm_Report_ABT.frx":2B608
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   130
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
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
            Left            =   4680
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   129
            ToolTipText     =   "�ϽX"
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmdSaveToText 
            BackColor       =   &H00C0E0FF&
            Caption         =   "�ˮ֪�"
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
            Left            =   11280
            Picture         =   "frm_Report_ABT.frx":2B912
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   128
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CheckBox optNormal 
            Caption         =   "���`ñ��"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox optAbnormal 
            Caption         =   "���`ñ��"
            Height          =   255
            Left            =   1200
            TabIndex        =   126
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox optNotYet 
            Caption         =   "���T�{ñ��"
            Height          =   255
            Left            =   2280
            TabIndex        =   125
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "�ө���"
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
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   124
            Top             =   1680
            Visible         =   0   'False
            Width           =   2325
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
            ItemData        =   "frm_Report_ABT.frx":2BC4E
            Left            =   6480
            List            =   "frm_Report_ABT.frx":2BC50
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   123
            ToolTipText     =   "�f�B���q"
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ListBox List3 
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
            ItemData        =   "frm_Report_ABT.frx":2BC52
            Left            =   9000
            List            =   "frm_Report_ABT.frx":2BC54
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   122
            ToolTipText     =   "�q�����O"
            Top             =   240
            Visible         =   0   'False
            Width           =   975
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
            Index           =   33
            Left            =   2655
            TabIndex        =   146
            Top             =   660
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���@���"
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
            Index           =   32
            Left            =   120
            TabIndex        =   145
            Top             =   645
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ϰ�"
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
            Left            =   360
            TabIndex        =   144
            Top             =   300
            Width           =   480
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
            Index           =   30
            Left            =   120
            TabIndex        =   143
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
            Left            =   2640
            TabIndex        =   142
            Top             =   1020
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ݧ@���X���T�{"
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
            Left            =   2880
            TabIndex        =   141
            Top             =   240
            Width           =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ƨ�"
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
         TabIndex        =   80
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
            TabIndex        =   95
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            Picture         =   "frm_Report_ABT.frx":2BC56
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   85
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
            Picture         =   "frm_Report_ABT.frx":2BF60
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   84
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
            Picture         =   "frm_Report_ABT.frx":2D25A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   83
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
            Picture         =   "frm_Report_ABT.frx":2D564
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   82
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
            Picture         =   "frm_Report_ABT.frx":57176
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   81
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
            TabIndex        =   89
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
            TabIndex        =   88
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
         TabIndex        =   66
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
            TabIndex        =   96
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
            Picture         =   "frm_Report_ABT.frx":57488
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   73
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
            Picture         =   "frm_Report_ABT.frx":5779A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   72
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
            Picture         =   "frm_Report_ABT.frx":813AC
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   71
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
            Picture         =   "frm_Report_ABT.frx":816B6
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   70
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
            Picture         =   "frm_Report_ABT.frx":829B0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   98
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
            TabIndex        =   75
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
            TabIndex        =   74
            Top             =   1020
            Width           =   360
         End
      End
      Begin VB.Frame Frame11 
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
         TabIndex        =   56
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
            TabIndex        =   97
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            Picture         =   "frm_Report_ABT.frx":82CBA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   61
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
            Picture         =   "frm_Report_ABT.frx":82FC4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   60
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
            Picture         =   "frm_Report_ABT.frx":842BE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   59
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
            Picture         =   "frm_Report_ABT.frx":845C8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   58
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
            Picture         =   "frm_Report_ABT.frx":AE1DA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   57
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabCaption(0)   =   "��L���Ӫ�"
            TabPicture(0)   =   "frm_Report_ABT.frx":AE4EC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "dgMainT4"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "���@���Ӫ�"
            TabPicture(1)   =   "frm_Report_ABT.frx":AE508
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "dgMainT4_1"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "���d���Ӫ�"
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
      Begin VB.Frame Frame9 
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
         TabIndex        =   43
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtNotCarNo 
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
            TabIndex        =   151
            Top             =   600
            Width           =   3165
         End
         Begin VB.OptionButton optAll 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Caption         =   "���T�{"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Caption         =   "�w�T�{"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
               Name            =   "�s�ө���"
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
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Picture         =   "frm_Report_ABT.frx":AE540
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   50
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
            Picture         =   "frm_Report_ABT.frx":AE852
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   49
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
            Picture         =   "frm_Report_ABT.frx":D8464
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   48
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
            Picture         =   "frm_Report_ABT.frx":D876E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   47
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
            Left            =   5640
            Picture         =   "frm_Report_ABT.frx":D9A68
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   46
            Top             =   960
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtAddDateET4 
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
         Begin VB.TextBox txtAddDateST4 
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
            TabIndex        =   44
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�A"
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
            Index           =   35
            Left            =   360
            TabIndex        =   153
            Top             =   660
            Width           =   8280
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ư�����"
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
            Index           =   34
            Left            =   120
            TabIndex        =   152
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��ƽT�{"
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
            TabIndex        =   116
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�t�e����"
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
            Left            =   120
            TabIndex        =   112
            Top             =   300
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
            Index           =   25
            Left            =   5175
            TabIndex        =   110
            Top             =   2100
            Visible         =   0   'False
            Width           =   360
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
            Index           =   24
            Left            =   2640
            TabIndex        =   109
            Top             =   2085
            Visible         =   0   'False
            Width           =   960
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
            Index           =   23
            Left            =   120
            TabIndex        =   108
            Top             =   1725
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
            Index           =   22
            Left            =   2640
            TabIndex        =   107
            Top             =   1740
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
            Index           =   21
            Left            =   2640
            TabIndex        =   106
            Top             =   1380
            Width           =   360
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
            Index           =   19
            Left            =   120
            TabIndex        =   105
            Top             =   1365
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "������"
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
            TabIndex        =   52
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
            TabIndex        =   51
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
         TabIndex        =   28
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtOrderDateST3 
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Alignment       =   2  '�m�����
            BeginProperty Font 
               Name            =   "�ө���"
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
            Top             =   1320
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
            TabIndex        =   34
            Top             =   1320
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
            Picture         =   "frm_Report_ABT.frx":D9D72
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   33
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
            Picture         =   "frm_Report_ABT.frx":DA07C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   32
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
            Picture         =   "frm_Report_ABT.frx":DB376
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   31
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
            Picture         =   "frm_Report_ABT.frx":DB680
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   30
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
            Picture         =   "frm_Report_ABT.frx":105292
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   29
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
            Index           =   29
            Left            =   2640
            TabIndex        =   120
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
            Index           =   28
            Left            =   120
            TabIndex        =   119
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
            Index           =   9
            Left            =   2640
            TabIndex        =   37
            Top             =   1380
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q����"
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
            TabIndex        =   36
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
         Caption         =   "�t�e���`��"
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
         TabIndex        =   15
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
            TabIndex        =   91
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
            TabIndex        =   90
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
            Picture         =   "frm_Report_ABT.frx":1055A4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   26
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            Picture         =   "frm_Report_ABT.frx":1058B6
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   19
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
            Picture         =   "frm_Report_ABT.frx":105BC0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   18
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
            Picture         =   "frm_Report_ABT.frx":106EBA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   17
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
            Picture         =   "frm_Report_ABT.frx":1071C4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   16
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
            TabIndex        =   93
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
            TabIndex        =   92
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   1005
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000004&
         Caption         =   "�ХI�ڸ�Ʃ���"
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
         TabIndex        =   6
         Top             =   660
         Width           =   8295
         Begin VB.CheckBox optRepack 
            Caption         =   "�[�u�p�O"
            Height          =   255
            Left            =   3360
            TabIndex        =   94
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox optTMS 
            Caption         =   "�B��д�"
            Height          =   255
            Left            =   1200
            TabIndex        =   41
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox optWMS 
            Caption         =   "���x�д�"
            Height          =   255
            Left            =   2280
            TabIndex        =   40
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
            Picture         =   "frm_Report_ABT.frx":130DD6
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
            Index           =   2
            Left            =   7080
            Picture         =   "frm_Report_ABT.frx":1310E8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   14
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
            Picture         =   "frm_Report_ABT.frx":15ACFA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   11
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
            Picture         =   "frm_Report_ABT.frx":15B004
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   10
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
            Picture         =   "frm_Report_ABT.frx":15C2FE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            Text            =   "���A"
            TextSave        =   "���A"
            Object.ToolTipText     =   "���A"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   17489
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

SSTab.Tab = 0

'������
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select distinct vehicle_id_no from trp05t order by vehicle_id_no "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then tmp_Rs.MoveFirst

Do While Not tmp_Rs.EOF
    cboCarT4.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    tmp_Rs.MoveNext
Loop
cboCarT4.AddItem "�ݱƨ�"

cboCarT4 = ""

tmp_Rs.Close

'�f�D
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

Combo1.AddItem "�H�t"
Combo1.AddItem "����"
Combo1.AddItem "�w��"
Combo1.ListIndex = 0
    
''�ϰ�
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
''�f�B���q
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
''��O
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

'���]
txtDeliveryDateS = "": txtDeliveryDateE = ""

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
txtOrderDateST3 = "": txtOrderDateET3 = ""
txtDeliveryDateST3 = "": txtDeliveryDateET3 = ""
End Sub
Private Sub cmdResetT4_Click()
'���]
cboCarT4 = ""
txtAddDateST4 = "": txtAddDateET4 = ""
txtDeliveryDateST4 = "": txtDeliveryDateET4 = ""
txtDeliveryST4 = "": txtDeliveryET4 = ""
txtRouteST4 = "": txtRouteET4 = ""
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
Recordset2Excel "LABT�^���ˮ֪�", rsMain
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
    Call OffLineRecordset(rsMainT1, rsTmp)
    
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
    
    Call WriteOut_RunLog("�B��дڡG3/5.��X�t�e�O...")
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
    
    Call WriteOut_RunLog("�B��дڡG5/5.��X�t�e�O...")
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
    '�i�f
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
    
        Call WriteOut_RunLog("�[�u�p�O�G1/3.��XNPP�[�u�p�O")
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
    
    Call WriteOut_RunLog("�[�u�p�O�G2/3.��X�DNPP�[�u�p�O")
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

'�@��[�u�p�O
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
    
    Call WriteOut_RunLog("�[�u�p�O�G3/3.��X�[�u�p�O����")
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
Recordset2Excel "LABT01�N���f��", rsMainT3
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2ExcelT4_Click()

'��ƱƧ�
Screen.MousePointer = 11
dgMainT4.Visible = False
dgMainT4_1.Visible = False
dgMainT4_2.Visible = False
Recordset2Excel_ABT "��L�q�����", rsMainT4
Recordset2Excel_ABT "���@�q�����", rsMainT4_1
Recordset2Excel_ABT "���d�q�����", rsMainT4_2
dgMainT4.Visible = True
dgMainT4_1.Visible = True
dgMainT4_2.Visible = True
'..�b���s��EXCEL
Screen.MousePointer = 0
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
'Dim chc_Orderdate As String, chc_DeliveryDate As String, i As Integer, strSelected As String
Dim strSelected As String
strSelected = ""


''�ϽX
'For i = 0 To List1.ListCount - 1
'    If List1.Selected(i) Then strSelected = strSelected & "'" & Left(List1.List(i), 2) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t1m.area_code in ( " & strSelected & "'') "
'
''�f�B���q
'strSelected = ""
'For i = 0 To List2.ListCount - 1
'    If List2.Selected(i) Then strSelected = strSelected & "'" & mySplit(List2.List(i), "_", 0) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t8m.company_code in ( " & strSelected & "'') "
'
''��O
'strSelected = ""
'For i = 0 To List3.ListCount - 1
'    If List3.Selected(i) Then strSelected = strSelected & "'" & mySplit(List3.List(i), "_", 0) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and isnull(s2.priority,'') in ( " & strSelected & "'') "
'
''���@���
'chc_Orderdate = ""
'If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
'   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
'ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
'   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateS.Text & "' "
'ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
'   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateE.Text & "' "
'End If
'
''��f���
'chc_DeliveryDate = ""
'If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
'   chc_DeliveryDate = "and convert(Char(8),s2.arrive_date,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
'ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
'   chc_DeliveryDate = "and convert(Char(8),arrive_date,112) = '" & txtDeliveryDateS.Text & "' "
'ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
'   chc_DeliveryDate = "and convert(Char(8),arrive_date,112) = '" & txtDeliveryDateE.Text & "' "
'End If
'
''ñ�����O
'If optNormal = 0 And optAbnormal = 0 And optNotYet = 0 Then GoTo NextStep
'Dim strStatus As String
'
'strStatus = "and s2.confirm_notes in ("
'
'If optNormal = 1 Then strStatus = strStatus & "'���`�q��',"
'If optAbnormal = 1 Then strStatus = strStatus & "'���`�q��','���X�q��',"
'If optNotYet = 1 Then strStatus = strStatus & "'',"
'
'str_SQL = str_SQL & Left(strStatus, Len(strStatus) - 1) & ")"
'
'NextStep:
'
''�f�D
'If Len(RTrim(Combo1.Text)) > 0 Then str_SQL = str_SQL & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' "
'
'If Combo2.Text = "�ϥΪ̡B���@�ɶ�" Then
'    str_SQL = str_SQL & "order by s2.confirm_userid,isnull(convert(char(19),s2.confirm_date,121),'') "
'Else
'    str_SQL = str_SQL & "order by isnull(t1m.channel,''),isnull(t1m.short_name,'') "
'End If

If Combo1 = "" Then MsgBox "�п�ܰt�e���q!", 16, Me.Caption: Screen.MousePointer = 0: Exit Sub
If txtDeliveryDateS.Text = "" Then MsgBox "�п�J��f���!", 16, Me.Caption: Screen.MousePointer = 0: Exit Sub

'If Combo1 = "�H�t" Then str_SQL = "exec gs_LABT01SdnList '" & txtDeliveryDateS.Text & "','002-10'"
'If Combo1 = "����" Then str_SQL = "exec gs_LABT01SdnList '" & txtDeliveryDateS.Text & "','000-31'"
'If Combo1 = "�w��" Then str_SQL = "exec gs_LABT01SdnList '" & txtDeliveryDateS.Text & "','000-70'"

'��������edit by Eric 20150112,���ϥ�SP
If Combo1 = "�H�t" Then strSelected = strSelected & "and s2.vehicle_id_no = '002-10' "
If Combo1 = "����" Then strSelected = strSelected & "and s2.vehicle_id_no = '000-31' "
If Combo1 = "�w��" Then strSelected = strSelected & "and s2.vehicle_id_no = '000-70' "

'�������
strSelected = strSelected & "and rtrim(isnull(s2.arrive_date,'')) = '" & txtDeliveryDateS.Text & "'"

str_SQL = "select " & _
            "�Ȥ�W�� = rtrim(isnull(t1m.short_name,'')) " & _
            ",��f�� = rtrim(isnull(s2.arrive_date,'')) " & _
            ",�q�����O = rtrim(isnull(s2.priority,'')) " & _
            ",�X�f�c�� = isnull((select sum(otqty) from ort02t where ort02t.receipt_no = s2.receipt_no),0) " & _
            ",�q�渹�X = rtrim(isnull(s2.extern,'')) " & _
            ",�禬�渹 = rtrim(isnull(s2.customerorderkey1,'')) " & _
            ",�h�f�c�� = isnull(rtrim(o.goodsback),0) " & _
            ",'�{��/�䲼' = o.cash " & _
            ",���`���p = case when s2.confirm_notes = '���`�q��' then 'N' when len(rtrim(isnull(s2.confirm_notes,''))) =0 then 'N' else 'Y' end " & _
            "from trp01m t1m right join sdn02t s2 on s2.consigneekey = t1m.consigneekey and s2.storerkey = t1m.storerkey " & _
            "join orders o on o.orderkey = s2.c_receipt_no " & _
            "where s2.storerkey = 'LABT01' " & strSelected & _
            "order by isnull(s2.extern,'')"


Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Call Replication_Recordset(tmp_Rs, rsMain)

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

End With

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
If Len(txtDeliveryDateST3) = 0 And Len(txtDeliveryDateET3) = 0 And Len(txtOrderDateST3) = 0 And Len(txtOrderDateET3) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT3.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String

str_SQL = "exec [gs_LABT01receive] '" & txtOrderDateST3 & "','" & txtOrderDateET3 & "','" & txtDeliveryDateST3 & "','" & txtDeliveryDateET3 & "'"

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
'If Len(txtDeliveryDateST4) = 0 Or Len(txtDeliveryDateET4) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
Screen.MousePointer = 11
Set dgMainT4.DataSource = Nothing: Set dgMainT4_1.DataSource = Nothing: Set dgMainT4_2.DataSource = Nothing: StatusBar.Panels(2).Text = ""
Dim strWhere As Integer, intloop As Integer, strTmp As String, chkDeliveryDate As String, chkCar As String, chkAdddateDate As String, chkDelivery As String, chkRoute As String, chkStatus As String, tmp_data() As String

'����
chkCar = ""
If RTrim(cboCarT4) <> "" Then chkCar = "and o2.vehicle_id_no = '" & cboCarT4 & "' "

'�ư�����
If Len(txtNotCarNo) > 0 Then
   '�x��s���G�s���A�H�r������
   tmp_data = Split(txtNotCarNo, ",", -1, vbTextCompare)    '���ο�J������
   strTmp = ""
   '�N�w���Φr��[�H�զX (�D�ťզr��~�[�H�զX)
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

'�����
chkAdddateDate = ""
If Len(RTrim(txtAddDateST4.Text)) > 0 And Len(RTrim(txtAddDateET4.Text)) > 0 Then
   chkAdddateDate = "and convert(char,o.adddate,112) between '" & txtAddDateST4.Text & "' and '" & txtAddDateET4.Text & "' "
ElseIf Len(RTrim(txtAddDateST4.Text)) > 0 And Len(RTrim(txtAddDateET4.Text)) = 0 Then
   chkAdddateDate = "and convert(char,o.adddate,112) = '" & txtAddDateST4.Text & "' "
ElseIf Len(RTrim(txtAddDateST4.Text)) = 0 And Len(RTrim(txtAddDateET4.Text)) > 0 Then
   chkAdddateDate = "and convert(char,o.adddate,112) = '" & txtAddDateET4.Text & "' "
End If

'�X����
chkDelivery = ""
If Len(RTrim(txtDeliveryST4.Text)) > 0 And Len(RTrim(txtDeliveryET4.Text)) > 0 Then
   chkDelivery = "and '20' + substring(o2.route_no,2,6) between '" & txtDeliveryST4.Text & "' and '" & txtDeliveryET4.Text & "' "
ElseIf Len(RTrim(txtDeliveryST4.Text)) > 0 And Len(RTrim(txtDeliveryET4.Text)) = 0 Then
   chkDelivery = "and '20' + substring(o2.route_no,2,6) = '" & txtDeliveryST4.Text & "' "
ElseIf Len(RTrim(txtDeliveryST4.Text)) = 0 And Len(RTrim(txtDeliveryET4.Text)) > 0 Then
   chkDelivery = "and '20' + substring(o2.route_no,2,6) = '" & txtDeliveryET4.Text & "' "
End If

'��f��
chkDeliveryDate = ""
If Len(RTrim(txtDeliveryDateST4.Text)) > 0 And Len(RTrim(txtDeliveryDateET4.Text)) > 0 Then
   chkDeliveryDate = "and convert(char(8),o2.arrive_Date,112) between '" & txtDeliveryDateST4.Text & "' and '" & txtDeliveryDateET4.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateST4.Text)) > 0 And Len(RTrim(txtDeliveryDateET4.Text)) = 0 Then
   chkDeliveryDate = "and convert(char(8),o2.arrive_Date,112) = '" & txtDeliveryDateST4.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateST4.Text)) = 0 And Len(RTrim(txtDeliveryDateET4.Text)) > 0 Then
   chkDeliveryDate = "and convert(char(8),o2.arrive_Date,112) = '" & txtDeliveryDateET4.Text & "' "
End If

'���s
chkRoute = ""
If Len(RTrim(txtRouteST4.Text)) > 0 And Len(RTrim(txtRouteET4.Text)) > 0 Then
   chkRoute = "and o2.route_no between '" & txtRouteST4.Text & "' and '" & txtRouteET4.Text & "' "
ElseIf Len(RTrim(txtRouteST4.Text)) > 0 And Len(RTrim(txtRouteET4.Text)) = 0 Then
   chkRoute = "and o2.route_no = '" & txtRouteST4.Text & "' "
ElseIf Len(RTrim(txtRouteST4.Text)) = 0 And Len(RTrim(txtRouteET4.Text)) > 0 Then
   chkRoute = "and o2.route_no = '" & txtRouteET4.Text & "' "
End If

'��ƪ��A
chkStatus = ""
If optNo = True Then chkStatus = "and len(rtrim(isnull(convert(char(20),o2.OTconfirmdate,120),''))) = 0 "
If optYes = True Then chkStatus = "and len(rtrim(isnull(convert(char(20),o2.OTconfirmdate,120),''))) > 0 "

str_SQL = "set nocount on if object_id ('tempdb..#2') is not null drop table #2  " & _
"select  Rtrim(isnull(o2.Extern,''))+'�B'+Rtrim(isnull(o.CustomerOrderkey,''))+'�B' + Rtrim(isnull(o.InvoiceNo,'')) + '�B' + Rtrim(isnull(o.B_Contact2,'')) as 'DN�渹' " & _
",Convert(char(10),o2.arrive_date,111) as '��f��' " & _
",Rtrim(t1m.full_name) as '�Ȥ�W��',isnull(Rtrim(t1m.Address),'') as '�Ȥ�a�}' " & _
",Rtrim(t1m.Phone) as '�Ȥ�q��',Rtrim(convert(char(1000),o.Notes)) as '�q��Ƶ�' " & _
",���u�s�� = o2.route_no,���� = rtrim(o2.vehicle_id_no) " & _
", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '�c��' " & _
",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '���ҽX' " & _
",�N���f�� = o.Cash,���h�f = case when isnull(o.GoodsBack,0) = 1 then '����' else '' end " & _
",��ƽT�{ = isnull(o2.OTconfirmuser,'���T�{'),���� = rtrim(isnull(o.B_City,'')),�ϰ� = left(t1m.area_code,1) into #2 " & _
"from trp02t o2 join trp03t o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
"inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
"inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
"left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
"left join trp02m tm on t1m.zip=tm.zip " & _
"where o.storerkey='LABT01' and o.type<>'�R��' and isnull(o.B_Phone1,'')<>'01' " & chkCar & chkAdddateDate & chkDelivery & chkDeliveryDate & chkRoute & chkStatus & _
"group by left(t1m.area_code,1),o2.OTconfirmuser,o2.route_no,o2.vehicle_id_no,o2.Extern ,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "

str_SQL = str_SQL & "union " & _
"select  Rtrim(isnull(o2.Extern,''))+'�B'+Rtrim(isnull(o.CustomerOrderkey,''))+'�B' + Rtrim(isnull(o.InvoiceNo,'')) + '�B' + Rtrim(isnull(o.B_Contact2,'')) as 'DN�渹' " & _
",Convert(char(10),o2.arrive_date,111) as '��f��' " & _
",Rtrim(t1m.full_name) as '�Ȥ�W��',isnull(Rtrim(t1m.Address),'') as '�Ȥ�a�}' " & _
",Rtrim(t1m.Phone) as '�Ȥ�q��',Rtrim(convert(char(1000),o.Notes)) as '�q��Ƶ�' " & _
",���u�s�� = o2.route_no,���� = rtrim(o2.vehicle_id_no) " & _
", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '�c��' " & _
",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '���ҽX' " & _
",�N���f�� = o.Cash,���h�f = case when isnull(o.GoodsBack,0) = 1 then '����' else '' end " & _
",��ƽT�{ = isnull(o2.OTconfirmuser,'���T�{'),���� = rtrim(isnull(o.B_City,'')),�ϰ� = left(t1m.area_code,1) " & _
"from ort02t o2 join ort03t o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
"inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
"inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
"left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
"left join trp02m tm on t1m.zip=tm.zip " & _
"where o.storerkey='LABT01' and o.type<>'�R��' and isnull(o.B_Phone1,'')<>'01' " & chkCar & chkAdddateDate & chkDelivery & chkDeliveryDate & chkRoute & chkStatus & _
"group by left(t1m.area_code,1),o2.OTconfirmuser,o2.route_no,o2.vehicle_id_no,o2.Extern ,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "


If RTrim(cboCarT4) = "�ݱƨ�" Or RTrim(cboCarT4) = "" Then
    str_SQL = str_SQL & "union select  Rtrim(isnull(o2.Extern,''))+'�B'+Rtrim(isnull(o.CustomerOrderkey,''))+'�B' + Rtrim(isnull(o.InvoiceNo,'')) + '�B' + Rtrim(isnull(o.B_Contact2,'')) as 'DN�渹' " & _
    ",Convert(char(10),o2.arrive_date,111) as '��f��' " & _
    ",Rtrim(t1m.full_name) as '�Ȥ�W��',isnull(Rtrim(t1m.Address),'') as '�Ȥ�a�}' " & _
    ",Rtrim(t1m.Phone) as '�Ȥ�q��',Rtrim(convert(char(1000),o.Notes)) as '�q��Ƶ�' " & _
    ",���u�s�� = '�ݱƨ�',���� = '�ݱƨ�' " & _
    ", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '�c��' " & _
    ",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '���ҽX' " & _
    ",�N���f�� = o.Cash,���h�f = case when isnull(o.GoodsBack,0) = 1 then '����' else '' end " & _
    ",��ƽT�{ = isnull(o2.OTconfirmuser,'���T�{'),���� = rtrim(isnull(o.B_City,'')),�ϰ� = left(t1m.area_code,1) " & _
    "from ort02w o2 join ort03w o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
    "inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
    "inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
    "left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
    "left join trp02m tm on t1m.zip=tm.zip " & _
    "where o.storerkey='LABT01' and o.type<>'�R��' and isnull(o.B_Phone1,'')<>'01' " & chkAdddateDate & chkDeliveryDate & chkRoute & chkStatus & _
    "group by left(t1m.area_code,1),o2.OTconfirmuser,o2.Extern,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2 ,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "

    str_SQL = str_SQL & "union select  Rtrim(isnull(o2.Extern,''))+'�B'+Rtrim(isnull(o.CustomerOrderkey,''))+'�B' + Rtrim(isnull(o.InvoiceNo,'')) + '�B' + Rtrim(isnull(o.B_Contact2,'')) as 'DN�渹' " & _
    ",Convert(char(10),o2.arrive_date,111) as '��f��' " & _
    ",Rtrim(t1m.full_name) as '�Ȥ�W��',isnull(Rtrim(t1m.Address),'') as '�Ȥ�a�}' " & _
    ",Rtrim(t1m.Phone) as '�Ȥ�q��',Rtrim(convert(char(1000),o.Notes)) as '�q��Ƶ�' " & _
    ",���u�s�� = '�ݱƨ�',���� = '�ݱƨ�' " & _
    ", case when p.casecnt>0 then Ceiling(Sum(o3.order_qty)/ p.casecnt)  else 1 end as  '�c��' " & _
    ",rtrim(t1m.Zip) as 'Zip',isnull(rtrim(tm.DCODE),'') as '���ҽX' " & _
    ",�N���f�� = o.Cash,���h�f = case when isnull(o.GoodsBack,0) = 1 then '����' else '' end " & _
    ",��ƽT�{ = isnull(o2.OTconfirmuser,'���T�{'),���� = rtrim(isnull(o.B_City,'')),�ϰ� = left(t1m.area_code,1) " & _
    "from trp02w o2 join trp03w o3 on o2.receipt_no = o3.receipt_no join orders o on o.orderkey=o2.c_receipt_no " & _
    "inner join Exceed_ABT..sku s on o3.product_no=s.sku and s.storerkey = 'LABT01' and s.storerkey = o.storerkey " & _
    "inner join Exceed_ABT..pack p on p.packkey=s.packkey " & _
    "left join trp01m t1m on t1m.storerkey = 'LABT01' and t1m.consigneekey = o.consigneekey " & _
    "left join trp02m tm on t1m.zip=tm.zip " & _
    "where o.storerkey='LABT01' and o.type<>'�R��' and isnull(o.B_Phone1,'')<>'01' " & chkAdddateDate & chkDeliveryDate & chkRoute & chkStatus & _
    "group by left(t1m.area_code,1),o2.OTconfirmuser,o2.Extern,o.CustomerOrderkey,o.InvoiceNo,o.B_Contact2 ,Convert(char(10),o2.arrive_date,111) ,Rtrim(t1m.full_name),isnull(Rtrim(t1m.Address),''),tm.DCODE ,Rtrim(t1m.Phone) ,Rtrim(convert(char(1000),o.Notes)), p.casecnt,t1m.Zip,o.Cash,o.GoodsBack,o2.otqty,rtrim(isnull(o.B_City,'')) "

End If

'��L���Ӫ���
str_SQL = str_SQL & "select ���u�s��,DN�渹,��f��,�Ȥ�W��,�Ȥ�a�},�Ȥ�q��,�q��Ƶ�,sum(�c��) as �c��,Zip,���ҽX,�N���f��,���h�f " & _
",���=isnull((select sum(isnull(otqty,0)) from trp02t where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1)  and OTconfirmdate is not null),0) " & _
"+isnull((select sum(isnull(otqty,0)) from ort02t where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1)  and OTconfirmdate is not null),0) " & _
"+isnull((select sum(isnull(otqty,0)) from trp02w where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1)  and OTconfirmdate is not null),0) " & _
"+isnull((select sum(isnull(otqty,0)) from ort02w where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1)  and OTconfirmdate is not null),0) " & _
",��ƽT�{,�ϰ� from  #2 where ���� not like '%���@%' and ���� not like '%���d%' group by DN�渹,��f��,�Ȥ�W��,�Ȥ�a�},�Ȥ�q��,�q��Ƶ�,Zip,���ҽX ,�N���f��,���h�f,��ƽT�{,�ϰ�,���u�s�� order by SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If tmp_Rs.EOF = True Then
    Set rsMainT4 = Nothing
    Screen.MousePointer = 0: MsgBox "�d�L��L���Ӫ��ơI", vbOKOnly + vbInformation, Me.Caption:
Else
    '�a�X���Ӹ��
    Call ReDim_Recordset(rsMainT4)
    Call Replication_Recordset(tmp_Rs, rsMainT4)
    tmp_Rs.Close
    rsMainT4.MoveFirst
    Do While Not rsMainT4.EOF
        str_SQL = "select �~�W²�� = case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end,�ƶq = isnull(sum(od.originalqty),0) from orders o (nolock) join orderdetail od (nolock)  on o.orderkey = od.orderkey and o.storerkey = 'LABT01' and o.type <> '�R��' " & _
                  "join Exceed_ABT..sku s (nolock)  on s.storerkey = od.storerkey and s.sku = od.sku " & _
                  "where CONVERT(varchar(12),o.deliverydate, 111) = '" & RTrim(rsMainT4.Fields("��f��")) & "' and o.externorderkey = '" & RTrim(mySplit(rsMainT4.Fields("DN�渹"), "�B", 0)) & "' group by  case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end "
                  
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                
                If RTrim(rsMainT4.Fields("�q��Ƶ�")) = "" Then
                    rsMainT4.Fields("�q��Ƶ�") = rsMainT4.Fields("�q��Ƶ�") & "�X�f���~�G"
                Else
                    rsMainT4.Fields("�q��Ƶ�") = rsMainT4.Fields("�q��Ƶ�") & "�@�X�f���~�G"
                End If
                
                Do While Not tmp_Rs.EOF
                    rsMainT4.Fields("�q��Ƶ�") = rsMainT4.Fields("�q��Ƶ�") & " " & RTrim(tmp_Rs.Fields("�~�W²��")) & "*" & RTrim(tmp_Rs.Fields("�ƶq")) & "�B"
                    tmp_Rs.MoveNext
                Loop
                rsMainT4.Fields("�q��Ƶ�") = Left(rsMainT4.Fields("�q��Ƶ�"), Len(rsMainT4.Fields("�q��Ƶ�")) - 1)
                tmp_Rs.Close
        rsMainT4.MoveNext
    Loop
    rsMainT4.MoveFirst
    Set dgMainT4.DataSource = rsMainT4
    SSTab1.Tab = 0
    StatusBar.Panels(2).Text = StatusBar.Panels(2).Text & "��L����:" & rsMainT4.RecordCount & " ����ƦC                   "
    SetDataGridColWidth Me.Caption, dgMainT4
End If

'���@���Ӫ�
str_SQL = "select ���u�s��,DN�渹,��f��,�Ȥ�W��,�Ȥ�a�},�Ȥ�q��,�q��Ƶ�,sum(�c��) as �c��,Zip,���ҽX,�N���f��,���h�f,���=isnull((select sum(isnull(otqty,0)) from trp02t where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1)  and OTconfirmdate is not null),0) " & _
    "+isnull((select sum(isnull(otqty,0)) from ort02t where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1) and OTconfirmdate is not null),0) " & _
    ",��ƽT�{,�ϰ� from  #2 where ���� like '%���@%' group by DN�渹,��f��,�Ȥ�W��,�Ȥ�a�},�Ȥ�q��,�q��Ƶ�,Zip,���ҽX ,�N���f��,���h�f,��ƽT�{,�ϰ�,���u�s�� order by SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = True Then
    Set rsMainT4_1 = Nothing
    Screen.MousePointer = 0: MsgBox "�d�L���@���Ӫ��ơI", vbOKOnly + vbInformation, Me.Caption:
Else
    '�a�X���Ӹ��
    Call ReDim_Recordset(rsMainT4_1)
    Call Replication_Recordset(tmp_Rs, rsMainT4_1)
    tmp_Rs.Close
     rsMainT4_1.MoveFirst
    Do While Not rsMainT4_1.EOF
        str_SQL = "select �~�W²�� = case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end,�ƶq = isnull(sum(od.originalqty),0) from orders o (nolock) join orderdetail od (nolock)  on o.orderkey = od.orderkey and o.storerkey = 'LABT01' and o.type <> '�R��' " & _
                  "join Exceed_ABT..sku s (nolock)  on s.storerkey = od.storerkey and s.sku = od.sku " & _
                  "where CONVERT(varchar(12),o.deliverydate, 111) = '" & RTrim(rsMainT4_1.Fields("��f��")) & "' and o.externorderkey = '" & RTrim(mySplit(rsMainT4_1.Fields("DN�渹"), "�B", 0)) & "' group by  case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end "
                
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.CursorLocation = 3 '�i�H�ק�recordset
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If RTrim(rsMainT4_1.Fields("�q��Ƶ�")) = "" Then
                    rsMainT4_1.Fields("�q��Ƶ�") = rsMainT4_1.Fields("�q��Ƶ�") & "�X�f���~�G"
                Else
                    rsMainT4_1.Fields("�q��Ƶ�") = rsMainT4_1.Fields("�q��Ƶ�") & "�@�X�f���~�G"
                End If
                Do While Not tmp_Rs.EOF
                    rsMainT4_1.Fields("�q��Ƶ�") = rsMainT4_1.Fields("�q��Ƶ�") & " " & RTrim(tmp_Rs.Fields("�~�W²��")) & "*" & RTrim(tmp_Rs.Fields("�ƶq")) & "�B"
                    tmp_Rs.MoveNext
                Loop
                rsMainT4_1.Fields("�q��Ƶ�") = Left(rsMainT4_1.Fields("�q��Ƶ�"), Len(rsMainT4_1.Fields("�q��Ƶ�")) - 1)
                tmp_Rs.Close
        rsMainT4_1.MoveNext
    Loop
        rsMainT4_1.MoveFirst
    Set dgMainT4_1.DataSource = rsMainT4_1
    SetDataGridColWidth Me.Caption, dgMainT4_1
    StatusBar.Panels(2).Text = StatusBar.Panels(2).Text & "���@����:" & rsMainT4_1.RecordCount & " ����ƦC                 "
    SSTab1.Tab = 1
End If


'���d���Ӫ�
str_SQL = "select ���u�s��,DN�渹,��f��,�Ȥ�W��,�Ȥ�a�},�Ȥ�q��,�q��Ƶ�,sum(�c��) as �c��,Zip,���ҽX,�N���f��,���h�f,���=isnull((select sum(isnull(otqty,0)) from trp02t where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1)  and OTconfirmdate is not null),0) " & _
    "+isnull((select sum(isnull(otqty,0)) from ort02t where extern =  SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1) and OTconfirmdate is not null),0) " & _
    ",��ƽT�{,�ϰ� from  #2 where ���� like '%���d%' group by DN�渹,��f��,�Ȥ�W��,�Ȥ�a�},�Ȥ�q��,�q��Ƶ�,Zip,���ҽX ,�N���f��,���h�f,��ƽT�{,�ϰ�,���u�s�� order by SUBSTRING(DN�渹 , 1, CHARINDEX('�B',  DN�渹 )-1) if object_id ('tempdb..#2') is not null drop table #2 set nocount off "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = True Then
    Set rsMainT4_2 = Nothing
    Screen.MousePointer = 0: MsgBox "�d�L���d���Ӫ��ơI", vbOKOnly + vbInformation, Me.Caption:
Else
    '�a�X���Ӹ��
    Call ReDim_Recordset(rsMainT4_2)
    Call Replication_Recordset(tmp_Rs, rsMainT4_2)
    tmp_Rs.Close
    rsMainT4_2.MoveFirst
    Do While Not rsMainT4_2.EOF
        str_SQL = "select �~�W²�� = case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end,�ƶq = isnull(sum(od.originalqty),0) from orders o (nolock) join orderdetail od (nolock)  on o.orderkey = od.orderkey and o.storerkey = 'LABT01'  and o.type <> '�R��' " & _
                  "join Exceed_ABT..sku s (nolock)  on s.storerkey = od.storerkey and s.sku = od.sku " & _
                  "where CONVERT(varchar(12),o.deliverydate, 111) = '" & RTrim(rsMainT4_2.Fields("��f��")) & "' and o.externorderkey = '" & RTrim(mySplit(rsMainT4_2.Fields("DN�渹"), "�B", 0)) & "' group by  case when len(rtrim(isnull(s.altsku,''))) = 0 then rtrim(isnull(s.descr,'')) else rtrim(isnull(s.altsku,'')) end "
                  
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If RTrim(rsMainT4_2.Fields("�q��Ƶ�")) = "" Then
                    rsMainT4_2.Fields("�q��Ƶ�") = rsMainT4_2.Fields("�q��Ƶ�") & "�X�f���~�G"
                Else
                    rsMainT4_2.Fields("�q��Ƶ�") = rsMainT4_2.Fields("�q��Ƶ�") & "�@�X�f���~�G"
                End If
                Do While Not tmp_Rs.EOF
                    rsMainT4_2.Fields("�q��Ƶ�") = rsMainT4_2.Fields("�q��Ƶ�") & " " & RTrim(tmp_Rs.Fields("�~�W²��")) & "*" & RTrim(tmp_Rs.Fields("�ƶq")) & "�B"
                    tmp_Rs.MoveNext
                Loop
                rsMainT4_2.Fields("�q��Ƶ�") = Left(rsMainT4_2.Fields("�q��Ƶ�"), Len(rsMainT4_2.Fields("�q��Ƶ�")) - 1)
                tmp_Rs.Close
        rsMainT4_2.MoveNext
    Loop
    rsMainT4_2.MoveFirst
    Set dgMainT4_2.DataSource = rsMainT4_2
    SetDataGridColWidth Me.Caption, dgMainT4_2
    StatusBar.Panels(2).Text = StatusBar.Panels(2).Text & "���d����:" & rsMainT4_2.RecordCount & " ����ƦC                 "
    SSTab1.Tab = 2
End If

    Screen.MousePointer = 0: dgMainT4.Visible = True: dgMainT4_1.Visible = True: dgMainT4_2.Visible = True

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
'��ƱƧ�
Recordset2Excel "�ըƹF���y�^���ˮ֪�", rsMain

'..�b���s��EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
'        .Columns("L").Select
'        .Selection.ClearContents
        .Range("B3").Value = Combo1
        .Range("A1").Select
        '�ƥ��ɮ�
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
'Create by Gemini @20061102 4 Recordset�ץXExcel
'�ϥλ���
'1.�s�WEXCEL�d����
'2.���פJ��ƪ��u�@��R�W��DATA
'3.�N���}�l��m��ƪ��x�s���J�{�����Y�r��(App.Title)�A��m��m����A100-Z100����
'4.�N�d���ɮM�ζ���1.�{���ؿ��U"XLT"�d�Ҹ�Ƨ�2.ini�ɩҫ��w�����|
'�Ѽƻ���
'frm:�ӷ�From����
'rs:�ӷ�Recordset
'�d��
'    Recordset2Excel Me, rs_Cust
'    '..�b���s��EXCEL
'    Set MyXlsApp = Nothing'�פ�Excel����
'�ŧi��Ҳ�
'Public MyXlsApp As Excel.Application
'**************************************************
On Error GoTo err_Handle
If rs Is Nothing Then MsgBox str & "�L��ƥi�����ɡI", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
If rs.RecordCount > 65535 Then MsgBox str & "��X��ƶW�LExcel����(65535)�I", 16, "Save2Excel�פ�": Exit Sub
If rs.RecordCount = -1 Then MsgBox str & "�L��ƥi�����ɡI", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String

'MsgBox "�t�ζi������Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel" 'add @ 20110402

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .DisplayAlerts = False '�u�@��s�W�ƻs�R������ܴ��ܵ���add @ 20110402

    If Dir(App.Path & "\XLT\" & str & ".xlt") = "" Then '�䤣�쥻���d����
        
        '���d���ɸ��|
        Dim objIni As vbIniFile, strXltPath As String
        Set objIni = New vbIniFile
        
        With objIni
        
            .FileName = striniFileName_FullPath
            strXltPath = RTrim(.ReadData("EXCEL", "XLTPATH", "")) '���䴩�����Ƨ��W��
            
        End With
        Set objIni = Nothing

    End If

    '�L���w���|�νd���ɦW�A���ϥνd����
    If Len(RTrim(strXltPath)) = 0 Then strXltPath = App.Path & "\XLT"
    If Dir(strXltPath, vbDirectory) = "" Or Len(RTrim(str)) = 0 Then GoTo Run
    
    '�d����
    If Dir(strXltPath & "\" & str & ".xlt") <> "" Then
'        If MsgBox("�O�_�ϥνd����?(" & strXltPath & "\" & str & ".xlt), vbQuestion + vbYesNo, "��Excel") = vbNo Then GoTo Run
        
        '�}�ҽd����
        .Workbooks.Open (strXltPath & "\" & str & ".xlt")
        
        '�M��DATA�u�@��
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets("Data").Select: Exit For '��wDATA�u�@��
        Next
        
        '�䤣��s�WDATA�u�@��
        If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then
            .Sheets.Add
            .ActiveSheet.Name = "DATA"
        Else
            '���j�M�s���x�s��
            For k = 65 To 90
                For j = 1 To 100
                    If UCase(.Range(Chr(k) & j).Value) = "BESTTRP" Then GoTo NextStep
                Next j
            Next k
            k = 65: j = 2 '�S���ɫ��w��A1(J=2�O�]���U���|-1)
        End If
        .ActiveSheet.Name = str
NextStep:
        '�g�J���D�C
        If j > 1 Then '�p�G�b�Ĥ@�C�A�h�������W��
            For i = 1 To rs.Fields.Count - 1
                l = i Mod 26
                .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Name
                '���W�L26
                If Chr(65 + l) = "Z" Then
                    If strCol = "" Then
                        strCol = "A"
                    Else
                        strCol = Chr(Asc(strCol) + 1)
                    End If
                End If
            Next i
            '�g�Jrecordset���
            rs.MoveFirst: j = 3
            Do While Not rs.EOF
                For i = 1 To rs.Fields.Count - 1
                    l = i Mod 26
                    .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Value
                    If RTrim(rs.Fields(i).Value) = "���T�{" Then .Range(strCol & Chr(k + l - 1 - 1) & j - 1).Value = ""
                    '���W�L26
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

        '��Ƽg�J
        '.Range(Chr(k) & j).CopyFromRecordset rs
        
    Else '���ϥνd����
Run:
        '�s�WExcel
        .Workbooks.Add: .ActiveSheet.Name = str
              '�g�J���D�C
        j = 2:  k = 65
        If j > 1 Then '�p�G�b�Ĥ@�C�A�h�������W��
            For i = 1 To rs.Fields.Count - 1
                l = i Mod 26
                .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Name
                '���W�L26
                If Chr(65 + l) = "Z" Then
                    If strCol = "" Then
                        strCol = "A"
                    Else
                        strCol = Chr(Asc(strCol) + 1)
                    End If
                End If
            Next i
            '�g�Jrecordset���
            rs.MoveFirst: j = 3
            Do While Not rs.EOF
                For i = 1 To rs.Fields.Count - 1
                    l = i Mod 26
                    .Range(strCol & Chr(k + l - 1) & j - 1).Value = rs.Fields(i).Value
                If RTrim(rs.Fields(i).Value) = "���T�{" Then
                    .Range(strCol & Chr(k + l - 1 - 1) & j - 1).Value = ""
                End If
                    '���W�L26
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
Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, "��EXECL���~!!")
End Sub

