VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Query_Charge 
   Caption         =   "�ХI�ڤ����"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
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
   ScaleHeight     =   8430
   ScaleWidth      =   11130
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   5280
      TabIndex        =   0
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
      StartOfWeek     =   93454337
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "��L�f�D"
      TabPicture(0)   =   "frm_Query_Charge.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "LABT01�Ȱ�"
      TabPicture(1)   =   "frm_Query_Charge.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " LVTL01���L�S"
      TabPicture(2)   =   "frm_Query_Charge.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
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
         TabIndex        =   39
         Top             =   360
         Width           =   8295
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
            Picture         =   "frm_Query_Charge.frx":0054
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   48
            Top             =   240
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
            Picture         =   "frm_Query_Charge.frx":035E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   47
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtOrderDateST2 
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
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateET2 
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
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.ComboBox Combo3 
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
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   44
            Top             =   240
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
            TabIndex        =   43
            Top             =   960
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
            TabIndex        =   42
            Top             =   960
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
            Index           =   2
            Left            =   7080
            Picture         =   "frm_Query_Charge.frx":1658
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   41
            Top             =   1200
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
            Picture         =   "frm_Query_Charge.frx":2B26A
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
            Index           =   14
            Left            =   2655
            TabIndex        =   53
            Top             =   660
            Visible         =   0   'False
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
            Index           =   13
            Left            =   120
            TabIndex        =   52
            Top             =   645
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D"
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
            Left            =   360
            TabIndex        =   51
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
            Index           =   11
            Left            =   120
            TabIndex        =   50
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
            Left            =   2655
            TabIndex        =   49
            Top             =   1020
            Width           =   360
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
         TabIndex        =   37
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT2 
            Height          =   2295
            Left            =   120
            TabIndex        =   38
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
         TabIndex        =   33
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMainT1 
            Height          =   2295
            Left            =   120
            TabIndex        =   34
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
         TabIndex        =   20
         Top             =   360
         Width           =   8295
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
            Picture         =   "frm_Query_Charge.frx":2B57C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   36
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
            Picture         =   "frm_Query_Charge.frx":2B88E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   35
            Top             =   1200
            Width           =   1065
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
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   960
            Width           =   1485
         End
         Begin VB.ComboBox Combo2 
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
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   25
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateET1 
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
            TabIndex        =   24
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateST1 
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
            TabIndex        =   23
            Top             =   600
            Width           =   1485
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
            Picture         =   "frm_Query_Charge.frx":554A0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   22
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
            Picture         =   "frm_Query_Charge.frx":5679A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   21
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
            Left            =   2655
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   1005
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D"
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
            Left            =   360
            TabIndex        =   30
            Top             =   300
            Width           =   480
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
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   645
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
            Index           =   5
            Left            =   2655
            TabIndex        =   28
            Top             =   660
            Width           =   360
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   8295
         Begin VB.CommandButton cmdQueryAll 
            BackColor       =   &H00FFFFC0&
            Caption         =   "All�d��"
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
            Picture         =   "frm_Query_Charge.frx":56AA4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   54
            Top             =   1200
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
            Picture         =   "frm_Query_Charge.frx":56DAE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   14
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
            Picture         =   "frm_Query_Charge.frx":570B8
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
            Index           =   0
            Left            =   7080
            Picture         =   "frm_Query_Charge.frx":573CA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   12
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
            Left            =   5880
            Picture         =   "frm_Query_Charge.frx":80FDC
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   11
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtOrderDateS 
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
            TabIndex        =   10
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtOrderDateE 
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
            TabIndex        =   9
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.ComboBox Combo1 
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
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   8
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateE 
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
            TabIndex        =   7
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDateS 
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
            TabIndex        =   6
            Top             =   960
            Width           =   1485
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
            TabIndex        =   19
            Top             =   660
            Visible         =   0   'False
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
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   645
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D"
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
            Left            =   360
            TabIndex        =   17
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
            Index           =   3
            Left            =   120
            TabIndex        =   16
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
            Left            =   2655
            TabIndex        =   15
            Top             =   1020
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
         Left            =   120
         TabIndex        =   3
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMain 
            Height          =   2295
            Left            =   120
            TabIndex        =   4
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
      TabIndex        =   1
      Top             =   8160
      Width           =   11130
      _ExtentX        =   19632
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
            Object.Width           =   12991
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
Attribute VB_Name = "frm_Query_Charge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private rsMainAll As ADODB.Recordset
Private rsMainT1 As ADODB.Recordset
Private rsMainT2 As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmd2Excel_Click()
If RTrim(Combo1) = "LLFA01" Then
    Call cmd2Excel_LLFA01
ElseIf RTrim(Combo1) = "LHYI01" Then Call cmd2Excel_LHYI01
Else
    Call cmd2Excel_Normal_Click
End If

End Sub
Private Sub cmd2Excel_LHYI01()

MsgBox "�t�ζi��j�q�����Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel"

Call WriteOut_RunLog("1/5.��X�p�O���Ӹ��")

On Error GoTo err_Handle

'��ƱƧ�
Recordset2Excel "�B�O����", rsMain
If rsMain Is Nothing Then Call Unload_RunLogForm: Exit Sub

'..�b���s��EXCEL
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String
With MyXlsApp: .Visible = False
 
Dim rsTmp As New ADODB.Recordset

'�����
    .Sheets.Add: .ActiveSheet.Name = "�����"

    str_SQL = "exec gs_Charge '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
                
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)

    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("2/5.��X�������")
    Call OffLineRecordset(tmp_Rs, rsTmp)

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
    
'�������Ӹ��
strSheet = "��������"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "��������" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "��������" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_ap '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("3/5.��X�������Ӹ��")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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
strSheet = "�����I"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "�����I" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "�����I" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_OtherARP '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("4/5.��X�����I���")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�s�i�~����
strSheet = "�s�i�~����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "�s�i�~����" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "�s�i�~����" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "if object_id ('tempdb..#temp') is not null drop table #temp if object_id ('tempdb..#temp1') is not null drop table #temp1 set nocount on select extern = s2.extern,costcode = s5.costcode,sumreceivable = sumreceivable / 0.8 " & _
"into #temp from sdn02t s2 join sdn03t s3 on s3.receipt_no = s2.receipt_no and s2.storerkey = 'LHYI01' " & _
"join sdn05t s5 on s5.sdn_no = s2.receipt_no and s2.arrive_date between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' and costcode like '%_POSM%' and sumreceivable > 0 join " & strWMSDB & "..sku s on s.sku = s3.product_no and s.storerkey = s3.storerkey " & _
"join " & strWMSDB & "..pack p on p.packkey = s.packkey group by s2.extern,s5.CostCode,s5.SumReceivable " & _
"select extern = s2.extern,POSMCS = sum(case when isnull(cast(s.notes1 as varchar(100)),'') <> 'POSM' then 0 when p.casecnt = 0 then 1 else ceiling(s3.ship_qty/p.casecnt)end) " & _
"into #temp1 from sdn02t s2 join sdn03t s3 on s3.receipt_no = s2.receipt_no and s2.storerkey = 'LHYI01' " & _
"and s2.arrive_date between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' join " & strWMSDB & "..sku s on s.sku = s3.product_no and s.storerkey = s3.storerkey " & _
"join " & strWMSDB & "..pack p on p.packkey = s.packkey group by s2.extern " & _
"select �f�D�渹 = rtrim(s2.extern),�p�O�ϰ� = left(t1m.area_code,1),POSM�c�� = (select POSMCS from #temp1 where Extern = s2.Extern ) " & _
",���I = isnull((select sum(sumreceivable) from #temp where Extern = s2.Extern and costcode like '%OD_POSM%'),0) " & _
",�[�X = isnull((select sum(sumreceivable) from #temp where Extern = s2.Extern and costcode like '%CS_POSM%'),0) " & _
",�p�p = isnull((select sum(sumreceivable) from #temp where Extern = s2.Extern and costcode like '%POSM%'),0) " & _
",��f��� = s2.arrive_date,�Ȥ�W�� = rtrim(t1m.short_name),�e�f�a�} = rtrim(t1m.address) " & _
"from sdn02t s2 join trp01m t1m on t1m.consigneekey = s2.consigneekey and s2.storerkey = t1m.storerkey and s2.storerkey = 'LHYI01' " & _
"join #temp t on t.Extern = s2.Extern group by s2.extern,s2.arrive_date,t1m.short_name,t1m.address,left(t1m.area_code,1) order by s2.arrive_date,s2.extern "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("5/5.��X�s�i�~�������")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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

Set MyXlsApp = Nothing
.Visible = True: End With
Call Unload_RunLogForm
Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cmd2Excel_LLFA01()

MsgBox "�t�ζi��j�q�����Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel"

Call WriteOut_RunLog("1/5.��X�p�O���Ӹ��")

On Error GoTo err_Handle

'��ƱƧ�
Recordset2Excel "�B�O����", rsMain
If rsMain Is Nothing Then Call Unload_RunLogForm: Exit Sub

'..�b���s��EXCEL
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String
With MyXlsApp: .Visible = False
 
    Dim rsTmp As New ADODB.Recordset
'�����
    .Sheets.Add: .ActiveSheet.Name = "�����"

    str_SQL = "exec gs_Charge '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
                
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)

    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("2/5.��X�������")
    Call OffLineRecordset(tmp_Rs, rsTmp)

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
    
'�������Ӹ��
strSheet = "��������"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "��������" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "��������" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_ap '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("3/5.��X�������Ӹ��")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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
strSheet = "�����I"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "�����I" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "�����I" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_OtherARP '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("4/5.��X�����I���")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�s�i�~����
strSheet = "�s�i�~����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "�s�i�~����" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "�s�i�~����" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "if object_id ('tempdb..#temp') is not null drop table #temp if object_id ('tempdb..#temp1') is not null drop table #temp1 set nocount on select extern = s2.extern,costcode = s5.costcode,sumreceivable = sumreceivable / 0.8 " & _
"into #temp from sdn02t s2 join sdn03t s3 on s3.receipt_no = s2.receipt_no and s2.storerkey = 'LLFA01' " & _
"join sdn05t s5 on s5.sdn_no = s2.receipt_no and s2.arrive_date between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' and costcode like '%_POSM%' and sumreceivable > 0 join " & strWMSDB & "..sku s on s.sku = s3.product_no and s.storerkey = s3.storerkey " & _
"join " & strWMSDB & "..pack p on p.packkey = s.packkey group by s2.extern,s5.CostCode,s5.SumReceivable " & _
"select extern = s2.extern,POSMCS = sum(case when isnull(cast(s.notes1 as varchar(100)),'') <> 'POSM' then 0 when p.casecnt = 0 then 1 else ceiling(s3.ship_qty/p.casecnt)end) " & _
"into #temp1 from sdn02t s2 join sdn03t s3 on s3.receipt_no = s2.receipt_no and s2.storerkey = 'LLFA01' " & _
"and s2.arrive_date between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' join " & strWMSDB & "..sku s on s.sku = s3.product_no and s.storerkey = s3.storerkey " & _
"join " & strWMSDB & "..pack p on p.packkey = s.packkey group by s2.extern " & _
"select �f�D�渹 = rtrim(s2.extern),�p�O�ϰ� = left(t1m.area_code,1),POSM�c�� = (select POSMCS from #temp1 where Extern = s2.Extern ) " & _
",���I = isnull((select sum(sumreceivable) from #temp where Extern = s2.Extern and costcode like '%OD_POSM%'),0) " & _
",�[�X = isnull((select sum(sumreceivable) from #temp where Extern = s2.Extern and costcode like '%CS_POSM%'),0) " & _
",�p�p = isnull((select sum(sumreceivable) from #temp where Extern = s2.Extern and costcode like '%POSM%'),0) " & _
",��f��� = s2.arrive_date,�Ȥ�W�� = rtrim(t1m.short_name),�e�f�a�} = rtrim(t1m.address) " & _
"from sdn02t s2 join trp01m t1m on t1m.consigneekey = s2.consigneekey and s2.storerkey = t1m.storerkey and s2.storerkey = 'LLFA01' " & _
"join #temp t on t.Extern = s2.Extern group by s2.extern,s2.arrive_date,t1m.short_name,t1m.address,left(t1m.area_code,1) order by s2.arrive_date,s2.extern "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("5/5.��X�s�i�~�������")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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

Set MyXlsApp = Nothing
.Visible = True: End With
Call Unload_RunLogForm
Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cmd2ExcelT2_Click()
On Error GoTo err_Handle
If rsMainT2 Is Nothing Then MsgBox "�L��ƥi�����ɡI", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
Screen.MousePointer = 11
Call WriteOut_RunLog("1/6.��X�p�O���Ӹ��")
Recordset2Excel "LVTL01�����b�ک��Ӫ�", rsMainT2

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
str_SQL = "exec gs_Charge 'LVTL01','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "

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

tmp_Rs.Open "exec [gs_LVTL01AR2_VK] '" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "','A280'", cn
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
tmp_Rs.Open "exec [gs_LVTL01AR2_VK] '" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "','A145'", cn
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
strSheet = "���ưt�e"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open "exec [gs_LVTL01AR2_Drink] '" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "'", cn
Call WriteOut_RunLog("5/6.��X���ưt�e���")
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

tmp_Rs.Open "exec es_OtherARP 'LVTL01','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "'", cn
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

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
Set rsMainAll = Nothing
Set rsMainT1 = Nothing
Set rsMainT2 = Nothing
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    SSTab.Height = Me.ScaleHeight - StatusBar.Height
    Frame2.Height = SSTab.Height - Frame1.Height - Frame1.Top - 120: dgMain.Height = Frame2.Height - 360
    Frame4.Height = SSTab.Height - Frame3.Height - Frame1.Top - 120: dgMainT1.Height = Frame4.Height - 360
    Frame6.Height = SSTab.Height - Frame5.Height - Frame1.Top - 120: dgMainT2.Height = Frame6.Height - 360
'    Frame8.Height = SSTab.Height - Frame7.Height - Frame1.Top - 120: dgMainT3.Height = Frame8.Height - 360
'    Frame10.Height = SSTab.Height - Frame9.Height - Frame1.Top - 120: dgMainT4.Height = Frame10.Height - 360
'    Frame12.Height = SSTab.Height - Frame11.Height - Frame1.Top - 120: dgMainT5.Height = Frame12.Height - 360
'    Frame14.Height = SSTab.Height - Frame13.Height - Frame1.Top - 120: dgMainT6.Height = Frame14.Height - 360
'    Frame16.Height = SSTab.Height - Frame15.Height - Frame1.Top - 120: dgMainT7.Height = Frame16.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab.Width = Me.ScaleWidth
    Frame2.Width = SSTab.Width - 360: dgMain.Width = Frame2.Width - 240
    Frame4.Width = SSTab.Width - 360: dgMainT1.Width = Frame4.Width - 240
    Frame6.Width = SSTab.Width - 360: dgMainT2.Width = Frame6.Width - 240
'    Frame8.Width = SSTab.Width - 360: dgMainT3.Width = Frame8.Width - 240
'    Frame10.Width = SSTab.Width - 360: dgMainT4.Width = Frame10.Width - 240
'    Frame12.Width = SSTab.Width - 360: dgMainT5.Width = Frame12.Width - 240
'    Frame14.Width = SSTab.Width - 360: dgMainT6.Width = Frame14.Width - 240
'    Frame16.Width = SSTab.Width - 360: dgMainT7.Width = Frame16.Width - 240
End If

End Sub

Private Sub dgmain_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub cmdExit_Click(Index As Integer)
Unload Me '�������{��
'End �������ε{��
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id
SSTab.Tab = 0

'�f�D
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(storerkey) from trp16M where storerkey not in ('LABT01')", cn, adOpenKeyset, adLockPessimistic

If Not tmp_Rs.EOF Then

    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        Combo1.AddItem tmp_Rs("storerkey")
        tmp_Rs.MoveNext
    Next
    tmp_Rs.Close: Set tmp_Rs = Nothing
    Combo1.ListIndex = 0

End If

Combo2.AddItem "LABT01": Combo2.ListIndex = 0
Combo3.AddItem "LVTL01": Combo3.ListIndex = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)

Me.mvDate.Visible = False
If Len(Trim(SSTab.Caption)) = 0 Then SSTab.Tab = PreviousTab: Exit Sub

StatusBar.Panels(2).Text = "0 ����ƦC"
If SSTab.Tab = 0 And (rsMain Is Nothing) = False Then StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
If SSTab.Tab = 1 And (rsMainT1 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT1.RecordCount & " ����ƦC"
If SSTab.Tab = 2 And (rsMainT2 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT2.RecordCount & " ����ƦC"
'If SSTab.Tab = 3 And (rsMainT3 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT3.RecordCount & " ����ƦC"
'If SSTab.Tab = 4 And (rsMainT4 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT4.RecordCount & " ����ƦC"
'If SSTab.Tab = 5 And (rsMainT5 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT5.RecordCount & " ����ƦC"
'If SSTab.Tab = 6 And (rsMainT6 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT6.RecordCount & " ����ƦC"
'If SSTab.Tab = 7 And (rsMainT7 Is Nothing) = False Then StatusBar.Panels(2).Text = rsMainT7.RecordCount & " ����ƦC"

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
Private Sub cmdQueryAll_Click()
Dim chc_Orderdate As String, chc_DeliveryDate As String
If (Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) = 0) And (Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) = 0) Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub

'��f���
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_Orderdate = "and ��f�� between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateE.Text & "' "
End If


MsgBox "�t�ζi��j�q�����Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel"

Call WriteOut_RunLog("1/2.��X�p�O���Ӹ��")
'��ƱƧ�
str_SQL = "select * from gv_sdn05tdetail where 1 = 1 " & chc_Orderdate
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic
If tmp_Rs.EOF = True Then tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
If tmp_Rs.RecordCount > 65535 Then tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "��ƶW�XExcel���ƭ���I�Э��s�d��", vbOKOnly + vbCritical, Me.Caption: Exit Sub
tmp_Rs.Sort = "��f��,���u�s��,�f�D�渹"

Set rsMainAll = New ADODB.Recordset

Call OffLineRecordset(tmp_Rs, rsMainAll)
tmp_Rs.Sort = "": tmp_Rs.Close
rsMainAll.MoveFirst

Recordset2Excel "���f�D�B�O����", rsMainAll

'..�b���s��EXCEL
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String
With MyXlsApp: .Visible = False
 
    Dim rsTmp As New ADODB.Recordset
'All�����I
strSheet = "�����I"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "�����I" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "�����I" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_AllARP '" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("2/2.��X���������I���")

tmp_Rs.Open str_SQL, cn

If tmp_Rs.RecordCount > 65535 Then tmp_Rs.Close:  Set MyXlsApp = Nothing: Call Unload_RunLogForm: Screen.MousePointer = 0: MsgBox "��ƶW�XExcel���ƭ���I�Э��s�d��", vbOKOnly + vbCritical, Me.Caption: Exit Sub

Call OffLineRecordset(tmp_Rs, rsTmp)

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

Set MyXlsApp = Nothing
.Visible = True: End With
Call Unload_RunLogForm
Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdQuery_Click()
If (Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) = 0) And (Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) = 0) Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub

'�f�D�ˬd
If Len(RTrim(Combo1.Text)) = 0 Then MsgBox "�п�J�f�D�s��", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String, chc_DeliveryDate As String, chc_Storerkey As String

''�f�D�s��
'If Len(RTrim(Combo1.Text)) > 0 Then chc_Storerkey = " and �f�D ='" & Combo1.Text & "' "

'�q����
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and �q��� between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and �q��� = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and �q��� = '" & txtOrderDateE.Text & "' "
End If

'��f���
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_Orderdate = "and ��f�� between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateE.Text & "' "
End If

If RTrim(Combo1.Text) = "LCHF01" Then
    str_SQL = "select * from gv_sdn05tdetail_LCHF01 where 1 = 1 " & chc_Orderdate & chc_DeliveryDate
Else
    str_SQL = "select * from gv_sdn05tdetail where 1 = 1 " & chc_Orderdate & chc_DeliveryDate
End If

'�f�D
If Len(RTrim(Combo1.Text)) > 0 Then str_SQL = str_SQL & "and �f�D = '" & RTrim(Combo1.Text) & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
tmp_Rs.Sort = "��f��,���u�s��,�f�D�渹"

Set rsMain = New ADODB.Recordset

Call OffLineRecordset(tmp_Rs, rsMain)
tmp_Rs.Sort = "": tmp_Rs.Close
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
Private Sub cmd2Excel_Normal_Click()

MsgBox "�t�ζi��j�q�����Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel"

Call WriteOut_RunLog("1/4.��X�p�O���Ӹ��")

On Error GoTo err_Handle
'��ƱƧ�
Recordset2Excel "�B�O����", rsMain
If rsMain Is Nothing Then Call Unload_RunLogForm: Exit Sub

'..�b���s��EXCEL
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String
With MyXlsApp: .Visible = False
 
Dim rsTmp As New ADODB.Recordset
'Dim xlsWB, xlsSht As Object
'xlsWB = MyXlsApp.Workbooks.Add
'xlsSht = xlsWB.Worksheets(1)
'�����
    .Sheets.Add: .ActiveSheet.Name = "�����"

    str_SQL = "exec gs_Charge '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
                
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)

    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("2/4.��X�������")
    Call OffLineRecordset(tmp_Rs, rsTmp)

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
    
'    If Combo1.Text = "LCHF01" Then
'        xlsSht.Columns("I:I").HorizontalAlignment = xlRight
'    End If

    .Range("A2").CopyFromRecordset rsTmp

    rsTmp.Close
    
'�������Ӹ��
strSheet = "��������"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "��������" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "��������" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_ap '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("3/4.��X�������Ӹ��")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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
strSheet = "�����I"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "�����I" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "�����I" Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_OtherARP '" & Combo1.Text & "','" & txtDeliveryDateS & "','" & txtDeliveryDateE & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("4/4.��X�����I���")

tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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

Set MyXlsApp = Nothing
.Visible = True: End With
Call Unload_RunLogForm
Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

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
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 200 Then Exit Sub
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
Private Sub cmdReset_Click()
'Call ClearForm_AllField(Me)
txtOrderDateS = ""
txtOrderDateE = ""
txtDeliveryDateS = ""
txtDeliveryDateE = ""

End Sub
Private Sub cmdQueryT1_Click()

If (Len(txtDeliveryDateST1) = 0 And Len(txtDeliveryDateET1) = 0) And (Len(txtOrderDateST1) = 0 And Len(txtOrderDateET1) = 0) Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub

'�f�D�ˬd
If Len(RTrim(Combo2.Text)) = 0 Then MsgBox "�п�J�f�D�s��", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMainT1.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String, chc_DeliveryDate As String, chc_Storerkey As String

''�f�D�s��
'If Len(RTrim(Combo1.Text)) > 0 Then chc_Storerkey = " and �f�D ='" & Combo1.Text & "' "

'�q����
chc_Orderdate = ""
If Len(txtOrderDateST1) > 0 And Len(txtOrderDateET1) > 0 Then
   chc_Orderdate = "and �q��� between '" & txtOrderDateST1 & "' and '" & txtOrderDateET1 & "' "
ElseIf Len(txtOrderDateST1) > 0 And Len(txtOrderDateET1) = 0 Then
   chc_Orderdate = "and �q��� = '" & txtOrderDateST1 & "' "
ElseIf Len(txtOrderDateST1) = 0 And Len(txtOrderDateET1) > 0 Then
   chc_Orderdate = "and �q��� = '" & txtOrderDateET1 & "' "
End If

'��f���
chc_DeliveryDate = ""
If Len(txtDeliveryDateST1) > 0 And Len(txtDeliveryDateET1) > 0 Then
   chc_Orderdate = "and ��f�� between '" & txtDeliveryDateST1 & "' and '" & txtDeliveryDateET1 & "' "
ElseIf Len(txtDeliveryDateST1) > 0 And Len(txtDeliveryDateET1) = 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateST1 & "' "
ElseIf Len(txtDeliveryDateST1) = 0 And Len(txtDeliveryDateET1) > 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateET1 & "' "
End If

str_SQL = "select * from gv_sdn05tdetail where 1 = 1 " & chc_Orderdate & chc_DeliveryDate

'�f�D
If Len(RTrim(Combo2.Text)) > 0 Then str_SQL = str_SQL & "and �f�D = '" & RTrim(Combo2.Text) & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
tmp_Rs.Sort = "��f��,���u�s��,�f�D�渹"

Set rsMainT1 = New ADODB.Recordset

Call OffLineRecordset(tmp_Rs, rsMainT1)
tmp_Rs.Sort = "": tmp_Rs.Close

rsMainT1.MoveFirst
Set dgMainT1.DataSource = rsMainT1

SetDataGridColWidth Me.Caption, dgMainT1
StatusBar.Panels(2).Text = rsMainT1.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT1.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdQueryT2_Click()

If (Len(txtDeliveryDateST2) = 0 And Len(txtDeliveryDateET2) = 0) Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub

'�f�D�ˬd
If Len(RTrim(Combo3.Text)) = 0 Then MsgBox "�п�J�f�D�s��", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMainT2.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String, chc_DeliveryDate As String, chc_Storerkey As String

''�f�D�s��
'If Len(RTrim(Combo1.Text)) > 0 Then chc_Storerkey = " and �f�D ='" & Combo1.Text & "' "

'�q����
chc_Orderdate = ""
If Len(txtOrderDateST2) > 0 And Len(txtOrderDateET2) > 0 Then
   chc_Orderdate = "and �q��� between '" & txtOrderDateST2 & "' and '" & txtOrderDateET2 & "' "
ElseIf Len(txtOrderDateST2) > 0 And Len(txtOrderDateET2) = 0 Then
   chc_Orderdate = "and �q��� = '" & txtOrderDateST2 & "' "
ElseIf Len(txtOrderDateST2) = 0 And Len(txtOrderDateET2) > 0 Then
   chc_Orderdate = "and �q��� = '" & txtOrderDateET2 & "' "
End If

'��f���
chc_DeliveryDate = ""
If Len(txtDeliveryDateST2) > 0 And Len(txtDeliveryDateET2) > 0 Then
   chc_Orderdate = "and ��f�� between '" & txtDeliveryDateST2 & "' and '" & txtDeliveryDateET2 & "' "
ElseIf Len(txtDeliveryDateST2) > 0 And Len(txtDeliveryDateET2) = 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateST2 & "' "
ElseIf Len(txtDeliveryDateST2) = 0 And Len(txtDeliveryDateET2) > 0 Then
   chc_Orderdate = "and ��f�� = '" & txtDeliveryDateET2 & "' "
End If

str_SQL = "select * from gv_sdn05tdetail where 1 = 1 " & chc_Orderdate & chc_DeliveryDate

'�f�D
If Len(RTrim(Combo3.Text)) > 0 Then str_SQL = str_SQL & "and �f�D = '" & RTrim(Combo3.Text) & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
tmp_Rs.Sort = "��f��,���u�s��,�f�D�渹"

Set rsMainT2 = New ADODB.Recordset

Call OffLineRecordset(tmp_Rs, rsMainT2)
tmp_Rs.Sort = "": tmp_Rs.Close

rsMainT2.MoveFirst
Set dgMainT2.DataSource = rsMainT2

SetDataGridColWidth Me.Caption, dgMainT2
StatusBar.Panels(2).Text = rsMainT2.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMainT2.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd2ExcelT1_Click()

If (Len(txtDeliveryDateST1) = 0 And Len(txtDeliveryDateET1) = 0) And (Len(txtOrderDateST1) = 0 And Len(txtOrderDateET1) = 0) Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub

MsgBox "�t�ζi��j�q�����Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel"

On Error GoTo err_Handle
Screen.MousePointer = 11
Dim rsTmp As New ADODB.Recordset
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    
    If Dir(App.Path & "\XLT\�Ȱ��ХI�ک���.xlt") = "" Then '�䤣�쥻���d����
        
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
    If Dir(strXltPath & "\�Ȱ��ХI�ک���.xlt") <> "" Then
        
        '�}�ҽd����
        .Workbooks.Open (strXltPath & "\�Ȱ��ХI�ک���.xlt")
    Else
        '�s�WExcel
        .Workbooks.Add
    End If
    
.ActiveWorkbook.Author = User_id

'���_�p�O���Ӹ��
'�M��u�@��
strSheet = "�B�O����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then .Sheets.Add: .ActiveSheet.Name = strSheet

Call WriteOut_RunLog("�B��дڡG1/7.�B�O���Ӹ��..")
rsMainT1.MoveFirst

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

'�����
'�M��u�@��
strSheet = "�����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_Charge2 '" & Combo2 & "' , '" & txtOrderDateST1 & "','" & txtOrderDateET1 & "','" & txtDeliveryDateST1 & "','" & txtDeliveryDateET1 & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("�B��дڡG2/7.��X�����..")
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
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

'�t�e
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�t�e"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LABT01ShipAR '" & txtOrderDateST1 & "' , '" & txtOrderDateET1 & "','" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG3/7.��X�t�e�O...")
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
strSheet = "�h�f"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LABT01ROrderAR '" & txtOrderDateST1 & "' , '" & txtOrderDateET1 & "','" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG4/7.��X�h�f�O....")
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

'�g�P�t�e
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�g�P"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LABT01XShipAR '" & txtOrderDateST1 & "' , '" & txtOrderDateET1 & "','" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG5/7.��X�g�P�t�e�O...")
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

'����
Screen.MousePointer = 11
'�M��u�@��
strSheet = "����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LABT01XShipAR2 '" & txtOrderDateST1 & "' , '" & txtOrderDateET1 & "','" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG6/7.��X���ƱM���O....")
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


'�����I
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�����I"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LABT01ARP '" & Trim(Combo2.Text) & "','" & txtOrderDateST1 & "' , '" & txtOrderDateET1 & "','" & txtDeliveryDateST1 & "' , '" & txtDeliveryDateET1 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG7/7.��X�����I....")
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
.Visible = True: End With

Call Unload_RunLogForm
Set MyXlsApp = Nothing
Screen.MousePointer = 0

Exit Sub

err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cmd2ExcelT2APP_Click()

If (Len(txtDeliveryDateST2) = 0 And Len(txtDeliveryDateET2) = 0) Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub

MsgBox "�t�ζi��j�q�����Excel�ɡA�Фžާ@��LExcel�@�~�A�H�K�����X���~�I", vbOKOnly + vbInformation, "Save2Excel"

On Error GoTo err_Handle
Screen.MousePointer = 11
Dim rsTmp As New ADODB.Recordset
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String
Dim countS As Long, countE As Long
countS = 1: countE = 18
'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    
    If Dir(App.Path & "\XLT\�����@�ХI�ک���.xlt") = "" Then '�䤣�쥻���d����
        
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
    If Dir(strXltPath & "\�����@�ХI�ک���.xlt") <> "" Then
        
        '�}�ҽd����
        .Workbooks.Open (strXltPath & "\�����@�ХI�ک���.xlt")
    Else
        '�s�WExcel
        .Workbooks.Add
    End If
    
.ActiveWorkbook.Author = User_id

'�B�O���Ӹ��
'�M��u�@��
strSheet = "�B�O����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then .Sheets.Add: .ActiveSheet.Name = strSheet

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " �B�O���Ӹ��..")
countS = countS + 1
rsMainT2.MoveFirst

'�g�J���D�C
k = 65: j = 1
For i = 0 To rsMainT2.Fields.Count - 1
    l = i Mod 26
    .Range(strCol & Chr(k + l) & j).Value = rsMainT2.Fields(i).Name
    '���W�L26
    If Chr(65 + l) = "Z" Then
        If strCol = "" Then
            strCol = "A"
        Else
            strCol = Chr(Asc(strCol) + 1)
        End If
    End If
Next i

.Range("A2").CopyFromRecordset rsMainT2

'�����
'�M��u�@��
strSheet = "�����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_Charge '" & Combo3 & "' , '" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�����..")
countS = countS + 1
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
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

'�������Ӹ��
strSheet = "�t�e"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01IAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
            
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�t�e�O")
countS = countS + 1
tmp_Rs.Open str_SQL, cn
Call OffLineRecordset(tmp_Rs, rsTmp)

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


'�t�e(�~��)
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�t�e(�~��)"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_LAPP01IAR_Lotion '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X �t�e(�~��)....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�ռ��O
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�ռ��O"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01AAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�ռ��O....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�ռ��O(�~��)
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�ռ��O(�~��)"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_LAPP01AAR_Lotion '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X �ռ��O(�~��)....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'���t
Screen.MousePointer = 11
'�M��u�@��
strSheet = "���t"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01A2BAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X���t...")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'���t(�~��)
Screen.MousePointer = 11
'�M��u�@��
strSheet = "���t(�~��)"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_LAPP01A2BAR_Lotion '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X ���t(�~��)....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�i�f
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�i�f"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01RCAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�i�f�O....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�i�f(�~��)
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�i�f(�~��)"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_LAPP01RCAR_Lotion '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X �i�f(�~��)....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�P�h����
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�P�h����"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01RAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�P�h���ӶO....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�P�h����(�~��)
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�P�h����(�~��)"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_LAPP01RAR_Lotion '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X �P�h����(�~��)....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�ڦ�
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�ڦ�"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01CancelAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�ڦ��O....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�ڦ�(�~��)
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�ڦ�(�~��)"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_LAPP01CancelAR_Lotion '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X �ڦ�(�~��)....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�ͬ��ʳf
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�ͬ��ʳf"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01LifeStoreAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�ͬ��ʳf....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

'�ͬ��ʳf(�~��)
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�ͬ��ʳf(�~��)"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_LAPP01LifeStoreAR_Lotion '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X �ͬ��ʳf(�~��)....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

''�W�U�ӸɶK
'Screen.MousePointer = 11
''�M��u�@��
'strSheet = "�W�U�ӸɶK"
'For i = 1 To .Sheets.Count
'    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
'Next
'
''�䤣��s�W�u�@��
'If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
'
'str_SQL = "exec gs_LAPP01StairsAR '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
'
'Call Confirm_Recordset_Closed(tmp_Rs)
'
'Call WriteOut_RunLog("�B��дڡG10/10��X�W�U�ӸɶK....")
'tmp_Rs.Open str_SQL, cn
'
'Call OffLineRecordset(tmp_Rs, rsTmp)
'
''�g�J���D�C
'k = 65: j = 1: strCol = ""
'For i = 0 To rsTmp.Fields.Count - 1
'    l = i Mod 26
'    .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
'    '���W�L26
'    If Chr(65 + l) = "Z" Then
'        If strCol = "" Then
'            strCol = "A"
'        Else
'            strCol = Chr(Asc(strCol) + 1)
'        End If
'    End If
'Next i
'
'.Range("A2").CopyFromRecordset rsTmp
'
'rsTmp.Close

'�i�f���ӰѦ�
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�i�f���ӰѦ�"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_LAPP01ReceiptDetail '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�i�f���ӰѦ�....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�����I"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec es_OtherARP '" & Combo3.Text & "','" & txtDeliveryDateST2 & "','" & txtDeliveryDateET2 & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

Call WriteOut_RunLog("�B��дڡG" & countS & "/" & countE & " ��X�����I....")
countS = countS + 1
tmp_Rs.Open str_SQL, cn

Call OffLineRecordset(tmp_Rs, rsTmp)

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

.Visible = True: End With

Call Unload_RunLogForm
Set MyXlsApp = Nothing
Screen.MousePointer = 0

Exit Sub

err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub txtOrderDateST1_Click()

Set objMvdateTarget = txtOrderDateST1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateST2_Click()

Set objMvdateTarget = txtOrderDateST2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtOrderDateET1_Click()

Set objMvdateTarget = txtOrderDateET1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateET2_Click()

Set objMvdateTarget = txtOrderDateET2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateST1_Click()

Set objMvdateTarget = txtDeliveryDateST1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtDeliveryDateST2_Click()

Set objMvdateTarget = txtDeliveryDateST2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET1_Click()

Set objMvdateTarget = txtDeliveryDateET1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateET2_Click()

Set objMvdateTarget = txtDeliveryDateET2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtOrderDateST1_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtOrderDateST2_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub

Private Sub txtOrderDateET1_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtOrderDateET2_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtDeliveryDateST1_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtDeliveryDateST2_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtDeliveryDateET1_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtDeliveryDateET2_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub

Private Sub dgMainT1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT1
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 200 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub
Private Sub dgMainT2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT2
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 200 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
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
Private Sub cmdResetT1_Click()
txtOrderDateST1 = ""
txtOrderDateET1 = ""
txtDeliveryDateST1 = ""
txtDeliveryDateET1 = ""
End Sub
Private Sub cmdResetT2_Click()
txtOrderDateST2 = ""
txtOrderDateET2 = ""
txtDeliveryDateST2 = ""
txtDeliveryDateET2 = ""
End Sub
