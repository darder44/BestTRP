VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_OP_CarControl 
   Caption         =   "�����i�X�ި�@�~"
   ClientHeight    =   7140
   ClientLeft      =   180
   ClientTop       =   795
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11475
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   4680
      TabIndex        =   90
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
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      StartOfWeek     =   92667905
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "��������"
      TabPicture(0)   =   "frm_OP_CarControl.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fam_Tab0_CarData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dg_Tab0_CarCheckin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd_Exit(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ado_CarCheckin"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd_Tab0_CarList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fam_Tab0_CarCheckin"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fam_Tab0_Query"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_Tab0_ShowQuery"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "��������"
      TabPicture(1)   =   "frm_OP_CarControl.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fam_Tab1_Query"
      Tab(1).Control(1)=   "cmd_Tab1_ShowQuery"
      Tab(1).Control(2)=   "fam_Tab1_CarData"
      Tab(1).Control(3)=   "cmd_Exit(1)"
      Tab(1).Control(4)=   "cmd_Tab1_CarList"
      Tab(1).Control(5)=   "fam_Tab1_CarCheckout"
      Tab(1).Control(6)=   "ado_CarCheckout"
      Tab(1).Control(7)=   "dg_Tab1_CarCheckout"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "�d����ƾ�z"
      TabPicture(2)   =   "frm_OP_CarControl.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fam_SelectedOrders"
      Tab(2).Control(1)=   "fam_SrcOrders"
      Tab(2).ControlCount=   2
      Begin VB.Frame fam_SrcOrders 
         Caption         =   "�ݽT�{���s"
         Height          =   2835
         Left            =   -74880
         TabIndex        =   99
         Top             =   4080
         Width           =   11220
         Begin MSDataGridLib.DataGrid dg_Tab2_Route 
            Height          =   2520
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   4445
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
      Begin VB.Frame fam_SelectedOrders 
         Caption         =   "�ݽT�{�d�����"
         Height          =   3735
         Left            =   -74880
         TabIndex        =   91
         Top             =   360
         Width           =   11220
         Begin VB.TextBox txt_Tab2_OutTime 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6780
            TabIndex        =   105
            Top             =   3315
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab2_InTime 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3900
            TabIndex        =   103
            Top             =   3315
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab2_Route_NO 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1050
            TabIndex        =   101
            Top             =   3315
            Width           =   1470
         End
         Begin VB.CommandButton cmd_Tab2_Selected 
            BackColor       =   &H00FF8080&
            Caption         =   "V �s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   97
            Top             =   2790
            Width           =   945
         End
         Begin VB.CommandButton cmd_Tab2_srcOrderReset 
            Appearance      =   0  '����
            BackColor       =   &H00808080&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7770
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   95
            Top             =   2790
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab2_SelectedCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3280
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   94
            Top             =   2790
            Width           =   1260
         End
         Begin VB.CommandButton cmd_Tab2_ImportCard 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���J�ݾ�z���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1205
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   93
            Top             =   2790
            Width           =   2055
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H008080FF&
            Caption         =   "��  �}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4560
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   92
            Top             =   2790
            Width           =   1110
         End
         Begin MSDataGridLib.DataGrid dg_Tab2_CardIn 
            Height          =   2505
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   4419
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
         Begin MSDataGridLib.DataGrid dg_Tab2_CardOut 
            Height          =   2505
            Left            =   5640
            TabIndex        =   107
            Top             =   240
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   4419
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
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
         Begin VB.CommandButton cmd_Tab2_RouteQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���s�j�M"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6645
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   96
            Top             =   2790
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���t�ɶ����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   18
            Left            =   5520
            TabIndex        =   106
            Top             =   3360
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����ɶ����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   17
            Left            =   2640
            TabIndex        =   104
            Top             =   3360
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���u�s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   16
            Left            =   150
            TabIndex        =   102
            Top             =   3360
            Width           =   840
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   435
            Index           =   1
            Left            =   120
            Top             =   3240
            Width           =   8175
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   435
            Index           =   0
            Left            =   120
            Top             =   2760
            Width           =   5655
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00FF0000&
            FillStyle       =   0  '���
            Height          =   435
            Left            =   6615
            Top             =   2760
            Width           =   1680
         End
      End
      Begin VB.Frame fam_Tab1_Query 
         BackColor       =   &H00404000&
         Caption         =   "�z�����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   -74610
         TabIndex        =   82
         Top             =   1185
         Visible         =   0   'False
         Width           =   2910
         Begin VB.CommandButton cmd_Tab1_Default 
            BackColor       =   &H00FFC0FF&
            Caption         =   "�w�]"
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
            Left            =   2250
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   88
            Top             =   195
            Width           =   570
         End
         Begin VB.CheckBox chk_Tab1_Checkin 
            Caption         =   "�z��w���쨮��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   300
            TabIndex        =   85
            Top             =   330
            Value           =   1  '�֨�
            Width           =   1800
         End
         Begin VB.TextBox txt_Tab1_QueryDate 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1170
            TabIndex        =   84
            Top             =   600
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab1_QueryCarID 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1155
            TabIndex        =   83
            Top             =   960
            Width           =   1470
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   15
            Left            =   255
            TabIndex        =   87
            Top             =   645
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���P���X"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   86
            Top             =   1020
            Width           =   840
         End
      End
      Begin VB.CommandButton cmd_Tab1_ShowQuery 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�H"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72720
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   81
         Top             =   825
         Width           =   345
      End
      Begin VB.CommandButton cmd_Tab0_ShowQuery 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�H"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   80
         Top             =   825
         Width           =   360
      End
      Begin VB.Frame fam_Tab0_Query 
         BackColor       =   &H00404000&
         Caption         =   "�z�����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1455
         Left            =   525
         TabIndex        =   74
         Top             =   1185
         Visible         =   0   'False
         Width           =   2910
         Begin VB.CommandButton cmd_Tab0_Default 
            BackColor       =   &H00FFC0FF&
            Caption         =   "�w�]"
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
            Left            =   2205
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   89
            Top             =   210
            Width           =   570
         End
         Begin VB.TextBox txt_Tab0_QueryCarID 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1155
            TabIndex        =   78
            Top             =   960
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab0_QueryDate 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1155
            TabIndex        =   76
            Top             =   600
            Width           =   1470
         End
         Begin VB.CheckBox chk_Tab0_Checkin 
            Caption         =   "�z��ݳ��쨮��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   270
            TabIndex        =   75
            Top             =   270
            Value           =   1  '�֨�
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���P���X"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   79
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   14
            Left            =   255
            TabIndex        =   77
            Top             =   645
            Width           =   840
         End
      End
      Begin VB.Frame fam_Tab1_CarData 
         Appearance      =   0  '����
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   -74760
         TabIndex        =   47
         Top             =   1380
         Width           =   10920
         Begin VB.TextBox txt_Tab1_RouteNo 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9750
            TabIndex        =   61
            Top             =   495
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_CarID 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   60
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Driver 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   59
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   58
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_DriveTimes 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   57
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab1_Phone 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5190
            TabIndex        =   56
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Checkin 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   55
            Top             =   495
            Width           =   2355
         End
         Begin VB.TextBox txt_Tab1_Company 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5010
            TabIndex        =   54
            Top             =   840
            Width           =   2745
         End
         Begin VB.TextBox txt_Tab1_VehicleType 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7770
            TabIndex        =   53
            Top             =   840
            Width           =   3075
         End
         Begin VB.TextBox txt_Tab1_CaseQty 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   52
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Palletin 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   51
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab1_PalletQty 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   50
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Volumn 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   49
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab1_Weight 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckout"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   48
            Top             =   495
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���P���X"
            Height          =   180
            Index           =   1
            Left            =   195
            TabIndex        =   71
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�r�p�H"
            Height          =   180
            Index           =   13
            Left            =   2490
            TabIndex        =   70
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            Height          =   180
            Index           =   7
            Left            =   195
            TabIndex        =   69
            Top             =   555
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   180
            Index           =   12
            Left            =   525
            TabIndex        =   68
            Top             =   900
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            Height          =   180
            Index           =   11
            Left            =   4755
            TabIndex        =   67
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����ɶ�"
            Height          =   180
            Index           =   6
            Left            =   2310
            TabIndex        =   66
            Top             =   555
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�餽�q/����"
            Height          =   180
            Index           =   10
            Left            =   3840
            TabIndex        =   65
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�c�� / �O��"
            Height          =   180
            Index           =   9
            Left            =   7065
            TabIndex        =   64
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��J�̪O��"
            Height          =   180
            Index           =   5
            Left            =   2160
            TabIndex        =   63
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n / ���q"
            Height          =   180
            Index           =   8
            Left            =   7065
            TabIndex        =   62
            Top             =   555
            Width           =   855
         End
      End
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "��  �}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -65025
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   45
         Top             =   705
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Tab1_CarList 
         BackColor       =   &H8000000A&
         Caption         =   "���J�w���쨮��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -74625
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   44
         Top             =   585
         Width           =   1905
      End
      Begin VB.Frame fam_Tab1_CarCheckout 
         Height          =   930
         Left            =   -72165
         TabIndex        =   37
         Top             =   405
         Width           =   6120
         Begin VB.CommandButton cmd_Tab1_CheckOutSave 
            BackColor       =   &H00FF8080&
            Caption         =   "�s  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4875
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   73
            Top             =   135
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab1_PalletOut 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1290
            TabIndex        =   41
            Top             =   510
            Width           =   660
         End
         Begin VB.TextBox txt_Tab1_Checkout 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   360
            Left            =   1290
            TabIndex        =   40
            Top             =   135
            Width           =   2340
         End
         Begin VB.CommandButton cmd_Tab1_Checkout 
            BackColor       =   &H008080FF&
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3660
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   39
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab1_ClearCheckin 
            BackColor       =   &H00FF8080&
            Caption         =   "��  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2445
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   38
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���ܤ��/�ɶ�"
            Height          =   180
            Index           =   4
            Left            =   105
            TabIndex        =   43
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��X�̪O��"
            Height          =   180
            Index           =   7
            Left            =   330
            TabIndex        =   42
            Top             =   600
            Width           =   900
         End
      End
      Begin VB.Frame fam_Tab0_CarCheckin 
         BackColor       =   &H00400000&
         Height          =   930
         Left            =   2835
         TabIndex        =   18
         Top             =   420
         Width           =   6135
         Begin VB.CommandButton cmd_Tab0_CheckinSave 
            BackColor       =   &H00FF8080&
            Caption         =   "�s  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4875
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   72
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab0_ClearCheckin 
            BackColor       =   &H00FF8080&
            Caption         =   "��  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2430
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   35
            Top             =   495
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab0_Checkin 
            BackColor       =   &H008080FF&
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3660
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   34
            Top             =   135
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab0_Checkin 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   360
            Left            =   1290
            TabIndex        =   22
            Top             =   135
            Width           =   2340
         End
         Begin VB.TextBox txt_Tab0_PalletIN 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1290
            TabIndex        =   21
            Top             =   510
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��J�̪O��"
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   3
            Left            =   330
            TabIndex        =   20
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "������/�ɶ�"
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   2
            Left            =   105
            TabIndex        =   19
            Top             =   225
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmd_Tab0_CarList 
         BackColor       =   &H8000000A&
         Caption         =   "���J�ݳ��쨮��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   525
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
         Top             =   570
         Width           =   1905
      End
      Begin MSAdodcLib.Adodc ado_CarCheckin 
         Height          =   405
         Left            =   330
         Top             =   6420
         Visible         =   0   'False
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "ado_CarCheckin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "��  �}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   9960
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   720
         Width           =   1035
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_CarCheckin 
         Height          =   4245
         Left            =   225
         TabIndex        =   1
         Top             =   2595
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   7488
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
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
      Begin VB.Frame fam_Tab0_CarData 
         Appearance      =   0  '����
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   225
         TabIndex        =   4
         Top             =   1395
         Width           =   10920
         Begin VB.TextBox txt_Tab0_Weight 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   33
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_Volumn 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   31
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_PalletQty 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9120
            TabIndex        =   30
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_DockNo 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   29
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab0_CaseQty 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7965
            TabIndex        =   26
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_VehicleType 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7770
            TabIndex        =   25
            Top             =   840
            Width           =   3075
         End
         Begin VB.TextBox txt_Tab0_Company 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5010
            TabIndex        =   24
            Top             =   840
            Width           =   2745
         End
         Begin VB.TextBox txt_Tab0_ExpectTime 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5190
            TabIndex        =   17
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_ExpectDate 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4035
            TabIndex        =   16
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_Phone 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5190
            TabIndex        =   14
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_DriveTimes 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   12
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txt_Tab0_DeliveryDate 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   10
            Top             =   495
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_Driver 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            TabIndex        =   8
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_CarID 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   975
            TabIndex        =   6
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_RouteNo 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            DataSource      =   "ado_CarCheckin"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9750
            TabIndex        =   36
            Top             =   495
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n / ���q"
            Height          =   180
            Index           =   6
            Left            =   7065
            TabIndex        =   32
            Top             =   555
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�Y�Ȧs"
            Height          =   180
            Index           =   3
            Left            =   2340
            TabIndex        =   28
            Top             =   900
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�c�� / �O��"
            Height          =   180
            Index           =   5
            Left            =   7065
            TabIndex        =   27
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�餽�q/����"
            Height          =   180
            Index           =   4
            Left            =   3840
            TabIndex        =   23
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�w�p������/�ɶ�"
            Height          =   180
            Index           =   1
            Left            =   2490
            TabIndex        =   15
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            Height          =   180
            Index           =   2
            Left            =   4755
            TabIndex        =   13
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   525
            TabIndex        =   11
            Top             =   900
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            Height          =   180
            Index           =   0
            Left            =   195
            TabIndex        =   9
            Top             =   555
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�r�p�H"
            Height          =   180
            Index           =   0
            Left            =   2490
            TabIndex        =   7
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���P���X"
            Height          =   180
            Index           =   0
            Left            =   195
            TabIndex        =   5
            Top             =   240
            Width           =   720
         End
      End
      Begin MSAdodcLib.Adodc ado_CarCheckout 
         Height          =   405
         Left            =   -74655
         Top             =   6405
         Visible         =   0   'False
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "ado_CarCheckOut"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_CarCheckout 
         Height          =   4245
         Left            =   -74760
         TabIndex        =   46
         Top             =   2595
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   7488
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
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
End
Attribute VB_Name = "frm_OP_CarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private blTab0CarListEvent As Boolean    '��������G�ƥ󱱨�X��
Private blCancelChangeRecord As Boolean  '��������G�O�_�i�H����U�����O��
Private blTab1CarListEvent As Boolean    '��������G�ƥ󱱨�X��
Private CardChange As Boolean            '�ƥ󱱨�X��

Private rs_Tab0_CarCheckin As ADODB.Recordset    '�ݳ��줧�����C��
Private rs_Tab1_CarCheckOut As ADODB.Recordset   '�ݳ��줧�����C��
Private rs_Tab2_CardIn As ADODB.Recordset
Private rs_Tab2_CardOut As ADODB.Recordset
Private rs_Tab2_Route As ADODB.Recordset


Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_Tab0_CarList_Click()
'�������� >> ���J�ݳ��쨮��
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
fam_Tab0_Query.Visible = False
Set dg_Tab0_CarCheckin.DataSource = Nothing
Set rs_Tab0_CarCheckin = Nothing
txt_Tab0_Checkin.Text = ""          '������/�ɶ�
txt_Tab0_PalletIN.Text = ""         '��J�̪O��

str_SQL = "Select �X�����,���P���X,����,����h��,����ɶ�,��J�̪O��,�r�p�H,�q��,�c��,�O��,���n,���q,�B�餽�q," & _
          "    ����,�w�p������,�w�p����ɶ�,�X�Y�Ȧs,���u�s��,���ܮɶ� " & _
          "From CarControl_srcCheckin "

Dim strWhere As String, strTmp As String, tmp_data() As String, intloop As Integer
strWhere = ""
'�z�����ɶ�
strTmp = ""
If chk_Tab0_Checkin.Value = vbChecked Then
   strTmp = " ����ɶ� = '' "
   If Len(strTmp) > 0 Then
      If Len(strWhere) = 0 Then
         strWhere = strTmp
      Else
         strWhere = strWhere & " and " & strTmp
      End If
   End If
End If
'�X�����
strTmp = ""
If Len(txt_Tab0_QueryDate.Text) > 0 Then
   strTmp = " �X����� = '" & strTmp & "' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'���P���X
strTmp = ""
If Len(txt_Tab0_QueryCarID.Text) > 0 Then
   tmp_data = Split(txt_Tab0_QueryCarID.Text, ",", -1, vbTextCompare)
   For intloop = LBound(tmp_data) To UBound(tmp_data)
       If Len(strTmp) = 0 Then
          strTmp = "'" & tmp_data(intloop) & "'"
       Else
          strTmp = strTmp & ",'" & tmp_data(intloop) & "'"
       End If
   Next intloop
   strTmp = " ���P���X in (" & strTmp & ") "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
If Len(strWhere) > 0 Then
   str_SQL = str_SQL & " Where " & strWhere
End If
str_SQL = str_SQL & " Order by �X�����,���P���X "

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݳ�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call CreateRS_Tab0_Checkin
Do While Not tmp_Rs.EOF
   With rs_Tab0_CarCheckin
     .AddNew
     .Fields("�s��") = .RecordCount
     .Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
     .Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
     .Fields("����").Value = tmp_Rs.Fields("����").Value
     .Fields("����h��").Value = tmp_Rs.Fields("����h��").Value
     .Fields("����ɶ�").Value = tmp_Rs.Fields("����ɶ�").Value
     .Fields("��J�̪O").Value = tmp_Rs.Fields("��J�̪O��").Value
     .Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
     .Fields("�q��").Value = tmp_Rs.Fields("�q��").Value
     .Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
     .Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
     .Fields("���n").Value = tmp_Rs.Fields("���n").Value
     .Fields("���q").Value = tmp_Rs.Fields("���q").Value
     .Fields("�B�餽�q").Value = tmp_Rs.Fields("�B�餽�q").Value
     .Fields("����").Value = tmp_Rs.Fields("����").Value
     .Fields("�w�p������").Value = tmp_Rs.Fields("�w�p������").Value
     .Fields("�w�p����ɶ�").Value = tmp_Rs.Fields("�w�p����ɶ�").Value
     .Fields("�X�Y�Ȧs").Value = tmp_Rs.Fields("�X�Y�Ȧs").Value
     .Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
     .Fields("���ܮɶ�").Value = tmp_Rs.Fields("���ܮɶ�").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab0_CarCheckin.MoveFirst
Set dg_Tab0_CarCheckin.DataSource = rs_Tab0_CarCheckin
With dg_Tab0_CarCheckin
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
blTab0CarListEvent = False
Set dg_Tab0_CarCheckin.DataSource = rs_Tab0_CarCheckin
'�]�w������
With dg_Tab0_CarCheckin
    .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
    .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
    .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
    .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
    .Columns(0).Width = 500         '�s��
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000        '�X�����
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900         '���P���X
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 450         '����
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 850         '����h��
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1700        '����ɶ�
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800         '��J�̪O
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 700         '�r�p�H
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1000        '�q��
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 800         '�c��
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800        '�O��
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 800        '���n
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800        '���q
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 1500       '�B�餽�q
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1500       '����
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1300       '�w�p������
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1300       '�w�p����ɶ�
    .Columns(16).Alignment = dbgLeft
    .Columns(17).Width = 900        '�X�Y�Ȧs
    .Columns(17).Alignment = dbgLeft
    .Columns(18).Width = 1200       '���u�s��
    .Columns(18).Alignment = dbgLeft

End With
Set ado_CarCheckin.Recordset = rs_Tab0_CarCheckin
txt_Tab0_CarID.DataField = "���P���X"
txt_Tab0_Driver.DataField = "�r�p�H"
txt_Tab0_Phone.DataField = "�q��"
txt_Tab0_DeliveryDate.DataField = "�X�����"
txt_Tab0_ExpectDate.DataField = "�w�p������"
txt_Tab0_ExpectTime.DataField = "�w�p����ɶ�"
txt_Tab0_DriveTimes.DataField = "����"
txt_Tab0_DockNo.DataField = "�X�Y�Ȧs"
txt_Tab0_Company.DataField = "�B�餽�q"
txt_Tab0_VehicleType.DataField = "����"
txt_Tab0_CaseQty.DataField = "�c��"
txt_Tab0_PalletQty.DataField = "�O��"
txt_Tab0_Volumn.DataField = "���n"
txt_Tab0_Weight.DataField = "���q"
txt_Tab0_Checkin.DataField = "����ɶ�"
txt_Tab0_PalletIN.DataField = "��J�̪O"
txt_Tab0_RouteNo.DataField = "���u�s��"

'�w�]���
If Not rs_Tab0_CarCheckin.EOF Then
   dg_Tab0_CarCheckin.SelBookmarks.Add dg_Tab0_CarCheckin.Bookmark
   txt_Tab0_Checkin.Text = rs_Tab0_CarCheckin.Fields("����ɶ�").Value
   txt_Tab0_PalletIN.Text = rs_Tab0_CarCheckin.Fields("��J�̪O").Value
End If

blTab0CarListEvent = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-���J�ݳ��쨮��", Me.Caption, "cmd_Tab0_CarList_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Checkin_Click()
'�������� >> ��������
If rs_Tab0_CarCheckin Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab0_CarCheckin.SelBookmarks.Count = 0 Then
   msg_text = "������ݳ��쨮��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
'���o sql serer �t�ήɶ�
str_SQL = "Select Convert(varchar,Getdate(),111) as CheckinDate , Convert(varchar,Getdate(),108) as CheckinTime "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
txt_Tab0_Checkin.Text = tmp_Rs.Fields("CheckinDate").Value & " " & tmp_Rs.Fields("CheckinTime").Value
tmp_Rs.Close
'�� [�s��] �~�N�ɶ��B�̪O�� ��ܩ� [�ݳ��쨮���C��]
'rs_Tab0_CarCheckin.Fields("����ɶ�").Value = txt_Tab0_Checkin.Text
'rs_Tab0_CarCheckin.Fields("��J�̪O").Value = Val(txt_Tab0_PalletIN.Text)
txt_Tab0_PalletIN.SelStart = 0: txt_Tab0_PalletIN.SelLength = Len(txt_Tab0_PalletIN.Text)
txt_Tab0_PalletIN.SetFocus
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-��������", Me.Caption, "cmd_Tab0_Checkin_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_CheckinSave_Click()
'�������� >> �s��
If rs_Tab0_CarCheckin Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab0_CarCheckin.SelBookmarks.Count = 0 Then
   msg_text = "������ݳ��쨮��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

   rs_Tab0_CarCheckin.Fields("����ɶ�").Value = txt_Tab0_Checkin.Text
   rs_Tab0_CarCheckin.Fields("��J�̪O").Value = Val(txt_Tab0_PalletIN.Text)

   If Len(Trim(txt_Tab0_Checkin.Text)) > 0 Then
      '�ˬd�������
      If Fun_ChkDateFormat2(Left(txt_Tab0_Checkin.Text, 10)) = 1 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  ��������ɶ���Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ�������ɶ��@����J�Ѧ�-[���]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�ˬd�ɶ������G��
      If Val(Mid(txt_Tab0_Checkin.Text, 12, 2)) < 0 Or Val(Mid(txt_Tab0_Checkin.Text, 12, 2)) >= 24 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  ��������ɶ���Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ�������ɶ��@����J�Ѧ�-[��]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�ˬd�ɶ������G��
      If Val(Mid(txt_Tab0_Checkin.Text, 15, 2)) < 0 Or Val(Mid(txt_Tab0_Checkin.Text, 15, 2)) > 59 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  ��������ɶ���Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ�������ɶ��@����J�Ѧ�-[��]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�ˬd�ɶ������G��
      If Val(Mid(txt_Tab0_Checkin.Text, 18, 2)) < 0 Or Val(Mid(txt_Tab0_Checkin.Text, 18, 2)) > 59 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  ��������ɶ��G��Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ�������ɶ��@����J�Ѧ�-[��]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�N�����Ƽg�^ TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If

      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_IN"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab0_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab0_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab0_DriveTimes.Text)
      'VEHICLE_CHECK_IN
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_CHECK_IN", adChar, adParamInput, 20)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_CHECK_IN").Value = txt_Tab0_Checkin.Text
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("PALLET_IN", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("PALLET_IN").Value = Val(txt_Tab0_PalletIN.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '�D�P�B����
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
      Loop
      Set tmp_Cmd = Nothing
   Else
      msg_text = "��ƿ��~�G����J��������ɶ�"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      Exit Sub
   End If
Exit Sub

err_Handle:
    If Not (tmp_Cmd Is Nothing) Then
       Set tmp_Cmd = Nothing
    End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-�s��", Me.Caption, "cmd_Tab0_CheckinSave_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Default_Click()
'���J�ݳ��쨮�� >> �z�����]�w >> �w�]
chk_Tab0_Checkin.Value = vbChecked
txt_Tab0_QueryDate.Text = ""
txt_Tab0_QueryCarID.Text = ""
End Sub

Private Sub cmd_Tab0_ImportOrders_Click()
    
End Sub

Private Sub cmd_Tab0_ShowQuery_Click()
'���J�ݳ��쨮�� >> ��ܿz�����]�w
fam_Tab0_Query.Visible = Not fam_Tab0_Query.Visible
End Sub

Private Sub cmd_Tab1_CheckOutSave_Click()
'�������� >> �s��
If rs_Tab1_CarCheckOut Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab1_CarCheckout.SelBookmarks.Count = 0 Then
   msg_text = "����������ܨ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
   rs_Tab1_CarCheckOut.Fields("���ܮɶ�").Value = txt_Tab1_Checkout.Text
   rs_Tab1_CarCheckOut.Fields("��X�̪O").Value = Val(txt_Tab1_PalletOut.Text)

   If Len(Trim(txt_Tab1_Checkout.Text)) > 0 Then
      '�ˬd�������
      If Fun_ChkDateFormat2(Left(txt_Tab1_Checkout.Text, 10)) = 1 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  �������ܮɶ���Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ������ܮɶ��@����J�Ѧ�-[���]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�ˬd�ɶ������G��
      If Val(Mid(txt_Tab1_Checkout.Text, 12, 2)) < 0 Or Val(Mid(txt_Tab1_Checkout.Text, 12, 2)) >= 24 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  �������ܮɶ���Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ������ܮɶ��@����J�Ѧ�-[��]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�ˬd�ɶ������G��
      If Val(Mid(txt_Tab1_Checkout.Text, 15, 2)) < 0 Or Val(Mid(txt_Tab1_Checkout.Text, 15, 2)) > 59 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  �������ܮɶ���Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ������ܮɶ��@����J�Ѧ�-[��]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�ˬd�ɶ������G��
      If Val(Mid(txt_Tab1_Checkout.Text, 18, 2)) < 0 Or Val(Mid(txt_Tab1_Checkout.Text, 18, 2)) > 59 Then
         msg_text = "��ƿ��~�G" & vbCrLf & "  �������ܮɶ��G��Ʈ榡���� yyyy/mm/dd hh:nn:ss�A�Y�����Ѥ��B" & vbCrLf & _
                    "  �Ы� [��������] �s�۰ʲ��ͨ����ɶ��@����J�Ѧ�-[��]" & vbCrLf & _
                    "  �`�N�G[���] �P [�ɶ�] ���j�������Ů�"
         MsgBox msg_text, vbOKOnly + vbCritical, msg_title
         Exit Sub
      End If
      '�N�����Ƽg�^ TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If

      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_OUT"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab1_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab1_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab1_DriveTimes.Text)
      'VEHICLE_CHECK_OUT
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_CHECK_OUT", adChar, adParamInput, 20)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_CHECK_OUT").Value = txt_Tab1_Checkout.Text
      'PALLET_OUT
      Set tmp_para = tmp_Cmd.CreateParameter("PALLET_OUT", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("PALLET_OUT").Value = Val(txt_Tab1_PalletOut.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '�D�P�B����
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
      Loop
      Set tmp_Cmd = Nothing
   Else
      msg_text = "��ƿ��~�G����J�������ܮɶ�"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
Exit Sub

err_Handle:
    If Not (tmp_Cmd Is Nothing) Then
       Set tmp_Cmd = Nothing
    End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-�s��", Me.Caption, "cmd_Tab1_CheckinSave_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_ClearCheckin_Click()
'�������� >> �M����������ɶ�
On Error GoTo err_Handle
If rs_Tab0_CarCheckin Is Nothing Then Exit Sub
If dg_Tab0_CarCheckin.SelBookmarks.Count = 0 Then
   msg_text = "��������i��������������"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
txt_Tab0_Checkin.Text = " "
txt_Tab0_PalletIN.Text = 0

      '�N�����ƲM�� TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If
      
      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_IN_Cancel"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab0_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab0_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab0_DriveTimes.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '�D�P�B����
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
      Loop
      Set tmp_Cmd = Nothing
      rs_Tab0_CarCheckin.Fields("����ɶ�").Value = " "
      rs_Tab0_CarCheckin.Fields("��J�̪O").Value = 0
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-����", Me.Caption, "cmd_Tab0_ClearCheckin_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_CarList_Click()
'�������� >> ���J�w���쨮��
On Error GoTo err_Handle
fam_Tab1_Query.Visible = False
Screen.MousePointer = vbHourglass
Set dg_Tab1_CarCheckout.DataSource = Nothing
Set rs_Tab1_CarCheckOut = Nothing

str_SQL = "Select �X�����,���P���X,����,����h��,���ܮɶ�,��X�̪O��,����ɶ�,��J�̪O��,�r�p�H,�q��,�c��,�O��,���n,���q,�B�餽�q,����,���u�s�� " & _
          "From CarControl_srcCheckOut "
Dim strWhere As String, strTmp As String, tmp_data() As String, intloop As Integer
strWhere = ""
'�z�����ɶ�
strTmp = ""
If chk_Tab1_Checkin.Value = vbChecked Then
   strTmp = " ���ܮɶ� = '' "
   If Len(strTmp) > 0 Then
      If Len(strWhere) = 0 Then
         strWhere = strTmp
      Else
         strWhere = strWhere & " and " & strTmp
      End If
   End If
End If
'�X�����
strTmp = ""
If Len(txt_Tab1_QueryDate.Text) > 0 Then
   tmp_data = Split(txt_Tab1_QueryDate.Text, ",", -1, vbTextCompare)
   For intloop = LBound(tmp_data) To UBound(tmp_data)
       If Len(strTmp) = 0 Then
          strTmp = "'" & tmp_data(intloop) & "'"
       Else
          strTmp = strTmp & ",'" & tmp_data(intloop) & "'"
       End If
   Next intloop
   strTmp = " �X����� in (" & strTmp & ") "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'���P���X
strTmp = ""
If Len(txt_Tab1_QueryCarID.Text) > 0 Then
   tmp_data = Split(txt_Tab1_QueryCarID.Text, ",", -1, vbTextCompare)
   For intloop = LBound(tmp_data) To UBound(tmp_data)
       If Len(strTmp) = 0 Then
          strTmp = "'" & tmp_data(intloop) & "'"
       Else
          strTmp = strTmp & ",'" & tmp_data(intloop) & "'"
       End If
   Next intloop
   strTmp = " ���P���X in (" & strTmp & ") "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
If Len(strWhere) > 0 Then
   str_SQL = str_SQL & " Where " & strWhere
End If
str_SQL = str_SQL & " Order by �X�����,���P���X "

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧤w���쨮�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   
   txt_Tab1_CarID.Text = ""
   txt_Tab1_Driver.Text = ""
   txt_Tab1_Phone.Text = ""
   txt_Tab1_DeliveryDate.Text = ""
   txt_Tab1_DriveTimes.Text = ""
   txt_Tab1_Checkin.Text = ""
   txt_Tab1_Palletin.Text = ""
   txt_Tab1_Company.Text = ""
   txt_Tab1_VehicleType.Text = ""
   txt_Tab1_CaseQty.Text = ""
   txt_Tab1_PalletQty.Text = ""
   txt_Tab1_Volumn.Text = ""
   txt_Tab1_Weight.Text = ""
   txt_Tab1_Checkout.Text = ""
   txt_Tab1_PalletOut.Text = ""
   txt_Tab1_RouteNo.Text = ""
   
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call CreateRS_Tab1_CheckOut
Do While Not tmp_Rs.EOF
   With rs_Tab1_CarCheckOut
     .AddNew
     .Fields("�s��") = .RecordCount
     .Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
     .Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
     .Fields("����").Value = tmp_Rs.Fields("����").Value
     .Fields("����h��").Value = tmp_Rs.Fields("����h��").Value
     .Fields("���ܮɶ�").Value = tmp_Rs.Fields("���ܮɶ�").Value
     .Fields("��X�̪O").Value = tmp_Rs.Fields("��X�̪O��").Value
     .Fields("����ɶ�").Value = tmp_Rs.Fields("����ɶ�").Value
     .Fields("��J�̪O").Value = tmp_Rs.Fields("��J�̪O��").Value
     .Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
     .Fields("�q��").Value = tmp_Rs.Fields("�q��").Value
     .Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
     .Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
     .Fields("���n").Value = tmp_Rs.Fields("���n").Value
     .Fields("���q").Value = tmp_Rs.Fields("���q").Value
     .Fields("�B�餽�q").Value = tmp_Rs.Fields("�B�餽�q").Value
     .Fields("����").Value = tmp_Rs.Fields("����").Value
     .Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab1_CarCheckOut.MoveFirst
Set dg_Tab1_CarCheckout.DataSource = rs_Tab1_CarCheckOut
With dg_Tab1_CarCheckout
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
blTab1CarListEvent = False
Set dg_Tab1_CarCheckout.DataSource = rs_Tab1_CarCheckOut
'�]�w������
With dg_Tab1_CarCheckout
    .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
    .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
    .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
    .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
    .Columns(0).Width = 500         '�s��
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000        '�X�����
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900         '���P���X
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 450         '����
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 850         '����h��
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1700        '���ܮɶ�
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800         '��X�̪O��
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 1700        '����ɶ�
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 800         '��J�̪O��
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700         '�r�p�H
    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 1000       '�q��
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 800        '�c��
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800        '�O��
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 800        '���n
    .Columns(13).Alignment = dbgRight
    .Columns(14).Width = 800        '���q
    .Columns(14).Alignment = dbgRight
    .Columns(15).Width = 1500       '�B�餽�q
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1500       '����
    .Columns(16).Alignment = dbgLeft
    .Columns(17).Width = 1200       '���u�s��
    .Columns(17).Alignment = dbgLeft
End With

Set ado_CarCheckout.Recordset = rs_Tab1_CarCheckOut
txt_Tab1_CarID.DataField = "���P���X"
txt_Tab1_Driver.DataField = "�r�p�H"
txt_Tab1_Phone.DataField = "�q��"
txt_Tab1_DeliveryDate.DataField = "�X�����"
txt_Tab1_DriveTimes.DataField = "����"
txt_Tab1_Checkin.DataField = "����ɶ�"
txt_Tab1_Palletin.DataField = "��J�̪O"
txt_Tab1_Company.DataField = "�B�餽�q"
txt_Tab1_VehicleType.DataField = "����"
txt_Tab1_CaseQty.DataField = "�c��"
txt_Tab1_PalletQty.DataField = "�O��"
txt_Tab1_Volumn.DataField = "���n"
txt_Tab1_Weight.DataField = "���q"
txt_Tab1_Checkout.DataField = "���ܮɶ�"
txt_Tab1_PalletOut.DataField = "��X�̪O"
txt_Tab1_RouteNo.DataField = "���u�s��"

'�ϥ���ܲĤ@�����
If Not rs_Tab1_CarCheckOut.EOF Then
   dg_Tab1_CarCheckout.SelBookmarks.Add dg_Tab1_CarCheckout.Bookmark
   txt_Tab1_Checkout.Text = rs_Tab1_CarCheckOut.Fields("���ܮɶ�").Value
   txt_Tab1_PalletOut.Text = rs_Tab1_CarCheckOut.Fields("��X�̪O").Value
End If
blTab1CarListEvent = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-���J�w���쨮��", Me.Caption, "cmd_Tab1_CarList_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Checkout_Click()
'�������� >> ��������
If rs_Tab1_CarCheckOut Is Nothing Then Exit Sub

On Error GoTo err_Handle
If dg_Tab1_CarCheckout.SelBookmarks.Count = 0 Then
   msg_text = "����������ܨ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'�H DB Server �ɶ�����
str_SQL = "Select Convert(varchar,Getdate(),111) as CheckoutDate , Convert(varchar,Getdate(),108) as CheckoutTime "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
txt_Tab1_Checkout.Text = tmp_Rs.Fields("CheckoutDate").Value & " " & tmp_Rs.Fields("CheckoutTime").Value
tmp_Rs.Close
'�� [�s��] �~�N�ɶ��B�̪O�� ��ܩ� [�w���쨮���C��]
'rs_Tab1_CarCheckOut.Fields("���ܮɶ�").Value = txt_Tab1_Checkout.Text
'rs_Tab1_CarCheckOut.Fields("��X�̪O").Value = txt_Tab1_PalletOut.Text
txt_Tab1_PalletOut.SelStart = 0: txt_Tab1_PalletOut.SelLength = Len(txt_Tab1_PalletOut.Text)
txt_Tab1_PalletOut.SetFocus
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-��������", Me.Caption, "cmd_Tab1_Checkout_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_ClearCheckin_Click()
'�������� >> �M���������ܮɶ�
On Error GoTo err_Handle
If rs_Tab1_CarCheckOut Is Nothing Then Exit Sub
If dg_Tab1_CarCheckout.SelBookmarks.Count = 0 Then
   msg_text = "����������ܨ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
txt_Tab1_Checkout.Text = " "
txt_Tab1_PalletOut.Text = 0

      '�N�����ƲM�� TRP05T
      If Not (tmp_Cmd Is Nothing) Then
         Set tmp_Cmd = Nothing
      End If
      Set tmp_Cmd = New ADODB.Command
      If tmp_para Is Nothing Then
         Set tmp_para = New ADODB.Parameter
      End If
      
      tmp_Cmd.ActiveConnection = cn
      tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
      tmp_Cmd.CommandType = adCmdStoredProc
      tmp_Cmd.CommandText = "VEHICLE_CHECK_OUT_Cancel"
      'ROUTE_NO
      Set tmp_para = tmp_Cmd.CreateParameter("ROUTE_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("ROUTE_NO").Value = txt_Tab1_RouteNo.Text
      'VEHICLE_ID_NO
      Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adVarChar, adParamInput, 15)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = txt_Tab1_CarID.Text
      'DRIVE_TIMES
      Set tmp_para = tmp_Cmd.CreateParameter("DRIVE_TIMES", adDouble, adParamInput)
      tmp_Cmd.Parameters.Append tmp_para
      tmp_Cmd.Parameters("DRIVE_TIMES").Value = Val(txt_Tab1_DriveTimes.Text)
      
      Call Confirm_Recordset_Closed(tmp_Rs)
      Call DB_CheckConnectStatus
      '�D�P�B����
      Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
      Do While tmp_Cmd.State = adStateExecuting
         DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
      Loop
      Set tmp_Cmd = Nothing
      rs_Tab1_CarCheckOut.Fields("���ܮɶ�").Value = " "
      rs_Tab1_CarCheckOut.Fields("��X�̪O").Value = 0
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��������-����", Me.Caption, "cmd_Tab1_ClearCheckin_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Default_Click()
'���J�w���쨮�� >> �z�����]�w >> �w�]
chk_Tab1_Checkin.Value = vbChecked
txt_Tab1_QueryDate.Text = ""
txt_Tab1_QueryCarID.Text = ""
End Sub

Private Sub cmd_Tab1_ShowQuery_Click()
'���J�w���쨮�� >> ��ܿz�����]�w
fam_Tab1_Query.Visible = Not fam_Tab1_Query.Visible

End Sub

Private Sub cmd_Tab2_ImportCard_Click()
    '�����i�X�@�~>>�פJ�ݾ�z�d��
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab2_CardIn.DataSource = Nothing
    Set dg_Tab2_CardOut.DataSource = Nothing
    Set dg_Tab2_Route.DataSource = Nothing
    
    '�ƨ��@�~�G�ݱƨ��q��
    Call CreateRS_Tab2_CardIn
    Call CreateRS_Tab2_CardOut
    Call CreateRS_Tab2_Route
    CardChange = False
    DoEvents
    
    '���^�ݱƨ��q��
    str_SQL = "select YMD,HM,Door,isnull(Port,'') as Port,Number,Username,Nickmane,CardNo,CardKey from gt_door where Status='0' and left(Port,3)='Car'"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    dg_Tab2_CardIn.Visible = False
    dg_Tab2_CardOut.Visible = False
    Do While Not tmp_Rs.EOF
        If Trim(tmp_Rs.Fields("Door").Value) = "1" Then
            rs_Tab2_CardIn.AddNew
            rs_Tab2_CardIn.Fields("���").Value = tmp_Rs.Fields("YMD").Value
            rs_Tab2_CardIn.Fields("�ɶ�").Value = tmp_Rs.Fields("HM").Value
            rs_Tab2_CardIn.Fields("���x").Value = tmp_Rs.Fields("Door").Value
            rs_Tab2_CardIn.Fields("����").Value = Trim(tmp_Rs.Fields("Port").Value)
            rs_Tab2_CardIn.Fields("�t�νs��").Value = tmp_Rs.Fields("CardKey").Value
            rs_Tab2_CardIn.Fields("�ϥΪ�").Value = tmp_Rs.Fields("Username").Value
            rs_Tab2_CardIn.Fields("�O�W").Value = tmp_Rs.Fields("Nickmane").Value
            rs_Tab2_CardIn.Fields("�d��").Value = tmp_Rs.Fields("CardNo").Value
            rs_Tab2_CardIn.Update
        Else
            rs_Tab2_CardOut.AddNew
            rs_Tab2_CardOut.Fields("���").Value = tmp_Rs.Fields("YMD").Value
            rs_Tab2_CardOut.Fields("�ɶ�").Value = tmp_Rs.Fields("HM").Value
            rs_Tab2_CardOut.Fields("���x").Value = tmp_Rs.Fields("Door").Value
            rs_Tab2_CardOut.Fields("����").Value = tmp_Rs.Fields("Port").Value
            rs_Tab2_CardOut.Fields("�t�νs��").Value = tmp_Rs.Fields("CardKey").Value
            rs_Tab2_CardOut.Fields("�ϥΪ�").Value = tmp_Rs.Fields("Username").Value
            rs_Tab2_CardOut.Fields("�O�W").Value = tmp_Rs.Fields("Nickmane").Value
            rs_Tab2_CardOut.Fields("�d��").Value = tmp_Rs.Fields("CardNo").Value
            rs_Tab2_CardOut.Update
        End If
       tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab2_CardIn.MoveFirst
    rs_Tab2_CardOut.MoveFirst
    dg_Tab2_CardIn.Visible = True
    dg_Tab2_CardOut.Visible = True
    
    '�פJ�ݽT�{���s
    str_SQL = "select �e�f��,���u�s��,����,�r�p�H,�w�p������,�w�p����ɶ�,����ɶ�,���t�ɶ� from CarControl_Card where ����ɶ�='' or ���t�ɶ�=''"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    dg_Tab2_Route.Visible = False
    Do While Not tmp_Rs.EOF
       rs_Tab2_Route.AddNew
       rs_Tab2_Route.Fields("�e�f��").Value = tmp_Rs.Fields("�e�f��").Value
       rs_Tab2_Route.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
       rs_Tab2_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
       rs_Tab2_Route.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
       rs_Tab2_Route.Fields("�w�p������").Value = tmp_Rs.Fields("�w�p������").Value
       rs_Tab2_Route.Fields("�w�p����ɶ�").Value = tmp_Rs.Fields("�w�p����ɶ�").Value
       rs_Tab2_Route.Fields("����ɶ�").Value = tmp_Rs.Fields("����ɶ�").Value
       rs_Tab2_Route.Fields("���t�ɶ�").Value = tmp_Rs.Fields("���t�ɶ�").Value
       rs_Tab2_Route.Update
       tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab2_Route.MoveFirst
    'Set dg_Tab2_Route.DataSource = rs_Tab2_Route
    txt_Tab2_Route_NO.Text = rs_Tab2_Route.Fields("���u�s��")
    dg_Tab2_Route.Visible = True
    
    CardChange = True
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_Handle:
       Dim tmpString As String
       msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
       tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
       CreateErrorLog Me.Name & "-�ƨ��C��-�פJ�ݱƨ��q��", Me.Caption, "cmd_Tab0_ImportOrders_Click", tmpString
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Selected_Click()
    If Len(Trim(txt_Tab2_InTime.Text)) = 0 And Len(Trim(txt_Tab2_OutTime.Text)) = 0 Then Exit Sub
    cn.BeginTrans
        str_SQL = "update trp05t set VEHICLE_CHECK_IN='" & Left(Trim(txt_Tab2_InTime.Text), 6) & " " & Mid(Trim(txt_Tab2_InTime.Text), 7, 2) & ":" & Right(Trim(txt_Tab2_InTime.Text), 2) & "'" & _
                ",VEHICLE_CHECK_OUT='" & Left(Trim(txt_Tab2_OutTime.Text), 6) & " " & Mid(Trim(txt_Tab2_OutTime.Text), 7, 2) & ":" & Right(Trim(txt_Tab2_OutTime.Text), 2) & "' where ROUTE_NO='" & Trim(txt_Tab2_Route_NO.Text) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        If Len(Trim(txt_Tab2_InTime.Text)) > 0 Then
            str_SQL = "update gt_door set Status='1' where CardKey='" & Trim(rs_Tab2_CardIn.Fields(4)) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            rs_Tab2_CardIn.Delete
            rs_Tab2_CardIn.Update
        End If
        If Len(Trim(txt_Tab2_OutTime.Text)) > 0 Then
            str_SQL = "update gt_door set Status='1' where CardKey='" & Trim(rs_Tab2_CardOut.Fields(4)) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            rs_Tab2_CardOut.Delete
            rs_Tab2_CardOut.Update
        End If
        rs_Tab2_Route.Delete
        rs_Tab2_Route.Update
        txt_Tab2_InTime.Text = ""
        txt_Tab2_OutTime.Text = ""
    cn.CommitTrans
End Sub

Private Sub cmd_Tab2_SelectedCancel_Click()
    txt_Tab2_InTime.Text = ""
    txt_Tab2_OutTime.Text = ""
End Sub

Private Sub cmd_Tab2_srcOrderReset_Click()
    If rs_Tab2_CardIn Is Nothing Then Exit Sub
    rs_Tab2_CardIn.Filter = adFilterNone
    rs_Tab2_CardIn.Sort = "��� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_Tab2_CardOut.Filter = adFilterNone
    rs_Tab2_CardOut.Sort = "��� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    txt_Tab2_InTime.Text = ""
    txt_Tab2_OutTime.Text = ""
End Sub

Private Sub dg_Tab0_CarCheckin_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�������� >> �ݳ��쨮���C�� >> ���
If blTab0CarListEvent Then
   With dg_Tab0_CarCheckin
        '�ϥ���ܿ������ƦC
        If Not rs_Tab0_CarCheckin.EOF Then
           dg_Tab0_CarCheckin.SelBookmarks.Add dg_Tab0_CarCheckin.Bookmark
           txt_Tab0_Checkin.Text = rs_Tab0_CarCheckin.Fields("����ɶ�").Value
           txt_Tab0_PalletIN.Text = rs_Tab0_CarCheckin.Fields("��J�̪O").Value
        End If
   End With
End If
End Sub

Private Sub dg_Tab1_CarCheckout_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�������� >> �w���쨮���C��
If blTab1CarListEvent Then
    With dg_Tab1_CarCheckout
        '�ϥ���ܿ������ƦC
        If Not rs_Tab1_CarCheckOut.EOF Then
            dg_Tab1_CarCheckout.SelBookmarks.Add dg_Tab1_CarCheckout.Bookmark
            txt_Tab1_Checkout.Text = rs_Tab1_CarCheckOut.Fields("���ܮɶ�").Value
            txt_Tab1_PalletOut.Text = rs_Tab1_CarCheckOut.Fields("��X�̪O").Value
        End If
    End With
End If
End Sub

Private Sub dg_Tab1_CarCheckout_SelChange(Cancel As Integer)
'�������� >> �w���쨮���C��
If blCancelChangeRecord Then
   Cancel = True
End If
End Sub

Private Sub dg_Tab2_CardIn_Click()
    If CardChange = False Then Exit Sub
    txt_Tab2_InTime.Text = rs_Tab2_CardIn.Fields("���").Value & rs_Tab2_CardIn.Fields("�ɶ�").Value
End Sub

Private Sub dg_Tab2_CardOut_Click()
    If CardChange = False Then Exit Sub
    Me.txt_Tab2_OutTime.Text = rs_Tab2_CardOut.Fields("���").Value & rs_Tab2_CardOut.Fields("�ɶ�").Value
End Sub


Private Sub dg_Tab2_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If CardChange = False Then Exit Sub
    txt_Tab2_Route_NO.Text = rs_Tab2_Route.Fields("���u�s��")
    str_SQL = "(�ϥΪ� LIKE '" & rs_Tab2_Route.Fields(3).Value & "' or �O�W LIKE '" & rs_Tab2_Route.Fields(2).Value & "')"
    rs_Tab2_CardIn.Filter = str_SQL
    If rs_Tab2_CardIn.RecordCount = 0 Then
         rs_Tab2_CardIn.Filter = adFilterNone
         rs_Tab2_CardIn.Sort = "��� ASC"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    End If
    rs_Tab2_CardOut.Filter = str_SQL
    If rs_Tab2_CardOut.RecordCount = 0 Then
        rs_Tab2_CardOut.Filter = adFilterNone
        rs_Tab2_CardOut.Sort = "��� ASC"
    End If
    txt_Tab2_InTime.Text = ""
    txt_Tab2_OutTime.Text = ""
End Sub


Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�����i�X�ި�@�~"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'��������
Call CreateRS_Tab0_Checkin
'��������
Call CreateRS_Tab0_Checkin

Call CreateRS_Tab2_CardIn
Call CreateRS_Tab2_CardOut
Call CreateRS_Tab2_Route
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�d�I��Ӫ����L����ƥ�
'�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������������
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
   fam_Tab0_Query.Visible = False
End If
End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   
   cmd_Tab0_CarList.Left = cmd_Tab0_CarList.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab0_CarData.Left = fam_Tab0_CarData.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   cmd_Exit(0).Left = cmd_Exit(0).Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab0_CarCheckin.Left = fam_Tab0_CarCheckin.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab0_CarCheckin.Height = dg_Tab0_CarCheckin.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Tab0_CarCheckin.Width = dg_Tab0_CarCheckin.Width - (dbsrcFormWidth - Me.ScaleWidth)
   cmd_Tab0_ShowQuery.Left = cmd_Tab0_ShowQuery.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab0_Query.Left = fam_Tab0_Query.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
      
   cmd_Tab1_CarList.Left = cmd_Tab1_CarList.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_CarData.Left = fam_Tab1_CarData.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   cmd_Exit(1).Left = cmd_Exit(1).Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_CarCheckout.Left = fam_Tab1_CarCheckout.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab1_CarCheckout.Height = dg_Tab1_CarCheckout.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Tab1_CarCheckout.Width = dg_Tab1_CarCheckout.Width - (dbsrcFormWidth - Me.ScaleWidth)
   cmd_Tab1_ShowQuery.Left = cmd_Tab1_ShowQuery.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_Query.Left = fam_Tab1_Query.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)

   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   
   cmd_Tab0_CarList.Left = cmd_Tab0_CarList.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab0_CarData.Left = fam_Tab0_CarData.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   cmd_Exit(0).Left = cmd_Exit(0).Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab0_CarCheckin.Left = fam_Tab0_CarCheckin.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab0_CarCheckin.Height = dg_Tab0_CarCheckin.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Tab0_CarCheckin.Width = dg_Tab0_CarCheckin.Width + (Me.ScaleWidth - dbsrcFormWidth)
   cmd_Tab0_ShowQuery.Left = cmd_Tab0_ShowQuery.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab0_Query.Left = fam_Tab0_Query.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   
   cmd_Tab1_CarList.Left = cmd_Tab1_CarList.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_CarData.Left = fam_Tab1_CarData.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   cmd_Exit(1).Left = cmd_Exit(1).Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_CarCheckout.Left = fam_Tab1_CarCheckout.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab1_CarCheckout.Height = dg_Tab1_CarCheckout.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Tab1_CarCheckout.Width = dg_Tab1_CarCheckout.Width + (Me.ScaleWidth - dbsrcFormWidth)
   cmd_Tab1_ShowQuery.Left = cmd_Tab1_ShowQuery.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_Query.Left = fam_Tab1_Query.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
      
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
End If
End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_OP_CarControl = Nothing
End Sub

Private Sub CreateRS_Tab0_Checkin()
'��������G�ݳ��쨮���C��
Call ReDim_Recordset(rs_Tab0_CarCheckin)
With rs_Tab0_CarCheckin
     .Fields.Append "�s��", adVarChar, 10
     .Fields.Append "�X�����", adVarChar, 10
     .Fields.Append "���P���X", adVarChar, 10
     .Fields.Append "����", adDouble
     .Fields.Append "����h��", adVarChar, 10
     .Fields.Append "����ɶ�", adVarChar, 20
     .Fields.Append "��J�̪O", adDouble
     .Fields.Append "�r�p�H", adVarChar, 20
     .Fields.Append "�q��", adVarChar, 20
     .Fields.Append "�c��", adDouble
     .Fields.Append "�O��", adDouble
     .Fields.Append "���n", adDouble
     .Fields.Append "���q", adDouble
     .Fields.Append "�B�餽�q", adVarChar, 60
     .Fields.Append "����", adVarChar, 60
     .Fields.Append "�w�p������", adVarChar, 10
     .Fields.Append "�w�p����ɶ�", adVarChar, 10
     .Fields.Append "�X�Y�Ȧs", adVarChar, 10
     .Fields.Append "���u�s��", adVarChar, 10
     .Fields.Append "���ܮɶ�", adVarChar, 20
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
End Sub

Private Sub CreateRS_Tab1_CheckOut()
'�������ܡG�w���쨮���C��
Call ReDim_Recordset(rs_Tab1_CarCheckOut)
With rs_Tab1_CarCheckOut
     .Fields.Append "�s��", adVarChar, 10
     .Fields.Append "�X�����", adVarChar, 10
     .Fields.Append "���P���X", adVarChar, 10
     .Fields.Append "����", adDouble
     .Fields.Append "����h��", adVarChar, 10
     .Fields.Append "���ܮɶ�", adVarChar, 20
     .Fields.Append "��X�̪O", adDouble
     .Fields.Append "����ɶ�", adVarChar, 20
     .Fields.Append "��J�̪O", adDouble
     .Fields.Append "�r�p�H", adVarChar, 20
     .Fields.Append "�q��", adVarChar, 20
     .Fields.Append "�c��", adDouble
     .Fields.Append "�O��", adDouble
     .Fields.Append "���n", adDouble
     .Fields.Append "���q", adDouble
     .Fields.Append "�B�餽�q", adVarChar, 60
     .Fields.Append "����", adVarChar, 60
     .Fields.Append "���u�s��", adVarChar, 10
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
   Case "��������.���J�ݳ��쨮��.�z�����.�X�����"
        txt_Tab0_QueryDate.Text = Format(mvDate.Value, "yyyymmdd")
   Case "��������.���J�w���쨮��.�z�����.�X�����"
        txt_Tab1_QueryDate.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_Tab0_PalletIN_KeyPress(KeyAscii As Integer)
'�������� >> ��J�̪O��
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   cmd_Tab0_CheckinSave.SetFocus
End If
End Sub

Private Sub txt_Tab0_QueryDate_Click()
'�������� >> ���J�ݳ��쨮�� >> �z����� >> �X�����
If Trim(txt_Tab0_QueryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_QueryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_QueryDate.Text, 4) & "/" & Mid(txt_Tab0_QueryDate.Text, 5, 2) & "/" & Right(txt_Tab0_QueryDate.Text, 2))
   End If
End If
mvDate.Tag = "��������.���J�ݳ��쨮��.�z�����.�X�����"
mvDate.Top = SSTab1.Top + fam_Tab0_Query.Top + txt_Tab0_QueryDate.Top + txt_Tab0_QueryDate.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Query.Left + txt_Tab0_QueryDate.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_QueryDate_Click()
'�������� >> ���J�w���쨮�� >> �z����� >> �X�����
If Trim(txt_Tab1_QueryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_QueryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_QueryDate.Text, 4) & "/" & Mid(txt_Tab1_QueryDate.Text, 5, 2) & "/" & Right(txt_Tab1_QueryDate.Text, 2))
   End If
End If
mvDate.Tag = "��������.���J�w���쨮��.�z�����.�X�����"
mvDate.Top = SSTab1.Top + fam_Tab1_Query.Top + txt_Tab1_QueryDate.Top + txt_Tab1_QueryDate.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Query.Left + txt_Tab1_QueryDate.Left
mvDate.Visible = True

End Sub

Private Sub CreateRS_Tab2_Route()
    Call ReDim_Recordset(rs_Tab2_Route)
    With rs_Tab2_Route
         .Fields.Append "�e�f��", adVarChar, 10
         .Fields.Append "���u�s��", adVarChar, 10
         .Fields.Append "����", adVarChar, 10
         .Fields.Append "�r�p�H", adVarChar, 20
         .Fields.Append "�w�p������", adVarChar, 8
         .Fields.Append "�w�p����ɶ�", adVarChar, 4
         .Fields.Append "����ɶ�", adVarChar, 20
         .Fields.Append "���t�ɶ�", adVarChar, 20
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '���ݳs������
    End With
    Set dg_Tab2_Route.DataSource = rs_Tab2_Route
    '�]�w������
    With dg_Tab2_Route
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 1000
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1000
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1200
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1200
        .Columns(7).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab2_CardIn()
    Call ReDim_Recordset(rs_Tab2_CardIn)
    With rs_Tab2_CardIn
         .Fields.Append "���", adVarChar, 6
         .Fields.Append "�ɶ�", adVarChar, 4
         .Fields.Append "���x", adVarChar, 10
         .Fields.Append "����", adVarChar, 20
         .Fields.Append "�t�νs��", adVarChar, 6
         .Fields.Append "�ϥΪ�", adVarChar, 20
         .Fields.Append "�O�W", adVarChar, 20
         .Fields.Append "�d��", adVarChar, 10
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '���ݳs������
    End With
    Set dg_Tab2_CardIn.DataSource = rs_Tab2_CardIn
    '�]�w������
    With dg_Tab2_CardIn
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 700
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 500
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 400
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1000
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1000
        .Columns(7).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab2_CardOut()
    Call ReDim_Recordset(rs_Tab2_CardOut)
    With rs_Tab2_CardOut
         .Fields.Append "���", adVarChar, 6
         .Fields.Append "�ɶ�", adVarChar, 4
         .Fields.Append "���x", adVarChar, 10
         .Fields.Append "����", adVarChar, 20
         .Fields.Append "�t�νs��", adVarChar, 6
         .Fields.Append "�ϥΪ�", adVarChar, 20
         .Fields.Append "�O�W", adVarChar, 20
         .Fields.Append "�d��", adVarChar, 10
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '���ݳs������
    End With
    Set dg_Tab2_CardOut.DataSource = rs_Tab2_CardOut
    '�]�w������
    With dg_Tab2_CardOut
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 700
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 500
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 400
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1000
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1000
        .Columns(7).Alignment = dbgLeft
    End With
End Sub
