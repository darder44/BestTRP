VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_RouteConfirm 
   Caption         =   "�X���T�{"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11400
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   240
      TabIndex        =   116
      Top             =   4680
      Visible         =   0   'False
      Width           =   3465
      _ExtentX        =   6112
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
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   97320961
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14215660
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "�X���T�{"
      TabPicture(0)   =   "frm_OP_RouteConfirm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dg_Tab0_C_RouteList"
      Tab(0).Control(1)=   "fam_Tab0_Consignee"
      Tab(0).Control(2)=   "cmd_Tab0_ConsigneeShow"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "cmd_Tab0_Confirm"
      Tab(0).Control(5)=   "cmd_Tab0_Delete"
      Tab(0).Control(6)=   "chkPalletDefend1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "�ݭ��ձƨ�"
      TabPicture(1)   =   "frm_OP_RouteConfirm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_Tab1_Reset"
      Tab(1).Control(1)=   "cmd_Tab1_Add"
      Tab(1).Control(2)=   "cmd_Tab1_Query"
      Tab(1).Control(3)=   "fam_SelectedOrders"
      Tab(1).Control(4)=   "fam_SrcOrders"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "�w�T�{���s "
      TabPicture(2)   =   "frm_OP_RouteConfirm.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "dg_Tab2_Route"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dg_Tab2_RouteOrders"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   " �q�����"
      TabPicture(3)   =   "frm_OP_RouteConfirm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fam_Tab3_Orders"
      Tab(3).Control(1)=   "fam_Tab3_OrderDetail"
      Tab(3).Control(2)=   "cmd_Tab3_DisplaySelectedOrder"
      Tab(3).Control(3)=   "cmd_Tab3_DisplayOrders"
      Tab(3).Control(4)=   "dg_Tab3_SDN02W"
      Tab(3).ControlCount=   5
      Begin VB.CheckBox chkPalletDefend1 
         BackColor       =   &H8000000A&
         Caption         =   "�O�_���@�̪O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   -71280
         TabIndex        =   134
         Top             =   360
         Value           =   1  '�֨�
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Frame Frame6 
         Height          =   645
         Left            =   120
         TabIndex        =   119
         Top             =   360
         Width           =   9075
         Begin VB.TextBox txt_Tab2_Route 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   600
            TabIndex        =   128
            Top             =   150
            Width           =   1365
         End
         Begin VB.CommandButton cmd_Tab2_SelectCar 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3630
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   124
            Top             =   150
            Width           =   330
         End
         Begin VB.TextBox txt_Tab2_VehicleNo 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   123
            Top             =   150
            Width           =   1125
         End
         Begin VB.TextBox txt_Tab2_DELIVERY_DATE 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6645
            TabIndex        =   122
            Top             =   150
            Width           =   1110
         End
         Begin VB.TextBox txt_Tab2_Driver 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4560
            TabIndex        =   121
            Top             =   150
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab2_CreateRoute 
            Appearance      =   0  '����
            BackColor       =   &H00FF8080&
            Caption         =   "�T�w�s��"
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
            Left            =   7920
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   120
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "���s"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   31
            Left            =   120
            TabIndex        =   129
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
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
            Height          =   270
            Index           =   30
            Left            =   2040
            TabIndex        =   127
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   29
            Left            =   5760
            TabIndex        =   126
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   14
            Left            =   4080
            TabIndex        =   125
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame fam_Tab3_Orders 
         BackColor       =   &H8000000A&
         Caption         =   "�s�W�q����"
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   100
         Top             =   4920
         Width           =   10905
         Begin VB.CommandButton cmd_Tab3_Query 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   115
            Top             =   158
            Width           =   330
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   5325
            TabIndex        =   114
            Top             =   150
            Width           =   1230
         End
         Begin VB.CommandButton cmd_Tab3_DelOrders 
            BackColor       =   &H00FF8080&
            Caption         =   "�R���q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9600
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   113
            Top             =   465
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab3_CaseQty 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   9780
            TabIndex        =   112
            Top             =   120
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_Volumn 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   8835
            TabIndex        =   111
            Top             =   120
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_Weight 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   7890
            TabIndex        =   110
            Top             =   120
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_FullName 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   945
            TabIndex        =   103
            Top             =   465
            Width           =   5610
         End
         Begin VB.TextBox txt_Tab3_Extern 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   3390
            TabIndex        =   102
            Top             =   165
            Width           =   1050
         End
         Begin VB.TextBox txt_Tab3_OrderKey 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   945
            TabIndex        =   101
            Top             =   165
            Width           =   1170
         End
         Begin MSDataGridLib.DataGrid dg_Tab3_SDN03W 
            Height          =   1155
            Left            =   120
            TabIndex        =   109
            Top             =   840
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   2037
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483624
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���q/���n/�c��"
            Height          =   180
            Index           =   38
            Left            =   6645
            TabIndex        =   108
            Top             =   195
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�f���"
            Height          =   180
            Index           =   37
            Left            =   4560
            TabIndex        =   107
            Top             =   210
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�W��"
            Height          =   180
            Index           =   28
            Left            =   165
            TabIndex        =   106
            Top             =   525
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D�渹"
            Height          =   180
            Index           =   27
            Left            =   2610
            TabIndex        =   105
            Top             =   225
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��s��"
            Height          =   180
            Index           =   24
            Left            =   165
            TabIndex        =   104
            Top             =   210
            Width           =   720
         End
      End
      Begin VB.Frame fam_Tab3_OrderDetail 
         BackColor       =   &H80000000&
         Caption         =   "�q�����"
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   -74760
         TabIndex        =   86
         Top             =   2520
         Width           =   10875
         Begin VB.CommandButton cmd_Tab3_ClearQty 
            BackColor       =   &H00FF80FF&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9105
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   93
            Top             =   150
            Width           =   420
         End
         Begin VB.CommandButton cmd_Tab3_CutOrders 
            BackColor       =   &H00FF8080&
            Caption         =   "�q�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9600
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   92
            Top             =   165
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab3_CutQty 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�ƶq����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7920
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   91
            Top             =   150
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab3_CutCaseQty 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   7200
            TabIndex        =   90
            Top             =   210
            Width           =   700
         End
         Begin VB.TextBox txt_Tab3_SelectedCaseQty 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1995
            TabIndex        =   89
            Top             =   225
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_SelectedWeight 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   3555
            TabIndex        =   88
            Top             =   225
            Width           =   945
         End
         Begin VB.TextBox txt_Tab3_SelectedVolumn 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   5100
            TabIndex        =   87
            Top             =   225
            Width           =   945
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab3_SelectedOrderDetail 
            Height          =   1845
            Left            =   45
            TabIndex        =   94
            Top             =   525
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   3254
            _Version        =   393216
            Cols            =   9
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��������p�p"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   195
            TabIndex        =   99
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�c�Ƥ���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   19
            Left            =   6315
            TabIndex        =   98
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���q"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Left            =   3045
            TabIndex        =   97
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   4575
            TabIndex        =   96
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�c��"
            BeginProperty Font 
               Name            =   "�s�ө���"
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
            Left            =   1500
            TabIndex        =   95
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.CommandButton cmd_Tab3_DisplaySelectedOrder 
         BackColor       =   &H00C0E0FF&
         Caption         =   "�q����Ω���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72120
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   84
         Top             =   360
         Width           =   2250
      End
      Begin VB.CommandButton cmd_Tab3_DisplayOrders 
         BackColor       =   &H00FF8080&
         Caption         =   "�פJ�ݱƨ��q��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74760
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   83
         Top             =   360
         Width           =   2250
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '����
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   3345
         Left            =   9300
         TabIndex        =   77
         Top             =   360
         Width           =   1995
         Begin VB.CommandButton cmd_Tab2_Excel 
            BackColor       =   &H00FFFF80&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   133
            ToolTipText     =   "�R��"
            Top             =   2760
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab2_DeliveryDate_Start 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   173
            TabIndex        =   131
            Top             =   1290
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab2_RouteNoDelete 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���u�s���R��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   130
            ToolTipText     =   "�R��"
            Top             =   2280
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab2_Route_Start 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   173
            MaxLength       =   10
            TabIndex        =   79
            Top             =   525
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab2_RouteNoQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���u�s���d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   78
            Top             =   1800
            Width           =   1785
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Left            =   465
            TabIndex        =   132
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���u�s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Left            =   465
            TabIndex        =   80
            Top             =   195
            Width           =   1020
         End
      End
      Begin VB.Frame fam_SrcOrders 
         Height          =   2955
         Left            =   -74865
         TabIndex        =   45
         Top             =   4020
         Width           =   11220
         Begin VB.Frame fam_SelectedSum 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   5595
            Begin VB.TextBox txt_Tab1_srcSelected_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   4695
               TabIndex        =   56
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcSelected_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   2865
               TabIndex        =   55
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcSelected_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   990
               TabIndex        =   54
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "����G�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   25
               Left            =   75
               TabIndex        =   59
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   23
               Left            =   2475
               TabIndex        =   58
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   22
               Left            =   4320
               TabIndex        =   57
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Enabled         =   0   'False
            Height          =   525
            Left            =   5610
            TabIndex        =   46
            Top             =   0
            Width           =   5595
            Begin VB.TextBox txt_Tab1_srcTotal_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   975
               TabIndex        =   49
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcTotal_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   2865
               TabIndex        =   48
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_srcTotal_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   4680
               TabIndex        =   47
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   21
               Left            =   4305
               TabIndex        =   52
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   20
               Left            =   2475
               TabIndex        =   51
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�`�p�G�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   18
               Left            =   75
               TabIndex        =   50
               Top             =   210
               Width           =   900
            End
         End
         Begin MSDataGridLib.DataGrid dg_SDN02W 
            Height          =   2190
            Left            =   45
            TabIndex        =   60
            Top             =   600
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   3863
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
         Height          =   3495
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   11220
         Begin VB.CommandButton cmd_Tab1_CreateRoute 
            Appearance      =   0  '����
            BackColor       =   &H00FF8080&
            Caption         =   "�إ߸��u�s��"
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
            Left            =   8520
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   76
            Top             =   120
            Width           =   1470
         End
         Begin VB.TextBox txt_Tab1_Driver0 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4920
            TabIndex        =   74
            Top             =   150
            Width           =   1080
         End
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   67
            Top             =   2835
            Width           =   5595
            Begin VB.TextBox txt_Tab1_Selected_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   4695
               TabIndex        =   70
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_Selected_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   2865
               TabIndex        =   69
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab1_Selected_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   990
               TabIndex        =   68
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�֭p�G�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   9
               Left            =   75
               TabIndex        =   73
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   7
               Left            =   2475
               TabIndex        =   72
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   6
               Left            =   4320
               TabIndex        =   71
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.CommandButton cmd_Tab1_Selected 
            BackColor       =   &H00FF8080&
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
            Height          =   375
            Left            =   5655
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   66
            Top             =   2955
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab1_SelectedCancel_All 
            BackColor       =   &H00FF80FF&
            Caption         =   "�ݿ����(��)"
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
            Left            =   6630
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   65
            Top             =   2955
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton cmd_Tab1_Remove 
            BackColor       =   &H008080FF&
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
            Height          =   375
            Left            =   6015
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   64
            Top             =   2955
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab1_srcOrderQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�q��j�M"
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
            Left            =   9525
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   63
            Top             =   2955
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab1_srcOrderReset 
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
            Left            =   10650
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   62
            Top             =   2955
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab1_SelectedCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "�ݿ����"
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
            Left            =   8085
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   61
            Top             =   2955
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txt_Tab1_DELIVERY_DATE 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7005
            TabIndex        =   41
            Top             =   150
            Width           =   1110
         End
         Begin VB.TextBox txt_Tab1_VehicleNo 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2880
            TabIndex        =   40
            Top             =   150
            Width           =   1125
         End
         Begin VB.CommandButton cmd_Tab1_SelectCar 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3990
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   39
            Top             =   150
            Width           =   330
         End
         Begin VB.CommandButton cmd_Tab1_ImportOrders 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���J�ݭ��խq��"
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
            Left            =   30
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   38
            Top             =   105
            Width           =   1815
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
            Height          =   495
            Index           =   1
            Left            =   9960
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   37
            Top             =   120
            Width           =   1110
         End
         Begin MSDataGridLib.DataGrid dg_Tab1_SelectedOrders 
            Height          =   2235
            Left            =   0
            TabIndex        =   42
            Top             =   600
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   3942
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
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   10
            Left            =   4440
            TabIndex        =   75
            Top             =   240
            Width           =   900
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '���z��
            Height          =   435
            Left            =   5610
            Top             =   2925
            Width           =   795
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   435
            Index           =   0
            Left            =   6600
            Top             =   2925
            Visible         =   0   'False
            Width           =   2790
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   12
            Left            =   6120
            TabIndex        =   44
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label1 
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
            Height          =   270
            Index           =   11
            Left            =   2040
            TabIndex        =   43
            Top             =   240
            Width           =   1020
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00FF0000&
            FillStyle       =   0  '���
            Height          =   435
            Left            =   9495
            Top             =   2925
            Visible         =   0   'False
            Width           =   1680
         End
      End
      Begin VB.CommandButton cmd_Tab1_Query 
         BackColor       =   &H0080FF80&
         Caption         =   "�d  ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73800
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Tab1_Add 
         BackColor       =   &H00FFFF80&
         Caption         =   "�s  �W"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72600
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   34
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Tab1_Reset 
         BackColor       =   &H000080FF&
         Caption         =   "����"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71400
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   32
         Top             =   840
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Tab0_Delete 
         BackColor       =   &H0080FFFF&
         Caption         =   "�D�g��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72120
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   26
         Top             =   390
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_Tab0_Confirm 
         BackColor       =   &H000080FF&
         Caption         =   "�T �{"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   25
         Top             =   390
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '����
         BackColor       =   &H8000000C&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   -69600
         TabIndex        =   11
         Top             =   3960
         Width           =   8640
         Begin VB.CommandButton cmd_Tab0_Clear1 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   118
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txt_Driver1 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3585
            TabIndex        =   28
            Top             =   600
            Width           =   1080
         End
         Begin VB.TextBox txt_DELIVERY_DATE1 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   5505
            TabIndex        =   27
            Top             =   600
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab0_RouteSelect1 
            BackColor       =   &H00FF8080&
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
            Height          =   360
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txt_Tab0_C_Route_No1 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1665
            TabIndex        =   19
            Top             =   195
            Width           =   1980
         End
         Begin VB.CommandButton cmd_Tab0_Selected 
            BackColor       =   &H00FF8080&
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
            Height          =   360
            Left            =   3720
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   18
            Top             =   195
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab0_Remove 
            BackColor       =   &H008080FF&
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
            Height          =   360
            Left            =   4200
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   17
            Top             =   195
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab0_SelectCar1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   13
            Top             =   615
            Width           =   330
         End
         Begin VB.TextBox txt_VehicleNo1 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   615
            Width           =   1080
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_RouteList1 
            Height          =   1920
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3387
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
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   3120
            TabIndex        =   30
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�X����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   4800
            TabIndex        =   29
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label1 
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
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   16
            Top             =   735
            Width           =   900
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   15
            Top             =   315
            Width           =   840
         End
      End
      Begin VB.CommandButton cmd_Tab0_ConsigneeShow 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��ܥ��T�{���s"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74805
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   360
         Width           =   1830
      End
      Begin VB.Frame fam_Tab0_Consignee 
         Appearance      =   0  '����
         BackColor       =   &H8000000C&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   -69600
         TabIndex        =   1
         Top             =   840
         Width           =   8640
         Begin VB.CheckBox chkPalletDefend2 
            BackColor       =   &H8000000A&
            Caption         =   "�O�_���@�̪O"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   6600
            MaskColor       =   &H00808080&
            TabIndex        =   136
            Top             =   600
            Value           =   1  '�֨�
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CommandButton cmd_Tab0_RouteSelect0 
            BackColor       =   &H00FF8080&
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
            Height          =   360
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   135
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab0_Clear0 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   117
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "��  �}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   6120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   33
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Del 
            BackColor       =   &H0080FFFF&
            Caption         =   "��������"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4920
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   31
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox txt_DELIVERY_DATE0 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   5505
            TabIndex        =   22
            Top             =   600
            Width           =   1080
         End
         Begin VB.TextBox txt_Driver0 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3585
            TabIndex        =   20
            Top             =   600
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab0_SelectCar0 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   9
            Top             =   615
            Width           =   330
         End
         Begin VB.TextBox txt_VehicleNo0 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   1080
         End
         Begin VB.CommandButton cmd_Tab0_OK 
            BackColor       =   &H000080FF&
            Caption         =   "�T  �{"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   6
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab0_C_Route_No0 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1665
            TabIndex        =   2
            Top             =   195
            Width           =   1980
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_RouteList0 
            Height          =   1800
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3175
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
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�X����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   4800
            TabIndex        =   23
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   3120
            TabIndex        =   21
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label1 
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
            Height          =   270
            Index           =   13
            Left            =   720
            TabIndex        =   10
            Top             =   735
            Width           =   900
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   16
            Left            =   720
            TabIndex        =   3
            Top             =   315
            Width           =   840
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_C_RouteList 
         Height          =   6120
         Left            =   -74820
         TabIndex        =   5
         Top             =   840
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   10795
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
      Begin MSDataGridLib.DataGrid dg_Tab2_RouteOrders 
         Height          =   3240
         Left            =   120
         TabIndex        =   81
         Top             =   3735
         Width           =   11180
         _ExtentX        =   19711
         _ExtentY        =   5715
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
      Begin MSDataGridLib.DataGrid dg_Tab2_Route 
         Height          =   2625
         Left            =   120
         TabIndex        =   82
         Top             =   1080
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   4630
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
      Begin MSDataGridLib.DataGrid dg_Tab3_SDN02W 
         Height          =   1755
         Left            =   -74760
         TabIndex        =   85
         Top             =   720
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   3096
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
Attribute VB_Name = "frm_OP_RouteConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_Tab0_C_RouteList As ADODB.Recordset    '���X���T�{�����s
Private rs_Tab0_RouteList0 As ADODB.Recordset     'Tab0
Private rs_Tab0_RouteList1 As ADODB.Recordset
Private rs_Tab1_SelectedOrders As ADODB.Recordset
Private rs_Tab2_Route As ADODB.Recordset
Private rs_Tab2_RouteOrders As ADODB.Recordset
Private rs_SDN02W As ADODB.Recordset
Private rs_Tab3_SDN02W As ADODB.Recordset
Private rs_Tab3_SDN03W As ADODB.Recordset
Private str_route As String                      '�s�W���s
Private dbsrcFormHeight As Double                'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double                 'Form �]�p�ɴ����e
Private Tab0_RouteListEventEnable As Boolean     'Tab0�O�_�Ұʿ���ƥ�
Private Tab1_RouteListEventEnable As Boolean     'Tab1�O�_�Ұʿ���ƥ�
Private Tab2_RouteListEventEnable As Boolean     'Tab1�O�_�Ұʿ���ƥ�
Private CutOrderkey As String                    '�s���ΥX�Ӥ��q��s��
Private dbCut_TotalCaseQty As Double
Private dbCut_TotalWeight As Double
Private dbCut_TotalVolumn As Double
Private rs_Tab2_RouteEvent As Boolean
Private intColumnIndex As Integer

Private Sub cmd_Exit_Click(Index As Integer)
    '���}
    Unload Me
End Sub

Private Sub cmd_Tab0_Clear0_Click()
    Call clear_Tab0_RouteList0
End Sub

Private Sub cmd_Tab0_Clear1_Click()
    Call clear_Tab0_RouteList1
End Sub

Private Sub cmd_Tab0_Confirm_Click()
    Dim Str_Receiver As String
    '�X���T�{ -�T�{���s
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    
    'Terry 20180515 �s�W�O�_���@�̪O���s
    Dim str_PalletDefend As String
    str_PalletDefend = ""
    If chkPalletDefend1.Value = vbChecked Then
        str_PalletDefend = "Y"
    Else
        str_PalletDefend = "N"
    End If
    
    
    
    Tab0_RouteListEventEnable = False
    Call WriteOut_RunLog("1.�}�l >> �s�JSDN01T;SDN02T;SDN03T")
    
    rs_Tab0_C_RouteList.Filter = "��='V'"
    If Not rs_Tab0_C_RouteList.EOF Then

        Do While Not rs_Tab0_C_RouteList.EOF
        Tran_Level = cn.BeginTrans
        '�����s�O�_�w�T�{
        Str_Receiver = ""
        Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)
        str_SQL = "Select t05t.Route_No as ���u�s�� From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & rs_Tab0_C_RouteList("���u�s��") & "' Union All Select t05t.Route_No as ���u�s�� From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & rs_Tab0_C_RouteList("���u�s��") & "' "
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        If Not tmp_Rs.EOF Then MsgBox "���u�s�� " & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & " �w�@�L�X���T�{!", 16, "�`�N": tmp_Rs.Close: GoTo nextROUTE
        tmp_Rs.Close
        
        '����дڤH
        Call ReDim_Recordset(tmp_Rs)
        str_SQL = "select �дڤH=isnull(receiver,driver) from trp09m(nolock) where vehicle_id_no = '" & rs_Tab0_C_RouteList.Fields("���P���X").Value & "'"
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        If tmp_Rs.EOF Then Str_Receiver = "" Else Str_Receiver = RTrim(tmp_Rs.Fields("�дڤH"))
        tmp_Rs.Close
        
            '�p�X�Q�عL�Ӫ��ɭԡA�N�O�h�f�q��A�{�b�w�g��A2B
            If Trim(rs_Tab0_C_RouteList.Fields("���O").Value) = "�h�f�q��" Then
'                DoEvents: DoEvents
                Call WriteOut_RunLog("�T�{���s: " & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "")
                
                '�s���Y,SDN01T
                str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,receiver,PalletDefend) " & _
                        "Values ( '" & Trim(rs_Tab0_C_RouteList.Fields("�X�����").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("���P���X").Value) & "', " & _
                        "'" & Trim(rs_Tab0_C_RouteList.Fields("�r�p�H").Value) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("����").Value) & "','" & Str_Receiver & "','" & str_PalletDefend & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                                              
                '�s�q��,SDN02T
                str_SQL = "INSERT dbo.SDN02T(C_ROUTE_NO,ROUTE_NO,STORERKEY,EXTERN,RECEIPT_DATE,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,OnTimeDelivery,PODOnTime,RejectOrder,DESCRIPTION,CONFIRM_DATE,CONSIGNEEKEY,CONFIRM_USERID,CUSTSIGNDATE,RBCCode,RSCCode,CONFIRM_Notes,PRIORITY,SCHEDULEDATE,CustomerOrderkey1,Scan,SDNSendDate,CUST_Handle,TRP_Handle,Advance,INV_Handle,TRP_Cost,Sorting_Cost,Total_Cost,VEHICLE_ID_NO,ExpectReceiptOK,SdnFeedBack,InvBack,C_RECEIPT_NO) " & _
                        "SELECT ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO ,t2.STORERKEY, t2.EXTERN , CONVERT(varchar(8),t2.RECEIPT_DATE, 112) AS RECEIPT_DATE,CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                        "SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.ORDER_QTY / sp.Casecnt end, 3), 0)) AS SHIP_CS, " & _
                        "SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDCUBE, 3), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDGrossWGT, 3), 0)) AS SHIP_WT, " & _
                        "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO,0,0,0,isnull(t2.description,''),null,t2.consigneekey,'',null,'','','',t2.priority,null,'','N',null,'','','','',0,0,0,t2.vehicle_id_no ,'N',0,'N',t2.receipt_no " & _
                        "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                        "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                        "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                        "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                        "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' " & _
                        "GROUP BY t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO,t2.CONSIGNEEKEY,t2.PRIORITY,t2.VEHICLE_ID_NO,t2.STORERKEY,t2.description,CONVERT(varchar(8),t2.RECEIPT_DATE, 112),t2.receipt_no"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '�s����,SDN03T
                str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                        "select  '" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ORDER_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                        "from ORT03t where route_no in( " & _
                        "select  route_no from ORT01t where  isnull(c_route_no,route_no)='" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' and left(route_no,1) <>'S')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                'ORT05T
                str_SQL = "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Else
'                DoEvents: DoEvents
                Call WriteOut_RunLog("�T�{���s: " & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "")
                
                '�s���Y,SDN01T
                str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,receiver,PalletDefend) " & _
                        "Values ( '" & Trim(rs_Tab0_C_RouteList.Fields("�X�����").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "','" & Trim(rs_Tab0_C_RouteList.Fields("���P���X").Value) & "', " & _
                        "'" & Trim(rs_Tab0_C_RouteList.Fields("�r�p�H").Value) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("����").Value) & "','" & Str_Receiver & "', '" & str_PalletDefend & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

                '�s�q��,SDN02T
                str_SQL = "INSERT dbo.SDN02T(C_ROUTE_NO,ROUTE_NO,STORERKEY,EXTERN,RECEIPT_DATE,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,OnTimeDelivery,PODOnTime,RejectOrder,DESCRIPTION,CONFIRM_DATE,CONSIGNEEKEY,CONFIRM_USERID,CUSTSIGNDATE,RBCCode,RSCCode,CONFIRM_Notes,PRIORITY,SCHEDULEDATE,CustomerOrderkey1,Scan,SDNSendDate,CUST_Handle,TRP_Handle,Advance,INV_Handle,TRP_Cost,Sorting_Cost,Total_Cost,VEHICLE_ID_NO,ExpectReceiptOK,SdnFeedBack,InvBack,C_RECEIPT_NO) " & _
                        "SELECT  ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO ,t2.STORERKEY , t2.EXTERN ,CONVERT(varchar(8),t2.RECEIPT_DATE, 112) AS RECEIPT_DATE, CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                        "SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.SHIP_QTY / sp.Casecnt end, 2), 0)) AS SHIP_CS, " & _
                        "SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.stdcube, 2), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDGrossWGT, 2), 0)) AS SHIP_WT, " & _
                        "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO,0,0,0,t2.description " & _
                        ",null,t2.consigneekey,'',null,'','','',t2.priority,t2.scheduledate,t2.customerorderkey1,'N',null,'','','','',0,0,0,t2.vehicle_id_no,'N',0,'N',t2.c_receipt_no " & _
                        "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                        "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                        "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                        "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                        "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' " & _
                        "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO,t2.CONSIGNEEKEY,t2.PRIORITY,t2.SCHEDULEDATE,t2.VEHICLE_ID_NO,t2.CustomerOrderkey1,t2.STORERKEY,t2.description,CONVERT(varchar(8),t2.RECEIPT_DATE, 112),t2.c_receipt_no "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '�s����,SDN03T
                str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                        "select  '" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ship_qty,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                        "from trp03t where  route_no in( " & _
                        "select  route_no from trp01t where  isnull(c_route_no,route_no)='" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' and left(route_no,1) <>'S')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '��sTRP05T���A
                str_SQL = "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '��sSDN01T���A
                str_SQL = "Update SDN01T set SDN01T.APPStatus = TRP05T.APPStatus ,SDN01T.VLListCount = TRP05T.VLListCount ,SDN01T.VLListPrintDate = TRP05T.VLListPrintDate from trp05t join sdn01t on isnull(trp05t.c_route_no,trp05t.route_no) = sdn01t.c_route_no and sdn01t.c_Route_No= '" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '��sAPPOrderDate���A
                str_SQL = "update AppOrderDate set status = '5' where status < '6' and c_Route_No= '" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "' "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                               
            End If
            
                '��sOrderType Add by Gemini @20190604
                str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��")) & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��")) & "' "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '��s�дڤH edit by Eric ����X�дڤH�A�binsert�N�ɶi�h�A�קK��s�⦸
                'cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & rs_Tab0_C_RouteList.Fields("���P���X").Value & "') where c_route_no = '" & rs_Tab0_C_RouteList.Fields("���u�s��").Value & "'", RowsAffect, adExecuteNoRecords
            
nextROUTE:
        cn.CommitTrans: Tran_Level = 0
            rs_Tab0_C_RouteList.MoveNext
            
        Loop

        
        '[�R���w������q��
        Call WriteOut_RunLog("3.�R���w��������u�s��")
'        DoEvents: DoEvents
        rs_Tab0_C_RouteList.MoveFirst
        Do While Not rs_Tab0_C_RouteList.EOF
            rs_Tab0_C_RouteList.Delete
            rs_Tab0_C_RouteList.MoveFirst
        Loop
    End If
  
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    Call WriteOut_RunLog("4.�T�{����")
'    DoEvents: DoEvents
    Call Unload_RunLogForm
    Call ReSet_Tab0_C_RouteList_SeqNo
    Call clear_Tab0_RouteList0
    Call clear_Tab0_RouteList1
    Tab0_RouteListEventEnable = True

    Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
    End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "�X���T�{-�T�{���s", Me.Caption, "cmd_Tab0_Confirm_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_ConsigneeShow_Click()
    '�X���T�{-�פJ���T�{���s
    Set dg_Tab0_C_RouteList.DataSource = Nothing
    Set rs_Tab0_C_RouteList = Nothing
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    str_SQL = "Select ' ' as '��',t01t.Route_No as ���u�s�� , Convert(varchar(8),t01t.Delivery_Date,112) as �X����� ,  " & _
            "Rtrim(t05t.Vehicle_ID_No) as ���P���X , t05t.Drive_Times as ���� , Rtrim(t05t.Driver) as �r�p�H , Rtrim(Isnull(t08m.SHORT_NAME,'')) as �B�餽�q,'���`�q��' as ���O " & _
            "From TRP01T t01t " & _
            "inner join TRP05T t05t on t05t.Route_No = t01t.Route_No " & _
            "left join TRP09M t09m on t09m.Vehicle_ID_No = t05t.Vehicle_ID_No " & _
            "Left outer join TRP08M t08m on t08m.Company_Code = t09m.TRP_Company_Code " & _
            "Where t01t.Route_No <> 'D' and  t05t.SDNStatus='0'  and t01t.C_ROUTE_NO is null " & _
            "Union All " & _
            "Select ' ' as '��',t01t.Route_No as ���u�s�� , Convert(varchar(8),t01t.Delivery_Date,112) as �X����� , " & _
            "Rtrim(t05t.Vehicle_ID_No) as ���P���X , t05t.Drive_Times as ���� , Rtrim(t05t.Driver) as �r�p�H , Rtrim(Isnull(t08m.SHORT_NAME,'')) as �B�餽�q,'�h�f�q��' as ���O " & _
            "From ORT01T t01t " & _
            "inner join ORT05T t05t on t05t.Route_No = t01t.Route_No " & _
            "left join TRP09M t09m on t09m.Vehicle_ID_No = t05t.Vehicle_ID_No " & _
            "Left outer join TRP08M t08m on t08m.Company_Code = t09m.TRP_Company_Code " & _
            "Where t01t.Route_No <> 'D' and  t05t.SDNStatus='0'  and t01t.C_ROUTE_NO is null order by t01t.Route_No "

    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab0_C_RouteList)
    tmp_Rs.Close
    
    With dg_Tab0_C_RouteList
         .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
         .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
         .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
         .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
    End With
    rs_Tab0_C_RouteList.MoveFirst
    Set dg_Tab0_C_RouteList.DataSource = rs_Tab0_C_RouteList
    With dg_Tab0_C_RouteList
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .Columns(0).Width = 400       '�Ǹ�
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 300       '�Ǹ�
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000      '���u�s��
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 900      '�X�����
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 900      '���P���X
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 400       '����
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 1000      '�r�p�H
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1200       '�B��²��
        .Columns(7).Alignment = dbgLeft
        .Columns(8).Width = 1200       '����ɶ�
        .Columns(8).Alignment = dbgLeft
    End With
    rs_Tab0_C_RouteList.MoveFirst
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = " �s�� "
    rs_Tab0_C_RouteList.MoveFirst
    blVLLReportEventEnable = True
    Tab0_RouteListEventEnable = True
    Screen.MousePointer = vbDefault
    Call cmd_Tab0_Clear0_Click
    Call cmd_Tab0_Clear1_Click
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-�פJ���T�{���s", Me.Caption, "cmd_Tab0_ConsigneeShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Del_Click()
    '�X���T�{-��������
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If Len(Trim(txt_Tab0_C_Route_No0.Text)) = 0 Then Exit Sub
    
    '�����s�O�_�w�T�{
    Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)

    str_SQL = "Select t05t.Route_No as ���u�s�� From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0 & "' Union All Select t05t.Route_No as ���u�s�� From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0 & "' "
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then MsgBox "���u�s�� " & txt_Tab0_C_Route_No0 & " �w�@�L�X���T�{!", 16, "�`�N": tmp_Rs.Close: cmd_Tab0_Del.Enabled = False: Exit Sub
     
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    Call WriteOut_RunLog("1.�}�l >> �s�JSDN02W")
    
    rs_Tab0_C_RouteList.Filter = "���u�s��='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
    If Not rs_Tab0_C_RouteList.EOF Then
        Tran_Level = cn.BeginTrans
        If Trim(rs_Tab0_C_RouteList.Fields("���O").Value) = "�h�f�q��" Then
            '�s�q��,SDN02W
            str_SQL = "INSERT dbo.SDN02W (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,C_RECEIPT_NO)" & _
                    "SELECT ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO , t2.EXTERN , CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                    "SUM(ISNULL(ROUND(case when s1.Casecnt = 0 then 0 else t3.SHIP_QTY / s1.Casecnt end , 2), 0)) AS SHIP_CS, " & _
                    "SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDCUBE, 2), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDGrossWGT, 2), 0)) AS SHIP_WT, " & _
                    "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO ,t2.receipt_no " & _
                    "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                    "INNER JOIN gv_SKUxpack s1 ON s1.Sku = t3.PRODUCT_NO and s1.storerkey = t3.storerkey " & _
                    "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                    "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                    "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and t2.RECEIPT_NO not in (select receipt_no from sdn02w ) " & _
                    "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�s����,SDN03W
            str_SQL = "Insert into SDN03W (ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME) " & _
                    "select  ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME " & _
                    "from ORT03t where route_no in( " & _
                    "select  route_no from ORT01t where  isnull(c_route_no,route_no)='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and left(route_no,1) <>'S')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��sORT05T���A
            str_SQL = "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            rs_Tab0_C_RouteList.MoveNext
        Else
            '�s�q��,SDN02W
            str_SQL = "INSERT dbo.SDN02W (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,C_RECEIPT_NO) " & _
                    "SELECT  ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS C_ROUTE_NO, t2.ROUTE_NO , t2.EXTERN , CONVERT(varchar(8),t2.ARRIVE_DATE, 112) AS ARRIVE_DATE, m1.FULL_NAME as CUST_NAME, " & _
                    "SUM(ISNULL(ROUND(case when s1.Casecnt = 0 then 0 else t3.SHIP_QTY / s1.Casecnt end , 2), 0)) AS SHIP_CS, " & _
                    "SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDCUBE, 2), 0)) AS SHIP_CBM,SUM(ISNULL(ROUND(t3.SHIP_QTY * s1.STDGrossWGT, 2), 0)) AS SHIP_WT, " & _
                    "ISNULL(t2.CAR_NOTES, '') AS CAR_NOTES,'0','','1','1',t2.RECEIPT_NO as RECEIPT_NO ,t2.c_receipt_no " & _
                    "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO " & _
                    "INNER JOIN gv_SKUxpack s1 ON s1.Sku = t3.PRODUCT_NO and s1.storerkey = t3.storerkey " & _
                    "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and t2.storerkey = m1.storerkey " & _
                    "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO " & _
                    "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and t2.RECEIPT_NO not in (select receipt_no from sdn02w ) " & _
                    "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO,t2.c_receipt_no"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�s����,SDN03W
            str_SQL = "Insert into SDN03W (ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME) " & _
                    "select  ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME " & _
                    "from trp03t where  route_no in( " & _
                    "select  route_no from trp01t where  isnull(c_route_no,route_no)='" & Trim(txt_Tab0_C_Route_No0.Text) & "' and left(route_no,1) <>'S')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��sTRP05T���A
            str_SQL = "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            rs_Tab0_C_RouteList.MoveNext
        End If
        
        cn.CommitTrans: Tran_Level = 0
        
        '[�R���w������q��
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.MoveFirst
        Do While Not rs_Tab0_C_RouteList.EOF
            rs_Tab0_C_RouteList.Delete
            rs_Tab0_C_RouteList.MoveFirst
        Loop
    End If
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    Call WriteOut_RunLog("4.�פJ����")
    DoEvents: DoEvents
    Call Unload_RunLogForm
    '�e���B�z
    Call ReSet_Tab0_C_RouteList_SeqNo
    Set dg_Tab0_RouteList0.DataSource = Nothing
    txt_Tab0_C_Route_No0.Text = ""
    txt_DELIVERY_DATE0.Text = ""
    txt_VehicleNo0.Text = ""
    txt_Driver0.Text = ""
    cmd_Tab0_OK.Enabled = False
    cmd_Tab0_Del.Enabled = False
    Tab0_RouteListEventEnable = True
    Call cmd_Tab0_Clear1_Click
    Exit Sub
    
err_Handle:
   If Tran_Level <> 0 Then cn.RollbackTrans
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-��������", Me.Caption, "cmd_Tab0_Del_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Delete_Click()
    '�D�g��
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    rs_Tab0_C_RouteList.Filter = "��='V'"
    If Not rs_Tab0_C_RouteList.EOF Then
        Do While Not rs_Tab0_C_RouteList.EOF
        
            '��sTRP05T���A
            str_SQL = "Update TRP05T set SDNStatus = '1' ,Receiver ='�D�g��' where Route_No='" & Trim(rs_Tab0_C_RouteList.Fields("���u�s��").Value) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            rs_Tab0_C_RouteList.MoveNext
        Loop
        '[�R���w������q��
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.MoveFirst
        Do While Not rs_Tab0_C_RouteList.EOF
            rs_Tab0_C_RouteList.Delete
            rs_Tab0_C_RouteList.MoveFirst
        Loop
    End If
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    DoEvents: DoEvents
    Call Unload_RunLogForm
    Call ReSet_Tab0_C_RouteList_SeqNo
    Call clear_Tab0_RouteList0
    Call clear_Tab0_RouteList1
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-�D�g��", Me.Caption, "cmd_Tab0_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmd_Tab0_OK_Click()
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
        
    '�����s�O�_�w�T�{
    Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)

    str_SQL = "Select t05t.Route_No as ���u�s�� From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0.Text & "' Union All Select t05t.Route_No as ���u�s�� From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No0.Text & "' "
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then MsgBox "���u�s�� " & txt_Tab0_C_Route_No0.Text & " �w�@�L�X���T�{!", 16, "�`�N": tmp_Rs.Close: cmd_Tab0_OK.Enabled = False: Exit Sub
    
    '�����s�O�_�w�T�{txt_Tab0_C_Route_No1.Text
    Call DB_CheckConnectStatus: Call ReDim_Recordset(tmp_Rs)

    str_SQL = "Select t05t.Route_No as ���u�s�� From TRP05T t05t Where t05t.SDNStatus= '1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No1 & "' Union All Select t05t.Route_No as ���u�s�� From ORT05T t05t where t05t.SDNStatus='1' and t05t.C_ROUTE_NO is null and t05t.Route_No = '" & txt_Tab0_C_Route_No1 & "' "
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then MsgBox "���u�s�� " & txt_Tab0_C_Route_No1 & " �w�@�L�X���T�{!", 16, "�`�N": tmp_Rs.Close: cmd_Tab0_OK.Enabled = False: Exit Sub
    
    On Error GoTo err_Handle
    
    'Terry 20180515 �s�W�O�_���@�̪O���s
    Dim str_PalletDefend As String
    str_PalletDefend = ""
    If chkPalletDefend2.Value = vbChecked Then
        str_PalletDefend = "Y"
    Else
        str_PalletDefend = "N"
    End If
    
    
    
    Tab0_RouteListEventEnable = False
    If Len(Trim(txt_Tab0_C_Route_No0.Text)) > 0 Then
        
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.Filter = "���u�s��='" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
        rs_Tab0_C_RouteList.MoveFirst
                
        If Not rs_Tab0_C_RouteList.EOF Then
        
        Tran_Level = cn.BeginTrans
        
        Call WriteOut_RunLog("1.�}�l >> �s�JSDN01T;SDN02T")
            If Trim(rs_Tab0_C_RouteList.Fields("���O").Value) = "�h�f�q��" Then     '�B�z�h�f�q��
                If rs_Tab0_RouteList0.RecordCount > 0 Then
                    '�s���Y,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE0.Text) & "','" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(txt_VehicleNo0.Text) & "', " & _
                            "'" & Trim(txt_Driver0.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("����").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    rs_Tab0_RouteList0.MoveFirst
                    Do While Not rs_Tab0_RouteList0.EOF
                        '�s�q��,SDN02T
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(rs_Tab0_RouteList0.Fields("���u�s��").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("�Ȥ�渹").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("���").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("���e�Ȥ�").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("�c��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("���n").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("���q").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("�h��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '�ɩһݸ�� by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = ort02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(char(8),ort02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = ort02t.consigneekey " & _
                                    ",sdn02t.description = ort02t.description " & _
                                    ",sdn02t.priority = ort02t.priority " & _
                                    ",sdn02t.scheduledate = ort02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = ort02t.vehicle_id_no " & _
                                    ",sdn02t.c_receipt_no = ort02t.c_receipt_no " & _
                                    "from ort02t join sdn02t on ort02t.receipt_no = sdn02t.receipt_no " & _
                                    "where ort02t.receipt_no = '" & Trim(rs_Tab0_RouteList0.Fields("�q�渹�X").Value) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '�s����,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No0.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ORDER_QTY,SIGN_QTY, SHIP_TIME ,weight,volumn_weight " & _
                            "from ORT03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList0.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        rs_Tab0_RouteList0.MoveNext
                    Loop
                End If
                              
            Else '�@��q��
                If rs_Tab0_RouteList0.RecordCount > 0 Then
                                  
                    '�s���Y,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE0.Text) & "','" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(txt_VehicleNo0.Text) & "', " & _
                            "'" & Trim(txt_Driver0.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("����").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    rs_Tab0_RouteList0.MoveFirst
                    Do While Not rs_Tab0_RouteList0.EOF
                        '�s�q��,SDN02T
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No0.Text) & "','" & Trim(rs_Tab0_RouteList0.Fields("���u�s��").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("�Ȥ�渹").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("���").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("���e�Ȥ�").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("�c��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("���n").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("���q").Value) & "','" & Trim(rs_Tab0_RouteList0.Fields("�h��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList0.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '�ɩһݸ�� by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = trp02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(varchar(8),trp02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = trp02t.consigneekey " & _
                                    ",sdn02t.description = trp02t.description " & _
                                    ",sdn02t.priority = trp02t.priority " & _
                                    ",sdn02t.scheduledate = trp02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = trp02t.vehicle_id_no " & _
                                    ",sdn02t.c_receipt_no = trp02t.c_receipt_no " & _
                                    "from trp02t join sdn02t on sdn02t.receipt_no = trp02t.receipt_no " & _
                                    "where sdn02t.receipt_no = '" & Trim(rs_Tab0_RouteList0.Fields("�q�渹�X").Value) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '�s����,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No0.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY, ship_qty,SIGN_QTY, SHIP_TIME ,weight,volumn_weight " & _
                            "from trp03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList0.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        rs_Tab0_RouteList0.MoveNext
                    Loop
                    
                End If
                
            End If
            
            '��sTRP05T���A
            cn.Execute "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'", RowsAffect, adExecuteNoRecords
                    
            '��sORT05T���A
            cn.Execute "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No0.Text) & "'", RowsAffect, adExecuteNoRecords
                    
            '��sSDN01T���A
            str_SQL = "Update SDN01T set SDN01T.APPStatus = TRP05T.APPStatus ,SDN01T.VLListCount = TRP05T.VLListCount ,SDN01T.VLListPrintDate = TRP05T.VLListPrintDate from trp05t join sdn01t on isnull(trp05t.c_route_no,trp05t.route_no) = sdn01t.c_route_no and sdn01t.c_Route_No= '" & Trim(txt_Tab0_C_Route_No0.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��sAPPOrderDate���A
            str_SQL = "update AppOrderDate set status = '5' where status < '6' and c_Route_No= '" & Trim(txt_Tab0_C_Route_No0.Text) & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��s�дڤH
            cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & Trim(txt_VehicleNo0.Text) & "') where c_route_no = '" & txt_Tab0_C_Route_No0.Text & "'", RowsAffect, adExecuteNoRecords
                        
            '��sOrderType Add by Gemini @20190604
            str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & txt_Tab0_C_Route_No0.Text & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & txt_Tab0_C_Route_No0.Text & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
            cn.CommitTrans: Tran_Level = 0
            
            '[�R���w������q��
            DoEvents: DoEvents
            rs_Tab0_C_RouteList.MoveFirst
            Do While Not rs_Tab0_C_RouteList.EOF
                rs_Tab0_C_RouteList.Delete
                rs_Tab0_C_RouteList.MoveFirst
            Loop
        End If
        
        '�e���B�z
        Call clear_Tab0_RouteList0
        rs_Tab0_C_RouteList.Filter = adFilterNone
        rs_Tab0_C_RouteList.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    End If
    
    If Len(Trim(txt_Tab0_C_Route_No1.Text)) > 0 Then
        Call WriteOut_RunLog("1.1.�}�l >> �s�JSDN01T")
        DoEvents: DoEvents
        rs_Tab0_C_RouteList.Filter = "���u�s��='" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
        rs_Tab0_C_RouteList.MoveFirst
        If Not rs_Tab0_C_RouteList.EOF Then
            cn.BeginTrans
            If Trim(rs_Tab0_C_RouteList.Fields("���O").Value) = "�h�f�q��" Then     '�B�z�h�f�q��
                If rs_Tab0_RouteList1.RecordCount > 0 Then
                    '�s���Y,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE1.Text) & "','" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(txt_VehicleNo1.Text) & "', " & _
                            "'" & Trim(txt_Driver1.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("����").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    '�s�q��,SDN02T
                    rs_Tab0_RouteList1.MoveFirst
                    Do While Not rs_Tab0_RouteList1.EOF
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(rs_Tab0_RouteList1.Fields("���u�s��").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("�Ȥ�渹").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("���").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("���e�Ȥ�").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("�c��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("���n").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("���q").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("�h��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '�ɩһݸ�� by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = ort02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(varchar(8),ort02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = ort02t.consigneekey " & _
                                    ",sdn02t.description = ort02t.description " & _
                                    ",sdn02t.priority = ort02t.priority " & _
                                    ",sdn02t.scheduledate = ort02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = '" & Trim(txt_VehicleNo1.Text) & "' " & _
                                    ",sdn02t.c_receipt_no = ort02t.c_receipt_no " & _
                                    "from ort02t join sdn02t on sdn02t.receipt_no = ort02t.receipt_no " & _
                                    "where ort02t.receipt_no = '" & Trim(rs_Tab0_RouteList1.Fields("�q�渹�X").Value) & "'"
                         cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                         
                        '�s����,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No1.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY, ship_qty,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                            "from ORT03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList1.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                                              
                       '�ɩһݸ�� by gemini
                       str_SQL = "Update sdn03t Set sdn03t.Weight = ort03t.Weight ,sdn03t.volumn_weight = ort03t.volumn_weight " & _
                            "from ort03t join sdn03t on sdn03t.receipt_no = ort03t.receipt_no and sdn03t.seq_no = ort03t.seq_no and isnull(sdn03t.subseq_no,'') = isnull(ort03t.subseq_no,'') " & _
                            "where sdn03t.receipt_no = '" & Trim(rs_Tab0_RouteList1.Fields("�q�渹�X").Value) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                       rs_Tab0_RouteList1.MoveNext
                    Loop
                End If
            
            Else
                If rs_Tab0_RouteList1.RecordCount > 0 Then
                    '�s���Y,SDN01T
                    str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times,PalletDefend) " & _
                            "Values ( '" & Trim(txt_DELIVERY_DATE1.Text) & "','" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(txt_VehicleNo1.Text) & "', " & _
                            "'" & Trim(txt_Driver1.Text) & "','0','" & User_id & "','" & Trim(rs_Tab0_C_RouteList.Fields("����").Value) & "', '" & str_PalletDefend & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    '�s�q��,SDN02T
                    rs_Tab0_RouteList1.MoveFirst
                    Do While Not rs_Tab0_RouteList1.EOF
                        str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO) " & _
                            "Values ( '" & Trim(txt_Tab0_C_Route_No1.Text) & "','" & Trim(rs_Tab0_RouteList1.Fields("���u�s��").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("�Ȥ�渹").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("���").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("���e�Ȥ�").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("�c��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("���n").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("���q").Value) & "','" & Trim(rs_Tab0_RouteList1.Fields("�h��").Value) & "', " & _
                            "'" & Trim(rs_Tab0_RouteList1.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '�ɩһݸ�� by gemini
                        str_SQL = "Update sdn02t " & _
                                    "Set sdn02t.storerkey = trp02t.storerkey " & _
                                    ",sdn02t.receipt_date = convert(varchar(8),trp02t.receipt_date,112) " & _
                                    ",sdn02t.consigneekey = trp02t.consigneekey " & _
                                    ",sdn02t.description = trp02t.description " & _
                                    ",sdn02t.priority = trp02t.priority " & _
                                    ",sdn02t.scheduledate = trp02t.scheduledate " & _
                                    ",sdn02t.scan = 'N' " & _
                                    ",sdn02t.vehicle_id_no = '" & Trim(txt_VehicleNo1.Text) & "' " & _
                                    ",sdn02t.c_receipt_no = trp02t.c_receipt_no " & _
                                    "from trp02t join sdn02t on sdn02t.receipt_no = trp02t.receipt_no " & _
                                    "where sdn02t.receipt_no = '" & Trim(rs_Tab0_RouteList1.Fields("�q�渹�X").Value) & "'"
                         cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        '�s����,SDN03T
                        str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight) " & _
                            "select  '" & Trim(txt_Tab0_C_Route_No1.Text) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME,weight,volumn_weight " & _
                            "from trp03t where  RECEIPT_NO in('" & Trim(rs_Tab0_RouteList1.Fields("�q�渹�X").Value) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
                        rs_Tab0_RouteList1.MoveNext
                    Loop
                End If
                
            End If
            
            '��sTRP05T���A
            str_SQL = "Update TRP05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��sORT05T���A
            str_SQL = "Update ORT05T set SDNStatus = '1' where Route_No='" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��sSDN01T���A
            str_SQL = "Update SDN01T set SDN01T.APPStatus = TRP05T.APPStatus ,SDN01T.VLListCount = TRP05T.VLListCount ,SDN01T.VLListPrintDate = TRP05T.VLListPrintDate from trp05t join sdn01t on isnull(trp05t.c_route_no,trp05t.route_no) = sdn01t.c_route_no and sdn01t.c_Route_No= '" & Trim(txt_Tab0_C_Route_No1.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��sAPPOrderDate���A
            str_SQL = "update AppOrderDate set status = '5' where status < '6' and c_Route_No= '" & Trim(txt_Tab0_C_Route_No1.Text) & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��s�дڤH
            cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & Trim(txt_VehicleNo1.Text) & "') where c_route_no = '" & txt_Tab0_C_Route_No1.Text & "'", RowsAffect, adExecuteNoRecords
            
            '��sOrderType Add by Gemini @20190604
            str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & txt_Tab0_C_Route_No1.Text & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & txt_Tab0_C_Route_No1.Text & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            cn.CommitTrans: Tran_Level = 0
            
            '[�R���w������q��
            DoEvents: DoEvents
            rs_Tab0_C_RouteList.MoveFirst
            Do While Not rs_Tab0_C_RouteList.EOF
                rs_Tab0_C_RouteList.Delete
                rs_Tab0_C_RouteList.MoveFirst
            Loop
        End If
        
        '�e���B�z
        Call clear_Tab0_RouteList1
        rs_Tab0_C_RouteList.Filter = adFilterNone
        rs_Tab0_C_RouteList.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    End If
    
    Call WriteOut_RunLog("4.�פJ����")
    DoEvents: DoEvents
    Call Unload_RunLogForm
    
    '�e���B�z
    Call ReSet_Tab0_C_RouteList_SeqNo
    cmd_Tab0_OK.Enabled = False
    cmd_Tab0_Del.Enabled = False
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   If Tran_Level <> 0 Then cn.RollbackTrans
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-�T�{", Me.Caption, "cmd_Tab0_Del_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Remove_Click()
    '�z��w�����
    
    '�ư��h�f�q�� Gemini @ 20060728
    If Left(txt_Tab0_C_Route_No0.Text, 1) = "R" Or Left(txt_Tab0_C_Route_No1.Text, 1) = "R" Then If Left(txt_Tab0_C_Route_No0.Text, 1) <> Left(txt_Tab0_C_Route_No1.Text, 1) Then MsgBox "�h�f���s�L�k�֤J��L���s�I", vbOKOnly, Me.Caption: Exit Sub
    
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList1.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    rs_Tab0_RouteList0.Filter = "��='V'"
    If Not rs_Tab0_RouteList0.EOF Then
       Do While Not rs_Tab0_RouteList0.EOF
          '�P�_�O�_�w�g����L
          rs_Tab0_RouteList1.Filter = adFilterNone
          rs_Tab0_RouteList1.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
          rs_Tab0_RouteList1.Filter = "�q�渹�X = '" & rs_Tab0_RouteList0.Fields("�q�渹�X").Value & "'"
          '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
          If blRouteModify Then blRouteChange = True
          If rs_Tab0_RouteList1.EOF Then
             '�s�W������q��
             rs_Tab0_RouteList1.AddNew
             rs_Tab0_RouteList1.Fields("�s��").Value = 999
             rs_Tab0_RouteList1.Fields("�G���ƨ�").Value = rs_Tab0_RouteList0.Fields("�G���ƨ�").Value
             rs_Tab0_RouteList1.Fields("�Ȥ�渹").Value = rs_Tab0_RouteList0.Fields("�Ȥ�渹").Value
             rs_Tab0_RouteList1.Fields("���u�s��").Value = rs_Tab0_RouteList0.Fields("���u�s��").Value
             rs_Tab0_RouteList1.Fields("���").Value = rs_Tab0_RouteList0.Fields("���").Value
             rs_Tab0_RouteList1.Fields("���e�Ȥ�").Value = rs_Tab0_RouteList0.Fields("���e�Ȥ�").Value
             rs_Tab0_RouteList1.Fields("�c��").Value = rs_Tab0_RouteList0.Fields("�c��").Value
             rs_Tab0_RouteList1.Fields("���q").Value = rs_Tab0_RouteList0.Fields("���q").Value
             rs_Tab0_RouteList1.Fields("���n").Value = rs_Tab0_RouteList0.Fields("���n").Value
             rs_Tab0_RouteList1.Fields("�h��").Value = rs_Tab0_RouteList0.Fields("�h��").Value
             rs_Tab0_RouteList1.Fields("�q�渹�X").Value = rs_Tab0_RouteList0.Fields("�q�渹�X").Value
             rs_Tab0_RouteList1.Update
          Else
             '��s������q����
             rs_Tab0_RouteList1.Fields("�G���ƨ�").Value = rs_Tab0_RouteList0.Fields("�G���ƨ�").Value
             rs_Tab0_RouteList1.Fields("�Ȥ�渹").Value = rs_Tab0_RouteList0.Fields("�Ȥ�渹").Value
             rs_Tab0_RouteList1.Fields("���u�s��").Value = rs_Tab0_RouteList0.Fields("���u�s��").Value
             rs_Tab0_RouteList1.Fields("���").Value = rs_Tab0_RouteList0.Fields("���").Value
             rs_Tab0_RouteList1.Fields("���e�Ȥ�").Value = rs_Tab0_RouteList0.Fields("���e�Ȥ�").Value
             rs_Tab0_RouteList1.Fields("�c��").Value = rs_Tab0_RouteList0.Fields("�c��").Value
             rs_Tab0_RouteList1.Fields("���q").Value = rs_Tab0_RouteList0.Fields("���q").Value
             rs_Tab0_RouteList1.Fields("���n").Value = rs_Tab0_RouteList0.Fields("���n").Value
             rs_Tab0_RouteList1.Fields("�h��").Value = rs_Tab0_RouteList0.Fields("�h��").Value
             rs_Tab0_RouteList1.Fields("�q�渹�X").Value = rs_Tab0_RouteList0.Fields("�q�渹�X").Value
          End If
          rs_Tab0_RouteList0.MoveNext
       Loop
       
       '[�R���w������q��
       rs_Tab0_RouteList0.MoveFirst
       Do While Not rs_Tab0_RouteList0.EOF
          rs_Tab0_RouteList0.Delete
          rs_Tab0_RouteList0.MoveFirst
       Loop
        
    End If
    rs_Tab0_RouteList1.Filter = adFilterNone
    rs_Tab0_RouteList1.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    rs_Tab0_RouteList0.Filter = adFilterNone
    rs_Tab0_RouteList0.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    Call ReSet_Tab0_RouteList1_SeqNo
    Call ReSet_Tab0_RouteList0_SeqNo
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-����V�U", Me.Caption, "cmd_Tab0_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_RouteSelect0_Click()
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If Trim(txt_Tab0_C_Route_No1.Text) = Trim(rs_Tab0_C_RouteList.Fields(2).Value) Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    str_route = Trim(rs_Tab0_C_RouteList.Fields(2).Value)
    txt_Tab0_C_Route_No0.Text = str_route
    txt_DELIVERY_DATE0.Text = Trim(rs_Tab0_C_RouteList.Fields(3).Value)
    txt_VehicleNo0.Text = Trim(rs_Tab0_C_RouteList.Fields(4).Value)
    txt_Driver0.Text = Trim(rs_Tab0_C_RouteList.Fields(6).Value)
    If Trim(rs_Tab0_C_RouteList.Fields("���O").Value) = "�h�f�q��" Then
                str_SQL = "SELECT  ' ' as '��',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS �G���ƨ�, t2.ROUTE_NO AS ���u�s��, t2.EXTERN AS �Ȥ�渹, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS ���, m1.FULL_NAME as ���e�Ȥ�,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.ORDER_QTY / sp.Casecnt end , 2), 0)) AS �c��, " & _
                  "SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDCUBE, 2), 0)) AS ���n,SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDGrossWGT, 2), 0)) AS ���q, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS �h��,t2.RECEIPT_NO as �q�渹�X  " & _
                  "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    Else
        str_SQL = "SELECT  ' ' as '��',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS �G���ƨ�, t2.ROUTE_NO AS ���u�s��, t2.EXTERN AS �Ȥ�渹, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS ���, m1.FULL_NAME as ���e�Ȥ�,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.SHIP_QTY / sp.Casecnt end , 2), 0)) AS �c��, " & _
                  "SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDCUBE, 2), 0)) AS ���n,SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDGrossWGT, 2), 0)) AS ���q, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS �h��,t2.RECEIPT_NO as �q�渹�X  " & _
                  "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    End If
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Call Replication_Recordset(tmp_Rs, rs_Tab0_RouteList0)
    Set dg_Tab0_RouteList0.DataSource = rs_Tab0_RouteList0
    tmp_Rs.Close
    With dg_Tab0_RouteList0
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500       '���
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000      '�G���ƨ�
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000      '�Ȥ�渹
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800      '���u�s��
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 800       '���
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1200      '���e�Ȥ�
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 800       '�c��
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '���n
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 800       '���q
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 800       '�h��
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 1000       '�q�渹�X
        .Columns(11).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    cmd_Tab0_Delete.Enabled = True
    cmd_Tab0_Del.Enabled = True
    cmd_Tab0_OK.Enabled = True
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-���s����W", Me.Caption, "cmd_Tab0_RouteSelect0_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_RouteSelect1_Click()
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If Trim(txt_Tab0_C_Route_No0.Text) = Trim(rs_Tab0_C_RouteList.Fields(2).Value) Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    str_route = Trim(rs_Tab0_C_RouteList.Fields(2).Value)
    txt_Tab0_C_Route_No1.Text = str_route
    txt_DELIVERY_DATE1.Text = Trim(rs_Tab0_C_RouteList.Fields(3).Value)
    txt_VehicleNo1.Text = Trim(rs_Tab0_C_RouteList.Fields(4).Value)
    txt_Driver1.Text = Trim(rs_Tab0_C_RouteList.Fields(6).Value)
    If Trim(rs_Tab0_C_RouteList.Fields("���O").Value) = "�h�f�q��" Then
                str_SQL = "SELECT  ' ' as '��',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS �G���ƨ�, t2.ROUTE_NO AS ���u�s��, t2.EXTERN AS �Ȥ�渹, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS ���, m1.FULL_NAME as ���e�Ȥ�,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.ORDER_QTY / sp.Casecnt end , 2), 0)) AS �c��, " & _
                  "SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDCUBE, 2), 0)) AS ���n,SUM(ISNULL(ROUND(t3.ORDER_QTY * sp.STDGrossWGT, 2), 0)) AS ���q, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS �h��,t2.RECEIPT_NO as �q�渹�X " & _
                  "FROM  dbo.ORT02T t2 INNER JOIN dbo.ORT03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.ORT05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    Else
        str_SQL = "SELECT  ' ' as '��',ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) AS �G���ƨ�, t2.ROUTE_NO AS ���u�s��, t2.EXTERN AS �Ȥ�渹, CONVERT(varchar(8), " & _
                  "t2.ARRIVE_DATE, 112) AS ���, m1.FULL_NAME as ���e�Ȥ�,SUM(ISNULL(ROUND(case when sp.Casecnt = 0 then 0 else t3.SHIP_QTY / sp.Casecnt end , 2), 0)) AS �c��, " & _
                  "SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDCUBE, 2), 0)) AS ���n,SUM(ISNULL(ROUND(t3.SHIP_QTY * sp.STDGrossWGT, 2), 0)) AS ���q, " & _
                  "ISNULL(t2.CAR_NOTES, '') AS �h��,t2.RECEIPT_NO as �q�渹�X  " & _
                  "FROM  dbo.TRP02T t2 INNER JOIN dbo.TRP03T t3 ON t2.ROUTE_NO = t3.ROUTE_NO AND t3.RECEIPT_NO = t2.RECEIPT_NO  " & _
                  "INNER JOIN gv_SKUxpack sp ON sp.Sku = t3.PRODUCT_NO and sp.storerkey = t3.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP01M m1 ON t2.CONSIGNEEKEY = m1.CONSIGNEEKEY and m1.storerkey = t2.storerkey " & _
                  "LEFT OUTER JOIN dbo.TRP05T t5 ON t2.ROUTE_NO = t5.ROUTE_NO  " & _
                  "WHERE ISNULL(t5.C_ROUTE_NO, t2.ROUTE_NO) ='" & str_route & "' " & _
                  "GROUP BY  t5.C_ROUTE_NO, t2.ROUTE_NO, t2.EXTERN, CONVERT(varchar(8),t2.ARRIVE_DATE, 112), m1.FULL_NAME, t2.CAR_NOTES,t2.RECEIPT_NO"
    End If
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    cmd_Tab0_Delete.Enabled = True
    Call Replication_Recordset(tmp_Rs, rs_Tab0_RouteList1)
    Set dg_Tab0_RouteList1.DataSource = rs_Tab0_RouteList1
    tmp_Rs.Close
    With dg_Tab0_RouteList1
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500       '���
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000      '�G���ƨ�
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000      '�Ȥ�渹
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800      '���u�s��
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 800       '���
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1200      '���e�Ȥ�
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 800       '�c��
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '���n
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 800       '���q
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 800       '�h��
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 1000       '�q�渹�X
        .Columns(11).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-���s����U", Me.Caption, "cmd_Tab0_RouteSelect1_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_SelectCar0_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar0.Name & "2")
End Sub

Private Sub cmd_Tab0_SelectCar1_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar1.Name & "2")
End Sub

Private Sub cmd_Tab0_Selected_Click()
    '�z��w�����
    
    '�ư��h�f�q�� Gemini @ 20060728
    If Left(txt_Tab0_C_Route_No0.Text, 1) = "R" Or Left(txt_Tab0_C_Route_No1.Text, 1) = "R" Then If Left(txt_Tab0_C_Route_No0.Text, 1) <> Left(txt_Tab0_C_Route_No1.Text, 1) Then MsgBox "�h�f���s�L�k�֤J��L���s�I", vbOKOnly, Me.Caption: Exit Sub
   
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If dg_Tab0_RouteList1.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab0_RouteListEventEnable = False
    rs_Tab0_RouteList1.Filter = "��='V'"
    If Not rs_Tab0_RouteList1.EOF Then
       Do While Not rs_Tab0_RouteList1.EOF
          '�P�_�O�_�w�g����L
          rs_Tab0_RouteList0.Filter = adFilterNone
          rs_Tab0_RouteList0.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
          rs_Tab0_RouteList0.Filter = "�q�渹�X = '" & rs_Tab0_RouteList1.Fields("�q�渹�X").Value & "'"
          '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
          If blRouteModify Then blRouteChange = True
          If rs_Tab0_RouteList0.EOF Then
             '�s�W������q��
             rs_Tab0_RouteList0.AddNew
             rs_Tab0_RouteList0.Fields("�s��").Value = 999
             rs_Tab0_RouteList0.Fields("�G���ƨ�").Value = rs_Tab0_RouteList1.Fields("�G���ƨ�").Value
             rs_Tab0_RouteList0.Fields("�Ȥ�渹").Value = rs_Tab0_RouteList1.Fields("�Ȥ�渹").Value
             rs_Tab0_RouteList0.Fields("���u�s��").Value = rs_Tab0_RouteList1.Fields("���u�s��").Value
             rs_Tab0_RouteList0.Fields("���").Value = rs_Tab0_RouteList1.Fields("���").Value
             rs_Tab0_RouteList0.Fields("���e�Ȥ�").Value = rs_Tab0_RouteList1.Fields("���e�Ȥ�").Value
             rs_Tab0_RouteList0.Fields("�c��").Value = rs_Tab0_RouteList1.Fields("�c��").Value
             rs_Tab0_RouteList0.Fields("���q").Value = rs_Tab0_RouteList1.Fields("���q").Value
             rs_Tab0_RouteList0.Fields("���n").Value = rs_Tab0_RouteList1.Fields("���n").Value
             rs_Tab0_RouteList0.Fields("�h��").Value = rs_Tab0_RouteList1.Fields("�h��").Value
             rs_Tab0_RouteList0.Fields("�q�渹�X").Value = rs_Tab0_RouteList1.Fields("�q�渹�X").Value
             rs_Tab0_RouteList0.Update
          Else
             '��s������q����
             rs_Tab0_RouteList0.Fields("�G���ƨ�").Value = rs_Tab0_RouteList1.Fields("�G���ƨ�").Value
             rs_Tab0_RouteList0.Fields("�Ȥ�渹").Value = rs_Tab0_RouteList1.Fields("�Ȥ�渹").Value
             rs_Tab0_RouteList0.Fields("���u�s��").Value = rs_Tab0_RouteList1.Fields("���u�s��").Value
             rs_Tab0_RouteList0.Fields("���").Value = rs_Tab0_RouteList1.Fields("���").Value
             rs_Tab0_RouteList0.Fields("���e�Ȥ�").Value = rs_Tab0_RouteList1.Fields("���e�Ȥ�").Value
             rs_Tab0_RouteList0.Fields("�c��").Value = rs_Tab0_RouteList1.Fields("�c��").Value
             rs_Tab0_RouteList0.Fields("���q").Value = rs_Tab0_RouteList1.Fields("���q").Value
             rs_Tab0_RouteList0.Fields("���n").Value = rs_Tab0_RouteList1.Fields("���n").Value
             rs_Tab0_RouteList0.Fields("�h��").Value = rs_Tab0_RouteList1.Fields("�h��").Value
             rs_Tab0_RouteList0.Fields("�q�渹�X").Value = rs_Tab0_RouteList1.Fields("�q�渹�X").Value
          End If
          rs_Tab0_RouteList1.MoveNext
       Loop
       
       '[�R���w������q��
       rs_Tab0_RouteList1.MoveFirst
       Do While Not rs_Tab0_RouteList1.EOF
          rs_Tab0_RouteList1.Delete
          rs_Tab0_RouteList1.MoveFirst
       Loop
       Call ReSet_Tab0_RouteList1_SeqNo
       Call ReSet_Tab0_RouteList0_SeqNo
    End If
    rs_Tab0_RouteList1.Filter = adFilterNone
    rs_Tab0_RouteList1.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    rs_Tab0_RouteList0.Filter = adFilterNone
    rs_Tab0_RouteList0.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    Tab0_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-����V�W", Me.Caption, "cmd_Tab0_Selected_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmd_Tab1_CreateRoute_Click()
    If rs_Tab1_SelectedOrders.RecordCount = 0 Then
        msg_text = "��ƿ��~�G�L�˸����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(txt_Tab1_DELIVERY_DATE.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J�X�����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    If Len(Trim(txt_Tab1_VehicleNo.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J���P���X"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_VehicleNo.SetFocus
        Exit Sub
    End If
    '�X������G�榡 yyyymmdd
    txt_Tab1_DELIVERY_DATE.Text = Trim(txt_Tab1_DELIVERY_DATE.Text)
    If Fun_ChkDateFormat(txt_Tab1_DELIVERY_DATE.Text) = 1 Then
        msg_text = "�X������G" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_DELIVERY_DATE.SelStart = 0: txt_Tab1_DELIVERY_DATE.SelLength = Len(txt_Tab1_DELIVERY_DATE.Text): txt_Tab1_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    On Error GoTo err_Handle
    Tab1_RouteListEventEnable = False
    
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    '�ˮ� [���P���X] �O�_����
    txt_Tab1_VehicleNo.Text = Trim(txt_Tab1_VehicleNo.Text)
    str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab1_VehicleNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       msg_text = "��ƿ��~�G���P���X " & txt_Tab1_VehicleNo.Text & " ������"
       MsgBox msg_text, vbOKOnly + vbCritical, msg_title
       txt_Tab1_VehicleNo.SelStart = 0: txt_Tab1_VehicleNo.SelLength = Len(txt_Tab1_VehicleNo.Text)
       txt_Tab1_VehicleNo.SetFocus
       Exit Sub
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    Dim intDriveTimes As Integer    '����
    Dim strRouteNo As String        '���u�s��
    
    '���ͨ���
    str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
              "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab1_DELIVERY_DATE.Text & "' and Vehicle_ID_No = '" & txt_Tab1_VehicleNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    tmp_Rs.Close
    
    '���͸��u�s��
    str_SQL = "Select Isnull(Max(Cast(Right(C_ROUTE_NO,3) as integer))+1,1) as RouteSN " & _
              "From SDN01T Where Substring(C_ROUTE_NO,2,6)='" & Mid(txt_Tab1_DELIVERY_DATE.Text, 3, 6) & "' and Left(C_ROUTE_NO,1) = 'N'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strRouteNo = "N" & Mid(txt_Tab1_DELIVERY_DATE, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
    tmp_Rs.Close
    DoEvents: DoEvents
    Tran_Level = cn.BeginTrans
    
        '�s���Y,SDN01T
        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,SDNStatus,AddUser,Drive_Times) " & _
            "Values ( '" & Trim(txt_Tab1_DELIVERY_DATE.Text) & "','" & Trim(strRouteNo) & "','" & Trim(txt_Tab1_VehicleNo.Text) & "', " & _
            "'" & Trim(txt_Tab1_Driver0.Text) & "','0','" & User_id & "','" & intDriveTimes & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
        '��s�дڤH
        cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & txt_Tab1_VehicleNo.Text & "') where c_route_no = '" & strRouteNo & "'", RowsAffect, adExecuteNoRecords
                    
        rs_Tab1_SelectedOrders.MoveFirst
        Do While Not rs_Tab1_SelectedOrders.EOF
            '�s�q��,SDN02T
            str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,RECEIPT_NO,C_RECEIPT_NO) " & _
                    "Values ( '" & Trim(strRouteNo) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("���u�s��").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("�Ȥ�渹").Value) & "', " & _
                    "'" & Trim(rs_Tab1_SelectedOrders.Fields("���").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("���e�Ȥ�").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("�c��").Value) & "', " & _
                    "'" & Trim(rs_Tab1_SelectedOrders.Fields("���n").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("���q").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("�h��").Value) & "', " & _
                    "'" & Trim(rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value) & "','" & Trim(rs_Tab1_SelectedOrders.Fields("C_RECEIPT_NO").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

            '�ɤ@��q��һݸ�� by gemini
            str_SQL = "Update sdn02t " & _
                        "Set sdn02t.storerkey = orders.storerkey " & _
                        ",sdn02t.receipt_date = convert(varchar(8),orders.orderdate,112) " & _
                        ",sdn02t.consigneekey = orders.consigneekey " & _
                        ",sdn02t.description = orders.notes " & _
                        ",sdn02t.priority = orders.priority " & _
                        ",sdn02t.scheduledate = (select top 1 trp02t.scheduledate from trp02t where trp02t.c_receipt_no = orders.orderkey order by trp02t.scheduledate ) " & _
                        ",sdn02t.scan = 'N' " & _
                        ",sdn02t.vehicle_id_no = '" & Trim(txt_Tab1_VehicleNo.Text) & "' " & _
                        "from orders join sdn02t on sdn02t.c_receipt_no = orders.orderkey " & _
                        "where sdn02t.receipt_no = '" & Trim(rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value) & "' and sdn02t.priority not in ('R','A2B','RC') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�ɰh�f�q��һݸ�� by gemini
            str_SQL = "Update sdn02t " & _
                        "Set sdn02t.storerkey = orders.storerkey " & _
                        ",sdn02t.receipt_date = convert(varchar(8),orders.orderdate,112) " & _
                        ",sdn02t.consigneekey = orders.consigneekey " & _
                        ",sdn02t.description = orders.notes " & _
                        ",sdn02t.priority = orders.priority " & _
                        ",sdn02t.scheduledate = (select top 1 ort02t.scheduledate from ort02t where ort02t.c_receipt_no = orders.orderkey order by ort02t.scheduledate ) " & _
                        ",sdn02t.scan = 'N' " & _
                        ",sdn02t.vehicle_id_no = '" & Trim(txt_Tab1_VehicleNo.Text) & "' " & _
                        "from orders join sdn02t on sdn02t.c_receipt_no = orders.orderkey " & _
                        "where sdn02t.receipt_no = '" & Trim(rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value) & "' and sdn02t.priority in ('R','A2B','RC') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�s����,SDN03T
            str_SQL = "Insert into SDN03T (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,SHIP_QTY,SIGN_QTY, SHIP_TIME) " & _
                "select  '" & Trim(strRouteNo) & "' as C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,ship_qty,SIGN_QTY, SHIP_TIME " & _
                "from SDN03W where  RECEIPT_NO in('" & Trim(rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�ɭq���� by gemini
             str_SQL = "Update sdn03t Set sdn03t.Weight = sdn03t.order_qty * sp.stdgrosswgt ,sdn03t.volumn_weight = sdn03t.order_qty * sp.stdcube " & _
                  "from gv_skuxpack sp join sdn03t on sdn03t.product_no = sp.sku and sp.storerkey = sdn03t.storerkey " & _
                  "where sdn03t.receipt_no = '" & Trim(rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value) & "'"
              cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                      
            '�R���q��,SDN02W
            str_SQL = "delete SDN02W where RECEIPT_NO in('" & Trim(rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�R������,SDN03W
            str_SQL = "delete SDN03W where RECEIPT_NO in('" & Trim(rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            rs_Tab1_SelectedOrders.MoveNext
        Loop
        
        '�R���w������q��
        DoEvents: DoEvents
        rs_Tab1_SelectedOrders.MoveFirst
        Do While Not rs_Tab1_SelectedOrders.EOF
            rs_Tab1_SelectedOrders.Delete
            rs_Tab1_SelectedOrders.MoveFirst
        Loop
        
'        '��sAPPOrderDate���A
'        str_SQL = "update AppOrderDate set status = '5' where and status < '6' and c_Route_No= '" & Trim(strRouteNo) & "' "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

            '��sOrderType Add by Gemini @20190604
            str_SQL = "update SDN01T set SDN01T.OrderType = isnull((SELECT distinct rtrim(s2.priority) + ',' FROM sdn02t s2 where s2.c_route_no = '" & strRouteNo & "' " & "order by rtrim(s2.priority)+ ',' FOR XML PATH('')),'')  where sdn01t.c_route_no = '" & strRouteNo & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
    cn.CommitTrans
    
    '�e���B�z
    txt_Tab1_DELIVERY_DATE.Text = ""
    txt_Tab1_VehicleNo.Text = ""
    txt_Tab1_Driver0.Text = ""
    txt_Tab1_Selected_Case.Text = "0"
    txt_Tab1_Selected_Volumn.Text = "0"
    txt_Tab1_Selected_Weight.Text = "0"
    '��ܷs�ؤ����s
    txt_Tab2_Route_Start.Text = strRouteNo
    Call cmd_Tab2_RouteNoQuery_Click
    Tab1_RouteListEventEnable = True
    SSTab1.Tab = 2
    Exit Sub
    
err_Handle:
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, Me.Name & "�X���T�{-�إ߸��u�s��")
End Sub

Private Sub cmd_Tab1_ImportOrders_Click()

 '��s�c�������
 cn.Execute "exec gs_UpdateSDNW", RowsAffect, adExecuteNoRecords
 
 '�X���T�{>>�פJ�ݭ���q��
 Screen.MousePointer = vbHourglass
 DoEvents: DoEvents
 Tab1_RouteListEventEnable = False
 Set dg_SDN02W.DataSource = Nothing
 strSourceFilter = adFilterNone
 DoEvents
 
 Call CreateRS_Tab1_SelectedOrders
 '�ݱƨ��q����J�G����p�p�G�k�s
 txt_Tab1_srcSelected_Case.Text = ""
 txt_Tab1_srcSelected_Volumn.Text = ""
 txt_Tab1_srcSelected_Weight.Text = ""
 
 '���^�ݱƨ��q��
 str_SQL = "SELECT  ' ' as '��',C_ROUTE_NO AS �G���ƨ�, ROUTE_NO AS ���u�s��,EXTERN AS �Ȥ�渹,ARRIVE_DATE AS ���,CUST_NAME as ���e�Ȥ�, " & _
         "SHIP_CS As �c��, SHIP_CBM As ���n, SHIP_WT As ���q, RECEIPT_NO As �q�渹�X, CAR_NOTES As �h�� , C_Receipt_no " & _
         "FROM dbo.SDN02W Order by �G���ƨ�,���u�s��,�Ȥ�渹,�c�� "
 strSourceOrderBy = " �G���ƨ�,���u�s��,�Ȥ�渹,�c "
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
 Call Replication_Recordset(tmp_Rs, rs_SDN02W)
 Set dg_SDN02W.DataSource = rs_SDN02W
 tmp_Rs.Close
 With dg_SDN02W
     .ColumnHeaders = True         '���D�����
     .RowHeight = 250
     .Columns(0).Width = 500       '�Ǹ�
     .Columns(0).Alignment = dbgLeft
     .Columns(1).Width = 500       '���
     .Columns(1).Alignment = dbgCenter
     .Columns(2).Width = 1000      '�G���ƨ�
     .Columns(2).Alignment = dbgLeft
     .Columns(3).Width = 1000      '�Ȥ�渹
     .Columns(3).Alignment = dbgLeft
     .Columns(4).Width = 800      '���u�s��
     .Columns(4).Alignment = dbgLeft
     .Columns(5).Width = 800       '���
     .Columns(5).Alignment = dbgLeft
     .Columns(6).Width = 1500      '���e�Ȥ�
     .Columns(6).Alignment = dbgLeft
     .Columns(7).Width = 800       '�c��
     .Columns(7).Alignment = dbgRight
     .Columns(8).Width = 800       '���n
     .Columns(8).Alignment = dbgRight
     .Columns(9).Width = 800       '���q
     .Columns(9).Alignment = dbgRight
     .Columns(10).Width = 1000       '�q�渹�X
     .Columns(10).Alignment = dbgLeft
     .Columns(11).Width = 800       '�h��
     .Columns(11).Alignment = dbgLeft
 End With
 DoEvents: DoEvents

 rs_SDN02W.MoveFirst
 '�ݱƨ��q���`�p��T
 Call Retrive_OrderSum
 Screen.MousePointer = vbDefault
 Tab1_RouteListEventEnable = True
 Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{>>�פJ�ݭ���q��", Me.Caption, "cmd_Tab1_ImportOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Remove_Click()
    '�z��w�����
    If dg_Tab1_SelectedOrders.DataSource Is Nothing Then Exit Sub
    If dg_SDN02W.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab1_RouteListEventEnable = False
    rs_Tab1_SelectedOrders.Filter = "��='V'"
    If Not rs_Tab1_SelectedOrders.EOF Then
        Do While Not rs_Tab1_SelectedOrders.EOF
            '�P�_�O�_�w�g����L
            rs_SDN02W.Filter = adFilterNone
            rs_SDN02W.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
            rs_SDN02W.Filter = "�q�渹�X = '" & rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value & "'"
            '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
            If blRouteModify Then blRouteChange = True
            If rs_SDN02W.EOF Then
                '�s�W������q��
                rs_SDN02W.AddNew
                rs_SDN02W.Fields("�s��").Value = 999
                rs_SDN02W.Fields("�G���ƨ�").Value = rs_Tab1_SelectedOrders.Fields("�G���ƨ�").Value
                rs_SDN02W.Fields("�Ȥ�渹").Value = rs_Tab1_SelectedOrders.Fields("�Ȥ�渹").Value
                rs_SDN02W.Fields("���u�s��").Value = rs_Tab1_SelectedOrders.Fields("���u�s��").Value
                rs_SDN02W.Fields("���").Value = rs_Tab1_SelectedOrders.Fields("���").Value
                rs_SDN02W.Fields("���e�Ȥ�").Value = rs_Tab1_SelectedOrders.Fields("���e�Ȥ�").Value
                rs_SDN02W.Fields("�c��").Value = rs_Tab1_SelectedOrders.Fields("�c��").Value
                rs_SDN02W.Fields("���q").Value = rs_Tab1_SelectedOrders.Fields("���q").Value
                rs_SDN02W.Fields("���n").Value = rs_Tab1_SelectedOrders.Fields("���n").Value
                rs_SDN02W.Fields("�h��").Value = rs_Tab1_SelectedOrders.Fields("�h��").Value
                rs_SDN02W.Fields("�q�渹�X").Value = rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value
                rs_SDN02W.Fields("C_Receipt_No").Value = rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value
                rs_SDN02W.Update
            Else
                '��s������q����
                rs_SDN02W.Fields("�G���ƨ�").Value = rs_Tab1_SelectedOrders.Fields("�G���ƨ�").Value
                rs_SDN02W.Fields("�Ȥ�渹").Value = rs_Tab1_SelectedOrders.Fields("�Ȥ�渹").Value
                rs_SDN02W.Fields("���u�s��").Value = rs_Tab1_SelectedOrders.Fields("���u�s��").Value
                rs_SDN02W.Fields("���").Value = rs_Tab1_SelectedOrders.Fields("���").Value
                rs_SDN02W.Fields("���e�Ȥ�").Value = rs_Tab1_SelectedOrders.Fields("���e�Ȥ�").Value
                rs_SDN02W.Fields("�c��").Value = rs_Tab1_SelectedOrders.Fields("�c��").Value
                rs_SDN02W.Fields("���q").Value = rs_Tab1_SelectedOrders.Fields("���q").Value
                rs_SDN02W.Fields("���n").Value = rs_Tab1_SelectedOrders.Fields("���n").Value
                rs_SDN02W.Fields("�h��").Value = rs_Tab1_SelectedOrders.Fields("�h��").Value
                rs_SDN02W.Fields("�q�渹�X").Value = rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value
                rs_SDN02W.Fields("C_Receipt_No").Value = rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value
            End If
            rs_Tab1_SelectedOrders.MoveNext
        Loop
        
        '[�R���w������q��
        rs_Tab1_SelectedOrders.MoveFirst
        Do While Not rs_Tab1_SelectedOrders.EOF
            rs_Tab1_SelectedOrders.Delete
            rs_Tab1_SelectedOrders.MoveFirst
        Loop
    End If
    rs_SDN02W.Filter = adFilterNone
    rs_SDN02W.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    rs_Tab1_SelectedOrders.Filter = adFilterNone
    rs_Tab1_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    Call ReSet_Tab1_SelectedOrders_SeqNo
    Call ReSet_Tab1_SDN02W_SeqNo
    Tab1_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-Tab1����V�U", Me.Caption, "cmd_Tab1_Selected_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_SelectCar_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab1_SelectCar.Name & "2")
End Sub

Private Sub cmd_Tab1_Selected_Click()
    '�z��w�����
    If dg_Tab1_SelectedOrders.DataSource Is Nothing Then Exit Sub
    If dg_SDN02W.DataSource Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    Tab1_RouteListEventEnable = False
    rs_SDN02W.Filter = "��='V'"
    If Not rs_SDN02W.EOF Then
        Do While Not rs_SDN02W.EOF
            '�P�_�O�_�w�g����L
            rs_Tab1_SelectedOrders.Filter = adFilterNone
            rs_Tab1_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
            rs_Tab1_SelectedOrders.Filter = "�q�渹�X = '" & rs_SDN02W.Fields("�q�渹�X").Value & "'"
            '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
            If blRouteModify Then blRouteChange = True
            If rs_Tab1_SelectedOrders.EOF Then
                '�s�W������q��
                rs_Tab1_SelectedOrders.AddNew
                rs_Tab1_SelectedOrders.Fields("�s��").Value = 999
                rs_Tab1_SelectedOrders.Fields("�G���ƨ�").Value = rs_SDN02W.Fields("�G���ƨ�").Value
                rs_Tab1_SelectedOrders.Fields("�Ȥ�渹").Value = rs_SDN02W.Fields("�Ȥ�渹").Value
                rs_Tab1_SelectedOrders.Fields("���u�s��").Value = rs_SDN02W.Fields("���u�s��").Value
                rs_Tab1_SelectedOrders.Fields("���").Value = rs_SDN02W.Fields("���").Value
                rs_Tab1_SelectedOrders.Fields("���e�Ȥ�").Value = rs_SDN02W.Fields("���e�Ȥ�").Value
                rs_Tab1_SelectedOrders.Fields("�c��").Value = rs_SDN02W.Fields("�c��").Value
                rs_Tab1_SelectedOrders.Fields("���q").Value = rs_SDN02W.Fields("���q").Value
                rs_Tab1_SelectedOrders.Fields("���n").Value = rs_SDN02W.Fields("���n").Value
                rs_Tab1_SelectedOrders.Fields("�h��").Value = rs_SDN02W.Fields("�h��").Value
                rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value = rs_SDN02W.Fields("�q�渹�X").Value
                rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value = rs_SDN02W.Fields("C_Receipt_No").Value
                rs_Tab1_SelectedOrders.Update
            Else
                '��s������q����
                rs_Tab1_SelectedOrders.Fields("�G���ƨ�").Value = rs_SDN02W.Fields("�G���ƨ�").Value
                rs_Tab1_SelectedOrders.Fields("�Ȥ�渹").Value = rs_SDN02W.Fields("�Ȥ�渹").Value
                rs_Tab1_SelectedOrders.Fields("���u�s��").Value = rs_SDN02W.Fields("���u�s��").Value
                rs_Tab1_SelectedOrders.Fields("���").Value = rs_SDN02W.Fields("���").Value
                rs_Tab1_SelectedOrders.Fields("���e�Ȥ�").Value = rs_SDN02W.Fields("���e�Ȥ�").Value
                rs_Tab1_SelectedOrders.Fields("�c��").Value = rs_SDN02W.Fields("�c��").Value
                rs_Tab1_SelectedOrders.Fields("���q").Value = rs_SDN02W.Fields("���q").Value
                rs_Tab1_SelectedOrders.Fields("���n").Value = rs_SDN02W.Fields("���n").Value
                rs_Tab1_SelectedOrders.Fields("�h��").Value = rs_SDN02W.Fields("�h��").Value
                rs_Tab1_SelectedOrders.Fields("�q�渹�X").Value = rs_SDN02W.Fields("�q�渹�X").Value
                rs_Tab1_SelectedOrders.Fields("C_Receipt_No").Value = rs_SDN02W.Fields("C_Receipt_No").Value
            End If
            txt_Tab1_Selected_Case.Text = Val(txt_Tab1_Selected_Case.Text) + Val(rs_SDN02W.Fields("�c��").Value)
            txt_Tab1_Selected_Weight.Text = Val(txt_Tab1_Selected_Weight.Text) + Val(rs_SDN02W.Fields("���q").Value)
            txt_Tab1_Selected_Volumn.Text = Val(txt_Tab1_Selected_Volumn.Text) + Val(rs_SDN02W.Fields("���n").Value)
            rs_SDN02W.MoveNext
        Loop
       
       '[�R���w������q��
        rs_SDN02W.MoveFirst
        Do While Not rs_SDN02W.EOF
            rs_SDN02W.Delete
            rs_SDN02W.MoveFirst
        Loop
        txt_Tab1_srcSelected_Case.Text = 0
        txt_Tab1_srcSelected_Volumn.Text = 0
        txt_Tab1_srcSelected_Weight.Text = 0
    End If
    rs_SDN02W.Filter = adFilterNone
    rs_SDN02W.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    rs_Tab1_SelectedOrders.Filter = adFilterNone
    rs_Tab1_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@�w�n���o���Ƥ~�|���s���
    Call ReSet_Tab1_SelectedOrders_SeqNo
    Call ReSet_Tab1_SDN02W_SeqNo
    Tab1_RouteListEventEnable = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-Tab1����V�W", Me.Caption, "cmd_Tab1_Selected_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_CreateRoute_Click()
    If Len(Trim(txt_Tab2_Route.Text)) = 0 Then Exit Sub
    
    If Len(Trim(txt_Tab2_DELIVERY_DATE.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J�X�����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    If Len(Trim(txt_Tab2_VehicleNo.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J���P���X"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab1_VehicleNo.SetFocus
        Exit Sub
    End If
    '�X������G�榡 yyyymmdd
    txt_Tab2_DELIVERY_DATE.Text = Trim(txt_Tab2_DELIVERY_DATE.Text)
    If Fun_ChkDateFormat(txt_Tab2_DELIVERY_DATE.Text) = 1 Then
        msg_text = "�X������G" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab2_DELIVERY_DATE.SelStart = 0: txt_Tab2_DELIVERY_DATE.SelLength = Len(txt_Tab2_DELIVERY_DATE.Text): txt_Tab2_DELIVERY_DATE.SetFocus
        Exit Sub
    End If
    On Error GoTo err_Handle
    Call DB_CheckConnectStatus
    
'    '�ˮ֬O�_ñ��T�{
'    Call ReDim_Recordset(tmp_Rs)
'    str_SQL = "Select Count(*) as Receiver From SDN02T Where C_Route_No = '" & Trim(rs_Tab2_Route.Fields("�G���ƨ�").Value) & "'  and confirm_notes <> '' "
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.Fields("Receiver").Value > 0 Then
'        tmp_Rs.Close
'        msg_text = "�w������ñ�槹�����@,�L�k�ק�G"
'        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    tmp_Rs.Close
'
'    'Terry 20190320 �s�W���b �w���@�̪O�����s���i�ק郞��
'    Call ReDim_Recordset(tmp_Rs)
'    str_SQL = "select count(*) from pallet_cds where checkno = '" & Trim(rs_Tab2_Route.Fields("�G���ƨ�").Value) & "'"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.Fields(0).Value > 0 Then
'        tmp_Rs.Close
'        MsgBox ("�����s�w���@�̪O�A�L�k�ܧ󨮸�!")
'        Exit Sub
'    End If
'    tmp_Rs.Close

    'Terry 20190327 �s�W���b �w���p�O��Ƥ����s���i�ק郞��
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from sdn05t where c_route_no = '" & Trim(rs_Tab2_Route.Fields("�G���ƨ�").Value) & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("�����s�w�����@��ơA�L�k�ܧ󨮸�!")
        Exit Sub
    End If
    tmp_Rs.Close
    
    
    
    '�ˮ� [���P���X] �O�_����
    Call ReDim_Recordset(tmp_Rs)
    txt_Tab1_VehicleNo.Text = Trim(txt_Tab1_VehicleNo.Text)
    str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab2_VehicleNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "��ƿ��~�G���P���X " & txt_Tab1_VehicleNo.Text & " ������"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        txt_Tab1_VehicleNo.SelStart = 0: txt_Tab1_VehicleNo.SelLength = Len(txt_Tab1_VehicleNo.Text)
        txt_Tab1_VehicleNo.SetFocus
        Exit Sub
    End If
    tmp_Rs.Close
    
    DoEvents: DoEvents
    cn.BeginTrans
        '�s���Y,SDN01T
        str_SQL = "Update SDN01T set DELIVERY_DATE='" & txt_Tab2_DELIVERY_DATE.Text & "',C_VEHICLE_ID_NO='" & Trim(txt_Tab2_VehicleNo.Text) & "',Driver='" & Trim(txt_Tab2_Driver.Text) & "',edituser = '" & User_id & _
                "',editdate = getdate() where C_Route_No='" & Trim(txt_Tab2_Route.Text) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '��s�дڤH
        cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & Trim(txt_Tab2_VehicleNo.Text) & "') where c_route_no = '" & Trim(txt_Tab2_Route.Text) & "'", RowsAffect, adExecuteNoRecords
        
    cn.CommitTrans

    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-Tab2�T�{�s��", Me.Caption, "cmd_Tab1_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Excel_Click()
    '�ƨ��@���� >> �� EXCEL

    If rs_Tab2_Route Is Nothing Then Exit Sub
    rs_Tab2_Route.MoveFirst
    On Error GoTo err_Handle
    Tab2_RouteListEventEnable = False
    '�N��Ƽg�Jexcel��
    Dim MyXlsApp As Excel.Application   '�}��excel��
    Dim objFld As Field
    Dim i, j As Integer
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '�s�WWookbooks
    MyXlsApp.Workbooks.Add
    '�s�WSheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "�X���T�{�@����"
    MyXlsApp.ActiveSheet.Name = "�X���T�{�@����"
    
    i = 1
    'select convert(char,s1.DELIVERY_DATE,112) as ���,s1.C_VEHICLE_ID_NO as ����,s1.Driver as �q��,s1.C_Route_No as �G���ƨ�, " & _
            "sum(s2.SHIP_CBM) as ���n,sum(s2.SHIP_WT) as ���q,Max(Distinct s2.CUST_NAME) as �Ȥ�²��
    MyXlsApp.Cells(i, 1).Value = "�s��"
    MyXlsApp.Cells(i, 2).Value = "�X�����"
    MyXlsApp.Cells(i, 3).Value = "���P���X"
    MyXlsApp.Cells(i, 4).Value = "�r�p�H"
    MyXlsApp.Cells(i, 5).Value = "���u�s��"
    MyXlsApp.Cells(i, 6).Value = "�B�e���q"
    MyXlsApp.Cells(i, 7).Value = "�B�e���n"
    MyXlsApp.Cells(i, 8).Value = "�Ȥ�²��"
    i = i + 1
    rs_Tab2_Route.MoveFirst
    '���,����,�渹,�Z�O,�ɥX,�٤J
    Do While Not rs_Tab2_Route.EOF
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab2_Route.Fields(0))
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab2_Route.Fields(1))
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rs_Tab2_Route.Fields(2)
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rs_Tab2_Route.Fields(3)
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rs_Tab2_Route.Fields(4)
        MyXlsApp.Cells(i, 6).Value = rs_Tab2_Route.Fields(5)
        MyXlsApp.Cells(i, 7).Value = rs_Tab2_Route.Fields(6)
        MyXlsApp.Cells(i, 8).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 8).Value = rs_Tab2_Route.Fields(7)
        rs_Tab2_Route.MoveNext
        i = i + 1
    Loop
    i = i + 1
    '�̾A��e
    MyXlsApp.Columns("A:H").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w
    MyXlsApp.Columns("F:G").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:H" & i - 1).Select
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
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Tab2_RouteListEventEnable = True
    Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�X���T�{-�� EXCEL", Me.Caption, "cmd_Tab2_Excel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_RouteNoDelete_Click()

    'If blAdmin = False Then MsgBox "�t�κ޲z���~���v�����榹�@�~!", 64, "�v������": Exit Sub
    '���u�s���C�� >> ���u�s���R��
    If rs_Tab2_Route.RecordCount = 0 Then Exit Sub
    If dg_Tab2_Route.SelBookmarks.Count = 0 Then MsgBox "��������u�s��!", 64, "���u�s���R��": Exit Sub
    On Error GoTo err_Handle
    Tab2_RouteListEventEnable = False
    Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
    strDeleteRouteNo = Trim(rs_Tab2_Route.Fields("�G���ƨ�").Value)
    strCarno = Trim(rs_Tab2_Route.Fields("����").Value)
    'dbDriveTimes = Trim(rs_Tab2_Route.Fields("����").Value)
        
'    '�ˬd�R�������s�O�_���w�^�Ǫ��q��
    rs_Tab2_RouteOrders.MoveFirst
    Do While Not rs_Tab2_RouteOrders.EOF
        Call ReDim_Recordset(tmp_Rs)
        str_SQL = "Select returnstatus From SDN02t Where receipt_no = '" & Trim(rs_Tab2_RouteOrders.Fields("�q�渹�X").Value) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        tmp_Rs.MoveFirst
        Do While Not tmp_Rs.EOF
            If tmp_Rs.Fields("returnstatus").Value > 0 Then
                msg_text = "�q�渹�X:" & Trim(rs_Tab2_RouteOrders.Fields("�q�渹�X").Value) & " ��Ƥw�^�ǡA�L�k�i��R��!"
                MsgBox msg_text, vbOKOnly + vbCritical, msg_title
                Screen.MousePointer = vbDefault
                blTab1RouteEventEnable = True
                Tab2_RouteListEventEnable = True
                tmp_Rs.Close
                Exit Sub
            End If
            tmp_Rs.MoveNext
        Loop
        rs_Tab2_RouteOrders.MoveNext
    Loop
    rs_Tab2_RouteOrders.MoveFirst
    
'    msg_text = "�T�{�R�����u�s���G" & strDeleteRouteNo
'    If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub

    'Terry 20190311 �s�W���b �w���@�̪O�����s���i��������
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from pallet_cds where checkno = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("�����s�w���@�̪O�A�L�k�R��!")
        Exit Sub
    End If
    tmp_Rs.Close
    
    'Terry 20190327 �s�W���b �w���p�O��Ƥ����s���i��������
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from sdn05t where c_route_no = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("�����s�w���p�O��ơA�L�k�R��!")
        Exit Sub
    End If
    tmp_Rs.Close
    

    If MsgBox("�����s�Ҧ��q�檺�B�O�Pñ��T�{�N�@�֧R���A�q��N��J�������աA�O�_�~��?", vbOKCancel, Trim(strDeleteRouteNo) & "==>��������") <> vbOK Then blTab1RouteEventEnable = True: Tab2RouteListEventEnable = True: Tab2_RouteListEventEnable = True: Exit Sub
    
    '�R�����s
    Call Delete_RouteNo(strDeleteRouteNo)
    
    '�R���d�ߵ��G���ӵ����u�s��--rs_Tab1_RouteOrders
    rs_Tab2_RouteOrders.Filter = adFilterNone
    rs_Tab2_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_Tab2_RouteOrders.Filter = "�G���ƨ�='" & strDeleteRouteNo & "'"
    If Not rs_Tab2_RouteOrders.EOF Then
        Do While Not rs_Tab2_RouteOrders.EOF
            rs_Tab2_RouteOrders.Delete
            rs_Tab2_RouteOrders.MoveFirst
        Loop
    End If
    rs_Tab2_RouteOrders.Filter = adFilterNone
    rs_Tab2_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    '(7).�R���d�ߵ��G���ӵ����u�s��--rs_Tab1_Route
    rs_Tab2_Route.Delete
    If Not rs_Tab2_Route.EOF Then rs_Tab2_Route.MoveFirst
    
    blTab1RouteEventEnable = True
    Tab2_RouteListEventEnable = True
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�X���T�{-���u�s���R��", Me.Caption, "cmd_Tab2_RouteNoDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_RouteNoQuery_Click()
    '�X���T�{ >> Tab2���u�s���d��
    'If Len(Trim(txt_Tab2_Route_Start.Text)) = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    '���u�s��
    txt_Tab2_Route_Start.Text = Trim(txt_Tab2_Route_Start.Text)
    strSubwhere = ""
    If Len(txt_Tab2_Route_Start.Text) > 0 Then
        strSubwhere = " s1.C_Route_No = '" & txt_Tab2_Route_Start.Text & "' "
    End If
    If Len(strSubwhere) > 0 Then
        If Len(str_Where) = 0 Then
            str_Where = strSubwhere
        Else
            str_Where = str_Where & " and " & strSubwhere
        End If
    End If
    '�X�����
    txt_Tab2_DeliveryDate_Start.Text = Trim(txt_Tab2_DeliveryDate_Start.Text)
    strSubwhere = ""
    If Len(txt_Tab2_DeliveryDate_Start.Text) > 0 Then
        strSubwhere = " s1.DELIVERY_DATE = '" & txt_Tab2_DeliveryDate_Start.Text & "' "
    End If
    If Len(strSubwhere) > 0 Then
        If Len(str_Where) = 0 Then
            str_Where = strSubwhere
        Else
            str_Where = str_Where & " and " & strSubwhere
        End If
    End If
    '��str_SQL
    'str_SQL = "Select C_Route_No as �G���ƨ�, Convert(varchar,DELIVERY_DATE,112) as ���, C_VEHICLE_ID_NO as ����, Driver as �q�� From SDN01T"
    str_SQL = "select convert(char(8),s1.DELIVERY_DATE,112) as ���,s1.C_VEHICLE_ID_NO as ����,s1.Driver as �q��,s1.C_Route_No as �G���ƨ�, " & _
            "sum(s2.SHIP_CBM) as ���n,sum(s2.SHIP_WT) as ���q,Max(Distinct s2.CUST_NAME) as �Ȥ�²�� from SDN01T s1 " & _
            "inner join SDN02T s2 on s1.C_Route_No=s2.C_Route_No "
    If Len(str_Where) = 0 Then
        str_SQL = str_SQL & "group by s1.DELIVERY_DATE order by s1.C_Route_No"
    Else
        str_SQL = str_SQL & " where" & str_Where & " group by s1.DELIVERY_DATE,s1.C_Route_No,s1.C_VEHICLE_ID_NO,s1.Driver order by s1.C_Route_No"
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����(SDN01T)"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab2_Route)
    Set dg_Tab2_Route.DataSource = rs_Tab2_Route
    'tmp_rs.Close
    With dg_Tab2_Route
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000      '�G���ƨ�
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '���
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1000      '����
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1500       '�q��
        .Columns(4).Alignment = dbgLeft
    End With
    rs_Tab2_Route.MoveFirst
    txt_Tab2_Route.Text = rs_Tab2_Route.Fields("�G���ƨ�").Value
    txt_Tab2_VehicleNo.Text = rs_Tab2_Route.Fields("����").Value
    txt_Tab2_Driver.Text = rs_Tab2_Route.Fields("�q��").Value
    txt_Tab2_DELIVERY_DATE.Text = rs_Tab2_Route.Fields("���").Value
    'SDN03T
    Call Display_Tab2_RouteOrders
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�X���T�{ >> Tab2���u�s���d��", Me.Caption, "cmd_Tab2_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_SelectCar_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab2_SelectCar.Name & "2")
End Sub

Private Sub cmd_Tab3_ClearQty_Click()
    '�ݤ��έq�� >> �M�� [�O�Ƥ���][�c�Ƥ���] ����
    txt_Tab3_CutCaseQty.Text = ""
    txt_Tab3_CutCaseQty.SetFocus
    'RUN Button [�ƶq����] Click
    Call cmd_Tab3_CutQty_Click
End Sub

Private Sub cmd_Tab3_CutOrders_Click()
    '�X���T�{ >> ���έq��
    If rs_Tab3_SDN02W Is Nothing Then Exit Sub
    If rs_Tab3_SDN02W.RecordCount = 0 Then Exit Sub
    
    Dim intTRP02WBookMark As String     '���b�i�� [�q����Χ@�~] ���q���ƦC
    Dim strCutOrder_SrcKey As String    '���b�i�� [�q����Χ@�~] ���q��s��
    Dim dbMaxKey As Double              '�s�q��s���G���X key
    Dim strCutOrder_NewKey As String    '�s���ΥX�Ӥ��q��� [�q��s��]
    Dim i As Double
    
    On Error GoTo err_Handle
    If Len(Trim(CutOrderkey)) = 0 Then Exit Sub
    
    '�ˬd�O���I������Τ��q��Ӷ�
    dg_Tab3_SelectedOrderDetail.Visible = False
    Dim dbCount As Double
    dbCount = 0
    With dg_Tab3_SelectedOrderDetail
        For i = 1 To .Rows - 2
            .Row = i: .Col = 1
            If Len(Trim(.Text)) <> 0 Then
                dbCount = dbCount + 1
            End If
        Next i
    End With
    dg_Tab3_SelectedOrderDetail.Visible = True
    If dbCount = 0 Then
        msg_text = "��ƿ��~�G����������Τ��q���"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    '��Ʈw���ʥ��--�_�I
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '���s���ΥX�Ӥ��q��M�w�� [�q��s��]
    strCutOrder_SrcKey = CutOrderkey
    str_SQL = "Select Cast(Code as integer) as AvailNo From CodeLKUP Where ListName = 'CUTORDERSNO'  "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        strCutOrder_NewKey = "CT" & Format(1, "00000000")
        str_SQL = "Insert into CodeLKUP (ListName,Code,Description,AddWho,EditWho) Values ('CUTORDERSNO',2,'����h�����s���ͭq�渹�X','" & User_id & "','" & User_id & "')"
    Else
        strCutOrder_NewKey = "CT" & Format(tmp_Rs.Fields("AvailNo").Value, "00000000")
        str_SQL = "Update CodeLKUP Set Code = " & (tmp_Rs.Fields("AvailNo").Value + 1) & " Where ListName = 'CUTORDERSNO'"
    End If
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    tmp_Rs.Close
    
    '���_��l�q��C�� DBGrid �� Event ����
    'blTRP02WEventEnable = False
    rs_Tab3_SDN02W.Filter = adFilterNone
    rs_Tab3_SDN02W.Filter = "�q�渹�X = '" & CutOrderkey & "'"
    If rs_Tab3_SDN02W.RecordCount = 0 Then
        msg_text = "��p���A�䤣��ŦX���󪺭�q���Ƴ�"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        rs_Tab3_SDN02W.Filter = adFilterNone
        rs_Tab3_SDN02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
        Exit Sub
    Else
       '���ͤ��έq��-Header
        intTRP02WBookMark = rs_Tab3_SDN02W.Bookmark
        txt_Tab3_OrderKey.Text = strCutOrder_NewKey     '�q��s��
        txt_Tab3_DeliveryDate.Text = rs_Tab3_SDN02W.Fields("���").Value    '�e�f��
        txt_Tab3_Extern.Text = rs_Tab3_SDN02W.Fields("�Ȥ�渹").Value  '�Ȥ�s��
        txt_Tab3_CaseQty.Text = txt_Tab3_SelectedCaseQty.Text '�c��
        txt_Tab3_Weight.Text = txt_Tab3_SelectedWeight.Text    '���q
        txt_Tab3_Volumn.Text = txt_Tab3_SelectedVolumn.Text    '���n
        txt_Tab3_FullName.Text = rs_Tab3_SDN02W.Fields("���e�Ȥ�").Value    '�Ȥ�W��
               
        '���ͷs���q����--SDN02W
        str_SQL = "Insert into SDN02W (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,c_RECEIPT_NO) " & _
                  "Select C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME, " & _
                  "" & txt_Tab3_SelectedCaseQty.Text & "," & txt_Tab3_SelectedVolumn.Text & "," & txt_Tab3_SelectedWeight.Text & ",CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,'" & strCutOrder_NewKey & "',c_RECEIPT_NO " & _
                  "From SDN02W Where  Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '��s��q�椧�έp�Ʀr--SDN02W
        str_SQL = "Update SDN02W Set SHIP_CS=SHIP_CS-" & txt_Tab3_SelectedCaseQty.Text & "," & _
                  "SHIP_WT=SHIP_WT-" & txt_Tab3_SelectedWeight.Text & ",SHIP_CBM=SHIP_CBM-" & txt_Tab3_SelectedVolumn.Text & " " & _
                  "Where Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    End If
    rs_Tab3_SDN02W.Filter = adFilterNone
    rs_Tab3_SDN02W.Sort = "�q�渹�X ASC"
    blTRP02WEventEnable = True

    '���έq�椧 OrderDetail
    Dim dbsrcQty As Double, dbCutQty As Double, dbSeqNo As String
    dbSeqNo = 0
    dg_Tab3_SelectedOrderDetail.Visible = False
    With dg_Tab3_SelectedOrderDetail
         For i = 1 To .Rows - 2
             .Row = i: .Col = 1
             If .Text <> "" Then   '�Ӷ��Q����i�����
                .Col = 0: dbSeqNo = .Text          '�O�d��w�涵���s���w������
                .Col = 4: dbsrcQty = Val(.Text)    '��q��c��
                .Col = 7: dbCutQty = Val(.Text)    '���νc��
                If dbsrcQty = dbCutQty Then        '�Y�������c�ƶi����ΡA���O�ǳƫ���R�����Ӷ�
                    .Col = 1: .Text = "X"
                    '������sSDN03W���q�渹�X
                    str_SQL = "Update SDN03W Set Receipt_No = '" & strCutOrder_NewKey & "' " & _
                              "Where Receipt_No = '" & CutOrderkey & "' and SEQ_NO = '" & dbSeqNo & "' "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                Else
                   '��s �ݤ��έq�����
                   .Col = 1: .Text = ""        '�M�����O�A��s����
                   .Col = 4: dbsrcQty = Val(.Text)    '��q��c��
                   .Col = 7: dbCutQty = Val(.Text)    '���νc��
                   .Col = 4: .Text = dbsrcQty - dbCutQty
                   .Col = 5: dbsrcQty = Val(.Text)    '��q�歫�q
                   .Col = 8: dbCutQty = Val(.Text)    '���έ��q
                   .Col = 5: .Text = dbsrcQty - dbCutQty
                   .Col = 6: dbsrcQty = Val(.Text)    '��q����n
                   .Col = 9: dbCutQty = Val(.Text)   '���Χ��n
                   .Col = 6: .Text = dbsrcQty - dbCutQty
                   
                   '��sSDN03W��ƶq
                   .Col = 7
                   str_SQL = "Update SDN03W Set SDN03W.SHIP_QTY = " & _
                   "SDN03W.SHIP_QTY - (" & .Text & " * sp.casecnt) ,SDN03W.ORDER_QTY= SDN03W.ORDER_QTY - ( " & .Text & " * sp.casecnt) " & _
                   "from sdn03w SDN03W join gv_skuxpack sp on sp.storerkey = SDN03W.storerkey and sp.sku = SDN03W.product_no Where SDN03W.Receipt_No = '" & CutOrderkey & "' and SDN03W.SEQ_NO = '" & dbSeqNo & "' "
                   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   
'                '�N�c�ƴ���^�Ӽ� by gemini 20071212
'               str_SQL = "Update SDN03W Set SDN03W.Order_Qty = SDN03W.Order_Qty * s1.casecnt,SDN03W.SHIP_QTY = SDN03W.SHIP_QTY * s1.casecnt " & _
'                        "from SDN03W SDN03W join sku s on SDN03W.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = SDN03W.storerkey " & _
'                        "Where Receipt_No = '" & CutOrderkey & "' and SEQ_NO = '" & dbSeqNo & "' "
'                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   
                   '�s�W�s�q�椧�q��Ӷ�
                   str_SQL = "Insert into SDN03W (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME) " & _
                             "Select s3.C_ROUTE_NO,s3.ROUTE_NO,s3.StorerKey,'" & strCutOrder_NewKey & "',s3.Seq_No,s3.SubSeq_No,s3.EXTERN,s3.Product_No,s3.Ship_Unit,"
                   .Col = 7: str_SQL = str_SQL & .Text & " * sp.casecnt,"
                   str_SQL = str_SQL & "s3.SIGN_QTY,s3.RSC_CODE,s3.RBC_CODE,s3.CONFIRM_DATE,s3.DESCRIPTION,"
                   .Col = 7: str_SQL = str_SQL & .Text & " * sp.casecnt,"
                   str_SQL = str_SQL & "s3.SHIP_TIME From SDN03W s3 join gv_skuxpack sp on sp.sku = s3.product_no and s3.storerkey = sp.storerkey Where s3.Receipt_No = '" & CutOrderkey & "' and s3.SEQ_NO = '" & dbSeqNo & "' "
                   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   
'                '�N�c�ƴ���^�Ӽ� by gemini 20071212
'               str_SQL = "Update SDN03W Set SDN03W.Order_Qty = SDN03W.Order_Qty * s1.casecnt ,SDN03W.SHIP_QTY = SDN03W.SHIP_QTY * s1.casecnt " & _
'                        "from SDN03W SDN03W join sku s on SDN03W.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = SDN03W.storerkey " & _
'                        "Where Receipt_No = '" & strCutOrder_NewKey & "' and SEQ_NO = '" & dbSeqNo & "' "
'                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                   .Col = 7: .Text = ""   '���νc��
                   .Col = 8: .Text = ""   '���έ��q
                   .Col = 9: .Text = ""  '���Χ��n
                End If
             End If
         Next i
    End With

    '�R���w���ƶq���Τ��q��Ӷ�
    Dim j As Double
    With dg_Tab3_SelectedOrderDetail
        For i = 1 To .Rows - 2
            For j = 1 To .Rows - 2
                .Row = j: .Col = 1
                If .Text = "X" Then
                    Call Delete_GridRow(j)
                    Exit For
                End If
            Next j
        Next i
        '���s���ͭq��[�`�έp���
        txt_Tab3_Weight.Text = 0
        txt_Tab3_Volumn.Text = 0
        For i = 1 To .Rows - 2
            .Row = i
            .Col = 5: txt_Tab3_Weight.Text = Val(txt_Tab3_Weight.Text) + Val(.Text)
            .Col = 6: txt_Tab3_Volumn.Text = Val(txt_Tab3_Volumn.Text) + Val(.Text)
        Next i
    End With
    dg_Tab3_SelectedOrderDetail.Visible = True
    Call Dispaly_dg_Tab3_SDN03W
    
    '�M������-��������έp
    txt_Tab3_SelectedCaseQty.Text = ""
    dbCut_TotalCaseQty = 0
    txt_Tab3_SelectedWeight.Text = ""
    dbCut_TotalWeight = 0
    txt_Tab3_SelectedVolumn.Text = ""
    dbCut_TotalVolumn = 0
    
    '�Ӷ����μƶq���G�O�ơA�c��
    txt_Tab3_CutCaseQty.Text = ""
    If dg_Tab3_SelectedOrderDetail.Rows = 2 And txt_Tab3_Weight.Text = 0 And txt_Tab3_Volumn.Text = 0 Then
        '�w�������Τ��q��G�R�� TRP02W & TRP03W
        str_SQL = "Delete From SDN02W Where Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Delete From SDN03W Where Receipt_No = '" & CutOrderkey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        DoEvents
        SSTab1.Tab = 0
        DoEvents
    End If
    cn.CommitTrans
    Tran_Level = 0
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then Tran_Level = 0: cn.RollbackTrans
   
   dg_Tab3_SelectedOrderDetail.Visible = True
   blTRP02WEventEnable = True
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�X���T�{ >> ���έq��", Me.Caption, "cmd_Tab3_CutOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_CutQty_Click()
    '�X���T�{ >> �ƶq����
    If rs_Tab3_SDN02W Is Nothing Then Exit Sub
    If rs_Tab3_SDN02W.RecordCount = 0 Then Exit Sub
    
    cmd_Tab3_CutOrders.Enabled = False
    cmd_Tab3_ClearQty.Enabled = False
    If Val(txt_Tab3_CutCaseQty.Text) = 0 Then
        '��ƶq�M����ܡG���������
        dg_Tab3_SelectedOrderDetail.Col = 1: dg_Tab3_SelectedOrderDetail.Text = ""   '�������
    End If
    Dim tmpQty As Double
    If Val(txt_Tab3_CutCaseQty.Text) > 0 Then
      
       dg_Tab3_SelectedOrderDetail.Col = 4: tmpQty = Val(dg_Tab3_SelectedOrderDetail.Text)
       If Val(txt_Tab3_CutCaseQty.Text) > tmpQty Then
            msg_text = "��ƿ��~�G���νc�� �j�� �~���`�c��"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            cmd_Tab3_ClearQty.Enabled = True
            cmd_Tab3_CutOrders.Enabled = True
            Exit Sub
       End If
    
       '��J���νc�ơG�c��
       dg_Tab3_SelectedOrderDetail.Col = 9
       dg_Tab3_SelectedOrderDetail.Text = ""
       dg_Tab3_SelectedOrderDetail.Col = 7
       dg_Tab3_SelectedOrderDetail.Text = txt_Tab3_CutCaseQty.Text
    End If
    '�p�������q��Ӷ����[�` [�c��] [���q] [�~�n] [�O��]
    Call Calculate_Tab3_SelectedPrderDetail
    
    '�M�����ζq����
    txt_Tab3_CutCaseQty.Text = ""
    
    cmd_Tab3_ClearQty.Enabled = True
    cmd_Tab3_CutOrders.Enabled = True
End Sub

Private Sub cmd_Tab3_DelOrders_Click()
    '�q����Ω��� >> �R��
    Dim dbDeleteRow As Double, strOrderkey As String, strStorerkey As String, strExtern As String
    strOrderkey = Trim(txt_Tab3_OrderKey.Text)      '�q��s�� Receipt_No
    strExtern = Trim(txt_Tab3_Extern.Text)        '�f�D�渹 Extern
    
    msg_text = "�R���@�~�G�T�{�R��������l�q��G" & strOrderkey
    If MsgBox(msg_text, vbOKCancel + vbInformation, msg_title) = vbCancel Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    
    '�ˮֱ��R�����q��G�H�f�D�渹���d�߱���
    str_SQL = "Select Count(*) as RecCount From SDN02W Where Extern = '" & strExtern & "' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields("RecCount").Value = 1 Then
        tmp_Rs.Close
        msg_text = "�q��s���G" & strOrderkey & " �����\�R���A�]��f�D�渹�u���������q����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf tmp_Rs.Fields("RecCount").Value = 0 Then
        tmp_Rs.Close
        msg_text = "�q��s���G" & strOrderkey & " �w���s�b�A�Э��s����d��"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    tmp_Rs.Close
    
    '�����̤p�q��s���G�����Q�R���q��Ҧ������ءB�ƶq
    Dim strToOrderKey As String
    str_SQL = "Select Min(Receipt_No) as �����q��s�� From SDN02W Where Extern = '" & strExtern & "'  and Receipt_No <> '" & strOrderkey & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        strToOrderKey = tmp_Rs.Fields("�����q��s��").Value
    Else
        tmp_Rs.Close
        msg_text = "�f�D�渹�䤣��������q��s���i�H�������R�����q�涵��"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    tmp_Rs.Close
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    '��s�����q�椧������� TRP02W
    str_SQL = "Update SDN02W Set SHIP_CS=SHIP_CS+" & Val(txt_Tab3_CaseQty.Text) & ",SHIP_CBM=SHIP_CBM+ " & Val(txt_Tab3_Volumn.Text) & ", " & _
           "SHIP_WT=SHIP_WT+" & Val(txt_Tab3_Weight.Text) & "Where  Receipt_No = '" & strToOrderKey & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    
    '��s�����q�椧������� TRP03W
    
    Do While Not rs_Tab3_SDN03W.EOF
        '���ݱ����q��s�����L�ۦP�����B�f�����q��Ӷ� SDN03W
        str_SQL = "Select Count(*) AS RecCount From SDN03W " & _
                  "Where  Receipt_No = '" & strToOrderKey & "' and " & _
                  "      Seq_No = " & Trim(rs_Tab3_SDN03W.Fields("����").Value) & " and Product_No = '" & Trim(rs_Tab3_SDN03W.Fields("�f��").Value) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.Fields("RecCount").Value = 0 Then
           '�s�W�Ӷ� SDN03W
           str_SQL = "Insert into SDN03W (C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME) " & _
                     "Select C_ROUTE_NO,ROUTE_NO,StorerKey,'" & strToOrderKey & "',Seq_No,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME " & _
                     "From SDN03W Where  Receipt_No = '" & strOrderkey & "' and " & _
                     "      Seq_No = " & Trim(rs_Tab3_SDN03W.Fields("����").Value) & " and Product_No = '" & Trim(rs_Tab3_SDN03W.Fields("�f��").Value) & "'"
        Else
           '��s�Ӷ� TRP03W
           str_SQL = "Update SDN03W Set Order_Qty = Order_Qty + " & Trim(rs_Tab3_SDN03W.Fields("�q��c��").Value) & ",SHIP_QTY=SHIP_QTY+" & Trim(rs_Tab3_SDN03W.Fields("�z�f�c��").Value) & " " & _
                     "Where  Receipt_No = '" & strToOrderKey & "' and " & _
                     "      Seq_No = " & Trim(rs_Tab3_SDN03W.Fields("����").Value) & " and Product_No = '" & Trim(rs_Tab3_SDN03W.Fields("�f��").Value) & "'"
        End If
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        tmp_Rs.Close
        rs_Tab3_SDN03W.MoveNext
    Loop
    '�R���Ӷ�
    rs_Tab3_SDN03W.MoveFirst
    Do While Not rs_Tab3_SDN03W.EOF
        rs_Tab3_SDN03W.Delete
        rs_Tab3_SDN03W.MoveFirst
    Loop
    str_SQL = "Delete From SDN03W Where  Receipt_No = '" & strOrderkey & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rs_Tab3_SDN03W.Filter = adFilterNone
    rs_Tab3_SDN03W.Sort = "���� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj

    '�R���q��D�� TRP02W
    str_SQL = "Delete From SDN02W Where  Receipt_No = '" & strOrderkey & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    cn.CommitTrans
    Tran_Level = 0
    txt_Tab3_OrderKey.Text = ""
    txt_Tab3_DeliveryDate.Text = ""
    txt_Tab3_Extern.Text = ""  '�Ȥ�s��
    txt_Tab3_CaseQty.Text = "" '�c��
    txt_Tab3_Weight.Text = ""    '���q
    txt_Tab3_Volumn.Text = ""  '���n
    txt_Tab3_FullName.Text = ""   '�Ȥ�W��
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�X���T�{-Tab3�q��R��", Me.Caption, "cmd_Tab3_DelOrders", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_DisplayOrders_Click()
    '�X���T�{ >> ��ܫݱƨ��q��
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab3_SDN02W.DataSource = Nothing
    Set rs_Tab3_SDN02W = Nothing
    On Error GoTo err_Handle
    str_SQL = "SELECT  C_ROUTE_NO AS �G���ƨ�, ROUTE_NO AS ���u�s��,EXTERN AS �Ȥ�渹,ARRIVE_DATE AS ���,CUST_NAME as ���e�Ȥ�, " & _
            "SHIP_CS As �c��, SHIP_CBM As ���n, SHIP_WT As ���q, RECEIPT_NO As �q�渹�X, CAR_NOTES As �h�� " & _
            "FROM dbo.SDN02W Order by �G���ƨ�,���u�s��,�Ȥ�渹,�c�� "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        Screen.MousePointer = vbDefault
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_SDN02W)
    tmp_Rs.Close
    rs_Tab3_SDN02W.MoveFirst
    Set dg_Tab3_SDN02W.DataSource = rs_Tab3_SDN02W
    With dg_Tab3_SDN02W
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000      '�G���ƨ�
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '�Ȥ�渹
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800      '���u�s��
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800       '���
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1200      '���e�Ȥ�
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 800       '�c��
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 800       '���n
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '���q
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000       '�q�渹�X
        .Columns(9).Alignment = dbgLeft
        .Columns(10).Width = 800       '�h��
        .Columns(10).Alignment = dbgLeft
    End With
    
'    '�M����
    Call SetGrid_Format_Tab3_SelectedOrderDetail
    Call Clear_CutOrderDetail
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��C��-��ܫݱƨ��q��", Me.Caption, "cmd_Tab0_DisplayOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_DisplaySelectedOrder_Click()
    '�q��C�� >> ��ܭq�����
    If rs_Tab3_SDN02W Is Nothing Then Exit Sub
    If rs_Tab3_SDN02W.RecordCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    '�]�w�w�������έq����Ӫ�
    Call Clear_CutOrderDetail
    On Error GoTo err_Handle
    DoEvents: DoEvents
    '�]�w�����έq�椧�q��W��
    CutOrderkey = Trim(rs_Tab3_SDN02W.Fields("�q�渹�X").Value)
    Call SetGrid_Format_Tab3_SelectedOrderDetail
    
    str_SQL = "Select rtrim(SEQ_NO) as ����,rtrim(PRODUCT_NO) as �f��,rtrim(sp.Descr) as �~�W,case when sp.casecnt = 0 then 0 else isnull (SHIP_QTY/sp.casecnt,0) end as �c��,(isnull(SHIP_QTY,0)*sp.Stdgrosswgt) as ���q,(isnull(SHIP_QTY,0)*sp.STDCUBE) as ���n " & _
            "from SDN03W inner join gv_skuxpack sp on sp.sku=PRODUCT_NO and sp.storerkey = sdn03w.storerkey " & _
            "where RECEIPT_NO='" & CutOrderkey & "' order by SEQ_NO"
            
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       Screen.MousePointer = vbDefault
       msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����Ӹ��"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Exit Sub
    End If
    Do While Not tmp_Rs.EOF
        With dg_Tab3_SelectedOrderDetail
             .Rows = .Rows + 1
             .Row = .Rows - 2
             .Col = 0    '�q����Ӷ���
             .Text = tmp_Rs.Fields("����").Value
             .Col = 2    '�f��
             .Text = tmp_Rs.Fields("�f��").Value
             .Col = 3    '�~�W
             .Text = tmp_Rs.Fields("�~�W").Value
             .Col = 4    '�c��
             .Text = tmp_Rs.Fields("�c��").Value
             .Col = 5    '���q
             .Text = tmp_Rs.Fields("���q").Value
             .Col = 6    '���n
             .Text = tmp_Rs.Fields("���n").Value
        End With
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    Set tmp_Rs = Nothing
    Screen.MousePointer = vbDefault
    dbCut_TotalCaseQty = 0
    dbCut_TotalWeight = 0
    dbCut_TotalVolumn = 0
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��C��-��ܭq��W��", Me.Caption, "cmd_Tab0_DisplaySelectedOrder_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   If Not (tmp_Rs Is Nothing) Then
      Set tmp_Rs = Nothing
   End If
End Sub

Private Sub cmd_Tab3_Query_Click()
    If Len(Trim(txt_Tab3_OrderKey.Text)) = 0 Then Exit Sub
    On Error GoTo err_Handle
    str_SQL = "SELECT  EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,RECEIPT_NO " & _
            "from SDN02W where  RECEIPT_NO= '" & Trim(txt_Tab3_OrderKey.Text) & "' "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        Screen.MousePointer = vbDefault
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    txt_Tab3_DeliveryDate.Text = tmp_Rs.Fields("ARRIVE_DATE").Value    '�e�f��
    txt_Tab3_Extern.Text = tmp_Rs.Fields("EXTERN").Value  '�Ȥ�s��
    txt_Tab3_CaseQty.Text = tmp_Rs.Fields("SHIP_CS").Value '�c��
    txt_Tab3_Weight.Text = tmp_Rs.Fields("SHIP_WT").Value   '���q
    txt_Tab3_Volumn.Text = tmp_Rs.Fields("SHIP_CBM").Value    '���n
    txt_Tab3_FullName.Text = tmp_Rs.Fields("CUST_NAME").Value    '�Ȥ�W��
    tmp_Rs.Close
    Call Dispaly_dg_Tab3_SDN03W
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{>>Tab3�d�߭q��", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab2_SelectCar.Name & "2")
End Sub

Private Sub dg_SDN02W_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_SDN02W.DataSource Is Nothing Then Exit Sub
    If Tab1_RouteListEventEnable = False Then Exit Sub
    '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
    If Trim(rs_SDN02W.Fields(1).Value) = "" Then
        rs_SDN02W.Fields(1).Value = "V"
        txt_Tab1_srcSelected_Case.Text = Val(txt_Tab1_srcSelected_Case.Text) + Val(rs_SDN02W.Fields("�c��").Value)
        txt_Tab1_srcSelected_Volumn.Text = Val(txt_Tab1_srcSelected_Volumn.Text) + Val(rs_SDN02W.Fields("���n").Value)
        txt_Tab1_srcSelected_Weight.Text = Val(txt_Tab1_srcSelected_Weight.Text) + Val(rs_SDN02W.Fields("���q").Value)
    Else
        rs_SDN02W.Fields(1).Value = " "
        txt_Tab1_srcSelected_Case.Text = Val(txt_Tab1_srcSelected_Case.Text) - Val(rs_SDN02W.Fields("�c��").Value)
        txt_Tab1_srcSelected_Volumn.Text = Val(txt_Tab1_srcSelected_Volumn.Text) - Val(rs_SDN02W.Fields("���n").Value)
        txt_Tab1_srcSelected_Weight.Text = Val(txt_Tab1_srcSelected_Weight.Text) - Val(rs_SDN02W.Fields("���q").Value)
    End If
End Sub

Private Sub dg_Tab0_C_RouteList_HeadClick(ByVal ColIndex As Integer)

If dg_Tab0_C_RouteList.Row = -1 Then Exit Sub

Tab0_RouteListEventEnable = False

If intColumnIndex = ColIndex Then
    rs_Tab0_C_RouteList.Sort = dg_Tab0_C_RouteList.Columns(ColIndex).Caption & " DESC"
    dg_Tab0_C_RouteList.ClearSelCols
    intColumnIndex = 255

Else
    rs_Tab0_C_RouteList.Sort = dg_Tab0_C_RouteList.Columns(ColIndex).Caption
    dg_Tab0_C_RouteList.ClearSelCols
    intColumnIndex = ColIndex

End If

Tab0_RouteListEventEnable = True

End Sub

Private Sub dg_Tab0_C_RouteList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab0_C_RouteList.DataSource Is Nothing Then Exit Sub
    If Tab0_RouteListEventEnable = False Then Exit Sub
    
    '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
    
    If Trim(rs_Tab0_C_RouteList.Fields(1).Value) = "" Then
        rs_Tab0_C_RouteList.Fields(1).Value = "V"
    Else
        rs_Tab0_C_RouteList.Fields(1).Value = " "
    End If
End Sub

Private Sub dg_Tab0_RouteList0_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab0_RouteList0.DataSource Is Nothing Then Exit Sub
    If Tab0_RouteListEventEnable = False Then Exit Sub
    '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
    If Trim(rs_Tab0_RouteList0.Fields(1).Value) = "" Then
        rs_Tab0_RouteList0.Fields(1).Value = "V"
    Else
        rs_Tab0_RouteList0.Fields(1).Value = " "
    End If
End Sub

Private Sub dg_Tab0_RouteList1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab0_RouteList1.DataSource Is Nothing Then Exit Sub
    If Tab0_RouteListEventEnable = False Then Exit Sub
    '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
    If Trim(rs_Tab0_RouteList1.Fields(1).Value) = "" Then
        rs_Tab0_RouteList1.Fields(1).Value = "V"
    Else
        rs_Tab0_RouteList1.Fields(1).Value = " "
    End If
End Sub

Private Sub dg_Tab1_SelectedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dg_Tab1_SelectedOrders.DataSource Is Nothing Then Exit Sub
    If rs_Tab1_SelectedOrders.RecordCount = 0 Then Exit Sub
    If Tab1_RouteListEventEnable = False Then Exit Sub
    '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
    If Trim(rs_Tab1_SelectedOrders.Fields(1).Value) = "" Then
        rs_Tab1_SelectedOrders.Fields(1).Value = "V"
    Else
        rs_Tab1_SelectedOrders.Fields(1).Value = " "
       
    End If
End Sub


Private Sub dg_Tab2_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Tab2_RouteListEventEnable = False Then Exit Sub
    txt_Tab2_Route.Text = rs_Tab2_Route.Fields("�G���ƨ�").Value
    txt_Tab2_VehicleNo.Text = rs_Tab2_Route.Fields("����").Value
    txt_Tab2_Driver.Text = rs_Tab2_Route.Fields("�q��").Value
    txt_Tab2_DELIVERY_DATE.Text = rs_Tab2_Route.Fields("���").Value
    Call Display_Tab2_RouteOrders
End Sub

Private Sub dg_Tab3_SelectedOrderDetail_Click()
    '�ݤ��Τ��q��G�q����Ӷ���
    '�I�@���G����A���D�M�� [���μƶq] �_�h�����O�� [���] ���A
    txt_Tab3_CutCaseQty.Text = ""
    Dim i As Integer
    Dim tmpQty As Double
    With dg_Tab3_SelectedOrderDetail
        .Col = 2   '�f��
        If Len(Trim(.Text)) = 0 Then Exit Sub
        .Col = 1
        If Len(.Text) = 0 Then
            .Text = "V"
            .Col = 4   '��ܩҿ�����c��
            tmpQty = .Text
            dbCut_TotalCaseQty = dbCut_TotalCaseQty + .Text
            txt_Tab3_SelectedCaseQty.Text = dbCut_TotalCaseQty
            .Col = 7: .Text = tmpQty
            txt_Tab3_CutCaseQty.Text = tmpQty
            
            .Col = 5   '��ܩҿ�������q
            tmpQty = .Text
            dbCut_TotalWeight = dbCut_TotalWeight + .Text
            txt_Tab3_SelectedWeight.Text = dbCut_TotalWeight
            .Col = 8: .Text = tmpQty
            
            .Col = 6   '��ܩҿ�������n
            tmpQty = .Text
            dbCut_TotalVolumn = dbCut_TotalVolumn + .Text
            txt_Tab3_SelectedVolumn.Text = dbCut_TotalVolumn
            .Col = 9: .Text = tmpQty
        Else
            .Col = 7   '���Τ��c��
            If Val(.Text) <> 0 Then
                txt_Tab3_CutCaseQty.Text = .Text
            End If
        End If
        '�ϥտ������Ʀ�
        .Col = 0
        For i = 0 To .Cols - 1
            .ColSel = i
        Next i
    End With
End Sub

Private Sub Form_Load()
    '�]�w Form �j�p�B��m
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 11475
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    SSTab1.Tab = 0
    Tab2_RouteListEventEnable = True
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    If Me.ScaleHeight < dbsrcFormHeight Then
        '�ܤp
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

Private Sub ReSet_Tab0_RouteList1_SeqNo()
    '���s���� [dg_Tab0_RouteList0] �� [�s��] ����
    dg_Tab0_RouteList1.Visible = False
    rs_Tab0_RouteList1.Filter = adFilterNone
    rs_Tab0_RouteList1.Sort = "�Ȥ�渹 asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab0_RouteList1.EOF Then rs_Tab0_RouteList1.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab0_RouteList1.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_RouteList1.Fields("�s��").Value = dbSeqNo
        rs_Tab0_RouteList1.MoveNext
    Loop
    If rs_Tab0_RouteList1.RecordCount > 0 Then rs_Tab0_RouteList1.MoveFirst
    dg_Tab0_RouteList1.Visible = True
End Sub

Private Sub ReSet_Tab0_RouteList0_SeqNo()
    '���s���� [dg_Tab0_RouteList1] �� [�s��] ����
    dg_Tab0_RouteList0.Visible = False
    rs_Tab0_RouteList0.Filter = adFilterNone
    rs_Tab0_RouteList0.Sort = "�Ȥ�渹 asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab0_RouteList0.EOF Then rs_Tab0_RouteList0.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab0_RouteList0.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_RouteList0.Fields("�s��").Value = dbSeqNo
        rs_Tab0_RouteList0.MoveNext
    Loop
    If rs_Tab0_RouteList0.RecordCount > 0 Then rs_Tab0_RouteList0.MoveFirst
    dg_Tab0_RouteList0.Visible = True
End Sub

Private Sub ReSet_Tab0_C_RouteList_SeqNo()
    '���s���� [dg_Tab0_RouteList1] �� [�s��] ����
    dg_Tab0_C_RouteList.Visible = False
    rs_Tab0_C_RouteList.Filter = adFilterNone
    rs_Tab0_C_RouteList.Sort = "���u�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab0_C_RouteList.EOF Then rs_Tab0_C_RouteList.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab0_C_RouteList.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_C_RouteList.Fields("�s��").Value = dbSeqNo
        rs_Tab0_C_RouteList.MoveNext
    Loop
    If rs_Tab0_C_RouteList.RecordCount > 0 Then rs_Tab0_C_RouteList.MoveFirst
    dg_Tab0_C_RouteList.Visible = True
End Sub

Private Sub ReSet_Tab1_SelectedOrders_SeqNo()
    '���s���� [dg_Tab0_RouteList1] �� [�s��] ����
    dg_Tab1_SelectedOrders.Visible = False
    rs_Tab1_SelectedOrders.Filter = adFilterNone
    rs_Tab1_SelectedOrders.Sort = "���u�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab1_SelectedOrders.EOF Then rs_Tab1_SelectedOrders.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_Tab1_SelectedOrders.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab1_SelectedOrders.Fields("�s��").Value = dbSeqNo
        rs_Tab1_SelectedOrders.MoveNext
    Loop
    If rs_Tab1_SelectedOrders.RecordCount > 0 Then rs_Tab1_SelectedOrders.MoveFirst
    dg_Tab1_SelectedOrders.Visible = True
    
End Sub

Private Sub ReSet_Tab1_SDN02W_SeqNo()
    '���s���� [dg_Tab0_RouteList1] �� [�s��] ����
    dg_SDN02W.Visible = False
    rs_SDN02W.Filter = adFilterNone
    rs_SDN02W.Sort = "���u�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_SDN02W.EOF Then rs_SDN02W.MoveFirst
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_SDN02W.EOF
        dbSeqNo = dbSeqNo + 1
        rs_SDN02W.Fields("�s��").Value = dbSeqNo
        rs_SDN02W.MoveNext
    Loop
    If rs_SDN02W.RecordCount > 0 Then rs_SDN02W.MoveFirst
    dg_SDN02W.Visible = True
    
End Sub

Private Sub Retrive_OrderSum()
    '�����ݱƨ��q��G�`�p��ƭ�
    txt_Tab1_srcTotal_Case.Text = ""
    txt_Tab1_srcTotal_Volumn.Text = ""
    txt_Tab1_srcTotal_Weight.Text = ""
    'SHIP_CS,SHIP_CBM,SHIP_WT
    str_SQL = "Select Isnull(Round(sum(SHIP_CS),0),0) as �`�c��,Isnull(Round(sum(SHIP_WT),0),0) as �`���q," & _
              "       Isnull(Round(sum(SHIP_CBM),0),0) as �`���n " & _
              "From SDN02W  "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
       txt_Tab1_srcTotal_Case.Text = tmp_Rs.Fields("�`�c��").Value
       txt_Tab1_srcTotal_Volumn.Text = tmp_Rs.Fields("�`���n").Value
       txt_Tab1_srcTotal_Weight.Text = tmp_Rs.Fields("�`���q").Value
    End If
    tmp_Rs.Close
End Sub


Private Sub CreateRS_Tab1_SelectedOrders()
    '�ƨ��@�~�G�w������ݱƨ��q��C��
    Set dg_Tab1_SelectedOrders.DataSource = Nothing
    Call ReDim_Recordset(rs_Tab1_SelectedOrders)
    With rs_Tab1_SelectedOrders
         .Fields.Append "�s��", adDouble
         .Fields.Append "��", adVarChar, 5
         .Fields.Append "�G���ƨ�", adVarChar, 10
         .Fields.Append "���u�s��", adVarChar, 10
         .Fields.Append "�Ȥ�渹", adVarChar, 30
         .Fields.Append "���", adVarChar, 20
         .Fields.Append "���e�Ȥ�", adVarChar, 60
         .Fields.Append "�c��", adDouble
         .Fields.Append "���n", adDouble
         .Fields.Append "���q", adDouble
         .Fields.Append "�q�渹�X", adVarChar, 20
         .Fields.Append "�h��", adVarChar, 50
         .Fields.Append "C_Receipt_No", adVarChar, 20
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '���ݳs������
    End With
    Set dg_Tab1_SelectedOrders.DataSource = rs_Tab1_SelectedOrders
    '�]�w������
    With dg_Tab1_SelectedOrders
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500       '�Ǹ�
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Width = 500       '���
            .Columns(1).Alignment = dbgCenter
            .Columns(2).Width = 1000      '�G���ƨ�
            .Columns(2).Alignment = dbgLeft
            .Columns(3).Width = 1000      '�Ȥ�渹
            .Columns(3).Alignment = dbgLeft
            .Columns(4).Width = 800      '���u�s��
            .Columns(4).Alignment = dbgLeft
            .Columns(5).Width = 800       '���
            .Columns(5).Alignment = dbgLeft
            .Columns(6).Width = 1500      '���e�Ȥ�
            .Columns(6).Alignment = dbgLeft
            .Columns(7).Width = 800       '�c��
            .Columns(7).Alignment = dbgRight
            .Columns(8).Width = 800       '���n
            .Columns(8).Alignment = dbgRight
            .Columns(9).Width = 800       '���q
            .Columns(9).Alignment = dbgRight
            .Columns(10).Width = 1000       '�q�渹�X
            .Columns(10).Alignment = dbgLeft
            .Columns(11).Width = 800       '�h��
            .Columns(11).Alignment = dbgLeft
    End With
End Sub


Private Sub SetGrid_Format_Tab3_SelectedOrderDetail()
    '����@���ݤ��έq�椧���ة���
    Dim sub_var1 As Integer, sub_var2 As Integer
    dg_Tab3_SelectedOrderDetail.Visible = False
    With dg_Tab3_SelectedOrderDetail
         .Rows = 2: .Cols = 10
         .FixedRows = 1
         '�]�w���\��C���
         .AllowBigSelection = True
         '�]�w�C����r�r��
         For sub_var1 = 0 To .Rows - 1
             .Row = sub_var1: .RowHeight(sub_var1) = 250
             For sub_var2 = 0 To .Cols - 1
                 .Col = sub_var2
                 .CellFontName = "�s�ө���": .CellFontSize = 9
             Next sub_var2
         Next sub_var1
         '�]�w�C�����e��
         .ColWidth(0) = 1000
         .ColWidth(1) = 300
         .ColWidth(2) = 1500
         .ColWidth(3) = 2200
         .ColWidth(4) = 600
         .ColWidth(5) = 1000
         .ColWidth(6) = 1000
         .ColWidth(7) = 850
         .ColWidth(8) = 1000
         .ColWidth(9) = 1000

         '�]�w�C�����D
         .Row = 0
         .Col = 0: .Text = "����"
         .Col = 1: .Text = "��"
         .Col = 2: .Text = "�f��"
         .Col = 3: .Text = "�~�W"
         .Col = 4: .Text = "�c��"
         .Col = 5: .Text = "���q"
         .Col = 6: .Text = "���n"
         .Col = 7: .Text = "���νc��"
         .Col = 8: .Text = "���έ��q"
         .Col = 9: .Text = "���Χ��n"
         '�]�w�C����r���
         .ColAlignment(0) = flexAlignLeftCenter
         .ColAlignment(1) = flexAlignCenterCenter
         .ColAlignment(2) = flexAlignLeftCenter
         .ColAlignment(3) = flexAlignLeftCenter
         .ColAlignment(4) = flexAlignRightCenter
         .ColAlignment(5) = flexAlignLeftCenter
         .ColAlignment(6) = flexAlignLeftCenter
         .ColAlignment(7) = flexAlignRightCenter
         .ColAlignment(8) = flexAlignLeftCenter
         .ColAlignment(9) = flexAlignLeftCenter
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1
             .CellAlignment = flexAlignCenterCenter
         Next sub_var1
         .Rows = 2: .Row = 1
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1: .Text = ""
         Next sub_var1
    End With
    dg_Tab3_SelectedOrderDetail.Visible = True
End Sub

Private Sub Calculate_Tab3_SelectedPrderDetail()
    '�p�������q��Ӷ��G�c�ơA���q�A�~�n�A�O��
    dbCut_TotalCaseQty = 0
    txt_Tab3_SelectedCaseQty.Text = ""
    dbCut_TotalWeight = 0
    txt_Tab3_SelectedWeight.Text = ""
    dbCut_TotalVolumn = 0
    txt_Tab3_SelectedVolumn.Text = ""
    
    Dim dbCaseQty As Double, dbWeight As Double, dbVolumn As Double, dbPalletQty As Double
    Dim dbCutPLQty As Double, dbCutCSQty As Double
    Dim i As Double
    With dg_Tab3_SelectedOrderDetail

        For i = 1 To .Rows - 2
            .Row = i
            .Col = 1
            If .Text <> "" Then   '�Q���
                .Col = 4: dbCaseQty = Val(.Text)     '�c��
                .Col = 5: dbWeight = Val(.Text)      '���q
                .Col = 6: dbVolumn = Val(.Text)      '���n
                .Col = 7   '���νc��
                If Val(.Text) <> 0 Then
                     dbCutCSQty = Val(.Text)
                     dbCut_TotalCaseQty = dbCut_TotalCaseQty + dbCutCSQty
                    .Col = 8   '���νc�ƴ��⤧���q
                    .Text = ((dbCutCSQty / dbCaseQty) * dbWeight)
                     dbCut_TotalWeight = dbCut_TotalWeight + ((dbCutCSQty / dbCaseQty) * dbWeight)
                    .Col = 9   '���νc�ƴ��⤧���n
                    .Text = ((dbCutCSQty / dbCaseQty) * dbVolumn)
                     dbCut_TotalVolumn = dbCut_TotalVolumn + ((dbCutCSQty / dbCaseQty) * dbVolumn)
                End If
            Else
                .Col = 7: .Text = ""
                .Col = 8: .Text = ""
                .Col = 9: .Text = ""
            End If
        Next i
    End With
    '��ܿ�����Ӷ��U��줧�[�`��
    txt_Tab3_SelectedCaseQty.Text = dbCut_TotalCaseQty
    txt_Tab3_SelectedWeight.Text = dbCut_TotalWeight
    txt_Tab3_SelectedVolumn.Text = dbCut_TotalVolumn
End Sub

Private Sub Delete_GridRow(ByVal intRow As Double)
    '�ݤ��έq�涵��(Detail) ��ƧR��
    If intRow = 0 Then Exit Sub
    Dim i As Double, j As Integer
    '1. �N�R���C��ƥѤU�@�C��ƨ��N
    '   �ӫ᪺��ƦC���W���@�C
    With dg_Tab3_SelectedOrderDetail
        For i = intRow To .Rows - 2   '�|���h�@��ťզC
            .Row = i
            For j = 0 To .Cols - 1
                .Col = j
                .Text = .TextArray((.Row + 1) * .Cols + .Col)
            Next j
            '����̫�Ĥ@�C���W�����̫�ĤG�C�ɡA�|�O�˥ո�ƦC�A[�Ǹ�] ��줣�঳��
            '����ƪ��C�A[�Ǹ�] �������s�s��
            .Col = 0
            If Val(.Text) = 0 Then .Text = ""   'Else .Text = .Row
        Next i
        '2. Grid �`�C�� - 1
        .Rows = .Rows - 1
        .Row = 1
        For i = 0 To .Cols - 1
            .ColSel = i
        Next i
    End With
End Sub

Private Sub Dispaly_dg_Tab3_SDN03W()
    '�X���T�{ >> Tab3��ܷs�W�q����ӭq��
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab3_SDN03W.DataSource = Nothing
    Set rs_Tab3_SDN03W = Nothing
    On Error GoTo err_Handle
    
    str_SQL = "SELECT  SEQ_NO as ���� " & _
        ",PRODUCT_NO as �f�� " & _
        ",sp.Descr as �~�W " & _
        ",isnull(case when sp.casecnt = 0 then 0 else SHIP_QTY/sp.casecnt end ,0) as �z�f�c�� " & _
        ",isnull(case when sp.casecnt = 0 then 0 else ORDER_QTY/sp.casecnt end ,0) as �q��c�� " & _
        ",(isnull(SHIP_QTY,0)*sp.Stdgrosswgt) as �z�f���q " & _
        ",(isnull(SHIP_QTY,0)*sp.STDCUBE) as �z�f���n " & _
        "from SDN03W inner join gv_skuxpack sp on sp.sku=PRODUCT_NO and sp.storerkey = sdn03w.storerkey " & _
        "where RECEIPT_NO= '" & Trim(txt_Tab3_OrderKey.Text) & "' order by SEQ_NO"
        
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
       tmp_Rs.Close
       Screen.MousePointer = vbDefault
       msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_SDN03W)
    tmp_Rs.Close
    rs_Tab3_SDN03W.MoveFirst
    Set dg_Tab3_SDN03W.DataSource = rs_Tab3_SDN03W
    With dg_Tab3_SDN03W
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 500      '����
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '�f��
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 2500      '�~�W
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1000      '�z�f�c��
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 1000       '�q��c��
        .Columns(5).Alignment = dbgRight
        .Columns(6).Width = 1000      '���q
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 1000       '���n
        .Columns(7).Alignment = dbgRight
    End With
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-Tab3��ܭq�����", Me.Caption, "Dispaly_dg_Tab3_SDN03W", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub Clear_CutOrderDetail()
    Set dg_Tab3_SDN03W.DataSource = Nothing
    txt_Tab3_OrderKey.Text = ""
    txt_Tab3_DeliveryDate.Text = ""
    txt_Tab3_Extern.Text = ""  '�Ȥ�s��
    txt_Tab3_CaseQty.Text = "" '�c��
    txt_Tab3_Weight.Text = ""    '���q
    txt_Tab3_Volumn.Text = ""  '���n
    txt_Tab3_FullName.Text = ""   '�Ȥ�W��
End Sub

Private Sub Delete_RouteNo(strRouteNo As String)
    Screen.MousePointer = vbHourglass
    blTab1RouteEventEnable = False
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
        On Error GoTo err_Handle
        '�R�� TRP01T ���u�s���D��
        Call DB_CheckConnectStatus
        
        '(1).�N SDN03T �g�^ SDN03W >> �R�� SDN03T
        str_SQL = "Insert into SDN03W( " & _
                  "C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME) " & _
                  "Select C_ROUTE_NO,ROUTE_NO,STORERKEY,RECEIPT_NO,SEQ_NO,SubSeq_No,EXTERN,PRODUCT_NO,SHIP_UNIT,SHIP_QTY,SIGN_QTY,RSC_CODE,RBC_CODE,CONFIRM_DATE,DESCRIPTION,ORDER_QTY,SHIP_TIME " & _
                  "From SDN03T  Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '(2).�N SDN02T �g�^ SDN02W >> �R�� SDN02T
        str_SQL = "Insert into SDN02W( " & _
                  "C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,c_receipt_no) " & _
                  "Select C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,C_Route_Time,C_Route_Total,RECEIPT_NO,c_receipt_no " & _
                  "From SDN02T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '(3).�R�� SDN03T & SDN02T & SDN01T
        str_SQL = "Delete From SDN03T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Delete From SDN02T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Delete From SDN01T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '�R���B�O
        str_SQL = "Delete From SDN05T Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans
    Tran_Level = 0
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�X���T�{-���u�s���R��", Me.Caption, "Form ���� SubProgram Delete_RouteNo", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub



Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    '������
    Select Case mvDate.Tag
        Case "Tab0_DELIVERY_DATE0"
             txt_DELIVERY_DATE0.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab0_DELIVERY_DATE1"
             txt_DELIVERY_DATE1.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab1_DELIVERY_DATE"
             txt_Tab1_DELIVERY_DATE.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab2_DELIVERY_DATE"
             txt_Tab2_DELIVERY_DATE.Text = Format(mvDate.Value, "YYYYMMDD")
        Case "Tab2_DELIVERYDATE_START"
             txt_Tab2_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
        Case Else
    End Select
    mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    mvDate.Visible = False
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_DELIVERY_DATE0_Click()
    'Tab0 >> �X�����
    If Trim(txt_DELIVERY_DATE0.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_DELIVERY_DATE0.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_DELIVERY_DATE0.Text, 4) & "/" & Mid(txt_DELIVERY_DATE0.Text, 5, 2) & "/" & Right(txt_DELIVERY_DATE0.Text, 2))
       End If
    End If
    mvDate.Tag = "Tab0_DELIVERY_DATE0"
    mvDate.Top = SSTab1.Top + fam_Tab0_Consignee.Top + txt_DELIVERY_DATE0.Top + txt_DELIVERY_DATE0.Height
    mvDate.Left = SSTab1.Left + fam_Tab0_Consignee.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub



Private Sub txt_DELIVERY_DATE1_Click()
    'Tab0 >> �X�����
    If Trim(txt_DELIVERY_DATE1.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_DELIVERY_DATE1.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_DELIVERY_DATE1.Text, 4) & "/" & Mid(txt_DELIVERY_DATE1.Text, 5, 2) & "/" & Right(txt_DELIVERY_DATE1.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab0_DELIVERY_DATE1"
    mvDate.Top = SSTab1.Top + Frame1.Top + txt_DELIVERY_DATE1.Top + txt_DELIVERY_DATE1.Height
    mvDate.Left = SSTab1.Left + fam_Tab0_Consignee.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DELIVERY_DATE_Click()
    'Tab0 >> �X�����
    If Trim(txt_Tab1_DELIVERY_DATE.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab1_DELIVERY_DATE.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab1_DELIVERY_DATE.Text, 4) & "/" & Mid(txt_Tab1_DELIVERY_DATE.Text, 5, 2) & "/" & Right(txt_Tab1_DELIVERY_DATE.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab1_DELIVERY_DATE"
    mvDate.Top = SSTab1.Top + fam_SelectedOrders.Top + txt_Tab1_DELIVERY_DATE.Top + txt_Tab1_DELIVERY_DATE.Height
    mvDate.Left = SSTab1.Left + fam_SelectedOrders.Left + txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub

Private Sub clear_Tab0_RouteList0()
    '�e���B�z
    Set dg_Tab0_RouteList0.DataSource = Nothing
    txt_Tab0_C_Route_No0.Text = ""
    txt_DELIVERY_DATE0.Text = ""
    txt_VehicleNo0.Text = ""
    txt_Driver0.Text = ""
End Sub
Private Sub clear_Tab0_RouteList1()
    '�e���B�z
    Set dg_Tab0_RouteList1.DataSource = Nothing
    txt_Tab0_C_Route_No1.Text = ""
    txt_DELIVERY_DATE1.Text = ""
    txt_VehicleNo1.Text = ""
    txt_Driver1.Text = ""
End Sub

Private Sub Display_Tab2_RouteOrders()
    'SDN03T
    If Tab2_RouteListEventEnable = False Then Exit Sub
    str_SQL = "SELECT  C_ROUTE_NO AS �G���ƨ�, ROUTE_NO AS ���u�s��,EXTERN AS �Ȥ�渹,ARRIVE_DATE AS ���,CUST_NAME as ���e�Ȥ�, " & _
            "SHIP_CS As �c��, SHIP_CBM As ���n, SHIP_WT As ���q, RECEIPT_NO As �q�渹�X, CAR_NOTES As �h�� " & _
            "FROM   SDN02T Where C_ROUTE_NO = '" & Trim(rs_Tab2_Route.Fields("�G���ƨ�").Value) & "' Order by C_ROUTE_NO,Receipt_No"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���u�s�����q����(TRP02T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab2_RouteOrders)
    Set dg_Tab2_RouteOrders.DataSource = rs_Tab2_RouteOrders
    tmp_Rs.Close
    With dg_Tab2_RouteOrders
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000      '�G���ƨ�
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '�Ȥ�渹
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800      '���u�s��
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800       '���
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1600      '���e�Ȥ�
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 800       '�c��
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 800       '���n
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 800       '���q
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000       '�h��
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 800       '�q�渹�X
        .Columns(10).Alignment = dbgRight
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�X���T�{ >> Tab2���u�s���d��", Me.Caption, "cmd_Tab2_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub txt_Tab1_VehicleNo_LostFocus()

If Len(Trim(txt_Tab1_VehicleNo)) = 0 Then Exit Sub

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

tmp_Rs.Open "select driver=isnull(driver,'') from trp09m where Vehicle_id_No = '" & Trim(txt_Tab1_VehicleNo) & "' ", cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then MsgBox "�L������!", 16, "�`�N": txt_Tab1_VehicleNo.SetFocus: Exit Sub

txt_Tab1_Driver0 = RTrim(tmp_Rs("driver")) & ""
tmp_Rs.Close

End Sub

Private Sub txt_Tab2_DELIVERY_DATE_Click()
    'Tab2 >> �X�����
    If Trim(txt_Tab2_DELIVERY_DATE.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab2_DELIVERY_DATE.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab2_DELIVERY_DATE.Text, 4) & "/" & Mid(txt_Tab2_DELIVERY_DATE.Text, 5, 2) & "/" & Right(txt_Tab2_DELIVERY_DATE.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab2_DELIVERY_DATE"
    mvDate.Top = SSTab1.Top + fam_Tab0_Consignee.Top + txt_Tab2_DELIVERY_DATE.Top + txt_Tab2_DELIVERY_DATE.Height
    mvDate.Left = SSTab1.Left + Frame6.Left + txt_Tab2_DELIVERY_DATE.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub

Private Sub txt_Tab2_DeliveryDate_Start_Click()
    'Tab2 >> �X�����
    If Trim(txt_Tab2_DeliveryDate_Start.Text) = "" Then
        mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_Start.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab2_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab2_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab2_DeliveryDate_Start.Text, 2))
        End If
    End If
    mvDate.Tag = "Tab2_DELIVERYDATE_START"
    mvDate.Top = SSTab1.Top + fam_Tab0_Consignee.Top + txt_Tab2_DeliveryDate_Start.Top + txt_Tab2_DeliveryDate_Start.Height
    mvDate.Left = SSTab1.Left + Frame6.Left + txt_Tab2_DELIVERY_DATE.Left     '+ txt_DELIVERY_DATE0.Left
    mvDate.Visible = True
End Sub
