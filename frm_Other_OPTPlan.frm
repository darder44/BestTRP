VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Other_OPTPlan 
   Caption         =   "�䥦�ƨ��@�~"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   12930
   WindowState     =   2  '�̤j��
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3960
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   4560
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
      StartOfWeek     =   104660993
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "��L�ƨ��@�~"
      TabPicture(0)   =   "frm_Other_OPTPlan.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fam_RouteData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fam_SelectedOrders"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fam_SrcOrders"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "���u�s���C��"
      TabPicture(1)   =   "frm_Other_OPTPlan.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dg_Tab1_Route"
      Tab(1).Control(1)=   "dg_Tab1_RouteOrders"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "�O�d�q��"
      TabPicture(2)   =   "frm_Other_OPTPlan.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dg_Tab2_ReservedOrders"
      Tab(2).Control(1)=   "cmd_Tab2_Delete"
      Tab(2).Control(2)=   "cmd_Tab2_FilterAndSort"
      Tab(2).Control(3)=   "cmd_Tab2_Reset"
      Tab(2).Control(4)=   "cmd_Tab2_ShowAll"
      Tab(2).Control(5)=   "cmd_Tab2_Remove"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frm_Other_OPTPlan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fam_Header"
      Tab(3).Control(1)=   "dgMain3"
      Tab(3).ControlCount=   2
      Begin MSDataGridLib.DataGrid dgMain3 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   10398
         _Version        =   393216
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
      Begin VB.Frame fam_Header 
         Height          =   705
         Left            =   -74880
         TabIndex        =   88
         Top             =   360
         Width           =   7935
         Begin VB.CommandButton cmdExport3 
            BackColor       =   &H8000000A&
            Caption         =   "��ƶץX"
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
            Height          =   525
            Left            =   4560
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
            Top             =   135
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryDate3 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   1290
            TabIndex        =   22
            Top             =   225
            Width           =   1350
         End
         Begin VB.CommandButton cmdRouteQuery3 
            BackColor       =   &H8000000A&
            Caption         =   "���s�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2895
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   23
            Top             =   135
            Width           =   1485
         End
         Begin VB.CommandButton cmdExit3 
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
            Height          =   525
            Index           =   1
            Left            =   6285
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   25
            Top             =   135
            Width           =   1485
         End
         Begin MSComDlg.CommonDialog CmnDialog 
            Left            =   120
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   22
            Left            =   195
            TabIndex        =   89
            Top             =   270
            Width           =   1020
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '����
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   -65595
         TabIndex        =   86
         Top             =   2475
         Width           =   1980
         Begin VB.CommandButton cmd_Tab1_RouteNoDelete 
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
            Height          =   825
            Left            =   105
            Picture         =   "frm_Other_OPTPlan.frx":0070
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   21
            ToolTipText     =   "�R��"
            Top             =   210
            Width           =   1785
         End
      End
      Begin VB.Frame fam_SrcOrders 
         Height          =   2835
         Left            =   120
         TabIndex        =   67
         Top             =   4320
         Width           =   12660
         Begin VB.Frame fam_SelectedSum 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   77
            Top             =   0
            Width           =   5595
            Begin VB.TextBox txt_Tab0_srcSelected_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   4695
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   3465
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Pallet 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   2220
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   990
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "����G�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   85
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�O��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   1845
               TabIndex        =   84
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   2
               Left            =   3075
               TabIndex        =   83
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   3
               Left            =   4320
               TabIndex        =   82
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   525
            Left            =   5610
            TabIndex        =   68
            Top             =   0
            Width           =   5595
            Begin VB.TextBox txt_Tab0_srcTotal_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   975
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Pallet 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   2220
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   3465
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   4680
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   8
               Left            =   4305
               TabIndex        =   76
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   9
               Left            =   3075
               TabIndex        =   75
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�O��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   10
               Left            =   1845
               TabIndex        =   74
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�`�p�G�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   11
               Left            =   75
               TabIndex        =   73
               Top             =   210
               Width           =   900
            End
         End
         Begin MSDataGridLib.DataGrid dg_TRP02W 
            Height          =   2160
            Left            =   45
            TabIndex        =   1
            Top             =   525
            Width           =   12435
            _ExtentX        =   21934
            _ExtentY        =   3810
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
         Height          =   3375
         Left            =   105
         TabIndex        =   46
         Top             =   1020
         Width           =   12660
         Begin VB.CommandButton cmd_Tab0_CreateRouteByAds 
            Appearance      =   0  '����
            BackColor       =   &H00FFFF00&
            Caption         =   "  �̦a�}  �ո��s"
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
            Left            =   11280
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   94
            Top             =   120
            Width           =   990
         End
         Begin VB.CheckBox chk_Tab0_Updateortw 
            Caption         =   "��s����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   405
            Left            =   5400
            TabIndex        =   93
            Top             =   135
            Width           =   750
         End
         Begin VB.Frame Frame1 
            Enabled         =   0   'False
            Height          =   525
            Left            =   15
            TabIndex        =   52
            Top             =   2820
            Width           =   5595
            Begin VB.TextBox txt_Tab0_Selected_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   4695
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   3465
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Pallet 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   2220
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   990
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�֭p�G�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   7
               Left            =   75
               TabIndex        =   60
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�O��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   6
               Left            =   1845
               TabIndex        =   59
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   5
               Left            =   3075
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
               Index           =   4
               Left            =   4320
               TabIndex        =   57
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel_All 
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
            Left            =   7830
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   2910
            Width           =   1530
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
            Height          =   375
            Left            =   6015
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   3
            Top             =   2910
            Width           =   345
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
            Height          =   375
            Left            =   5640
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   2
            Top             =   2910
            Width           =   345
         End
         Begin VB.TextBox txt_Tab0_TRPDate 
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
            Left            =   1605
            TabIndex        =   9
            Top             =   150
            Width           =   1110
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarNo 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3240
            TabIndex        =   10
            Top             =   150
            Width           =   1125
         End
         Begin VB.CommandButton cmd_Tab0_SelectCar 
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
            Left            =   4320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   90
            Top             =   150
            Width           =   330
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarType 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   9360
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   315
            Width           =   945
         End
         Begin VB.TextBox txt_Tab0_DeliveryDriver 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   7050
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab0_DeliveryCompany 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   6240
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   315
            Width           =   825
         End
         Begin VB.TextBox txt_Tab0_DeliveryPhone 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   8205
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   315
            Width           =   1170
         End
         Begin VB.CommandButton cmd_Tab0_srcOrderQuery 
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
            TabIndex        =   6
            Top             =   2910
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_srcOrderReset 
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
            TabIndex        =   7
            Top             =   2910
            Width           =   495
         End
         Begin VB.CheckBox chk_Tab0_DriveTimes 
            Caption         =   "��ܨ���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   405
            Left            =   4665
            TabIndex        =   11
            Top             =   150
            Width           =   750
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel 
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
            Left            =   8925
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   47
            Top             =   2910
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.CommandButton cmd_Tab0_Reserve 
            BackColor       =   &H00FF8080&
            Caption         =   "�O�d�q��"
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
            TabIndex        =   4
            Top             =   2910
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab0_ImportOrders 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���J�ݱƨ��q��"
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
            TabIndex        =   0
            Top             =   105
            Width           =   1095
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
            Index           =   0
            Left            =   10320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   17
            Top             =   120
            Width           =   870
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_SelectedOrders 
            Height          =   2145
            Left            =   0
            TabIndex        =   8
            Top             =   600
            Width           =   12570
            _ExtentX        =   22172
            _ExtentY        =   3784
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
            Height          =   435
            Index           =   12
            Left            =   1170
            TabIndex        =   66
            Top             =   150
            Width           =   435
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
            Height          =   390
            Index           =   13
            Left            =   2790
            TabIndex        =   65
            Top             =   165
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   9510
            TabIndex        =   64
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�r�p�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   15
            Left            =   7335
            TabIndex        =   63
            Top             =   120
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�餽�q"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   6240
            TabIndex        =   62
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   8565
            TabIndex        =   61
            Top             =   120
            Width           =   540
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '���z��
            Height          =   435
            Left            =   5610
            Top             =   2880
            Width           =   795
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   435
            Index           =   0
            Left            =   6600
            Top             =   2880
            Width           =   2790
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00FF0000&
            FillStyle       =   0  '���
            Height          =   435
            Left            =   9495
            Top             =   2880
            Width           =   1680
         End
      End
      Begin VB.Frame fam_RouteData 
         Height          =   585
         Left            =   105
         TabIndex        =   36
         Top             =   420
         Width           =   11220
         Begin VB.TextBox txt_Tab0_Route 
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
            Left            =   540
            TabIndex        =   91
            Top             =   135
            Width           =   1380
         End
         Begin VB.CommandButton cmd_Tab0_CreateRoute 
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
            Left            =   10065
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   15
            Top             =   90
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_SelectedRemove_All 
            BackColor       =   &H000080FF&
            Caption         =   "�w��q�沾��(��)"
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
            Left            =   3855
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   41
            Top             =   75
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab0_DockNo 
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
            Left            =   5475
            TabIndex        =   12
            Top             =   135
            Width           =   1155
         End
         Begin VB.TextBox txt_Tab0_CarCheckInTime 
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
            Left            =   9285
            MaxLength       =   4
            TabIndex        =   14
            Top             =   135
            Width           =   750
         End
         Begin VB.TextBox txt_Tab0_CarCheckInDate 
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
            Left            =   7320
            TabIndex        =   13
            Top             =   135
            Width           =   1140
         End
         Begin VB.CommandButton cmd_Tab0_Query 
            Appearance      =   0  '����
            BackColor       =   &H00808000&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1980
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   39
            Top             =   75
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmd_Tab0_Save 
            BackColor       =   &H00FF8080&
            Caption         =   "�s��"
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
            Left            =   2595
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   38
            Top             =   75
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmd_Tab0_Clear 
            BackColor       =   &H008080FF&
            Caption         =   "�M��"
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
            Left            =   3210
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   37
            Top             =   75
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox txt_Tab0_RouteNo 
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
            Left            =   540
            TabIndex        =   40
            Top             =   150
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�ѦҸ��s"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   23
            Left            =   120
            TabIndex        =   92
            Top             =   135
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�X�Y�Ȧs"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   18
            Left            =   5040
            TabIndex        =   45
            Top             =   135
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�w�p����ɶ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   19
            Left            =   8610
            TabIndex        =   44
            Top             =   135
            Width           =   675
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   0
            Left            =   4980
            Top             =   105
            Width           =   1680
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   1
            Left            =   8565
            Top             =   105
            Width           =   1500
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   2
            Left            =   6675
            Top             =   105
            Width           =   1875
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�w�p������"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   20
            Left            =   6720
            TabIndex        =   43
            Top             =   135
            Width           =   675
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   3
            Left            =   45
            Top             =   105
            Width           =   1890
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   540
            Index           =   1
            Left            =   1950
            Top             =   45
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   540
            Index           =   2
            Left            =   3840
            Top             =   45
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label1 
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
            Height          =   435
            Index           =   21
            Left            =   105
            TabIndex        =   42
            Top             =   150
            Visible         =   0   'False
            Width           =   435
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '����
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   -65625
         TabIndex        =   34
         Top             =   510
         Width           =   1995
         Begin VB.TextBox txt_Tab1_RouteNo 
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
            Left            =   180
            TabIndex        =   19
            Top             =   525
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab1_RouteNoQuery 
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
            Height          =   825
            Left            =   105
            Picture         =   "frm_Other_OPTPlan.frx":037A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   20
            Top             =   975
            Width           =   1785
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
            TabIndex        =   35
            Top             =   195
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmd_Tab2_Remove 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���ܫݱƨ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65100
         Picture         =   "frm_Other_OPTPlan.frx":0684
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   32
         Top             =   2550
         Width           =   1440
      End
      Begin VB.CommandButton cmd_Tab2_ShowAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "���J�����q��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65100
         Picture         =   "frm_Other_OPTPlan.frx":098E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   31
         Top             =   825
         Width           =   1440
      End
      Begin VB.CommandButton cmd_Tab2_Reset 
         Appearance      =   0  '����
         BackColor       =   &H00808080&
         Caption         =   "�����q��"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65100
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   30
         Top             =   4185
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmd_Tab2_FilterAndSort 
         BackColor       =   &H00C0E0FF&
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
         Height          =   825
         Left            =   -65100
         Picture         =   "frm_Other_OPTPlan.frx":0C98
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmd_Tab2_Delete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�f�D�渹�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65100
         Picture         =   "frm_Other_OPTPlan.frx":1562
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   28
         ToolTipText     =   "�R��"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1440
      End
      Begin MSDataGridLib.DataGrid dg_Tab2_ReservedOrders 
         Height          =   6330
         Left            =   -74895
         TabIndex        =   33
         Top             =   510
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   11165
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid dg_Tab1_RouteOrders 
         Height          =   3240
         Left            =   -74880
         TabIndex        =   18
         Top             =   3645
         Width           =   11220
         _ExtentX        =   19791
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
      Begin MSDataGridLib.DataGrid dg_Tab1_Route 
         Height          =   3105
         Left            =   -74910
         TabIndex        =   16
         Top             =   510
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   5477
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
Attribute VB_Name = "frm_Other_OPTPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private blTRP02WEventEnable As Boolean
Private blORT02WEventEnable As Boolean              '�ݿ���q�� Event Ĳ�o���ı���
Private blTab0SelectedOrderEventEnable As Boolean   '�w����q�� Event Ĳ�o���ı���
Private blTab1RouteEventEnable As Boolean           '���u�s���C�� Event Ĳ�o���ı���
Private blTab2ReservedEventEnable As Boolean        '�O�d�q��C�� Event Ĳ�o���ı���

Private blRouteModify As Boolean                    '�ƨ��@�~ >> ���u�s�� �d�ߡG���ĸ��u�s��
Private blRouteChange As Boolean                    '�ƨ��@�~ >> ���u�s�� ��Ʋ����ѧO�X��
Private strDispRouteNo As String                    '�ƨ��@�~ >> ���u�s�� �d�ߡG���u�s��

Private rs_ORT02W As ADODB.Recordset                '�ƨ��@�~�G�פJ���ݱƨ��q��
Private rs_Tab0_SelectedOrders As ADODB.Recordset   '�ƨ��@�~�G�w������ݱƨ��q��
Private rs_Tab1_Route As ADODB.Recordset            '���s�C��G���u�s���C��
Private rs_Tab1_RouteOrders As ADODB.Recordset      '���s�C��G���u�s�����ݤ��q��
Private rs_Tab2_ReservedOrders As ADODB.Recordset   '�O�d�q��

Private strSourceFilter As String        '�ݱƨ��q��z��
Private strSourceOrderBy As String       '�ݱƨ��q��ƧǤ覡
Private dbsrcSelected_Case As Double     '�ݱƨ��q��: ����c��
Private dbsrcSelected_Pallet As Double   '�ݱƨ��q��: ����O��
Private dbsrcSelected_Volumn As Double   '�ݱƨ��q��: ������n
Private dbsrcSelected_Weight As Double   '�ݱƨ��q��: ������q
Private dbSelectedCount As Double        '����q�浧��
Private DelRecord

Private rsMain3 As ADODB.Recordset

Private Sub cmd_Exit_Click(Index As Integer)
    '���}
    Unload Me
End Sub

Private Sub cmd_Tab0_Clear_Click()
    '�ƨ��@�~ >> �M��
    If rs_Tab0_SelectedOrders.RecordCount <> 0 And blRouteModify = False Then
        '�s�W���u�s���Ҧ��G
        '�I�s [�w��q�沾��(��)] �ӳB�z�w�Q�Ȯɿ���� [�ݱƨ��q��] �٭�^ [�ݱƨ��q��]
        Call cmd_Tab0_SelectedRemove_All_Click
    End If
    If blRouteModify And blRouteChange Then
       '���ĸ��u�s�� & ��Ƥw�D���ʡA�n user �T�{�O�_�s��
        msg_text = "���u�s����ƬO�_�s�ɡH"
        If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
            '�I�s�s�ɵ{��
            Call cmd_Tab0_Save_Click
        Else
            '���s�ɡ��������s���J [�ݱƨ��q��] �w�٭� [���][����] �ާ@�� [�ݱƨ��q��] ���v�T
            Call cmd_Tab0_ImportOrders_Click
        End If
    End If
    '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
    Call Clear_RouteData
    txt_Tab0_RouteNo.Text = ""
End Sub

Private Sub cmd_Tab0_CreateRoute_Click()
    '�ƨ��@�~ >> �إ߸��u�s��
    
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then
        msg_text = "��ƿ��~�G�L�˸����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    'Terry 20190619 �ˬdReceiptno�O�_�w�s�b�إߦn���@�����s��
    Dim strReceiptNo As String
    strReceiptNo = ""
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        strReceiptNo = strReceiptNo & "'" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "',"
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    strReceiptNo = strReceiptNo & "''"
    
    str_SQL = "select receipt_no from ort02t where receipt_no in (" & strReceiptNo & ")"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        MsgBox ("���q��w�զ��@�����s�A�Э��s���J�ݱƨ��q��òM��[�w������@���q��]"), vbOKOnly + vbCritical
        tmp_Rs.Close
        Exit Sub
    End If
    tmp_Rs.Close
    
    '�ˬd�O�_�V�f�D�ը�
    Dim strStorerkey As String
    rs_Tab0_SelectedOrders.MoveFirst
    strStorerkey = Mid(rs_Tab0_SelectedOrders("�q��s��"), 12, 6)
    
    Do While Not rs_Tab0_SelectedOrders.EOF
        If strStorerkey <> Mid(rs_Tab0_SelectedOrders("�q��s��"), 12, 6) Then '�V�f�D
            If MsgBox("������t�����P�f�D�A�нT�{�O�_�~��إ߸��s?", vbYesNo, "�V�f�D�ը�") <> vbYes Then
                Exit Sub
            Else
                GoTo NextStep
            End If
        End If
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
NextStep:
    
    '�ˮָ��u�s����ƬO�_���T�A���~�N�b Function ������� MessageBox
    If RouteData_Check() = False Then Exit Sub
    
    On Error GoTo err_Handle
    
    cmd_Tab0_CreateRoute.Enabled = False
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    
    '�ˬd�i�����q
    Dim intableWT, intableCBM
    str_SQL = "select rtrim(isnull(LOADING_SIZE,0)),rtrim(isnull(MAX_CUBIC_CAPACITY,0)) from dbo.TRP09M where VEHICLE_ID_NO='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intableWT = tmp_Rs.Fields(0).Value
    intableCBM = tmp_Rs.Fields(1).Value
    tmp_Rs.Close
    If intableWT < Val(txt_Tab0_Selected_Weight.Text) Then
        msg_text = "�ƨ����q�W�L�����i����,�����i����:" & intableWT
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        cmd_Tab0_CreateRoute.Enabled = True
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    
    If intableCBM < Val(txt_Tab0_Selected_Volumn.Text) Then
        msg_text = "�ƨ����q�W�L�����i�����n,�����i�����n:" & intableCBM
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        cmd_Tab0_CreateRoute.Enabled = True
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Dim intDriveTimes As Integer    '����
    Dim strRouteNo As String        '���u�s��
    
    '1.���ͨ���
    str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
              "From ORT05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    tmp_Rs.Close
    
    '2.���͸��u�s��
    str_SQL = "Select Isnull(Max(Cast(Right(Route_No,3) as integer))+1,1) as RouteSN " & _
              "From ORT01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'R'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strRouteNo = "R" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
    tmp_Rs.Close
    
    '3.Insert into ORT01T ���u�s���D��
    '  ORT01T.EXE_CONFIRM = '0' �s���͸��u�s���A�|���^�ǹL exe
    str_SQL = "Insert into ORT01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
              strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
              txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '4.insert into ORT05T �����i�X�޲z
    str_SQL = "Insert into ORT05T (Route_No,Vehicle_ID_No,Drive_Times,Delivery_Date,Valid_Vehicle,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Dock_No,Expect_Time,Expect_Date) Values ('" & _
              strRouteNo & "','" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",'" & _
              Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "','1'," & _
              txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'" & _
              txt_Tab0_DockNo.Text & "','" & txt_Tab0_CarCheckInTime.Text & "','" & txt_Tab0_CarCheckInDate.Text & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�Ѩ����D�ɧ�s�����������
    str_SQL = "Update ORT05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From ORT05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and Route_No = '" & strRouteNo & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�g�� SSTab1.Tab 1 [���u�s���C��]
    blTab1RouteEventEnable = False
    rs_Tab1_Route.AddNew
    rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
    rs_Tab1_Route.Fields("���u�s��").Value = strRouteNo
    rs_Tab1_Route.Fields("�X�����").Value = txt_Tab0_TRPDate.Text
    rs_Tab1_Route.Fields("���P���X").Value = txt_Tab0_DeliveryCarNo.Text
    rs_Tab1_Route.Fields("����").Value = intDriveTimes
    rs_Tab1_Route.Fields("�r�p�H").Value = txt_Tab0_DeliveryDriver.Text
    rs_Tab1_Route.Fields("�c��").Value = txt_Tab0_Selected_Case.Text
    rs_Tab1_Route.Fields("�O��").Value = txt_Tab0_Selected_Pallet.Text
    rs_Tab1_Route.Fields("���n").Value = txt_Tab0_Selected_Volumn.Text
    rs_Tab1_Route.Fields("���q").Value = txt_Tab0_Selected_Weight.Text
    rs_Tab1_Route.Fields("����").Value = txt_Tab0_DeliveryCarType.Text
    rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = txt_Tab0_DockNo.Text
    rs_Tab1_Route.Fields("�w�p������").Value = txt_Tab0_CarCheckInDate.Text
    rs_Tab1_Route.Fields("�w�p����ɶ�").Value = txt_Tab0_CarCheckInTime.Text
    rs_Tab1_Route.Fields("EXE�^��").Value = "�s�ظ��s"
    rs_Tab1_Route.Fields("�ƨ���").Value = User_id
    rs_Tab1_Route.Update
    blTab1RouteEventEnable = True
    
    '5.insert into ORT02T [�ƨ��q����]
    '  �g�� SSTab1.Tab 1 [���u�s�����q��W�Ӫ�]
    blTab0SelectedOrderEventEnable = False
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.MoveFirst
    
    Do While Not rs_Tab0_SelectedOrders.EOF
        'insert into ORT02T
        str_SQL = "Insert into ORT02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                  "Select '" & strRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                  "From ORT02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '�g�J�ѦҸ��u�s����orders.containertype
        str_SQL = "Update orders Set containertype = '" & Trim(txt_Tab0_Route.Text) & "' , trafficCop = null Where orderkey = '" & Left(rs_Tab0_SelectedOrders("�q��s��"), 10) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '�g�� SSTab1.Tab 1 [���u�s�����q����Ӫ�]
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("�s��").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("���u�s��").Value = strRouteNo
        rs_Tab1_RouteOrders.Fields("���h��").Value = rs_Tab0_SelectedOrders.Fields("���h��").Value
        rs_Tab1_RouteOrders.Fields("�q��s��").Value = rs_Tab0_SelectedOrders.Fields("�q��s��").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = rs_Tab0_SelectedOrders.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("��f�Ȥ�²��").Value = rs_Tab0_SelectedOrders.Fields("��f�Ȥ�²��").Value
        rs_Tab1_RouteOrders.Fields("�c��").Value = rs_Tab0_SelectedOrders.Fields("�c��").Value
        rs_Tab1_RouteOrders.Fields("�O��").Value = rs_Tab0_SelectedOrders.Fields("�O��").Value
        rs_Tab1_RouteOrders.Fields("���n").Value = rs_Tab0_SelectedOrders.Fields("���n").Value
        rs_Tab1_RouteOrders.Fields("���q").Value = rs_Tab0_SelectedOrders.Fields("���q").Value
        rs_Tab1_RouteOrders.Fields("����").Value = rs_Tab0_SelectedOrders.Fields("����").Value
        rs_Tab1_RouteOrders.Fields("�q��Ƶ�").Value = rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value
        rs_Tab1_RouteOrders.Fields("�S��ݨD1").Value = rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value
        rs_Tab1_RouteOrders.Fields("�S��ݨD2").Value = rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE�^��").Value = rs_Tab0_SelectedOrders.Fields("EXE�^��").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = rs_Tab0_SelectedOrders.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("���A").Value = rs_Tab0_SelectedOrders.Fields("���A").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�²��").Value = rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value
        rs_Tab1_RouteOrders.Update
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
    '�T�{���u�s���� exe_confirm ���A
    '�D�n�ت��G�w�^�Ǥ����s�R����A���s���ͤ����s�A�Y�������O�H�^�ǭq��A�������s�]�w�� [�w�^��]
'Mark by Gemini @20111010
'    str_SQL = "Update ORT01T Set EXE_Confirm = (Select min(EXE_Confirm) From ORT02T Where ORT02T.Route_No = ORT01T.Route_No) " & _
'              "Where ORT01T.Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
    cn.CommitTrans
    Tran_Level = 0
    
    If dg_Tab1_Route.SelBookmarks.Count > 0 Then
        dg_Tab1_Route.SelBookmarks.Remove 0
    End If
    dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
    rs_Tab1_RouteOrders.Filter = " ���u�s�� = '" & strRouteNo & "' "
    blTab0SelectedOrderEventEnable = True
    
    '4.�� TRP02T Trigger [insert] �i��H�U�@�~
    '   a.�g�J TRP03T -- �ƨ��q�������
    '   b.�R�� TRP03W -- �ݱƨ��q�������
    '   c.�R�� TRP02W -- �ݱƨ��q��D��
    
    
    
    '5.�M�� [�w������ݱƨ��q��C��]
    blTab0SelectedOrderEventEnable = False
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w-ReSet
    Call CreateRS_Tab0_SelectedOrders
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    '6.�M���ƨ��@�~����
    txt_Tab0_DockNo.Text = ""               '�X�Y�Ȧs
    txt_Tab0_CarCheckInDate.Text = ""       '�����w�p������
    txt_Tab0_CarCheckInTime.Text = ""       '�����w�p����ɶ�
    txt_Tab0_TRPDate.Text = ""              '�X�����
    txt_Tab0_DeliveryCarNo.Text = ""        '���P���X
    txt_Tab0_DeliveryCompany.Text = ""      '�B�餽�q
    txt_Tab0_DeliveryDriver.Text = ""       '�r�p�H
    txt_Tab0_DeliveryPhone.Text = ""        '�q��
    txt_Tab0_DeliveryCarType.Text = ""      '����
    
    cmd_Tab0_CreateRoute.Enabled = True
    
    '�ݱƨ��q���`�p��T
    Call ReCaculate_OrderSum
    
    SSTab1.Tab = 1
    DoEvents: DoEvents
    
'    'Terry 20200212 �ƨ������JBestAPP Ĳ�o�����\�� �L�״��ϥ�
'    cn.Execute "exec Andys_BestTMSOrderImport", RowsAffect, adExecuteNoRecords
'    Dim HttpClient As Object
'
'    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
'    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/InsertWaybillList", False
'    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
'    HttpClient.Send
'
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
   End If
   
   '�D�J���~���ܡGlocal �� Recordset [���u�s���C��] ��ƥ����R��
   '�]�� [���u�s���C��] ���� DB connection.transaction ����
   blTab1RouteEventEnable = False
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_Route.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_RouteOrders.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_RouteOrders.EOF Then
      Do While Not rs_Tab1_RouteOrders.EOF
         rs_Tab1_RouteOrders.Delete
         rs_Tab1_RouteOrders.MoveFirst
      Loop
   End If
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
      
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-�إ߸��u�s��", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_CreateRoute.Enabled = True
End Sub

Private Sub cmd_Tab0_CreateRouteByAds_Click()
   '�ƨ��@�~ >> �إ߸��u�s��
   'Terry 20191107 �h�f�ƨ��s�W�@�a�}�ո��s
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then
        msg_text = "��ƿ��~�G�L�˸����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    '�ˮָ��u�s����ƬO�_���T�A���~�N�b Function ������� MessageBox
    If RouteData_Check() = False Then Exit Sub
    
    On Error GoTo err_Handle
    
    
    'Terry�ˬdReceiptno�O�_�w�s�b�إߦn���@�����s��
    Dim strReceiptNo As String
    strReceiptNo = ""
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        strReceiptNo = strReceiptNo & "'" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "',"
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    strReceiptNo = strReceiptNo & "''"
    
    str_SQL = "select receipt_no from ort02t where receipt_no in (" & strReceiptNo & ")"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        MsgBox ("���q��w�զ��@�����s�A�Э��s���J�ݱƨ��q��òM��[�w������@���q��]"), vbOKOnly + vbCritical
        tmp_Rs.Close
        Exit Sub
    End If
    tmp_Rs.Close
    
    '�ˬd�O�_�V�f�D�ը�
    Dim strStorerkey As String
    rs_Tab0_SelectedOrders.MoveFirst
    strStorerkey = Mid(rs_Tab0_SelectedOrders("�q��s��"), 12, 6)
    
    Do While Not rs_Tab0_SelectedOrders.EOF
        If strStorerkey <> Mid(rs_Tab0_SelectedOrders("�q��s��"), 12, 6) Then '�V�f�D
            If MsgBox("������t�����P�f�D�A�нT�{�O�_�~��إ߸��s?", vbYesNo, "�V�f�D�ը�") <> vbYes Then
                Exit Sub
            Else
                GoTo NextStep
            End If
        End If
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
NextStep:

    '�ˬd�i�����q
    Dim intableWT, intableCBM
    str_SQL = "select rtrim(isnull(LOADING_SIZE,0)),rtrim(isnull(MAX_CUBIC_CAPACITY,0)) from dbo.TRP09M where VEHICLE_ID_NO='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intableWT = tmp_Rs.Fields(0).Value
    intableCBM = tmp_Rs.Fields(1).Value
    tmp_Rs.Close
    If intableWT < Val(txt_Tab0_Selected_Weight.Text) Then
        msg_text = "�ƨ����q�W�L�����i����,�����i����:" & intableWT
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    
    If intableCBM < Val(txt_Tab0_Selected_Volumn.Text) Then
        msg_text = "�ƨ����q�W�L�����i�����n,�����i�����n:" & intableCBM
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    
    cmd_Tab0_CreateRouteByAds.Enabled = False
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Dim intDriveTimes As Integer    '����
    Dim strRouteNo As String        '���u�s��
    Dim strAddress As String        '���P�a�}���ͷs�����u�s��
    Dim strRouteNosum As String     '��sTRP01�BTRP05
    strAddress = ""
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "Zip,��f�a�}"
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        
        If Trim(strAddress) <> Trim(rs_Tab0_SelectedOrders.Fields("��f�a�}").Value) Then '�a�}���@��
            
            strAddress = Trim(rs_Tab0_SelectedOrders.Fields("��f�a�}").Value)
            '1.���ͨ���
            str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                      "From ORT05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
            tmp_Rs.Close
            
            '2.���͸��u�s��
            str_SQL = "Select Isnull(Max(Cast(Right(Route_No,3) as integer))+1,1) as RouteSN " & _
                      "From ORT01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'R'"
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            strRouteNo = "R" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
            tmp_Rs.Close
            
            If Len(strRouteNosum) = 0 Then strRouteNosum = "'" & strRouteNo & "'" Else strRouteNosum = strRouteNosum & ",'" & strRouteNo & "'"
            
            '3.Insert into ORT01T ���u�s���D��
            '  ORT01T.EXE_CONFIRM = '0' �s���͸��u�s���A�|���^�ǹL exe
            str_SQL = "Insert into ORT01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
                      strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
                      txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '4.insert into ORT05T �����i�X�޲z
            str_SQL = "Insert into ORT05T (Route_No,Vehicle_ID_No,Drive_Times,Delivery_Date,Valid_Vehicle,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Dock_No,Expect_Time,Expect_Date) Values ('" & _
                      strRouteNo & "','" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",'" & _
                      Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "','1'," & _
                      txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'" & _
                      txt_Tab0_DockNo.Text & "','" & txt_Tab0_CarCheckInTime.Text & "','" & txt_Tab0_CarCheckInDate.Text & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�Ѩ����D�ɧ�s�����������
            str_SQL = "Update ORT05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
                      "From ORT05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and Route_No = '" & strRouteNo & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
        End If
        
    '5.insert into ORT02T [�ƨ��q����]
    '  �g�� SSTab1.Tab 1 [���u�s�����q��W�Ӫ�]
    blTab0SelectedOrderEventEnable = False
    
    str_SQL = "Insert into ORT02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
              " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
              "Select '" & strRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
              " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
              "From ORT02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�g�J�ѦҸ��u�s����orders.containertype
    str_SQL = "Update orders Set containertype = '" & Trim(txt_Tab0_Route.Text) & "' , trafficCop = null Where orderkey = '" & Left(rs_Tab0_SelectedOrders("�q��s��"), 10) & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rs_Tab0_SelectedOrders.MoveNext
    
Loop
    
    
    '6. update trp01t,trp05t�A
    str_SQL = "update ORT01T set WEIGHT=(select sum(ORT02T.WEIGHT) from ORT02T where ORT02T.ROUTE_NO=ORT01T.ROUTE_NO) " & _
        "where   route_no in ( " & strRouteNosum & ") " & _
        "update ORT01T set CASE_CNT=(select sum(ORT02T.CASE_CNT) from ORT02T where ORT02T.ROUTE_NO=ORT01T.ROUTE_NO) " & _
        "where  route_no in ( " & strRouteNosum & ") " & _
        "update ORT01T set Pallet_Qty=(select sum(ORT02T.Pallet_Qty) from ORT02T where ORT02T.ROUTE_NO=ORT01T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") " & _
        "update ORT01T set VOLUMN_WEIGHT=(select sum(ORT02T.VOLUMN_WEIGHT) from ORT02T where ORT02T.ROUTE_NO=ORT01T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "update ORT05T set WEIGHT=(select sum(ORT02T.WEIGHT) from ORT02T where ORT02T.ROUTE_NO=ORT05T.ROUTE_NO) " & _
        "where   route_no in ( " & strRouteNosum & ") " & _
        "update ORT05T set CASE_CNT=(select sum(ORT02T.CASE_CNT) from ORT02T where ORT02T.ROUTE_NO=ORT05T.ROUTE_NO) " & _
        "where  route_no in ( " & strRouteNosum & ") " & _
        "update ORT05T set Pallet_Qty=(select sum(ORT02T.Pallet_Qty) from ORT02T where ORT02T.ROUTE_NO=ORT05T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") " & _
        "update ORT05T set VOLUMN_WEIGHT=(select sum(ORT02T.VOLUMN_WEIGHT) from ORT02T where ORT02T.ROUTE_NO=ORT05T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    
    cn.CommitTrans
    Tran_Level = 0

    If dg_Tab1_Route.SelBookmarks.Count > 0 Then
        dg_Tab1_Route.SelBookmarks.Remove 0
    End If
'    dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
'    rs_Tab1_RouteOrders.Filter = " ���u�s�� = '" & strRouteNo & "' "
    blTab0SelectedOrderEventEnable = True
    
    '7.�� ORT02T Trigger [insert] �i��H�U�@�~
    '   a.�g�J ORT02T -- �ƨ��q�������
    '   b.�R�� ORT02W -- �ݱƨ��q�������
    '   c.�R�� ORT02W -- �ݱƨ��q��D��
    
    '8.�M�� [�w������ݱƨ��q��C��]
    blTab0SelectedOrderEventEnable = False
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w-ReSet
    Call CreateRS_Tab0_SelectedOrders
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    
    '6.�M���ƨ��@�~����
    txt_Tab0_DockNo.Text = ""               '�X�Y�Ȧs
    txt_Tab0_CarCheckInDate.Text = ""       '�����w�p������
    txt_Tab0_CarCheckInTime.Text = ""       '�����w�p����ɶ�
    txt_Tab0_TRPDate.Text = ""              '�X�����
    txt_Tab0_DeliveryCarNo.Text = ""        '���P���X
    txt_Tab0_DeliveryCompany.Text = ""      '�B�餽�q
    txt_Tab0_DeliveryDriver.Text = ""       '�r�p�H
    txt_Tab0_DeliveryPhone.Text = ""        '�q��
    txt_Tab0_DeliveryCarType.Text = ""      '����
    
    cmd_Tab0_CreateRouteByAds.Enabled = True
    
    
    '�ݱƨ��q���`�p��T
    Call ReCaculate_OrderSum
    
    SSTab1.Tab = 1
    DoEvents: DoEvents
    
    '�d�߱ƨ����G
    '�]�w���u�s���C��
    blTab1RouteEventEnable = False
    Call CreateRS_Tab1_Route
    blTab1RouteEventEnable = True
    '�]�w���u�s�����q��C��
    Call CreateRS_Tab1_RouteOrders
    
    str_SQL = "Select ���u�s��,�X�����,���P���X,����,�r�p�H,�c��,�O��,���n,���q,����,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,EXE�^��,�ƨ��� " & _
              "From ORTPlan_RouteData Where ���u�s�� in ( " & strRouteNosum & ") order by ���u�s��"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����(ORT01T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTab1RouteEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab1_Route.AddNew
        rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
        rs_Tab1_Route.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
        rs_Tab1_Route.Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
        rs_Tab1_Route.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
        rs_Tab1_Route.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab1_Route.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab1_Route.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab1_Route.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = tmp_Rs.Fields("�X�Y�Ȧs").Value
        rs_Tab1_Route.Fields("�w�p������").Value = tmp_Rs.Fields("�w�p������").Value
        rs_Tab1_Route.Fields("�w�p����ɶ�").Value = tmp_Rs.Fields("�w�p����ɶ�").Value
        rs_Tab1_Route.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_Tab1_Route.Fields("�ƨ���").Value = tmp_Rs.Fields("�ƨ���").Value
        rs_Tab1_Route.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab1_Route.MoveFirst
    blTab1RouteEventEnable = True
    tmp_Rs.Close
    'TRP03W
    str_SQL = "Select ���u�s��,���h��,�q��s��,ZIP,��f�Ȥ�²��,��f�Ȥ�a�},�c��,�O��,���n,���q,Receipt_No,EXE�^��,Area,�Ȥ�²��,���A" & _
              " From ORTPlan_RouteOrders " & _
               "Where ���u�s�� in ( " & strRouteNosum & ") Order by ���u�s��,Receipt_No"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���u�s�����q����(ORT02T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Do While Not tmp_Rs.EOF
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("�s��").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
        rs_Tab1_RouteOrders.Fields("���h��").Value = tmp_Rs.Fields("���h��").Value
        rs_Tab1_RouteOrders.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("��f�Ȥ�²��").Value = tmp_Rs.Fields("��f�Ȥ�²��").Value
        rs_Tab1_RouteOrders.Fields("��f�Ȥ�a�}").Value = tmp_Rs.Fields("��f�Ȥ�a�}").Value
        rs_Tab1_RouteOrders.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab1_RouteOrders.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab1_RouteOrders.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab1_RouteOrders.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value
        rs_Tab1_RouteOrders.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_Tab1_RouteOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab1_RouteOrders.MoveFirst
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err_Handle:
   If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
   End If
   
   '�D�J���~���ܡGlocal �� Recordset [���u�s���C��] ��ƥ����R��
   '�]�� [���u�s���C��] ���� DB connection.transaction ����
   blTab1RouteEventEnable = False
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_Route.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_RouteOrders.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_RouteOrders.EOF Then
      Do While Not rs_Tab1_RouteOrders.EOF
         rs_Tab1_RouteOrders.Delete
         rs_Tab1_RouteOrders.MoveFirst
      Loop
   End If
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
      
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-�̦a�}�إ߸��u�s��", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_CreateRouteByAds.Enabled = True

End Sub

Private Sub cmd_Tab0_ImportOrders_Click()
On Error GoTo err_Handle
Dim strReceiptNo As String
strReceiptNo = ""
    '��s�c�O�������
    If chk_Tab0_Updateortw.Value = 1 Then
        cn.Execute "exec gs_UpdateORTW", RowsAffect, adExecuteNoRecords
    End If
    
'    '��sOrders���
'    str_SQL = "update ort02w set ort02w.otqty = orders.otqty from ort02w join orders on ort02w.receipt_no = orders.orderkey and ort02w.OTConfirmuser is null and ort02w.OTQTY is null and orders.OTQTY is not null "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�ƨ��@�~>>�פJ�ݱƨ��q��
    Screen.MousePointer = vbHourglass
     DoEvents: DoEvents
    Set dg_TRP02W.DataSource = Nothing

    '�ƨ��@�~�G�ݱƨ��q��
    Call CreateRS_Tab0_TRP02W
    
    strSourceFilter = adFilterNone
    DoEvents
    
    '���w����q��̡G�߰� user �O�_�n�M��
    If rs_Tab0_SelectedOrders.RecordCount <> 0 Then
       msg_text = "���J�ݱƨ��q��G[�w����q��] �O�_�i��M��"
       If MsgBox(msg_text, vbYesNo + vbInformation + vbDefaultButton2, msg_title) = vbYes Then
          '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
          Call Clear_RouteData
          txt_Tab0_RouteNo.Text = ""
        Else
            dg_Tab0_SelectedOrders.Enabled = False
            rs_Tab0_SelectedOrders.MoveFirst
            Do While Not rs_Tab0_SelectedOrders.EOF
                strReceiptNo = strReceiptNo & rs_Tab0_SelectedOrders.Fields("Receipt_no") & "','"
                rs_Tab0_SelectedOrders.MoveNext
            Loop
            
            dg_Tab0_SelectedOrders.Enabled = True
       End If
    End If
    
    '�ݱƨ��q����J�G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '���^�ݱƨ��q��
    str_SQL = "Select Convert(varchar(8),a1.Arrive_Date,112) as ���h�� , Rtrim(a1.Receipt_No) + '(' + Rtrim(a1.StorerKey)+'-'+Rtrim(a1.Extern)+')' as �q��s�� , " & _
            "�q�����A =isnull(a2.channel_type,''),Isnull(Round(a1.Case_cnt,2),0) as �c�� ,  Isnull(Round(a1.Pallet_Qty,2),0) as �O�� , " & _
            "Isnull(Round(a1.Weight,2),0) as ���q , Isnull(Round(a1.Volumn_Weight,2),0) as ���n , Rtrim(a1.ConsigneeKey) as �Ȥ�s�� , " & _
            "case when a1.priority = 'A2B' then (select isnull(rtrim(zip),'x') from trp01m where storerkey = a1.storerkey and rtrim(consigneekey) = rtrim(a1.bconsigneekey)) else Isnull(Rtrim(a2.ZIP),'x') end as ZIP ,��f�Ȥ�²�� = isnull((select TRP01M.short_name from TRP01M join orders on TRP01M.consigneekey = orders.b_company and orders.storerkey = TRP01M.storerkey and len(rtrim(isnull(orders.b_company,''))) > 0 and orders.orderkey = a1.c_receipt_no ),'') , isnull(Rtrim(a2.Address),'x')   as ���f�a�} , Rtrim(Isnull(a2.Vehicle_Type,'x')) as ���� , " & _
            "Case When b2.Description = '�L�S��ݨD' Then 'X' else Rtrim(Isnull(b2.Description,'')) End as �S��ݨD1 , " & _
            "Case When b3.Description = '�L�S��ݨD' Then 'X' else Rtrim(Isnull(b3.Description,'')) End as �S��ݨD2 , " & _
            "Rtrim(Isnull(a1.Urgent_Mark,'')) as ��� ,Rtrim(Isnull(a1.Reserve_Mark,'')) as �M�� ,Rtrim(Isnull(a1.Cold_Mark,'')) as �N��  , " & _
            "Rtrim(a1.Receipt_No) as Receipt_No , Rtrim(a1.StorerKey) as �f�D , Convert(varchar(8),a1.Receipt_Date,112) as �q��� , " & _
            "Rtrim(Isnull(a1.Extern,'')) as �f�D�渹 , " & _
            "Case When Isnull(Rtrim(Cast(c1.Notes as varchar(300))),'') = '' Then 'X' else Rtrim(Cast(c1.Notes as varchar(300))) End as �q��Ƶ� ,�t�e�ܧO = isnull(c1.facility,''), " & _
            "case when a1.priority = 'A2B' then (select Isnull(Rtrim(Area_Code),'') from trp01m where storerkey = a1.storerkey and rtrim(consigneekey) = rtrim(a1.bconsigneekey)) else Isnull(Rtrim(a2.Area_Code),'') end as Area , Rtrim(a2.Short_Name) as �Ȥ�²�� , Rtrim(Isnull(a1.Priority,'')) as ���A,��f�a�} = isnull((select TRP01M.address from TRP01M join orders on TRP01M.consigneekey = orders.b_company and orders.storerkey = TRP01M.storerkey and len(rtrim(isnull(orders.b_company,''))) > 0 and orders.orderkey = a1.c_receipt_no ),''),Rtrim(Isnull(c1.Type,'')) as �q�����O " & _
            ",�ѦҸ��s = (select top 1 trp02t.route_no from trp02t trp02t where a1.storerkey = trp02t.storerkey and a1.ConsigneeKey = trp02t.ConsigneeKey and trp02t.route_no <> 'D' and convert(char(8),trp02t.arrive_date,112) > = convert(char(8),getdate(),112) order by trp02t.ROUTE_NO desc) " & _
            "From ORT02W a1 " & _
            "left outer join TRP01M a2 on a2.ConsigneeKey = a1.ConsigneeKey and a1.storerkey = a2.storerkey " & _
            "Left outer join TRP04M b2 on b2.Extra_Demand_Code = a2.Extra_Demand_Code " & _
            "Left outer join TRP04M b3 on b3.Extra_Demand_Code = a2.Extra_Demand_Code2 " & _
            "Left outer join Orders c1 on c1.OrderKey = a1.c_receipt_no " & _
            " where a1.receipt_no not in ('" & strReceiptNo & "')"

    strSourceOrderBy = " �q��s�� "
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
    'blORT02WEventEnable = False
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    Do While Not tmp_Rs.EOF
        rs_ORT02W.AddNew
        rs_ORT02W.Fields("�s��").Value = rs_ORT02W.RecordCount
        rs_ORT02W.Fields("�ѦҸ��s").Value = tmp_Rs("�ѦҸ��s") & ""
        rs_ORT02W.Fields("���h��").Value = tmp_Rs.Fields("���h��").Value
        rs_ORT02W.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
        rs_ORT02W.Fields("�q�����A").Value = tmp_Rs.Fields("�q�����A").Value
        rs_ORT02W.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_ORT02W.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_ORT02W.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_ORT02W.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_ORT02W.Fields("�Ȥ�s��").Value = tmp_Rs.Fields("�Ȥ�s��").Value
        rs_ORT02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value & ""
        rs_ORT02W.Fields("zip").Value = tmp_Rs.Fields("zip").Value & ""
        rs_ORT02W.Fields("��f�Ȥ�²��").Value = tmp_Rs.Fields("��f�Ȥ�²��") & IIf(tmp_Rs("���A") = "A2B", "", tmp_Rs("�t�e�ܧO"))
        rs_ORT02W.Fields("���f�a�}").Value = tmp_Rs.Fields("���f�a�}").Value
        rs_ORT02W.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value
        rs_ORT02W("�t�e�ܧO") = tmp_Rs.Fields("�t�e�ܧO")
        rs_ORT02W.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_ORT02W.Fields("�S��ݨD1").Value = tmp_Rs.Fields("�S��ݨD1").Value
        rs_ORT02W.Fields("�S��ݨD2").Value = tmp_Rs.Fields("�S��ݨD2").Value
        rs_ORT02W.Fields("���").Value = tmp_Rs.Fields("���").Value
        rs_ORT02W.Fields("�M��").Value = tmp_Rs.Fields("�M��").Value
        rs_ORT02W.Fields("�N��").Value = tmp_Rs.Fields("�N��").Value
        rs_ORT02W.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_ORT02W.Fields("�f�D�渹").Value = tmp_Rs.Fields("�f�D�渹").Value
        rs_ORT02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value & ""
        rs_ORT02W.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value & ""
        rs_ORT02W.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_ORT02W.Fields("��f�a�}").Value = tmp_Rs.Fields("��f�a�}").Value
        rs_ORT02W.Fields("�q�����O").Value = tmp_Rs.Fields("�q�����O").Value
        rs_ORT02W.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_ORT02W.MoveFirst
    dg_TRP02W.Visible = True
    'blORT02WEventEnable = True
    blTRP02WEventEnable = True
    
    '�ݱƨ��q���`�p��T
    Call Retrive_OrderSum
    
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
    CreateErrorLog Me.Name & "-�h�f�ƨ�-�פJ�ݱƨ��q��", Me.Caption, "cmd_Tab0_ImportOrders_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Query_Click()
    '�ƨ��@�~ >> �d��
    If Len(txt_Tab0_RouteNo.Text) = 0 Then Exit Sub
    If rs_Tab0_SelectedOrders.RecordCount <> 0 And blRouteModify = False Then
        '�s�W���u�s���Ҧ��G
        '�I�s [�w��q�沾��(��)] �ӳB�z�w�Q�Ȯɿ���� [�ݱƨ��q��] �٭�^ [�ݱƨ��q��]
        Call cmd_Tab0_SelectedRemove_All_Click
    End If
    If blRouteModify And blRouteChange Then
        '���ĸ��u�s�� & ��Ƥw�D���ʡA�n user �T�{�O�_�s��
        msg_text = "���u�s����ƬO�_�s�ɡH"
        If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
            '�I�s�s�ɵ{��
            Call cmd_Tab0_Save_Click
        Else
            '���s�ɡ��������s���J [�ݱƨ��q��] �w�٭� [���][����] �ާ@�� [�ݱƨ��q��] ���v�T
            Call cmd_Tab0_ImportOrders_Click
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
    Call Clear_RouteData
    
    '���o���s���
    str_SQL = "Select �X�����,���P���X,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,�B�餽�q,�r�p�H,�r�p�q��,����,�c��,�O��,���n,���q " & _
              "From TRPPlan_RouteQuery Where ���u�s�� = '" & txt_Tab0_RouteNo.Text & "'"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧱ƨ����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
        '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
        Call Clear_RouteData
        
        Screen.MousePointer = vbDefault
        txt_Tab0_RouteNo.SelStart = 0: txt_Tab0_RouteNo.SelLength = Len(txt_Tab0_RouteNo.Text)
        txt_Tab0_RouteNo.SetFocus
        Exit Sub
    End If
    txt_Tab0_TRPDate.Text = tmp_Rs.Fields("�X�����").Value
    txt_Tab0_DeliveryCarNo.Text = tmp_Rs.Fields("���P���X").Value
    txt_Tab0_DockNo.Text = tmp_Rs.Fields("�X�Y�Ȧs").Value
    txt_Tab0_CarCheckInDate.Text = tmp_Rs.Fields("�w�p������").Value
    txt_Tab0_CarCheckInTime.Text = tmp_Rs.Fields("�w�p����ɶ�").Value
    txt_Tab0_DeliveryCompany.Text = tmp_Rs.Fields("�B�餽�q").Value
    txt_Tab0_DeliveryDriver.Text = tmp_Rs.Fields("�r�p�H").Value
    txt_Tab0_DeliveryPhone.Text = tmp_Rs.Fields("�r�p�q��").Value
    txt_Tab0_DeliveryCarType.Text = tmp_Rs.Fields("����").Value
    txt_Tab0_Selected_Case.Text = tmp_Rs.Fields("�c��").Value
    txt_Tab0_Selected_Pallet.Text = tmp_Rs.Fields("�O��").Value
    txt_Tab0_Selected_Volumn.Text = tmp_Rs.Fields("���n").Value
    txt_Tab0_Selected_Weight.Text = tmp_Rs.Fields("���q").Value
    tmp_Rs.Close
    
    '���o���s�q��
    str_SQL = "Select ���h��,�q��s��,ZIP,Area,���A,�Ȥ�²��,�c��,�O��,���n,���q,����,�q��Ƶ�,�S��ݨD1,�S��ݨD2,Receipt_No,EXE�^��,�Ȥ�W�� " & _
              "From TRPPlan_RouteQueryOrders Where ���u�s�� = '" & txt_Tab0_RouteNo.Text & "' Order by Receipt_No "
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q��W�Ӹ��"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
        '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
        Call Clear_RouteData
        
        Screen.MousePointer = vbDefault
        txt_Tab0_RouteNo.SelStart = 0: txt_Tab0_RouteNo.SelLength = Len(txt_Tab0_RouteNo.Text)
        txt_Tab0_RouteNo.SetFocus
        Exit Sub
    End If
    blTab0SelectedOrderEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab0_SelectedOrders.AddNew
        rs_Tab0_SelectedOrders.Fields("�s��").Value = rs_Tab0_SelectedOrders.RecordCount
        rs_Tab0_SelectedOrders.Fields("���h��").Value = tmp_Rs.Fields("���h��").Value
        rs_Tab0_SelectedOrders.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
        rs_Tab0_SelectedOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value & ""
        rs_Tab0_SelectedOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value & ""
        rs_Tab0_SelectedOrders.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value
        rs_Tab0_SelectedOrders.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab0_SelectedOrders.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab0_SelectedOrders.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab0_SelectedOrders.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab0_SelectedOrders.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value
        rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value = tmp_Rs.Fields("�S��ݨD1").Value
        rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value = tmp_Rs.Fields("�S��ݨD2").Value
        rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab0_SelectedOrders.Fields("�Ȥ�W��").Value = tmp_Rs.Fields("�Ȥ�W��").Value
        rs_Tab0_SelectedOrders.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab0_SelectedOrders.MoveFirst
    rs_Tab0_SelectedOrders.Sort = " �s�� asc "
    blTab0SelectedOrderEventEnable = True
    tmp_Rs.Close
    blRouteModify = True
    blRouteChange = False
    strDispRouteNo = Trim(txt_Tab0_RouteNo.Text)
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��C��-�d��", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Remove_Click()
    '�ƨ��@�~ >> �� �w����q�����
    If rs_ORT02W Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Sub
    '�w����q��Y�L�ϥտ���GDisable �w��������ʧ@�A����~�R
    If dg_Tab0_SelectedOrders.SelBookmarks.Count = 0 Then Exit Sub

    blTab0SelectedOrderEventEnable = False

    '���������q��s�� Receipt_No
    Dim strReceiptNo As String
    strReceiptNo = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value

    '�N���R���� [�w����q��] �[�J [�ݱƨ��q��]
    Call SelectedOrders_Removeto_TRP02W(strReceiptNo)
    '���s���� [�ݱƨ��q��] �� [�s��] ����
    Call ReSet_TRP02W_SeqNo

    '�R���ϥտ�����q��G�w����q�泡��
    rs_Tab0_SelectedOrders.Delete
    If Not rs_Tab0_SelectedOrders.EOF Then rs_Tab0_SelectedOrders.MoveFirst
    If dg_Tab0_SelectedOrders.SelBookmarks.Count > 0 Then dg_Tab0_SelectedOrders.SelBookmarks.Remove 0
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders

    blTRP02WEventEnable = False
    rs_ORT02W.Filter = adFilterNone
    If strSourceFilter <> "0" Then rs_ORT02W.Filter = strSourceFilter
    If rs_ORT02W.EOF Then
       strSourceFilter = adFilterNone
       rs_ORT02W.Filter = adFilterNone
    End If
    rs_ORT02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    blTRP02WEventEnable = True
    blTab0SelectedOrderEventEnable = True

    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum


End Sub

Private Sub cmd_Tab0_Reserve_Click()
    '�ݱƨ��q��G�O�d�q��
    cmd_Tab0_Reserve.Enabled = False
    
    '�ݱƨ��q��G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
    blTRP02WEventEnable = False
    rs_ORT02W.Filter = adFilterNone
    rs_ORT02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    Dim strRouteNo As String, intDriveTimes As Integer, dbOrderCnt As Double, iLoop As Double
    strRouteNo = "D"   '�S����u�s���A�κީҦ��O�d�q��
    intDriveTimes = 1
    dbOrderCnt = 0
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    blTab2ReservedEventEnable = False
    '�z��w�����
    rs_ORT02W.Filter = "��='V'"
    If Not rs_ORT02W.EOF Then
        Do While Not rs_ORT02W.EOF
            rs_Tab2_ReservedOrders.AddNew
            For iLoop = 0 To rs_ORT02W.Fields.Count - 1
                rs_Tab2_ReservedOrders.Fields(iLoop).Value = rs_ORT02W.Fields(iLoop).Value
            Next iLoop
            rs_Tab2_ReservedOrders.Fields(1).Value = " "
            rs_Tab2_ReservedOrders.Update
            
            'insert into ORT02T
            str_SQL = "Insert into ORT02T (Route_No,StorerKey,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,description,c_receipt_no,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                      "Select '" & strRouteNo & "',StorerKey,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " 'D'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,description,c_receipt_no,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                      "From ORT02W Where Receipt_No = '" & rs_ORT02W.Fields("Receipt_No").Value & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            '����� TRP02T Trigger [insert] �i��H�U�@�~
            '   a.�g�J TRP03T -- �ƨ��q�������
            '   b.�R�� TRP03W -- �ݱƨ��q�������
            '   c.�R�� TRP02W -- �ݱƨ��q��D��
            
            rs_ORT02W.MoveNext
        Loop
        '[�ݿ���q��] ���A�R���w������q��
        rs_ORT02W.MoveFirst
        Do While Not rs_ORT02W.EOF
            rs_ORT02W.Delete
            rs_ORT02W.MoveFirst
        Loop
    End If
    
    '��s trp01t & trp05t �� [�c��] [�O��] [���q] [���n]
    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    blTab2ReservedEventEnable = True
    
    If strSourceFilter <> "0" Then rs_ORT02W.Filter = strSourceFilter
    If rs_ORT02W.EOF Then
        strSourceFilter = adFilterNone
        rs_ORT02W.Filter = adFilterNone
    End If
    rs_ORT02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    '�����ϥտ�����A
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    blTRP02WEventEnable = True
    cmd_Tab0_Reserve.Enabled = True
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-�إ߸��u�s��", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_Reserve.Enabled = True
End Sub

Private Sub cmd_Tab0_Save_Click()
    '�ƨ��@�~ >> �u�s���ק�Ҧ��s��
    If blRouteModify = False Then
        msg_text = "�D�g [�d��] �{�ǩ���ܤ����� [���u�s��]"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    End If
    If blRouteChange = False Then
        msg_text = "[���u�s��] ����ƨå����ʡA�������� [�s��] �{��"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    Else
       '�q���Ʀ����ʡA�B�������Q�����A���P�R��
        If rs_Tab0_SelectedOrders.RecordCount = 0 Then
            msg_text = "�����u�s���ثe�w�L�q��A�O�_�R�������s�H"
            If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
                Call Delete_RouteNo(strDispRouteNo)
                Call Clear_RouteData
                txt_Tab0_RouteNo.Text = ""
                Exit Sub
            End If
        End If
    End If
    '�ˮָ��u�s����ƬO�_���T��J
    If RouteData_Check = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    cmd_Tab0_Save.Enabled = False
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Dim intDriveTimes As Integer    '����
    '1.�T�{ [�X�����] �P [���P���X] [�ק��v��] & [��ƬO�_�D����]
    '  �Y���ʫh�������s�p�⨮��
    str_SQL = "Select Rtrim(t05t.Vehicle_ID_No) as ���P���X,Convert(varchar(8),t01t.Delivery_Date,112) as �X�����,Rtrim(Isnull(t01t.AddWho,'')) as AddWho,t05t.Drive_Times as ���� " & _
              "From TRP05T t05t inner join TRP01T t01t on t01t.Route_No = t05t.Route_No " & _
              "Where t05t.Route_No = '" & strDispRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "���u�s�� [" & strDispRouteNo & "] �w�䤣���ƤF"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl = True Then
        tmp_Rs.Close
        msg_text = "�v�����ޡG���u�s�����ק�u���\�ѭ�Ʃw�̰���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    intDriveTimes = tmp_Rs.Fields("����").Value
    If tmp_Rs.Fields("�X�����").Value <> txt_Tab0_TRPDate.Text Or UCase(tmp_Rs.Fields("���P���X").Value) <> txt_Tab0_DeliveryCarNo.Text Then
        '�X����� or ���P���X�D���ʡG�������s�p�⨮��
        tmp_Rs.Close
        str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                  "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    End If
    tmp_Rs.Close
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '2.��s TRP05T & TRP01T & TRP03T
    str_SQL = "Update TRP01T Set Delivery_Date = '" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "' " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP05T Set Delivery_Date = '" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "', " & _
              "   Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & ",Dock_No = '" & txt_Tab0_DockNo.Text & "',Expect_Date = '" & txt_Tab0_CarCheckInDate.Text & "'," & _
              "   Expect_Time = '" & txt_Tab0_CarCheckInTime.Text & "' " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP03T Set Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '3.�Ѩ����D�ɧ�s TRP05T �����������
    str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strDispRouteNo & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '4..�N TRP02T ������s�Хܬ� [��s�X��] DeleteFlag = '1'
    str_SQL = "Update TRP02T Set DeleteFlag='1' Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '5.�N�q���s�X�� DeleteFalg �٭�^ 0
    '  �䤣�쪺�A��ܬO�s�[�J���A�i��s�W�{��
    blTab0SelectedOrderEventEnable = False
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        str_SQL = "Update TRP02T Set DeleteFlag='0' Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        If RowsAffect = 0 Then
            '�s�W�q��
            str_SQL = "Insert into TRP02T (Route_No,StorerKey,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                      "Select '" & strDispRouteNo & "',StorerKey,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                      "From TRP02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
        End If
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    blTab0SelectedOrderEventEnable = True
    
    '6.�N�����q���٭�^ TRP02W & TRP03W
    '(1).�N TRP03T �g�^ TRP03W >> �R�� TRP03T
    str_SQL = "Insert into TRP03W(" & _
              "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From TRP03T A INNER JOIN TRP02T B ON B.Receipt_No = a.Receipt_No and b.DeleteFlag = '1' and b.Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).�N TRP02T �g�^ TRP02W >> �R�� TRP02T
    str_SQL = "Insert into TRP02W(" & _
              "   RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
              "Select RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
              "From TRP02T Where DeleteFlag='1' and Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(3).�R�� TRP02T & TRP03T
    str_SQL = "Delete TRP03T FROM TRP02T Where TRP02T.Receipt_No = TRP03T.Receipt_No and TRP02T.DeleteFlag='1' and TRP02T.Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Delete From TRP02T Where DeleteFlag='1' and Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '7.��s TRP01T & TRP05T ���έp����
    str_SQL = "exec  ReservedOrders_Recalculate " & strDispRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
       cn.CommitTrans
       Tran_Level = 0
    End If
    
    '�M���ù�����
    Call Clear_RouteData
    txt_Tab0_RouteNo.Text = ""
    cmd_Tab0_Save.Enabled = True
    
    '�ݱƨ��q���`�p��T
    Call Retrive_OrderSum
    
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
   CreateErrorLog Me.Name & "-�ƨ��@�~-���u�s���ק�s��", Me.Caption, "cmd_Tab0_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_SelectCar_Click()
    '�ƨ��@�~ >> �q�����
    If Len(txt_Tab0_TRPDate.Text) = 0 Then
        msg_text = "�Х���J�G�X�����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SetFocus
        Exit Sub
    Else
        If chk_Tab0_DriveTimes.Value = vbChecked Then
            '��ܹB�e�����ݿ�M��--�]�t�w�Ʃw������������
            Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar.Name & "1")
        Else
            '��ܹB�e�����ݿ�M��--����ܨ������
            Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar.Name & "2")
        End If
    End If
End Sub

Private Sub cmd_Tab0_Selected_Click()
    '�ݱƨ��q��G���
    
    '�ݱƨ��q��G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    If Len(Trim(rs_ORT02W("�ѦҸ��s"))) > 0 Then txt_Tab0_Route.Text = Trim(rs_ORT02W("�ѦҸ��s"))
'    '���X�۹����ƨ���ƶ�J
'        str_SQL = "Select �X�����,���P���X,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,�B�餽�q = t9m.trp_company_code,�r�p�H,�r�p�q��,���� = t9m.vehicle_type " & _
'              "From TRPPlan_RouteQuery t join trp09m t9m on ���P���X = t9m.vehicle_id_no Where ���u�s�� = '" & rs_ORT02W("�ѦҸ��s") & "'"
'        Dim rsTmp As New ADODB.Recordset
'        rsTmp.Open str_SQL, cn
'        If rsTmp.EOF = 0 Then
'        txt_Tab0_TRPDate = rsTmp("�X�����")
'        txt_Tab0_DeliveryCarNo = rsTmp("���P���X")
'        txt_Tab0_DeliveryCompany = rsTmp("�B�餽�q")
'        txt_Tab0_DeliveryDriver = rsTmp("�r�p�H")
'        txt_Tab0_DeliveryPhone = rsTmp("�r�p�q��")
'        txt_Tab0_DeliveryCarType = rsTmp("����") & ""
'        txt_Tab0_DockNo = rsTmp("�X�Y�Ȧs")
'        txt_Tab0_CarCheckInDate = rsTmp("�w�p������")
'        txt_Tab0_CarCheckInTime = rsTmp("�w�p����ɶ�")
'        End If
'        rsTmp.Close: Set rsTmp = Nothing
'    End If
    
    '�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
    blTRP02WEventEnable = False
    rs_ORT02W.Filter = adFilterNone
    rs_ORT02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    '�z��w�����
    rs_ORT02W.Filter = "��='V'"
    If Not rs_ORT02W.EOF Then
        dg_Tab0_SelectedOrders.Visible = False
        blTab0SelectedOrderEventEnable = False
        Do While Not rs_ORT02W.EOF
            '�P�_�O�_�w�g����L
            rs_Tab0_SelectedOrders.Filter = adFilterNone
            rs_Tab0_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
            rs_Tab0_SelectedOrders.Filter = "Receipt_No = '" & rs_ORT02W.Fields("Receipt_No").Value & "'"
            '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
            If blRouteModify Then blRouteChange = True
            If rs_Tab0_SelectedOrders.EOF Then
                '�s�W������q��
                rs_Tab0_SelectedOrders.AddNew
                rs_Tab0_SelectedOrders.Fields("�s��").Value = 999
                rs_Tab0_SelectedOrders.Fields("���h��").Value = rs_ORT02W.Fields("���h��").Value
                rs_Tab0_SelectedOrders.Fields("�q��s��").Value = rs_ORT02W.Fields("�q��s��").Value
                rs_Tab0_SelectedOrders.Fields("�q�����A").Value = rs_ORT02W.Fields("�q�����A").Value
                rs_Tab0_SelectedOrders.Fields("ZIP").Value = rs_ORT02W.Fields("ZIP").Value
                rs_Tab0_SelectedOrders.Fields("Area").Value = rs_ORT02W.Fields("Area").Value
                rs_Tab0_SelectedOrders.Fields("���A").Value = rs_ORT02W.Fields("���A").Value
                rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value = rs_ORT02W.Fields("�Ȥ�²��").Value
                rs_Tab0_SelectedOrders.Fields("���f�a�}").Value = rs_ORT02W.Fields("���f�a�}").Value
                rs_Tab0_SelectedOrders.Fields("�c��").Value = rs_ORT02W.Fields("�c��").Value
                rs_Tab0_SelectedOrders.Fields("�O��").Value = rs_ORT02W.Fields("�O��").Value
                rs_Tab0_SelectedOrders.Fields("���n").Value = rs_ORT02W.Fields("���n").Value
                rs_Tab0_SelectedOrders.Fields("���q").Value = rs_ORT02W.Fields("���q").Value
                rs_Tab0_SelectedOrders.Fields("����").Value = rs_ORT02W.Fields("����").Value
                rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = rs_ORT02W.Fields("�q��Ƶ�").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value = rs_ORT02W.Fields("�S��ݨD1").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value = rs_ORT02W.Fields("�S��ݨD2").Value
                rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = rs_ORT02W.Fields("�q��Ƶ�").Value
                rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = rs_ORT02W.Fields("Receipt_No").Value
                rs_Tab0_SelectedOrders.Fields("��f�Ȥ�²��").Value = rs_ORT02W.Fields("��f�Ȥ�²��").Value
                
                'Terry 20191107 �]�̦a�}�ո��s�\��ݱƧ� �N�DA2B�q�檺��f�a�}����J���f�a�}��쪺��(�]�DA2B�q��S��f�a�}�A�ҥH�i�H�o�˩�)
                If rs_ORT02W.Fields("���A").Value = "A2B" Then
                    rs_Tab0_SelectedOrders.Fields("��f�a�}").Value = rs_ORT02W.Fields("��f�a�}").Value
                Else
                    rs_Tab0_SelectedOrders.Fields("��f�a�}").Value = rs_ORT02W.Fields("���f�a�}").Value
                End If
                
                rs_Tab0_SelectedOrders.Fields("�ѦҸ��s").Value = rs_ORT02W.Fields("�ѦҸ��s").Value
                
                rs_Tab0_SelectedOrders.Update
            Else
                '��s������q����
                rs_Tab0_SelectedOrders.Fields("���h��").Value = rs_ORT02W.Fields("���h��").Value
                rs_Tab0_SelectedOrders.Fields("�q��s��").Value = rs_ORT02W.Fields("�q��s��").Value
                rs_Tab0_SelectedOrders.Fields("�q�����A").Value = rs_ORT02W.Fields("�q�����A").Value
                rs_Tab0_SelectedOrders.Fields("ZIP").Value = rs_ORT02W.Fields("ZIP").Value
                rs_Tab0_SelectedOrders.Fields("Area").Value = rs_ORT02W.Fields("Area").Value
                rs_Tab0_SelectedOrders.Fields("���A").Value = rs_ORT02W.Fields("���A").Value
                rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value = rs_ORT02W.Fields("�Ȥ�²��").Value
                rs_Tab0_SelectedOrders.Fields("���f�a�}").Value = rs_ORT02W.Fields("���f�a�}").Value
                rs_Tab0_SelectedOrders.Fields("�c��").Value = rs_ORT02W.Fields("�c��").Value
                rs_Tab0_SelectedOrders.Fields("�O��").Value = rs_ORT02W.Fields("�O��").Value
                rs_Tab0_SelectedOrders.Fields("���n").Value = rs_ORT02W.Fields("���n").Value
                rs_Tab0_SelectedOrders.Fields("���q").Value = rs_ORT02W.Fields("���q").Value
                rs_Tab0_SelectedOrders.Fields("����").Value = rs_ORT02W.Fields("����").Value
                rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = rs_ORT02W.Fields("�q��Ƶ�").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value = rs_ORT02W.Fields("�S��ݨD1").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value = rs_ORT02W.Fields("�S��ݨD2").Value
                rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = rs_ORT02W.Fields("Receipt_No").Value
                rs_Tab0_SelectedOrders.Fields("��f�Ȥ�²��").Value = rs_ORT02W.Fields("��f�Ȥ�²��").Value
                
                'Terry 20191107 �]�̦a�}�ո��s�\��ݱƧ� �N�DA2B�q�檺��f�a�}����J���f�a�}��쪺�� (�]�DA2B�q��S��f�a�}�A�ҥH�i�H�o�˩�)
                If rs_ORT02W.Fields("���A").Value = "A2B" Then
                    rs_Tab0_SelectedOrders.Fields("��f�a�}").Value = rs_ORT02W.Fields("��f�a�}").Value
                Else
                    rs_Tab0_SelectedOrders.Fields("��f�a�}").Value = rs_ORT02W.Fields("���f�a�}").Value
                End If
                
                rs_Tab0_SelectedOrders.Fields("�q�����O").Value = rs_ORT02W.Fields("�q�����O").Value
                rs_Tab0_SelectedOrders.Fields("�ѦҸ��s").Value = rs_ORT02W.Fields("�ѦҸ��s").Value
            End If
            rs_ORT02W.MoveNext
        Loop
        '���s�� [�w����q��] ���� [�s��] �P������Ʋέp�G�c�ơA�O�ơA���n�A���q
        Call Calculate_SelectedOrders
        dg_Tab0_SelectedOrders.Visible = True
        blTab0SelectedOrderEventEnable = True
        
        '[�ݿ���q��] ���A�R���w������q��
        rs_ORT02W.MoveFirst
        Do While Not rs_ORT02W.EOF
            rs_ORT02W.Delete
            rs_ORT02W.MoveFirst
        Loop
    End If
    
    If strSourceFilter <> "0" Then rs_ORT02W.Filter = strSourceFilter
    If rs_ORT02W.EOF Then
        rs_ORT02W.Filter = adFilterNone
        strSourceFilter = adFilterNone
    End If
    rs_ORT02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    '�����ϥտ�����A
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    
    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum
    
    blTRP02WEventEnable = True

End Sub

Private Sub cmd_Tab0_SelectedCancel_All_Click()
    '�ƨ��@�~ >> X�ݿ��������
    
    '�ݱƨ��q��G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
    blTRP02WEventEnable = False
    rs_ORT02W.Filter = adFilterNone
    rs_ORT02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    '�z��w�����
    rs_ORT02W.Filter = "��='V'"
    If Not rs_ORT02W.EOF Then
        Do While Not rs_ORT02W.EOF
            rs_ORT02W.Fields("��").Value = " "
            rs_ORT02W.MoveNext
        Loop
    End If
    
    blTRP02WEventEnable = False
    If strSourceFilter <> "0" Then rs_ORT02W.Filter = strSourceFilter
    If rs_ORT02W.EOF Then
        strSourceFilter = adFilterNone
        rs_ORT02W.Filter = adFilterNone
    End If
    rs_ORT02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    blTRP02WEventEnable = True
    
    '�����ϥտ�����A
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    '�٭� [�ݱƨ��q��] �Ƨǳ]�w
    blTRP02WEventEnable = True
End Sub

Private Sub cmd_Tab0_SelectedCancel_Click()
    '�ƨ��@�~ >> X�ݿ����
    If rs_ORT02W Is Nothing Then Exit Sub
        '�ݿ���q��Y�L�ϥտ���GDisable �ݿ�����A����~�R
        If dg_TRP02W.SelBookmarks.Count = 0 Then Exit Sub
        
        If Trim(rs_ORT02W.Fields(1).Value) = "V" Then
        rs_ORT02W.Fields(1).Value = " "
        dbSelectedCount = dbSelectedCount - 1
        '�ݿ�w��G����p�p��s
        If dbSelectedCount = 0 Then
            dbsrcSelected_Case = 0
            dbsrcSelected_Pallet = 0
            dbsrcSelected_Volumn = 0
            dbsrcSelected_Weight = 0
            txt_Tab0_srcSelected_Case.Text = 0: txt_Tab0_srcSelected_Pallet.Text = 0
            txt_Tab0_srcSelected_Volumn.Text = 0: txt_Tab0_srcSelected_Weight.Text = 0
        Else
            dbsrcSelected_Case = dbsrcSelected_Case - rs_ORT02W.Fields("�c��").Value
            dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_ORT02W.Fields("�O��").Value
            dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_ORT02W.Fields("���n").Value
            dbsrcSelected_Weight = dbsrcSelected_Weight - rs_ORT02W.Fields("���q").Value
            txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
            txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
        End If
        '�����ϥտ�����A
        If dg_TRP02W.SelBookmarks.Count > 0 Then
            dg_TRP02W.SelBookmarks.Remove 0
        End If
    End If

End Sub


Private Sub cmd_Tab0_SelectedRemove_All_Click()
    '�ƨ��@�~ >> �� �w����q�����-����
    If rs_ORT02W Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then Exit Sub
    '���u�s���d�ߡG���ĸ��u�s��
    '���U [�w��q�沾��(��) ���P��R�����u�s��
    If blRouteModify Then
        msg_text = "�T�w�n�R�������u�s�� [" & txt_Tab0_RouteNo.Text & "]"
        If MsgBox(msg_text, vbCritical + vbOKCancel, msg_title) = vbOK Then
            '�R�����w���u�s��
            Call Delete_RouteNo(strDispRouteNo)
            '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
            Call Clear_RouteData
            txt_Tab0_RouteNo.Text = ""
        End If
        Exit Sub
    End If
    
    blTab0SelectedOrderEventEnable = False
    
    '���������q��s�� Receipt_No
    Dim strReceiptNo As String
    '�v���g�^ [�ݱƨ��q�� TRP02W]
    rs_Tab0_SelectedOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        strReceiptNo = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
        '�N���R���� [�w����q��] �[�J [�ݱƨ��q��]
        Call SelectedOrders_Removeto_TRP02W(strReceiptNo)
        rs_Tab0_SelectedOrders.MoveNext
    Loop
       
    '���s���� [�ݱƨ��q��] �� [�s��] ����
    Call ReSet_TRP02W_SeqNo
    
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w-ReSet
    Call CreateRS_Tab0_SelectedOrders
    
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    '�ƧǤ覡
    
    blTRP02WEventEnable = False
    rs_ORT02W.Filter = adFilterNone
    If strSourceFilter <> "0" Then rs_ORT02W.Filter = strSourceFilter
    If rs_ORT02W.EOF Then
        strSourceFilter = adFilterNone
        rs_ORT02W.Filter = adFilterNone
    End If
    rs_ORT02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    blTRP02WEventEnable = True
    
    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum
    
    blTab0SelectedOrderEventEnable = True
End Sub

Private Sub cmd_Tab0_srcOrderReset_Click()
    '�ƨ��@�~ >> �����ݱƨ��q��z��Ƨ�
    If rs_ORT02W Is Nothing Then Exit Sub
    '�����z�����A���]�ƧǨ̾�
     blTRP02WEventEnable = False
    '�z��w����̡G�������
    rs_ORT02W.Filter = "��='V'"
    If Not rs_ORT02W.EOF Then
        Do While Not rs_ORT02W.EOF
            rs_ORT02W.Fields(1).Value = " "
            rs_ORT02W.MoveNext
        Loop
    End If
    rs_ORT02W.Filter = adFilterNone
    strSourceFilter = adFilterNone
     'rs_ORT02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_ORT02W.Sort = strSourceOrderBy
    blTRP02WEventEnable = True
    
    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum

End Sub

Private Sub cmd_Tab1_RouteNoDelete_Click()
    '���u�s���C�� >> ���u�s���R��
    If rs_Tab1_Route.RecordCount = 0 Then Exit Sub
    If dg_Tab1_Route.SelBookmarks.Count = 0 Then Exit Sub
    
    Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
    strDeleteRouteNo = Trim(rs_Tab1_Route.Fields("���u�s��").Value)
    strCarno = Trim(rs_Tab1_Route.Fields("���P���X").Value)
    dbDriveTimes = Trim(rs_Tab1_Route.Fields("����").Value)
    
    '���R�������s�G�O�_�w�X���T�{
    Call Confirm_Recordset_Closed(tmp_Rs)
    'str_SQL = "Select c_Route_No  From SDN01T Where c_Route_No = '" & strDeleteRouteNo & "'"
    'Terry 20191127 �אּ�ˬd�X�����A
    str_SQL = "Select Route_No  From ORT05T Where Route_No = '" & strDeleteRouteNo & "' and sdnstatus = '1' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�`�N�G�����u�s���w�X���T�{�A�L�k�R��! "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    tmp_Rs.Close
    
    '���R�������s�G�O�_�w�������� Add by Terry 20191127
    str_SQL = "Select Route_No  From SDN02W Where Route_No = '" & strDeleteRouteNo & "' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�`�N�G�����u�s���w�������աA�L�k�R��! "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    tmp_Rs.Close
    
    
    msg_text = "�T�{�R�����u�s���G" & strDeleteRouteNo
    If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    '���ұ��R�������s�A�ƨ��̬O�_�����ɵn�J���ϥΪ�
    str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From ORT01T Where Route_No = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "��Ʋ��`�G�䤣����R�������u�s��"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    Else
        If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl = True Then
            tmp_Rs.Close
            msg_text = "�v�����ޡG���u�s�����R���u���\�ѭ�ƨ��̰���"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Exit Sub
        End If
    End If
    tmp_Rs.Close
    
    '���R�������s�G��������B���ܮɶ��O�_�w�n��
    str_SQL = "Select EXE_CONFIRM  From ORT01T Where Route_No = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields("EXE_CONFIRM").Value = "1" Or tmp_Rs.Fields("EXE_CONFIRM").Value = "2" Then
        tmp_Rs.Close
        msg_text = "��Ʋ��`�G�����u�s���w�^��ids "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    tmp_Rs.Close
    
    '�R�����s
    Call Delete_RouteNo(strDeleteRouteNo)
    
    '�R���d�ߵ��G���ӵ����u�s��--rs_Tab1_RouteOrders
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_Tab1_RouteOrders.Filter = "���u�s��='" & strDeleteRouteNo & "'"
    If Not rs_Tab1_RouteOrders.EOF Then
        Do While Not rs_Tab1_RouteOrders.EOF
        
        '�R���f�t�ѦҸ��s�渹
        str_SQL = "update orders set containertype = '',trafficCop=null where orderkey ='" & Left(rs_Tab1_RouteOrders("�q��s��"), 10) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
     
        rs_Tab1_RouteOrders.Delete
        rs_Tab1_RouteOrders.MoveFirst
        Loop
    End If
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    '(7).�R���d�ߵ��G���ӵ����u�s��--rs_Tab1_Route
    rs_Tab1_Route.Delete
    If Not rs_Tab1_Route.EOF Then rs_Tab1_Route.MoveFirst
    
    blTab1RouteEventEnable = True
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���u�s���C��-���u�s���R��", Me.Caption, "cmd_Tab1_RouteNoDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_RouteNoQuery_Click()
    '���u�s���C�� >> ���u�s���d��
    If Len(Trim(txt_Tab1_RouteNo.Text)) = 0 Then MsgBox "�п�J���u�s���I", vbOKOnly, "���u�s���d��": Exit Sub
    
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    
    '�]�w���u�s���C��
    blTab1RouteEventEnable = False
    Call CreateRS_Tab1_Route
    blTab1RouteEventEnable = True
    '�]�w���u�s�����q��C��
    Call CreateRS_Tab1_RouteOrders
    
    str_SQL = "Select ���u�s��,�X�����,���P���X,����,�r�p�H,�c��,�O��,���n,���q,����,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,EXE�^��,�ƨ��� " & _
              "From ORTPlan_RouteData Where ���u�s�� like '%" & txt_Tab1_RouteNo.Text & "%' order by ���u�s��"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����(ORT01T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTab1RouteEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab1_Route.AddNew
        rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
        rs_Tab1_Route.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
        rs_Tab1_Route.Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
        rs_Tab1_Route.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
        rs_Tab1_Route.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab1_Route.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab1_Route.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab1_Route.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = tmp_Rs.Fields("�X�Y�Ȧs").Value
        rs_Tab1_Route.Fields("�w�p������").Value = tmp_Rs.Fields("�w�p������").Value
        rs_Tab1_Route.Fields("�w�p����ɶ�").Value = tmp_Rs.Fields("�w�p����ɶ�").Value
        rs_Tab1_Route.Fields("�ƨ���").Value = tmp_Rs.Fields("�ƨ���").Value
        rs_Tab1_Route.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab1_Route.MoveFirst
    blTab1RouteEventEnable = True
    tmp_Rs.Close
    
    'TRP03W
    str_SQL = "Select ���u�s��,���h��,�q��s��,ZIP,��f�Ȥ�²��,�c��,�O��,���n,���q,Receipt_No,EXE�^��,Area,���A,�Ȥ�²�� " & _
              "From ORTPlan_RouteOrders " & _
               "Where ���u�s�� like '%" & txt_Tab1_RouteNo.Text & "%' Order by ���u�s��,Receipt_No"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���u�s�����q����(ORT02T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Do While Not tmp_Rs.EOF
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("�s��").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("���u�s��").Value = tmp_Rs("���u�s��")
        rs_Tab1_RouteOrders.Fields("���h��").Value = tmp_Rs("���h��")
        rs_Tab1_RouteOrders.Fields("�q��s��").Value = tmp_Rs("�q��s��")
        rs_Tab1_RouteOrders.Fields("ZIP").Value = tmp_Rs("ZIP") & ""
        rs_Tab1_RouteOrders.Fields("��f�Ȥ�²��").Value = tmp_Rs("��f�Ȥ�²��") & ""
        rs_Tab1_RouteOrders.Fields("�c��").Value = tmp_Rs("�c��")
        rs_Tab1_RouteOrders.Fields("�O��").Value = tmp_Rs("�O��")
        rs_Tab1_RouteOrders.Fields("���n").Value = tmp_Rs("���n")
        rs_Tab1_RouteOrders.Fields("���q").Value = tmp_Rs("���q")
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = tmp_Rs("Receipt_No")
        rs_Tab1_RouteOrders.Fields("Area").Value = tmp_Rs("Area") & ""
        rs_Tab1_RouteOrders.Fields("���A").Value = tmp_Rs("���A")
        rs_Tab1_RouteOrders.Fields("�Ȥ�²��").Value = tmp_Rs("�Ȥ�²��")
        rs_Tab1_RouteOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab1_RouteOrders.MoveFirst
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���u�s���C��-���u�s���d��", Me.Caption, "cmd_Tab1_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Delete_Click()
        '�O�d�q�� >> ���� [�ݱƨ��q��]
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    DelRecord = MsgBox("�R�����ƵL�k�_��,�T�w�n�R��? ", vbQuestion + vbYesNo, "�R��")
    If DelRecord = vbNo Then
        Exit Sub
    End If
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    blTab2ReservedEventEnable = False
    blTRP02WEventEnable = False
    cmd_Tab2_Delete.Enabled = False
    
    '�z��w�����
    rs_Tab2_ReservedOrders.Filter = "��='V'"
    If rs_Tab2_ReservedOrders.EOF Then
        rs_Tab2_ReservedOrders.Filter = adFilterNone
        rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
        blTab2ReservedEventEnable = True
        blTRP02WEventEnable = True
        cmd_Tab2_Remove.Enabled = True
        Exit Sub
    End If
    
    Dim strRouteNo As String, iLoop As Double
    strRouteNo = "D"
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        str_SQL = "delete  TRP02T where Extern ='" & rs_Tab2_ReservedOrders.Fields("�f�D�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP03T where Extern ='" & rs_Tab2_ReservedOrders.Fields("�f�D�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP02W where Extern ='" & rs_Tab2_ReservedOrders.Fields("�f�D�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP03W where Extern ='" & rs_Tab2_ReservedOrders.Fields("�f�D�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP02W_TEMP where Extern ='" & rs_Tab2_ReservedOrders.Fields("�f�D�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "update orders set B_PHONE2='00',trafficCop=null,type='�R��'  where externorderkey='" & rs_Tab2_ReservedOrders.Fields("�f�D�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        rs_Tab2_ReservedOrders.MoveNext
    Loop
    
    '[�ݿ���q��] ���A�R���w������q��
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        rs_Tab2_ReservedOrders.Delete
        rs_Tab2_ReservedOrders.MoveFirst
    Loop
    
'    '��s trp01t & trp05t �� [�c��] [�O��] [���q] [���n]
'    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    
    If Not (rs_ORT02W Is Nothing) Then
        If strSourceFilter <> "0" Then rs_ORT02W.Filter = strSourceFilter
        If rs_ORT02W.EOF Then
            strSourceFilter = adFilterNone
            rs_ORT02W.Filter = adFilterNone
        End If
        rs_ORT02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    End If
    
    rs_Tab2_ReservedOrders.Filter = adFilterNone
    rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
    blTab2ReservedEventEnable = True
    blTRP02WEventEnable = True
    
    cmd_Tab2_Delete.Enabled = True
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�O�d�q��-���ܫݱƨ��q��C��", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   blTab2ReservedEventEnable = True
   blTRP02WEventEnable = True
   cmd_Tab2_Delete.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Remove_Click()
    '�O�d�q�� >> ���� [�ݱƨ��q��]
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    blTab2ReservedEventEnable = False
    blTRP02WEventEnable = False
    cmd_Tab2_Remove.Enabled = False
    
    '�z��w�����
    rs_Tab2_ReservedOrders.Filter = "��='V'"
    If rs_Tab2_ReservedOrders.EOF Then
        rs_Tab2_ReservedOrders.Filter = adFilterNone
        rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
        blTab2ReservedEventEnable = True
        blTRP02WEventEnable = True
        cmd_Tab2_Remove.Enabled = True
        Exit Sub
    End If
    
    Dim strRouteNo As String, iLoop As Double
    strRouteNo = "D"
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
       If Not (rs_ORT02W Is Nothing) Then
            rs_ORT02W.AddNew
            For iLoop = 0 To rs_Tab2_ReservedOrders.Fields.Count - 1
                rs_ORT02W.Fields(iLoop).Value = rs_Tab2_ReservedOrders.Fields(iLoop).Value
            Next iLoop
            rs_ORT02W.Fields(0).Value = 999
            rs_ORT02W.Fields(1).Value = " "
            rs_ORT02W.Update
       End If
       
       '(1).�N ORT03T �g�^ ORT03W >> �R�� ORT03T
       str_SQL = "Insert into ORT03W(" & _
                 "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
                 "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
                 "From ORT03T A Where a.Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("Receipt_No").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
       '(2).�N ORT02T �g�^ ORT02W >> �R�� ORT02T
       str_SQL = "Insert into ORT02W(" & _
                 "   RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
                 "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                 "Select RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
                 "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                 "From ORT02T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("Receipt_No").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
       '(3).�R�� TRP02T & TRP03T
       str_SQL = "Delete From ORT03T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("Receipt_No").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
       
       str_SQL = "Delete From ORT02T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("Receipt_No").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
       rs_Tab2_ReservedOrders.MoveNext
    Loop
    
    '[�ݿ���q��] ���A�R���w������q��
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        rs_Tab2_ReservedOrders.Delete
        rs_Tab2_ReservedOrders.MoveFirst
    Loop
    
    '��s trp01t & trp05t �� [�c��] [�O��] [���q] [���n]
    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    
    If Not (rs_ORT02W Is Nothing) Then
        If strSourceFilter <> "0" Then rs_ORT02W.Filter = strSourceFilter
        If rs_ORT02W.EOF Then
            strSourceFilter = adFilterNone
            rs_ORT02W.Filter = adFilterNone
        End If
        rs_ORT02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    End If
    
    rs_Tab2_ReservedOrders.Filter = adFilterNone
    rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
    blTab2ReservedEventEnable = True
    blTRP02WEventEnable = True
    
    cmd_Tab2_Remove.Enabled = True
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�O�d�q��-���ܫݱƨ��q��C��", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   blTab2ReservedEventEnable = True
   blTRP02WEventEnable = True
   cmd_Tab2_Remove.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_ShowAll_Click()
    '�ƨ��@�~>>��ܩҦ��O�d�q����
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    
    '�O�d�q��C��
    blTab2ReservedEventEnable = False
    Call CreateRS_Tab2_ReservedOrders
    DoEvents
    
    '���^�O�d�q����
    str_SQL = "Select ' ' as '��',�ѦҸ��s,���h��,�q��s��,�q�����A,�c��,�O��,���n,���q,�Ȥ�s��,ZIP,�Ȥ�²��,Area,���A,���f�a�},�q��Ƶ�,�t�e�ܧO,����,�S��ݨD1,�S��ݨD2,���,�M��,�N��,Receipt_No,�f�D�渹,��f�Ȥ�²��,��f�a�},�q�����O " & _
              "From ORTPlan_ReservedOrder Order by �q��s�� "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫O�d�q����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Dim iLoop As Double
    Do While Not tmp_Rs.EOF
        rs_Tab2_ReservedOrders.AddNew
        For iLoop = 1 To rs_Tab2_ReservedOrders.Fields.Count - 1
            rs_Tab2_ReservedOrders.Fields(iLoop).Value = tmp_Rs.Fields(iLoop - 1).Value
        Next iLoop
        rs_Tab2_ReservedOrders.Fields(0).Value = rs_Tab2_ReservedOrders.RecordCount
        rs_Tab2_ReservedOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    blTab2ReservedEventEnable = True
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�O�d�q��-��ܥ����q��", Me.Caption, "cmd_Tab2_ShowAll_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_FilterAndSort_Click()
    '�ƨ��@�~ >> �O�d�q��j�M
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    
    strFormName_FilterAndSort = Me.Name
    strRSName_FilterAndSort = "rs_Tab2_ReservedOrders"
    
    If ShowForm_RS_FilterAndSort(rs_Tab2_ReservedOrders, "�O�d�q��", Me.Tag) = False Then
        MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Me.WindowState = 2

End Sub

Private Sub cmd_Tab2_Reset_Click()
    '�ƨ��@�~ >> �����O�d�q��z��Ƨ�
    '�����z�����A���]�ƧǨ̾�
     blTab2ReservedEventEnable = False
     rs_Tab2_ReservedOrders.Filter = adFilterNone
     rs_ORT02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
     blTab2ReservedEventEnable = True
End Sub



Private Sub cmdExit3_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdExport3_Click()
Dim strFileLine As String, strExternorderkey As String
Dim i As Integer, j As Integer
Dim arrLen, ConfirmYN

If Dir("C:\From_ids\UTLR\XRSLUPL.TXT") <> "" Then
    ConfirmYN = MsgBox("�ɮפw�g�s�b�A�O�_�мg?", vbQuestion + vbYesNo, Me.Caption)
    If ConfirmYN = vbNo Then Screen.MousePointer = 0: Exit Sub
End If

i = 0: j = 0
'If (Right(App.Path, 1) = "/" Or Right(App.Path, 1) = "\") Then strFilePathName = App.Path & "BestTransaction.csv"
arrLen = Array(12, 12, 8, 8, 10, 30, 30, 30, 30, 30, 30, 30, 30, 30, 3, 14, 30, 7, 7, 7, 7, 4, 4, 4, 7, 7, 7, 7, 60, 8, 2, 12, 1, 1, 11, 11, 10, 1, 12, 1, 7, 4, 16, 10)

rsMain3.MoveFirst

cn.BeginTrans

Open "C:\From_ids\UTLR\XRSLUPL.TXT" For Output As #1
Do While Not rsMain3.EOF
    strFileLine = ""
    
    For i = 0 To rsMain3.Fields.Count - 1
        strFileLine = strFileLine & GetWord(rsMain3(i) & "", 1, arrLen(i))
    Next i
    
      strFileLine = strFileLine & Val(Format(Now(), "yyyymmddhhmmss"))
    '�g�J���
    Print #1, strFileLine
    j = j + 1
    
    '��s���p���w�^��
    If strExternorderkey <> "R" & rsMain3("ORDERNO") Then
    str_SQL = "update orders set B_fax2 = '1',trafficCop=null where externorderkey ='R" & rsMain3("ORDERNO") & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    strExternorderkey = rsMain3("ORDERNO")
    End If

    rsMain3.MoveNext
    
Loop
''�����X
'Print #1, Chr(26)
Close #1
cn.CommitTrans
'�ƥ��ɮ�
If Dir("C:\from_ids\backup\UTLR\", vbDirectory) = "" Then MkDir "C:\From_ids\Backup\UTLR\"

    FileCopy "C:\From_ids\UTLR\XRSLUPL.TXT", "C:\from_ids\backup\UTLR\XRSLUPL" & Format(Now(), "yyyymmddhhmmss") & ".TXT"

MsgBox "�ɮ׶ץX���� (C:\From_ids\UTLR\XRSLUPL.TXT)�A�@ " & j & " ����ƦC�C", 64, Me.Caption


End Sub

Private Sub cmdRouteQuery3_Click()
Dim i As Long, strSql As String
Dim chcDeliveryDate As String, chcOrderby As String

Screen.MousePointer = 11
Set dgMain3.DataSource = Nothing

strSql = "select * from rordersexport2utl "
        
chcOrderby = "order by loadno , orderno , ultorderline"

'�X�����
chcDeliveryDate = ""
If Len(txtDeliveryDate3.Text) > 0 Then chcDeliveryDate = "where left(loadno,7) = 'R" & Mid(txtDeliveryDate3.Text, 3, 6) & "' "

'�զX�r��
strSql = strSql & chcDeliveryDate & chcOrderby

Set rsMain3 = New ADODB.Recordset
rsMain3.CursorLocation = adUseClient
cn.CommandTimeout = 0
rsMain3.Open strSql, cn

If rsMain3.EOF = True Then Screen.MousePointer = 0: MsgBox "�L��ƥi��ܡI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Set dgMain3.DataSource = rsMain3

SetDataGridColWidth Me.Caption, dgMain3

'���D��
With dgMain3

    .ColumnHeaders = True        '���D�����
    .RowHeight = 300
    .Columns(17).Alignment = dbgRight
    .Columns(18).Alignment = dbgRight
    .Columns(19).Alignment = dbgRight
    .Columns(20).Alignment = dbgRight
    .Columns(22).Alignment = dbgRight
    .Columns(24).Alignment = dbgRight
    .Columns(25).Alignment = dbgRight
    .Columns(26).Alignment = dbgRight
    .Columns(27).Alignment = dbgRight
    .Columns(34).Alignment = dbgRight
    .Columns(35).Alignment = dbgRight
    .Columns(42).Alignment = dbgRight

End With

cmdExport3.Enabled = True

Screen.MousePointer = 0
End Sub

Private Sub dg_Tab0_SelectedOrders_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�� [�w����q��] dg_Tab0_SelectedOrders �����D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_Tab0_SelectedOrders) <> "Nothing" Then
        OrderFieldName = "[" & dg_Tab0_SelectedOrders.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_Tab0_SelectedOrders.Sort = OrderFieldName & " DESC "
        Else
            strOrder = "ASC"
            rs_Tab0_SelectedOrders.Sort = OrderFieldName & " ASC "
        End If
    End If
End Sub

Private Sub dg_Tab0_SelectedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '�ƨ��@�~ >> �w����q�� DBGrid
    If blTab0SelectedOrderEventEnable Then
        With dg_Tab0_SelectedOrders
            '�ϥ���ܿ������ƦC
            If Not rs_Tab0_SelectedOrders.EOF Then
                dg_Tab0_SelectedOrders.SelBookmarks.Add rs_Tab0_SelectedOrders.Bookmark
            End If
        End With
    End If
End Sub

Private Sub dg_Tab1_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '���u�s���C��G�����
    If blTab1RouteEventEnable Then
        If Not rs_Tab1_Route.EOF Then
            dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
            rs_Tab1_RouteOrders.Filter = " ���u�s�� = '" & rs_Tab1_Route.Fields("���u�s��").Value & "' "
        End If
    End If
End Sub

Private Sub dg_Tab2_ReservedOrders_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�� [�O�d�q��] dg_Tab2_ReservedOrder �����D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_Tab2_ReservedOrders) <> "Nothing" Then
        OrderFieldName = "[" & dg_Tab2_ReservedOrders.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_Tab2_ReservedOrders.Sort = OrderFieldName & " DESC "
        Else
            strOrder = "ASC"
            rs_Tab2_ReservedOrders.Sort = OrderFieldName & " ASC "
        End If
    End If
End Sub

Private Sub dg_Tab2_ReservedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '�ƨ��@�~ >> �O�d�q�� DBGrid
    If rs_Tab2_ReservedOrders.EOF Then Exit Sub
    If blTab2ReservedEventEnable Then
        With dg_Tab2_ReservedOrders
            '�I�@�U����A���I�h [����]
            If Trim(rs_Tab2_ReservedOrders.Fields(1).Value) = "" Then
                rs_Tab2_ReservedOrders.Fields(1).Value = "V"
            Else
                rs_Tab2_ReservedOrders.Fields(1).Value = " "
            End If
            '�ϥ���ܿ������ƦC
            If Not rs_Tab2_ReservedOrders.EOF Then
                dg_Tab2_ReservedOrders.SelBookmarks.Add rs_Tab2_ReservedOrders.Bookmark
            End If
        End With
    End If
End Sub

Private Sub dg_TRP02W_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim objDataGrid As Object: Set objDataGrid = dg_TRP02W
If Len(objDataGrid.Columns(ColIndex).DataField) = 0 Or objDataGrid.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, "��L�ƨ��ݱƨ��q��" & objDataGrid.Name, objDataGrid.Columns(ColIndex).DataField, objDataGrid.Columns(ColIndex).Width
End Sub

Private Sub dg_TRP02W_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�� [�ݱƨ��q��] dg_TRP02W �����D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_ORT02W) <> "Nothing" Then
        '�קK���� [���] ���ʧ@
        blTRP02WEventEnable = False
        OrderFieldName = "[" & dg_TRP02W.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_ORT02W.Sort = OrderFieldName & " DESC "
            strSourceOrderBy = OrderFieldName & " desc "
        Else
            strOrder = "ASC"
            rs_ORT02W.Sort = OrderFieldName & " ASC "
            strSourceOrderBy = OrderFieldName & " asc "
        End If
        blTRP02WEventEnable = True
    End If
End Sub

Private Sub dg_TRP02W_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '�ƨ��@�~ >> �ݱƨ��q�� DBGrid
    If blTRP02WEventEnable Then
        With dg_TRP02W
            '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
            If Trim(rs_ORT02W.Fields(1).Value) = "" Then
                rs_ORT02W.Fields(1).Value = "V"
                dbSelectedCount = dbSelectedCount + 1
                '����p�p��s
                dbsrcSelected_Case = dbsrcSelected_Case + rs_ORT02W.Fields("�c��").Value
                dbsrcSelected_Pallet = dbsrcSelected_Pallet + rs_ORT02W.Fields("�O��").Value
                dbsrcSelected_Volumn = dbsrcSelected_Volumn + rs_ORT02W.Fields("���n").Value
                dbsrcSelected_Weight = dbsrcSelected_Weight + rs_ORT02W.Fields("���q").Value
                txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
                txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
            Else
                dbSelectedCount = dbSelectedCount - 1
                rs_ORT02W.Fields(1).Value = " "
                '����p�p��s
                If dbSelectedCount = 0 Then
                    dbsrcSelected_Case = 0
                    dbsrcSelected_Pallet = 0
                    dbsrcSelected_Volumn = 0
                    dbsrcSelected_Weight = 0
                    txt_Tab0_srcSelected_Case.Text = 0: txt_Tab0_srcSelected_Pallet.Text = 0
                    txt_Tab0_srcSelected_Volumn.Text = 0: txt_Tab0_srcSelected_Weight.Text = 0
                Else
                    dbsrcSelected_Case = dbsrcSelected_Case - rs_ORT02W.Fields("�c��").Value
                    dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_ORT02W.Fields("�O��").Value
                    dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_ORT02W.Fields("���n").Value
                    dbsrcSelected_Weight = dbsrcSelected_Weight - rs_ORT02W.Fields("���q").Value
                    txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
                    txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
                End If
            End If
            '�ϥ���ܿ������ƦC
            If Not rs_ORT02W.EOF Then
                dg_TRP02W.SelBookmarks.Add rs_ORT02W.Bookmark
            End If
        End With
    End If
End Sub

Private Sub cmd_Tab0_srcOrderQuery_Click()
    '�ƨ��@�~ >> �ݱƨ��q��j�M
    If rs_ORT02W Is Nothing Then Exit Sub
    If rs_ORT02W.RecordCount = 0 Then Exit Sub
    
    strFormName_FilterAndSort = Me.Name
    strRSName_FilterAndSort = "rs_ORT02W"
    
    If ShowForm_RS_FilterAndSort(rs_ORT02W, "�ݱƨ��q��", Me.Tag) = False Then
        MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Me.WindowState = 2

End Sub

Private Sub dgMain3_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain3
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub Form_Activate()
    '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
    Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
    msg_title = "�ƨ��@�~"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '�d�I��Ӫ����L����ƥ�
    '�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������������
    If KeyCode = vbKeyEscape Then
        mvDate.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '�]�w Form �j�p�B��m
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 13170
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    
    '�ƨ��@�~�G�ݱƨ��q��
    Call CreateRS_Tab0_TRP02W
    strSourceFilter = adFilterNone
    strSourceOrderBy = " �s�� asc "
    
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w
    Call CreateRS_Tab0_SelectedOrders
    
    '�w���ͤ����u�s���C��
    Call CreateRS_Tab1_Route
    Call CreateRS_Tab1_RouteOrders
    
    '�O�d�q��C��
    Call CreateRS_Tab2_ReservedOrders
    blTab2ReservedEventEnable = True
    SSTab1.Tab = 0
End Sub

Private Sub Form_Resize()
    '�����j�p�ܰ�
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    If Me.ScaleHeight < dbsrcFormHeight Then
        '�ܤp
        SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
        SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
        
        fam_SelectedOrders.Width = fam_SelectedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dg_Tab0_SelectedOrders.Width = dg_Tab0_SelectedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        fam_SrcOrders.Height = fam_SrcOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        fam_SrcOrders.Width = fam_SrcOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dg_TRP02W.Height = dg_TRP02W.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_TRP02W.Width = dg_TRP02W.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        Frame3.Left = Frame3.Left - (dbsrcFormWidth - Me.ScaleWidth)
        Frame4.Left = Frame4.Left - (dbsrcFormWidth - Me.ScaleWidth)
        dg_Tab1_Route.Width = dg_Tab1_Route.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        dg_Tab1_RouteOrders.Height = dg_Tab1_RouteOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_Tab1_RouteOrders.Width = dg_Tab1_RouteOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dg_Tab2_ReservedOrders.Height = dg_Tab2_ReservedOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_Tab2_ReservedOrders.Width = dg_Tab2_ReservedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dgMain3.Height = SSTab1.Height - 1300
        dgMain3.Width = SSTab1.Width - 240
        
        cmd_Tab2_Remove.Left = cmd_Tab2_Remove.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_FilterAndSort.Left = cmd_Tab2_FilterAndSort.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Reset.Left = cmd_Tab2_Reset.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_ShowAll.Left = cmd_Tab2_ShowAll.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Delete.Left = cmd_Tab2_Delete.Left - (dbsrcFormWidth - Me.ScaleWidth)
        dbsrcFormHeight = Me.ScaleHeight
        dbsrcFormWidth = Me.ScaleWidth
    Else
       SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
       SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
       
       fam_SelectedOrders.Width = fam_SelectedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       dg_Tab0_SelectedOrders.Width = dg_Tab0_SelectedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       fam_SrcOrders.Width = fam_SrcOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       fam_SrcOrders.Height = fam_SrcOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_TRP02W.Height = dg_TRP02W.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_TRP02W.Width = dg_TRP02W.Width + (Me.ScaleWidth - dbsrcFormWidth)
       
       Frame3.Left = Frame3.Left + (Me.ScaleWidth - dbsrcFormWidth)
       dg_Tab1_Route.Width = dg_Tab1_Route.Width + (Me.ScaleWidth - dbsrcFormWidth)
       Frame4.Left = Frame4.Left + (Me.ScaleWidth - dbsrcFormWidth)
       
       dg_Tab1_RouteOrders.Height = dg_Tab1_RouteOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_Tab1_RouteOrders.Width = dg_Tab1_RouteOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       dg_Tab2_ReservedOrders.Height = dg_Tab2_ReservedOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_Tab2_ReservedOrders.Width = dg_Tab2_ReservedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       dgMain3.Height = SSTab1.Height - 1300
       dgMain3.Width = SSTab1.Width - 240
       
       cmd_Tab2_Remove.Left = cmd_Tab2_Remove.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_FilterAndSort.Left = cmd_Tab2_FilterAndSort.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Reset.Left = cmd_Tab2_Reset.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_ShowAll.Left = cmd_Tab2_ShowAll.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Delete.Left = cmd_Tab2_Delete.Left + (Me.ScaleWidth - dbsrcFormWidth)
       
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
    Set frm_OP_TRPPlan = Nothing
End Sub

Private Sub CreateRS_Tab0_TRP02W()
    '�ƨ��@�~�G�ݱƨ��q��
    Call ReDim_Recordset(rs_ORT02W)
    With rs_ORT02W
        .Fields.Append "�s��", adDouble
        .Fields.Append "��", adVarChar, 2
        .Fields.Append "�ѦҸ��s", adVarChar, 10
        .Fields.Append "���h��", adVarChar, 10
        .Fields.Append "�q��s��", adVarChar, 60
        .Fields.Append "�q�����A", adVarChar, 60
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "�Ȥ�s��", adVarChar, 30
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "�Ȥ�²��", adVarChar, 60
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "���A", adVarChar, 10
        .Fields.Append "���f�a�}", adVarChar, 120
        .Fields.Append "�q��Ƶ�", adVarChar, 300
        .Fields.Append "�t�e�ܧO", adVarChar, 120
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�S��ݨD1", adVarChar, 60
        .Fields.Append "�S��ݨD2", adVarChar, 60
        .Fields.Append "���", adVarChar, 10
        .Fields.Append "�M��", adVarChar, 10
        .Fields.Append "�N��", adVarChar, 10
        .Fields.Append "Receipt_No", adVarChar, 10
        .Fields.Append "�f�D�渹", adVarChar, 40
        .Fields.Append "��f�Ȥ�²��", adVarChar, 120
        .Fields.Append "��f�a�}", adVarChar, 120
        .Fields.Append "�q�����O", adVarChar, 10
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_TRP02W.DataSource = rs_ORT02W
'    '�]�w������
    With dg_TRP02W
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .RowHeight = 250
'        .Columns(0).Width = 500         '�Ǹ�
'        .Columns(0).Alignment = dbgCenter
'        .Columns(1).Width = 300         '����ѧO���
'        .Columns(1).Alignment = dbgCenter
'        .Columns(2).Width = 1000        '���u�s��
'        .Columns(2).Alignment = dbgCenter
'        .Columns(3).Width = 800         '���h��
'        .Columns(3).Alignment = dbgCenter
'        .Columns(4).Width = 2100        '�q��s���G�q��s��+�f�D�渹+�f�D
'        .Columns(4).Alignment = dbgLeft
'        .Columns(5).Width = 800        '�q���O
'        .Columns(5).Alignment = dbgLeft
'        .Columns(6).Width = 600         '�c��
'        .Columns(6).Alignment = dbgRight
'        .Columns(7).Width = 600         '�O��
'        .Columns(7).Alignment = dbgRight
'        .Columns(8).Width = 600         '���n
'        .Columns(8).Alignment = dbgRight
'        .Columns(9).Width = 600         '���q
'        .Columns(9).Alignment = dbgRight
'        .Columns(10).Width = 1100        '�Ȥ�s��
'        .Columns(10).Alignment = dbgLeft
'        .Columns(11).Width = 400         'zip
'        .Columns(11).Alignment = dbgCenter
'        .Columns(12).Width = 1000       '�Ȥ�²��
'        .Columns(12).Alignment = dbgLeft
'        .Columns(13).Width = 450        'Area_Code
'        .Columns(13).Alignment = dbgCenter
'        .Columns(14).Width = 450        '���A�GPriority
'        .Columns(14).Alignment = dbgCenter
'        .Columns(15).Width = 3000       '�B�e�a�}
'        .Columns(15).Alignment = dbgLeft
'        .Columns(16).Width = 1400       '�q��Ƶ�
'        .Columns(16).Alignment = dbgLeft
'        .Columns(17).Width = 500        '����
'        .Columns(17).Alignment = dbgCenter
'        .Columns(18).Width = 1500       '�S��ݨD1
'        .Columns(18).Alignment = dbgLeft
'        .Columns(19).Width = 1500       '�S��ݨD2
'        .Columns(19).Alignment = dbgLeft
'        .Columns(20).Width = 500        '���
'        .Columns(20).Alignment = dbgCenter
'        .Columns(21).Width = 500        '�M��
'        .Columns(21).Alignment = dbgCenter
'        .Columns(22).Width = 500        '�N��
'        .Columns(22).Alignment = dbgCenter
'        .Columns(23).Width = 1100       'Receipt_No
'        .Columns(23).Alignment = dbgLeft
'        .Columns(24).Width = 900        '�f�D�渹
'        .Columns(24).Alignment = dbgLeft
'        .Columns(25).Width = 1500       '�Ȥ�W��
'        .Columns(25).Alignment = dbgLeft
'        .Columns(26).Width = 1500       '���f��
'        .Columns(26).Alignment = dbgLeft
'        .Columns(27).Width = 500       '�q�����O
'        .Columns(27).Alignment = dbgLeft
    End With
    SetDataGridColWidth "��L�ƨ��ݱƨ��q��", dg_TRP02W
End Sub

Private Sub CreateRS_Tab0_SelectedOrders()
    '�ƨ��@�~�G�w������ݱƨ��q��C��
    Call ReDim_Recordset(rs_Tab0_SelectedOrders)
    With rs_Tab0_SelectedOrders
        .Fields.Append "�s��", adDouble
        .Fields.Append "�ѦҸ��s", adVarChar, 10
        .Fields.Append "���h��", adVarChar, 20
        .Fields.Append "�q��s��", adVarChar, 60
        .Fields.Append "�q�����A", adVarChar, 60
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "���A", adVarChar, 20
        .Fields.Append "�Ȥ�²��", adVarChar, 120
        .Fields.Append "���f�a�}", adVarChar, 120
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�q��Ƶ�", adVarChar, 300
        .Fields.Append "�S��ݨD1", adVarChar, 60
        .Fields.Append "�S��ݨD2", adVarChar, 60
        .Fields.Append "Receipt_No", adVarChar, 10
        .Fields.Append "EXE�^��", adVarChar, 20
        .Fields.Append "��f�Ȥ�²��", adVarChar, 120
        .Fields.Append "��f�a�}", adVarChar, 120
        .Fields.Append "�q�����O", adVarChar, 10
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_Tab0_SelectedOrders.DataSource = rs_Tab0_SelectedOrders
    '�]�w������
    With dg_Tab0_SelectedOrders
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500        '�s��
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000        '���u�s��
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 800         '���h��
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2100        '�q��s��
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 800        '�q���O
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 400         'ZIP
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 450         'Area
        .Columns(6).Alignment = dbgCenter
        .Columns(7).Width = 450         '���A�GOrders.Priority
        .Columns(7).Alignment = dbgCenter
        .Columns(8).Width = 1000        '�Ȥ�²��
        .Columns(8).Alignment = dbgLeft
        .Columns(9).Width = 600         '�c��
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 600         '�O��
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 600         '���n
        .Columns(11).Alignment = dbgRight
        .Columns(12).Width = 600        '���q
        .Columns(12).Alignment = dbgRight
        .Columns(13).Width = 450        '����
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1200       '�q��Ƶ�
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 1500       '�S��ݨD-1
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1500       '�S��ݨD-2
        .Columns(16).Alignment = dbgLeft
        .Columns(17).Width = 1000       'Receipt_No
        .Columns(17).Alignment = dbgLeft
        .Columns(18).Width = 1000       'EXE�^��
        .Columns(18).Alignment = dbgLeft
        .Columns(19).Width = 1500       '�Ȥ�W��
        .Columns(19).Alignment = dbgLeft
        .Columns(20).Width = 1500       '���f��
        .Columns(20).Alignment = dbgLeft
        .Columns(21).Width = 500        '�q�����O
        .Columns(21).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab1_Route()
    '�ƨ��@�~�G�w�s�������u�s���C��
    Call ReDim_Recordset(rs_Tab1_Route)
    With rs_Tab1_Route
        .Fields.Append "�s��", adDouble
        .Fields.Append "���u�s��", adVarChar, 10
        .Fields.Append "�X�����", adVarChar, 8
        .Fields.Append "���P���X", adVarChar, 10
        .Fields.Append "����", adDouble
        .Fields.Append "�r�p�H", adVarChar, 20
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�X�Y�Ȧs", adVarChar, 10
        .Fields.Append "�w�p������", adVarChar, 8
        .Fields.Append "�w�p����ɶ�", adVarChar, 4
        .Fields.Append "EXE�^��", adVarChar, 20
        .Fields.Append "�ƨ���", adVarChar, 30
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_Tab1_Route.DataSource = rs_Tab1_Route
    '�]�w������
    With dg_Tab1_Route
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500         '�s��
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000        '���u�s��
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 900         '�X�����
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 850         '���P���X
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 500         '����
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 900         '�r�p�H
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 700         '�c��
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 700         '�O��
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 700         '���n
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 700         '���q
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 450        '����
        .Columns(10).Alignment = dbgCenter
        .Columns(11).Width = 1000       '�X�Y�Ȧs
        .Columns(11).Alignment = dbgLeft
        .Columns(12).Width = 1400       '�w�p����������
        .Columns(12).Alignment = dbgLeft
        .Columns(13).Width = 1400       '�w�p��������ɶ�
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 900        'EXE �^�Ǫ��A
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 1200       '�ƨ���
        .Columns(15).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab1_RouteOrders()
    '�ƨ��@�~�G�w�s�����s���q��C��
    Call ReDim_Recordset(rs_Tab1_RouteOrders)
    With rs_Tab1_RouteOrders
        .Fields.Append "�s��", adDouble
        .Fields.Append "���u�s��", adVarChar, 10
        .Fields.Append "���h��", adVarChar, 20
        .Fields.Append "�q��s��", adVarChar, 60
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "�Ȥ�²��", adVarChar, 40
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "�q��Ƶ�", adVarChar, 300
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�S��ݨD1", adVarChar, 60
        .Fields.Append "�S��ݨD2", adVarChar, 60
        .Fields.Append "Receipt_No", adVarChar, 60
        .Fields.Append "EXE�^��", adVarChar, 20
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "���A", adVarChar, 10
        .Fields.Append "��f�Ȥ�²��", adVarChar, 120
        .Fields.Append "��f�Ȥ�a�}", adVarChar, 200
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_Tab1_RouteOrders.DataSource = rs_Tab1_RouteOrders
    '�]�w������
    With dg_Tab1_RouteOrders
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500         '�s��
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1050        '���u�s��
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 900         '���h��
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2150        '�q��s��
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 400         'ZIP
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 1500        '�Ȥ�W��
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 700         '�c��
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 700         '�O��
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 700         '���n
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 700         '���q
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 1500       '�q��Ƶ�
        .Columns(10).Alignment = dbgLeft
        .Columns(11).Width = 1200       '����
        .Columns(11).Alignment = dbgLeft
        .Columns(12).Width = 1500       '�S��ݨD-1
        .Columns(12).Alignment = dbgLeft
        .Columns(13).Width = 1500       '�S��ݨD-2
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1100       'Receipt_No
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 1100       'EXE�^��
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 450        'Area
        .Columns(16).Alignment = dbgCenter
        .Columns(17).Width = 450        '���A
        .Columns(17).Alignment = dbgCenter
        .Columns(18).Width = 1100       '�Ȥ�²��
        .Columns(18).Alignment = dbgLeft
        .Columns(19).Width = 1100       '�Ȥ�a�}
        .Columns(19).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab2_ReservedOrders()
    '�ƨ��@�~�G�O�d�q��
    Call ReDim_Recordset(rs_Tab2_ReservedOrders)
    With rs_Tab2_ReservedOrders
         .Fields.Append "�s��", adDouble
         .Fields.Append "��", adVarChar, 2
         .Fields.Append "�ѦҸ��s", adVarChar, 10
         .Fields.Append "���h��", adVarChar, 10
         .Fields.Append "�q��s��", adVarChar, 60
         .Fields.Append "�q�����A", adVarChar, 60
         .Fields.Append "�c��", adDouble
         .Fields.Append "�O��", adDouble
         .Fields.Append "���n", adDouble
         .Fields.Append "���q", adDouble
         .Fields.Append "�Ȥ�s��", adVarChar, 30
         .Fields.Append "ZIP", adVarChar, 10
         .Fields.Append "�Ȥ�²��", adVarChar, 60
         .Fields.Append "Area", adVarChar, 10
         .Fields.Append "���A", adVarChar, 10
         .Fields.Append "���f�a�}", adVarChar, 120
         .Fields.Append "�q��Ƶ�", adVarChar, 300
         .Fields.Append "�t�e�ܧO", adVarChar, 120
         .Fields.Append "����", adVarChar, 10
         .Fields.Append "�S��ݨD1", adVarChar, 60
         .Fields.Append "�S��ݨD2", adVarChar, 60
         .Fields.Append "���", adVarChar, 10
         .Fields.Append "�M��", adVarChar, 10
         .Fields.Append "�N��", adVarChar, 10
         .Fields.Append "Receipt_No", adVarChar, 10
         .Fields.Append "�f�D�渹", adVarChar, 40
         .Fields.Append "��f�Ȥ�²��", adVarChar, 120
         .Fields.Append "��f�a�}", adVarChar, 120
         .Fields.Append "�q�����O", adVarChar, 10
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '���ݳs������
    End With
    Set dg_Tab2_ReservedOrders.DataSource = rs_Tab2_ReservedOrders
    '�]�w������
    With dg_Tab2_ReservedOrders
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .RowHeight = 250
        .Columns(0).Width = 500         '�Ǹ�
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 300         '����ѧO���
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000        '���u�s��
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 800         '���h��
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 2100        '�q��s���G�q��s��+�f�D�渹+�f�D
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 800        '�q���O
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 600         '�c��
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 600         '�O��
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 600         '���n
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 600         '���q
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 1100        '�Ȥ�s��
        .Columns(10).Alignment = dbgLeft
        .Columns(11).Width = 400         'zip
        .Columns(11).Alignment = dbgCenter
        .Columns(12).Width = 1000       '�Ȥ�²��
        .Columns(12).Alignment = dbgLeft
        .Columns(13).Width = 450        'Area_Code
        .Columns(13).Alignment = dbgCenter
        .Columns(14).Width = 450        '���A�GPriority
        .Columns(14).Alignment = dbgCenter
        .Columns(15).Width = 3000       '�B�e�a�}
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1400       '�q��Ƶ�
        .Columns(16).Alignment = dbgLeft
        .Columns(17).Width = 450       '�t�e�ܧO
        .Columns(17).Alignment = dbgLeft
        .Columns(18).Width = 500        '����
        .Columns(18).Alignment = dbgCenter
        .Columns(19).Width = 1500       '�S��ݨD1
        .Columns(19).Alignment = dbgLeft
        .Columns(20).Width = 1500       '�S��ݨD2
        .Columns(20).Alignment = dbgLeft
        .Columns(21).Width = 500        '���
        .Columns(21).Alignment = dbgCenter
        .Columns(22).Width = 500        '�M��
        .Columns(22).Alignment = dbgCenter
        .Columns(23).Width = 500        '�N��
        .Columns(23).Alignment = dbgCenter
        .Columns(24).Width = 1100       'Receipt_No
        .Columns(24).Alignment = dbgLeft
        .Columns(25).Width = 900        '�f�D�渹
        .Columns(25).Alignment = dbgLeft
        .Columns(26).Width = 1500       '�Ȥ�W��
        .Columns(26).Alignment = dbgLeft
        .Columns(27).Width = 1500       '���f��
        .Columns(27).Alignment = dbgLeft
        .Columns(28).Width = 500       '�q�����O
        .Columns(28).Alignment = dbgLeft
    End With
End Sub

Private Sub Calculate_SelectedOrders()
    '�@�~���e�G
    '1.�w��w����q��C��A�̭q��s�����s���� [�s��] ����
    '2.�p��w����q�椧�֭p���
    Dim dbSeqNo As Double
    dbSeqNo = 0
    txt_Tab0_Selected_Case.Text = ""
    txt_Tab0_Selected_Pallet.Text = ""
    txt_Tab0_Selected_Volumn.Text = ""
    txt_Tab0_Selected_Weight.Text = ""
    
    rs_Tab0_SelectedOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "Receipt_No asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab0_SelectedOrders.EOF Then
       rs_Tab0_SelectedOrders.MoveFirst
    Else
        '�M�X�z�����A���L��ƪ̡A���� SubProgram ����
        Exit Sub
    End If
    Do While Not rs_Tab0_SelectedOrders.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_SelectedOrders.Fields("�s��").Value = dbSeqNo
        txt_Tab0_Selected_Case.Text = Val(txt_Tab0_Selected_Case.Text) + rs_Tab0_SelectedOrders.Fields("�c��").Value
        txt_Tab0_Selected_Pallet.Text = Val(txt_Tab0_Selected_Pallet.Text) + rs_Tab0_SelectedOrders.Fields("�O��").Value
        txt_Tab0_Selected_Volumn.Text = Val(txt_Tab0_Selected_Volumn.Text) + rs_Tab0_SelectedOrders.Fields("���n").Value
        txt_Tab0_Selected_Weight.Text = Val(txt_Tab0_Selected_Weight.Text) + rs_Tab0_SelectedOrders.Fields("���q").Value
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    rs_Tab0_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab0_SelectedOrders.EOF Then rs_Tab0_SelectedOrders.MoveFirst
End Sub

Private Sub SelectedOrders_Removeto_TRP02W(ByVal strReceiptNo As String)
    '�N���w�� [�q��s��] �[�J [�ݱƨ��q��]
    blTRP02WEventEnable = False
    
    rs_ORT02W.Filter = adFilterNone
    rs_ORT02W.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    If rs_ORT02W.RecordCount > 0 Then
        rs_ORT02W.Filter = "Receipt_No = '" & strReceiptNo & "'"
        If Not rs_ORT02W.EOF Then
            '�q��s���w�s�b���ܡA���i��s�W�A�]����s
            rs_ORT02W.Filter = adFilterNone
            rs_ORT02W.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
            blTRP02WEventEnable = True
            Exit Sub
        End If
    End If
    
    '���^�ݱƨ��q��
    If blRouteModify Then
        '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
        blRouteChange = True
        '�g�Ѭd�߸��u�s���ұo���q����
        str_SQL = "Select ���h��,�q��s��,�c��,�O��,���n,���q,�Ȥ�s��,ZIP,�Ȥ�W��,�B�e�a�},�q��Ƶ�,�t�e�ܧO,����,�S��ݨD1,�S��ݨD2,���,�M��,�N��,Receipt_No,�f�D�渹,EXE�^��,Area,�Ȥ�²��,���A " & _
                  "From TRPPlan_RouteQueryOrdersRemove Where Receipt_No = '" & strReceiptNo & "' Order by �q��s�� "
    Else
'        str_SQL = "Select ���h��,�q��s��,�c��,�O��,���n,���q,�Ȥ�s��,ZIP,�Ȥ�W��,�B�e�a�},�q��Ƶ�,����,�S��ݨD1,�S��ݨD2,���,�M��,�N��,Receipt_No,�f�D�渹,EXE�^��,Area,�Ȥ�²��,���A " & _
'                  "From TRPPlan_SourceOrder Where Receipt_No = '" & strReceiptNo & "' Order by �q��s�� "
        str_SQL = "Select Convert(varchar(8),a1.Arrive_Date,112) as ���h�� , Rtrim(a1.Receipt_No) + '(' + Rtrim(a1.StorerKey)+'-'+Rtrim(a1.Extern)+')' as �q��s�� ,�q�����A = isnull(a2.channel_type,''), " & _
            "Isnull(Round(a1.Case_cnt,2),0) as �c�� ,  Isnull(Round(a1.Pallet_Qty,2),0) as �O�� , " & _
            "Isnull(Round(a1.Weight,2),0) as ���q , Isnull(Round(a1.Volumn_Weight,2),0) as ���n , Rtrim(a1.ConsigneeKey) as �Ȥ�s�� , " & _
            "Isnull(Rtrim(a2.ZIP),'x') as ZIP,��f�Ȥ�²�� = isnull((select TRP01M.short_name from TRP01M join orders on TRP01M.consigneekey = orders.b_company and len(rtrim(isnull(orders.b_company,''))) > 0 and orders.orderkey = a1.c_receipt_no ),'') ,isnull( Rtrim(a2.Full_Name),'x')   as �Ȥ�W�� , isnull(Rtrim(a2.Address),'x')   as ���f�a�} , Rtrim(Isnull(a2.Vehicle_Type,'x')) as ���� , " & _
            "Case When b2.Description = '�L�S��ݨD' Then 'X' else Rtrim(Isnull(b2.Description,'')) End as �S��ݨD1 , " & _
            "Case When b3.Description = '�L�S��ݨD' Then 'X' else Rtrim(Isnull(b3.Description,'')) End as �S��ݨD2 , " & _
            "Rtrim(Isnull(a1.Urgent_Mark,'')) as ��� ,Rtrim(Isnull(a1.Reserve_Mark,'')) as �M�� ,Rtrim(Isnull(a1.Cold_Mark,'')) as �N��  , " & _
            "Rtrim(a1.Receipt_No) as Receipt_No , Rtrim(a1.StorerKey) as �f�D , Convert(varchar(8),a1.Receipt_Date,112) as �q��� , " & _
            "Rtrim(Isnull(a1.Extern,'')) as �f�D�渹 ,��f�a�} = isnull((select TRP01M.address from TRP01M join orders on TRP01M.consigneekey = orders.b_company and len(rtrim(isnull(orders.b_company,''))) > 0 and orders.orderkey = a1.c_receipt_no ),''), " & _
            "Case When Isnull(Rtrim(Cast(c1.Notes as varchar(300))),'') = '' Then 'X' else Rtrim(Cast(c1.Notes as varchar(300))) End as �q��Ƶ� ,�t�e�ܧO = isnull(c1.facility,''), " & _
            "Isnull(Rtrim(a2.Area_Code),'') as Area , Rtrim(a2.Short_Name) as �Ȥ�²�� , Rtrim(Isnull(a1.Priority,'')) as ���A,Rtrim(Isnull(c1.DischargePlace,'')) as ���f��,Rtrim(Isnull(c1.Type,'')) as �q�����O " & _
            ",�ѦҸ��s = (select top 1 route_no from  trp02t trp02t where a1.ConsigneeKey = trp02t.ConsigneeKey and substring(trp02t.route_no,2,6) > convert (char(8) , getdate() , 12)) " & _
            "From ORT02W a1 " & _
            "left outer join TRP01M a2 on a2.ConsigneeKey = a1.ConsigneeKey " & _
            "Left outer join TRP04M b2 on b2.Extra_Demand_Code = a2.Extra_Demand_Code " & _
            "Left outer join TRP04M b3 on b3.Extra_Demand_Code = a2.Extra_Demand_Code2 " & _
            "Left outer join Orders c1 on c1.OrderKey = a1.c_receipt_no " & _
            "where a1.Receipt_No = '" & strReceiptNo & "'"
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q���ƥi�H�٭�^ [�ݿ���q��]"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        blTRP02WEventEnable = True
        Exit Sub
    End If
    
    rs_ORT02W.AddNew
    rs_ORT02W.Fields("�s��").Value = 999
    rs_ORT02W.Fields("���h��").Value = tmp_Rs.Fields("���h��").Value
    rs_ORT02W.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
    rs_ORT02W.Fields("�q�����A").Value = tmp_Rs.Fields("�q�����A").Value
    rs_ORT02W.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
    rs_ORT02W.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
    rs_ORT02W.Fields("���n").Value = tmp_Rs.Fields("���n").Value
    rs_ORT02W.Fields("���q").Value = tmp_Rs.Fields("���q").Value
    rs_ORT02W.Fields("�Ȥ�s��").Value = tmp_Rs.Fields("�Ȥ�s��").Value
    rs_ORT02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
    rs_ORT02W.Fields("zip").Value = tmp_Rs.Fields("zip").Value
    rs_ORT02W.Fields("��f�Ȥ�²��").Value = tmp_Rs.Fields("��f�Ȥ�²��").Value
    rs_ORT02W.Fields("���f�a�}").Value = tmp_Rs.Fields("���f�a�}").Value
    rs_ORT02W.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value
    rs_ORT02W("�t�e�ܧO") = tmp_Rs("�t�e�ܧO")
    rs_ORT02W.Fields("����").Value = tmp_Rs.Fields("����").Value
    rs_ORT02W.Fields("�S��ݨD1").Value = tmp_Rs.Fields("�S��ݨD1").Value
    rs_ORT02W.Fields("�S��ݨD2").Value = tmp_Rs.Fields("�S��ݨD2").Value
    rs_ORT02W.Fields("���").Value = tmp_Rs.Fields("���").Value
    rs_ORT02W.Fields("�M��").Value = tmp_Rs.Fields("�M��").Value
    rs_ORT02W.Fields("�N��").Value = tmp_Rs.Fields("�N��").Value
    rs_ORT02W.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
    rs_ORT02W.Fields("�f�D�渹").Value = tmp_Rs.Fields("�f�D�渹").Value
    rs_ORT02W.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value & ""
    rs_ORT02W.Fields("���A").Value = tmp_Rs.Fields("���A").Value
    rs_ORT02W.Fields("��f�a�}").Value = tmp_Rs.Fields("��f�a�}").Value
    rs_ORT02W.Fields("�q�����O").Value = tmp_Rs.Fields("�q�����O").Value
    rs_ORT02W("�ѦҸ��s") = tmp_Rs("�ѦҸ��s") & ""
    rs_ORT02W.Update
    tmp_Rs.Close
    
    rs_ORT02W.Filter = adFilterNone
    rs_ORT02W.Sort = "�q��s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_ORT02W.EOF Then rs_ORT02W.MoveFirst
    blTRP02WEventEnable = True
End Sub

Private Sub ReSet_TRP02W_SeqNo()
    '���s���� [�ݱƨ��q��] �� [�s��] ����
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    rs_ORT02W.Filter = adFilterNone
    rs_ORT02W.Sort = "�q��s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_ORT02W.EOF Then rs_ORT02W.MoveFirst
    
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_ORT02W.EOF
        dbSeqNo = dbSeqNo + 1
        rs_ORT02W.Fields("�s��").Value = dbSeqNo
        rs_ORT02W.MoveNext
    Loop
    If rs_ORT02W.RecordCount > 0 Then rs_ORT02W.MoveFirst
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    '������
    Select Case mvDate.Tag
           Case "�X�����"
                txt_Tab0_TRPDate.Text = Format(mvDate.Value, "yyyymmdd")
           Case "�w�p������"
                txt_Tab0_CarCheckInDate.Text = Format(mvDate.Value, "yyyymmdd")
'
    End Select
    mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call Form_KeyDown(27, 0)
If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_Tab0_CarCheckInDate_Click()
    '�ƨ��@�~ >> �w�p������
    If Trim(txt_Tab0_CarCheckInDate.Text) = "" Then
        mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab0_CarCheckInDate.Text, 4) & "/" & Mid(txt_Tab0_CarCheckInDate.Text, 5, 2) & "/" & Right(txt_Tab0_CarCheckInDate.Text, 2))
        End If
    End If
    mvDate.Left = fam_RouteData.Left + txt_Tab0_CarCheckInDate.Left
    mvDate.Top = fam_RouteData.Top + txt_Tab0_CarCheckInDate.Top + txt_Tab0_CarCheckInDate.Height
    mvDate.Tag = "�w�p������"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab0_CarCheckInDate_KeyPress(KeyAscii As Integer)
    '�w�p������
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '�����\��J�r��
              KeyAscii = 0
         Case vbKeyReturn
              KeyAscii = 0
              txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
              txt_Tab0_CarCheckInTime.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_CarCheckInTime_KeyPress(KeyAscii As Integer)
    '�w�p����ɶ�
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '�����\��J�r��
              KeyAscii = 0
    End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_KeyPress(KeyAscii As Integer)
    '�ƨ��@�~ >> ���P���X
    Select Case KeyAscii
           Case 97 To 122   '�ഫ���j�g�r��
                KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_LostFocus()  'daniel--20040928<���uuser��J���~������>
    If Len(txt_Tab0_DeliveryCarNo.Text) = 0 Then Exit Sub
    str_SQL = "Select Vehicle_ID_No from trp09m where Vehicle_ID_No='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "' "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        'tmp_rs.Close
        msg_text = "�L���������"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SetFocus
    End If
    tmp_Rs.Close
End Sub

Private Sub txt_Tab0_DockNo_KeyPress(KeyAscii As Integer)
    '�X�Y�Ȧs
    Select Case KeyAscii
           Case 97 To 122   '�ഫ���j�g�r��
                KeyAscii = KeyAscii - 32
           Case vbKeyReturn
                KeyAscii = 0
                txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
                txt_Tab0_CarCheckInDate.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_Route_Change()
If Len(Trim(txt_Tab0_Route.Text)) <> 10 Then Exit Sub
    '���X�۹����ƨ���ƶ�J
        str_SQL = "Select �X�����,���P���X,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,�B�餽�q = isnull(t9m.trp_company_code,''),�r�p�H,�r�p�q��,���� = isnull(t9m.vehicle_type,'') " & _
              "From TRPPlan_RouteQuery t join trp09m t9m on ���P���X = t9m.vehicle_id_no Where ���u�s�� = '" & Trim(txt_Tab0_Route.Text) & "'"
        Dim rsTmp As New ADODB.Recordset
        rsTmp.Open str_SQL, cn
        
        If rsTmp.EOF = 0 Then
        txt_Tab0_TRPDate = rsTmp("�X�����")
        txt_Tab0_DeliveryCarNo = rsTmp("���P���X")
        txt_Tab0_DeliveryCompany = rsTmp("�B�餽�q")
        txt_Tab0_DeliveryDriver = rsTmp("�r�p�H")
        txt_Tab0_DeliveryPhone = rsTmp("�r�p�q��")
        txt_Tab0_DeliveryCarType = rsTmp("����") & ""
        txt_Tab0_DockNo = rsTmp("�X�Y�Ȧs")
        txt_Tab0_CarCheckInDate = rsTmp("�w�p������")
        txt_Tab0_CarCheckInTime = rsTmp("�w�p����ɶ�")
        End If
        rsTmp.Close: Set rsTmp = Nothing

End Sub

Private Sub txt_Tab0_Route_KeyPress(KeyAscii As Integer)
    '���u�s���C�� >> ���u�s��
    Select Case KeyAscii
         Case 97 To 122   '�ഫ�j�g�r��
              KeyAscii = KeyAscii - 32
         Case vbKeyReturn
              cmd_Tab1_RouteNoQuery.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_RouteNo_KeyPress(KeyAscii As Integer)
    '�ƨ��@�~ >> ���u�s��
    Select Case KeyAscii
        Case 97 To 122     '�p�g�r���אּ�j�g�r��
             KeyAscii = KeyAscii - 32
        Case vbKeyReturn
             cmd_Tab0_Query.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_TRPDate_Click()
    '�ƨ��@�~ >> �X�����
    If Trim(txt_Tab0_TRPDate.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2))
        End If
    End If
    mvDate.Left = fam_SelectedOrders.Left + txt_Tab0_TRPDate.Left
    mvDate.Top = fam_SelectedOrders.Top + txt_Tab0_TRPDate.Top + txt_Tab0_TRPDate.Height
    mvDate.Tag = "�X�����"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab0_TRPDate_KeyPress(KeyAscii As Integer)
    '�ƨ��@�~ > [�X�����] ��Ʈ榡�Gyyyymmdd
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '�����\��J�r��
              KeyAscii = 0
         Case vbKeyReturn
              If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
                 msg_text = "�X������G" & funRtn_msg
                 MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                 txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
                 Exit Sub
              Else
                 cmd_Tab0_SelectCar.SetFocus
              End If
    End Select
End Sub

Public Sub frm_OP_TRPPlan_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
    '��椽�ΰƵ{���A�� frm_RS_FilterAndSort ���I�s
    '�ǤJ�ȡGstrCode      �ʧ@�ѧO�X
    '                     [FILTER] �ۭq�z��    [SORT] �Ƨ�
    '        strReturn    �z�� or �Ƨ� ���]�w�r��
    
    Select Case strCode
           Case "FILTER"  '�ۭq�z��
                Select Case UCase(strRSName_FilterAndSort)
                       Case "RS_ORT02W"                '�ݱƨ��q����
                            blTRP02WEventEnable = False
                            '�z��w����̡G�������
                            rs_ORT02W.Filter = "��='V'"
                            If Not rs_ORT02W.EOF Then
                               Do While Not rs_ORT02W.EOF
                                  rs_ORT02W.Fields(1).Value = " "
                                  rs_ORT02W.MoveNext
                               Loop
                            End If
                            rs_ORT02W.Filter = adFilterNone
                            rs_ORT02W.Filter = strReturn
                            strSourceFilter = strReturn
                            If rs_ORT02W.RecordCount = 0 Then
                               msg_text = "��p���A�䤣��ŦX���󪺭q���"
                               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                               rs_ORT02W.Filter = adFilterNone
                               strSourceFilter = adFilterNone
                               rs_ORT02W.Sort = strSourceOrderBy   '�٭�ƧǤ覡
                               blTRP02WEventEnable = True
                               Exit Sub
                            End If
                            '���s�p�� [�ݱƨ��C��] ���`�p��T
                            Call ReCaculate_OrderSum
                            blTRP02WEventEnable = True
                       Case "RS_TAB2_RESERVEDORDERS"   '�O�d�q��
                            blTab2ReservedEventEnable = False
                            rs_Tab2_ReservedOrders.Filter = adFilterNone
                            rs_Tab2_ReservedOrders.Filter = strReturn
                            If rs_Tab2_ReservedOrders.RecordCount = 0 Then
                               msg_text = "��p���A�䤣��ŦX���󪺫O�d�q���"
                               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                               rs_Tab2_ReservedOrders.Filter = adFilterNone
                               rs_Tab2_ReservedOrders.Sort = strSourceOrderBy   '�٭�ƧǤ覡
                               blTRP02WEventEnable = True
                               Exit Sub
                            End If
                            blTab2ReservedEventEnable = True
                       
                End Select
           Case "SORT"    '�Ƨ�
                Select Case UCase(strRSName_FilterAndSort)
                       Case "rs_ORT02W"               '�ݱƨ��q����
                            If rs_ORT02W.EOF Then Exit Sub
                            blTRP02WEventEnable = False
                            rs_ORT02W.Sort = strReturn
                            strSourceOrderBy = strReturn
                            blTRP02WEventEnable = True
                       Case "RS_TAB2_RESERVEDORDERS"   '�O�d�q��
                            If rs_Tab2_ReservedOrders.EOF Then Exit Sub
                            blTab2ReservedEventEnable = False
                            rs_Tab2_ReservedOrders.Sort = strReturn
                            strSourceOrderBy = strReturn
                            blTab2ReservedEventEnable = True
                End Select
    End Select
End Sub

Private Sub txt_Tab1_RouteNo_KeyPress(KeyAscii As Integer)
    '���u�s���C�� >> ���u�s��
    Select Case KeyAscii
         Case 97 To 122   '�ഫ�j�g�r��
              KeyAscii = KeyAscii - 32
         Case vbKeyReturn
              cmd_Tab1_RouteNoQuery.SetFocus
    End Select
End Sub

Private Sub Clear_RouteData()
    '�ƨ��@�~�G�M�����u�s��������
    blRouteModify = False
    strDispRouteNo = ""
    blRouteChange = False
    
    blTab0SelectedOrderEventEnable = False
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w
    Call CreateRS_Tab0_SelectedOrders
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    txt_Tab0_TRPDate.Text = ""
    txt_Tab0_DeliveryCarNo.Text = ""
    txt_Tab0_DockNo.Text = ""
    txt_Tab0_CarCheckInDate.Text = ""
    txt_Tab0_CarCheckInTime.Text = ""
    txt_Tab0_DeliveryCompany.Text = ""
    txt_Tab0_DeliveryDriver.Text = ""
    txt_Tab0_DeliveryPhone.Text = ""
    txt_Tab0_DeliveryCarType.Text = ""
    txt_Tab0_Selected_Case.Text = ""
    txt_Tab0_Selected_Pallet.Text = ""
    txt_Tab0_Selected_Volumn.Text = ""
    txt_Tab0_Selected_Weight.Text = ""
End Sub

Private Function RouteData_Check() As Boolean

    '�ˮָ��u�s����ƬO�_���T
    RouteData_Check = False
    If Len(Trim(txt_Tab0_TRPDate.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J�X�����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SetFocus
        Exit Function
    End If
    If Len(Trim(txt_Tab0_DeliveryCarNo.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J���P���X"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Function
    End If
    
    '����ˮ�
    'a1.�X������G�榡 yyyymmdd
    txt_Tab0_TRPDate.Text = Trim(txt_Tab0_TRPDate.Text)
    If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
        msg_text = "�X������G" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
        Exit Function
    End If
    'a2.�X����� >= ����
    If txt_Tab0_TRPDate.Text < Format(Now, "yyyymmdd") Then
        msg_text = "�X��������o�p�󤵤�"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
        Exit Function
    End If
    
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    'b.�ˮ� [���P���X] �O�_����
    txt_Tab0_DeliveryCarNo.Text = Trim(txt_Tab0_DeliveryCarNo.Text)
    str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "��ƿ��~�G���P���X " & txt_Tab0_DeliveryCarNo.Text & " ������"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Function
    End If
    tmp_Rs.Close
    '���w�X�Y�Ȧs�G������J
    txt_Tab0_DockNo.Text = Trim(txt_Tab0_DockNo.Text)
    If Len(Trim(txt_Tab0_DockNo.Text)) = 0 Then
        msg_text = "��ƿ��~�G[�X�Y�Ȧs] ������J"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DockNo.SetFocus
        Exit Function
    End If
    '�w�p������
    txt_Tab0_CarCheckInDate.Text = Trim(txt_Tab0_CarCheckInDate.Text)
    If Len(Trim(txt_Tab0_CarCheckInDate.Text)) <> 8 Then
        msg_text = "�w�p�������G��Ʈ榡 yyyymmdd "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
        txt_Tab0_CarCheckInDate.SetFocus
        Exit Function
    End If
    If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
        msg_text = "�w�p�������G��ƿ��~ yyyymmdd�A" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
        txt_Tab0_CarCheckInDate.SetFocus
        Exit Function
    End If
    'a2.�w�p������ >= ����
    If txt_Tab0_CarCheckInDate.Text < Format(Now, "yyyymmdd") Then
       msg_text = "�w�p���������o�p�󤵤�"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text): txt_Tab0_CarCheckInDate.SetFocus
       Exit Function
    End If
    
    '�w�p����ɶ�
    txt_Tab0_CarCheckInTime.Text = Trim(txt_Tab0_CarCheckInTime.Text)
    If Len(Trim(txt_Tab0_CarCheckInTime.Text)) <> 4 Then
        msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
        txt_Tab0_CarCheckInTime.SetFocus
        Exit Function
    End If
    Select Case Left(txt_Tab0_CarCheckInTime.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
                txt_Tab0_CarCheckInTime.SetFocus
                Exit Function
    End Select
    Select Case Right(txt_Tab0_CarCheckInTime.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
                txt_Tab0_CarCheckInTime.SetFocus
                Exit Function
    End Select
    RouteData_Check = True
End Function

Private Sub Delete_RouteNo(strRouteNo As String)
    Screen.MousePointer = vbHourglass
    blTab1RouteEventEnable = False
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '�R�� ORT01T ���u�s���D��
    Call DB_CheckConnectStatus
    
    '(1).�N ORT03T �g�^ ORT03W >> �R�� ORT03T
    str_SQL = "Insert into ORT03W(" & _
              "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From ORT03T A Where a.Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).�N ORT02T �g�^ ORT02W >> �R�� ORT02T
    str_SQL = "Insert into ORT02W(" & _
              "   RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,C_RECEIPT_NO,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
              "Select RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,C_RECEIPT_NO,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
              "From ORT02T Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(3).�R�� ORT02T & ORT03T
    str_SQL = "Delete From ORT03T Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Delete From ORT02T Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
              
    '(4).�R�� ORT05T
    str_SQL = "Delete From ORT05T Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(5).�R�� ORT01T
    str_SQL = "Delete From ORT01T Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    
    '(6)��Ʈw���ʽT�{
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
   CreateErrorLog Me.Name & "-�ƨ��@�~-���u�s���R��", Me.Caption, "Form ���� SubProgram Delete_RouteNo", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub Retrive_OrderSum()
    '�����ݱƨ��q��G�`�p��ƭ�
    txt_Tab0_srcTotal_Case.Text = ""
    txt_Tab0_srcTotal_Pallet.Text = ""
    txt_Tab0_srcTotal_Volumn.Text = ""
    txt_Tab0_srcTotal_Weight.Text = ""
    str_SQL = "Select Isnull(Round(sum(�c��),0),0) as �`�c��,Isnull(Round(sum(���q),0),0) as �`���q," & _
              "       Isnull(Round(sum(���n),0),0) as �`���n,Isnull(Round(sum(�O��),0),0) as �`�O�� " & _
              "From RCutOrders_SourceOrder  "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        txt_Tab0_srcTotal_Case.Text = tmp_Rs.Fields("�`�c��").Value
        txt_Tab0_srcTotal_Pallet.Text = tmp_Rs.Fields("�`�O��").Value
        txt_Tab0_srcTotal_Volumn.Text = tmp_Rs.Fields("�`���n").Value
        txt_Tab0_srcTotal_Weight.Text = tmp_Rs.Fields("�`���q").Value
    End If
    tmp_Rs.Close
End Sub

Private Sub ReCaculate_OrderSum()
    '�����ݱƨ��q��G�`�p��ƭ�  >>  �ثe�ݿ�C���`�p
    txt_Tab0_srcTotal_Case.Text = ""
    txt_Tab0_srcTotal_Pallet.Text = ""
    txt_Tab0_srcTotal_Volumn.Text = ""
    txt_Tab0_srcTotal_Weight.Text = ""
    
    If rs_ORT02W.RecordCount = 0 Then Exit Sub
    
    Dim dbTotalCase As Double
    Dim dbTotalPallet As Double
    Dim dbTotalWeight As Double
    Dim dbTotalVolumn As Double
    dbTotalCase = 0: dbTotalPallet = 0: dbTotalVolumn = 0: dbTotalWeight = 0
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    rs_ORT02W.MoveFirst
    Do While Not rs_ORT02W.EOF
        dbTotalCase = dbTotalCase + rs_ORT02W.Fields("�c��").Value
        dbTotalPallet = dbTotalPallet + rs_ORT02W.Fields("�O��").Value
        dbTotalVolumn = dbTotalVolumn + rs_ORT02W.Fields("���n").Value
        dbTotalWeight = dbTotalWeight + rs_ORT02W.Fields("���q").Value
        rs_ORT02W.MoveNext
    Loop
    rs_ORT02W.MoveFirst
    txt_Tab0_srcTotal_Case.Text = dbTotalCase
    txt_Tab0_srcTotal_Pallet.Text = dbTotalPallet
    txt_Tab0_srcTotal_Volumn.Text = dbTotalVolumn
    txt_Tab0_srcTotal_Weight.Text = dbTotalWeight
    
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
End Sub

Private Sub txtDeliveryDate3_Click()
    '�ƨ��@�~ >> �X�����
    If Trim(txtDeliveryDate3.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txtDeliveryDate3.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txtDeliveryDate3.Text, 4) & "/" & Mid(txtDeliveryDate3.Text, 5, 2) & "/" & Right(txtDeliveryDate3.Text, 2))
        End If
    End If
    mvDate.Left = fam_SelectedOrders.Left + txtDeliveryDate3.Left
    mvDate.Top = fam_SelectedOrders.Top + txtDeliveryDate3.Top + txtDeliveryDate3.Height
    mvDate.Tag = "�X�����"
    mvDate.Visible = True: mvDate.ZOrder
End Sub

Private Sub txtDeliveryDate3_KeyPress(KeyAscii As Integer)
    '�ƨ��@�~ > [�X�����] ��Ʈ榡�Gyyyymmdd
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '�����\��J�r��
              KeyAscii = 0
         Case vbKeyReturn
              If Fun_ChkDateFormat(txtDeliveryDate3.Text) = 1 Then
                 msg_text = "�X������G" & funRtn_msg
                 MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                 txtDeliveryDate3.SelStart = 0: txtDeliveryDate3.SelLength = Len(txtDeliveryDate3.Text): txtDeliveryDate3.SetFocus
                 Exit Sub
              Else
                 cmdRouteQuery3.SetFocus
              End If
    End Select
End Sub
