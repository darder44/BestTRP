VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_OP_DCRouteMerge 
   Caption         =   "�G���ƨ��@�~"
   ClientHeight    =   7530
   ClientLeft      =   195
   ClientTop       =   855
   ClientWidth     =   13290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15630
   ScaleWidth      =   28560
   Begin TabDlg.SSTab SSTab1 
      Height          =   7440
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   13123
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "�G���ƨ�"
      TabPicture(0)   =   "frm_OP_DCRouteMerge.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fam_SrcRoute"
      Tab(0).Control(1)=   "fam_SelectedOrders"
      Tab(0).Control(2)=   "fam_RouteData"
      Tab(0).Control(3)=   "fra_ExtraQuery"
      Tab(0).Control(4)=   "mvDate"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "�G���ƨ����s�C��"
      TabPicture(1)   =   "frm_OP_DCRouteMerge.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fam_Tab1_Delete"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dg_Tab1_Route"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dg_Tab1_RouteDC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fam_Tab1_Query"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin MSComCtl2.MonthView mvDate 
         Height          =   2220
         Left            =   -73350
         TabIndex        =   70
         Top             =   4635
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
         StartOfWeek     =   114163713
         TitleBackColor  =   -2147483646
         TitleForeColor  =   16777215
         TrailingForeColor=   -2147483643
         CurrentDate     =   38232
         MaxDate         =   2958455
      End
      Begin VB.Frame fra_ExtraQuery 
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         Caption         =   "�d�߱���]�w"
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   3600
         Begin VB.TextBox txt_FPlanDate_Start 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1005
            TabIndex        =   78
            Top             =   225
            Width           =   1125
         End
         Begin VB.TextBox txt_FPlanDate_End 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2385
            TabIndex        =   77
            Top             =   225
            Width           =   1125
         End
         Begin VB.TextBox txt_FDeliveryDate_Start 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   990
            TabIndex        =   76
            Top             =   540
            Width           =   1125
         End
         Begin VB.TextBox txt_FDeliveryDate_End 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2385
            TabIndex        =   75
            Top             =   540
            Width           =   1125
         End
         Begin VB.CheckBox chk_AddWho 
            Caption         =   "�ƨ��H���z��"
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
            Left            =   990
            TabIndex        =   73
            Top             =   885
            Value           =   1  '�֨�
            Width           =   1875
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ƨ����"
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
            Left            =   90
            TabIndex        =   82
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label3 
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
            Left            =   2145
            TabIndex        =   81
            Top             =   255
            Width           =   240
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   80
            Top             =   585
            Width           =   840
         End
         Begin VB.Label Label3 
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
            Index           =   15
            Left            =   2145
            TabIndex        =   79
            Top             =   570
            Width           =   240
         End
      End
      Begin VB.Frame fam_Tab1_Query 
         BackColor       =   &H00404000&
         Height          =   2160
         Left            =   9420
         TabIndex        =   58
         Top             =   285
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
            Left            =   120
            TabIndex        =   60
            Top             =   630
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Tab1_RouteNoQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�G���ƨ��d��"
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
            Picture         =   "frm_OP_DCRouteMerge.frx":0038
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   59
            Top             =   1230
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
            ForeColor       =   &H0080FF80&
            Height          =   240
            Left            =   465
            TabIndex        =   61
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.Frame fam_RouteData 
         Height          =   540
         Left            =   -71295
         TabIndex        =   12
         Top             =   315
         Width           =   9915
         Begin VB.CommandButton cmd_ShowQuery 
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
            Left            =   1575
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   74
            Top             =   135
            Width           =   375
         End
         Begin VB.CommandButton cmd_Tab0_CreateRoute 
            Appearance      =   0  '����
            BackColor       =   &H00FF8080&
            Caption         =   "�إߤG�����s"
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
            Left            =   7260
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   69
            Top             =   75
            Width           =   1410
         End
         Begin VB.TextBox txt_Tab0_CarCheckInDate 
            Alignment       =   2  '�m�����
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
            Left            =   4455
            TabIndex        =   66
            Top             =   105
            Width           =   1140
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
            Left            =   2490
            TabIndex        =   53
            Top             =   105
            Width           =   1215
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
            Height          =   435
            Index           =   0
            Left            =   8745
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   15
            Top             =   90
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Tab0_ImportRoute 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�פJ�@�����s"
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
            Left            =   105
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   14
            Top             =   90
            Width           =   1440
         End
         Begin VB.TextBox txt_Tab0_CarCheckInTime 
            Alignment       =   2  '�m�����
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
            Left            =   6375
            TabIndex        =   13
            Top             =   105
            Width           =   660
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   2
            Left            =   3630
            Top             =   75
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
            ForeColor       =   &H00400000&
            Height          =   435
            Index           =   20
            Left            =   3795
            TabIndex        =   67
            Top             =   120
            Width           =   675
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
            ForeColor       =   &H00400000&
            Height          =   435
            Index           =   18
            Left            =   2040
            TabIndex        =   54
            Top             =   120
            Width           =   435
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   0
            Left            =   1995
            Top             =   75
            Width           =   1740
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
            ForeColor       =   &H00400000&
            Height          =   435
            Index           =   19
            Left            =   5700
            TabIndex        =   16
            Top             =   120
            Width           =   675
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   1
            Left            =   5640
            Top             =   75
            Width           =   1425
         End
      End
      Begin VB.Frame fam_SelectedOrders 
         Height          =   2790
         Left            =   -74895
         TabIndex        =   17
         Top             =   765
         Width           =   12300
         Begin VB.TextBox txt_Tab0_DeliveryCarTypeCode 
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
            Left            =   5955
            TabIndex        =   83
            Top             =   675
            Visible         =   0   'False
            Width           =   570
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
            Height          =   420
            Left            =   7875
            TabIndex        =   68
            Top             =   180
            Width           =   780
         End
         Begin VB.CommandButton cmd_Tab0_srcRouteReset 
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
            Height          =   435
            Left            =   9990
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   57
            Top             =   2190
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_srcRouteQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�����ƨ����s�j�M"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   9990
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   56
            Top             =   1455
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_SelectedRemove_All 
            BackColor       =   &H000080FF&
            Caption         =   "�w����s����(��)"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   9975
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   55
            Top             =   630
            Width           =   1140
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
            Left            =   4320
            TabIndex        =   24
            Top             =   150
            Width           =   1320
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
            Left            =   6165
            TabIndex        =   23
            Top             =   150
            Width           =   1230
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
            Left            =   7380
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   22
            Top             =   150
            Width           =   330
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarType 
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
            Left            =   11970
            TabIndex        =   21
            Top             =   315
            Width           =   555
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
            Left            =   9600
            TabIndex        =   20
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab0_DeliveryCompany 
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
            Left            =   8760
            TabIndex        =   19
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
            Left            =   10785
            TabIndex        =   18
            Top             =   315
            Width           =   1170
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_SelectedRoute 
            Height          =   1695
            Left            =   45
            TabIndex        =   25
            Top             =   1035
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   2990
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
         Begin VB.Frame Frame1 
            Enabled         =   0   'False
            Height          =   525
            Left            =   15
            TabIndex        =   44
            Top             =   525
            Width           =   5895
            Begin VB.TextBox txt_Tab0_Selected_Weight 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   4845
               TabIndex        =   48
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Volumn 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   3555
               TabIndex        =   47
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Pallet 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   2265
               TabIndex        =   46
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Case 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   1005
               TabIndex        =   45
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
               TabIndex        =   52
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
               Left            =   1890
               TabIndex        =   51
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
               Left            =   3165
               TabIndex        =   50
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
               Left            =   4470
               TabIndex        =   49
               Top             =   210
               Width           =   360
            End
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
            Left            =   5670
            TabIndex        =   30
            Top             =   165
            Width           =   420
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
            Left            =   3840
            TabIndex        =   31
            Top             =   150
            Width           =   435
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00400000&
            FillStyle       =   0  '���
            Height          =   1350
            Left            =   9930
            Top             =   1380
            Width           =   1245
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
            Left            =   11940
            TabIndex        =   29
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
            Left            =   9885
            TabIndex        =   28
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
            Left            =   8760
            TabIndex        =   27
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
            Left            =   11145
            TabIndex        =   26
            Top             =   120
            Width           =   540
         End
      End
      Begin VB.Frame fam_SrcRoute 
         Height          =   3795
         Left            =   -74895
         TabIndex        =   1
         Top             =   3585
         Width           =   12300
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
            Left            =   5985
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   43
            Top             =   135
            Width           =   345
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
            Left            =   6405
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   42
            Top             =   135
            Width           =   345
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
            Left            =   7245
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   41
            Top             =   135
            Width           =   1260
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel_All 
            BackColor       =   &H00FF80FF&
            Caption         =   "�ݿ����(��)"
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
            Left            =   8520
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   40
            Top             =   135
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Case 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   38
            Top             =   795
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Pallet 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   36
            Top             =   1365
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Volumn 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   34
            Top             =   1950
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab0_srcTotal_Weight 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   10005
            TabIndex        =   32
            Top             =   2505
            Width           =   1080
         End
         Begin MSDataGridLib.DataGrid dg_TRP01T 
            Height          =   2475
            Left            =   60
            TabIndex        =   11
            Top             =   525
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   4366
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
         Begin VB.Frame fam_SelectedSum 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   5895
            Begin VB.TextBox txt_Tab0_srcSelected_Weight 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   4830
               TabIndex        =   6
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Volumn 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   3555
               TabIndex        =   5
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Pallet 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   2265
               TabIndex        =   4
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Case 
               Alignment       =   1  '�a�k���
               Appearance      =   0  '����
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   1005
               TabIndex        =   3
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
               TabIndex        =   10
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
               Left            =   1890
               TabIndex        =   9
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
               Left            =   3165
               TabIndex        =   8
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
               Left            =   4440
               TabIndex        =   7
               Top             =   195
               Width           =   360
            End
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_Orders 
            Height          =   690
            Left            =   60
            TabIndex        =   71
            Top             =   3045
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   1217
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
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���t�w�X���T�{"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   21
            Left            =   10320
            TabIndex        =   84
            Top             =   240
            Width           =   1260
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   450
            Left            =   7200
            Top             =   90
            Width           =   2925
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '���z��
            Height          =   435
            Left            =   5940
            Top             =   105
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�c    ��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   11
            Left            =   10020
            TabIndex        =   39
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�O    ��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   10
            Left            =   10020
            TabIndex        =   37
            Top             =   1170
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��    �n"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   9
            Left            =   10020
            TabIndex        =   35
            Top             =   1755
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��    �q"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   8
            Left            =   10020
            TabIndex        =   33
            Top             =   2310
            Width           =   540
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_RouteDC 
         Height          =   3480
         Left            =   105
         TabIndex        =   62
         Top             =   3510
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   6138
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
         Left            =   105
         TabIndex        =   63
         Top             =   360
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
      Begin VB.Frame fam_Tab1_Delete 
         Appearance      =   0  '����
         BackColor       =   &H00000040&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   9450
         TabIndex        =   64
         Top             =   2385
         Width           =   1965
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
            Left            =   90
            Picture         =   "frm_OP_DCRouteMerge.frx":0342
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   65
            ToolTipText     =   "�R��"
            Top             =   180
            Width           =   1785
         End
      End
   End
End
Attribute VB_Name = "frm_OP_DCRouteMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private blTRP01TEventEnable As Boolean              '�ݿ�������ƨ����s Event Ĳ�o���ı���
Private blTab0SelectedRouteEventEnable As Boolean   '�w��������ƨ����s�� Event Ĳ�o���ı���
Private blTab1RouteEventEnable As Boolean           '�G���ƨ����s�� Event Ĳ�o���ı���

Private rs_TRP01T As ADODB.Recordset                '�G���ƨ��@�~�G�פJ�������ƨ�{DC}���u�s��
Private rs_Tab0_SelectedRoute As ADODB.Recordset    '�w������i��G���ƨ�{�֨�}�������ƨ�{DC}���u�s��
Private rs_Tab0_Orders As ADODB.Recordset           '���s�������q��
Private rs_Tab1_Route As ADODB.Recordset            '�����ƨ�{DC}���s�i��G���ƨ����ͤ����u�s��
Private rs_Tab1_RouteDC As ADODB.Recordset          '�����ƨ����u�s��

Private strSourceFilter As String        '�ݱƨ��������ƨ����s�z��
Private strSourceOrderBy As String       '�ݱƨ����@���ƨ����s�ƧǤ覡
Private dbsrcSelected_Case As Double     '�����ƨ��� [DC] ���s�G����c��
Private dbsrcSelected_Pallet As Double   '�����ƨ��� [DC] ���s�G����O��
Private dbsrcSelected_Volumn As Double   '�����ƨ��� [DC] ���s�G������n
Private dbsrcSelected_Weight As Double   '�����ƨ��� [DC] ���s�G������q
Private dbSelectedCount As Double        '��������ƨ����s����

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_ShowQuery_Click()
'�G���ƨ� >> �פJ�@���ƨ����s >> �d�߱���
fra_ExtraQuery.Visible = Not fra_ExtraQuery.Visible
End Sub

Private Sub cmd_Tab0_CreateRoute_Click()
'�G���ƨ�  >> �إߤG���ƨ����s
Dim Str_RouteNo As String
Dim str_FirstRouteNo As String
str_FirstRouteNo = ""
Str_RouteNo = ""
If rs_Tab0_SelectedRoute.RecordCount = 0 Then Exit Sub

'add by Terry 20190614 �ˬd���s�O�_�w�զ��G�����s
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
    str_FirstRouteNo = str_FirstRouteNo & "'" & rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value & "',"
    rs_Tab0_SelectedRoute.MoveNext
Loop

str_FirstRouteNo = str_FirstRouteNo & "''"
rs_Tab0_SelectedRoute.MoveFirst
str_SQL = "select c_route_no from trp01t where route_no in (" & str_FirstRouteNo & ") and c_route_no is not null"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
    MsgBox ("���@�����s�w�զ��G�����s�A�Э��s���J�@�����s�òM��[�w������@�����s]"), vbOKOnly + vbCritical
    tmp_Rs.Close
    Exit Sub
End If
tmp_Rs.Close


'add by Eric 20141211 �ˬd�O�_�w�g�Q�X���T�{
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
    Str_RouteNo = Str_RouteNo & "'" & rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value & "',"
    rs_Tab0_SelectedRoute.MoveNext
Loop
rs_Tab0_SelectedRoute.MoveFirst
str_SQL = "select route_no from trp05t t5 where t5.sdnstatus = 1 and t5.route_no in (" & Mid(Str_RouteNo, 1, Len(Str_RouteNo) - 1) & ")"

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Str_RouteNo = ""
If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    Do While Not tmp_Rs.EOF
        '���w�X�������u�s��
        Str_RouteNo = Str_RouteNo & tmp_Rs.Fields("route_no") & " , "
        tmp_Rs.MoveNext
    Loop
    msg_text = "�o�{�����s�w�g�X���T�{�A�нT�{�@�����s�O�_�w�Q�X���C" & Chr(13) + Chr(10) & "�í��s���J�@�����s�i��G�����s�@�~"
    MsgBox msg_text, vbOKOnly + vbCritical, msg_title

    msg_text = "�w�X�������u�s��:" & Chr(13) + Chr(10) & Str_RouteNo
    MsgBox msg_text, vbOKOnly + vbCritical, msg_title
    tmp_Rs.Close
    Exit Sub
Else
    tmp_Rs.Close
End If

If Len(Trim(txt_Tab0_TRPDate.Text)) = 0 Then
   msg_text = "��ƿ��~�G����J�X�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_TRPDate.SetFocus
   Exit Sub
End If
If Len(Trim(txt_Tab0_DeliveryCarNo.Text)) = 0 Then
   msg_text = "��ƿ��~�G����J���P���X"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_DeliveryCarNo.SetFocus
   Exit Sub
End If

'����ˮ�

'a.�X������G�榡 yyyymmdd
txt_Tab0_TRPDate.Text = Trim(txt_Tab0_TRPDate.Text)
If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
   msg_text = "�X������G" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
   Exit Sub
End If
'a2.�X����� >= ����
If txt_Tab0_TRPDate.Text < Format(Now, "yyyymmdd") Then
   msg_text = "�X��������o�p�󤵤�"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
   Exit Sub
End If

'b.�ˮ� [���P���X] �O�_����
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

txt_Tab0_DeliveryCarNo.Text = Trim(txt_Tab0_DeliveryCarNo.Text)
str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "��ƿ��~�G���P���X " & txt_Tab0_DeliveryCarNo.Text & " ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
   txt_Tab0_DeliveryCarNo.SetFocus
   Exit Sub
End If
Call ReDim_Recordset(tmp_Rs)

'�ˬd�i�����q
Dim intableWT, intableCBM As Long
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

'���w�X�Y�Ȧs�G������J
txt_Tab0_DockNo.Text = Trim(txt_Tab0_DockNo.Text)
If Len(Trim(txt_Tab0_DockNo.Text)) = 0 Then
   msg_text = "��ƿ��~�G[�X�Y�Ȧs] ������J"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_DockNo.SetFocus
End If

'�w�p������
txt_Tab0_CarCheckInDate.Text = Trim(txt_Tab0_CarCheckInDate.Text)
If Len(Trim(txt_Tab0_CarCheckInDate.Text)) <> 8 Then
   msg_text = "�w�p�������G��Ʈ榡 yyyymmdd "
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
   txt_Tab0_CarCheckInDate.SetFocus
   Exit Sub
End If

If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
   msg_text = "�w�p�������G��ƿ��~ yyyymmdd�A" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
   txt_Tab0_CarCheckInDate.SetFocus
   Exit Sub
End If

'�w�p������ >= ����
If txt_Tab0_CarCheckInDate.Text < Format(Now, "yyyymmdd") Then
   msg_text = "�w�p���������o�p�󤵤�"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text): txt_Tab0_CarCheckInDate.SetFocus
   Exit Sub
End If

'�w�p����ɶ�
txt_Tab0_CarCheckInTime.Text = Trim(txt_Tab0_CarCheckInTime.Text)
If Len(Trim(txt_Tab0_CarCheckInTime.Text)) <> 4 Then
   msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
   txt_Tab0_CarCheckInTime.SetFocus
   Exit Sub
End If
Select Case Left(txt_Tab0_CarCheckInTime.Text, 2)
       Case "00" To "24"
       Case Else
            msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
            txt_Tab0_CarCheckInTime.SetFocus
            Exit Sub
End Select
Select Case Right(txt_Tab0_CarCheckInTime.Text, 2)
       Case "00" To "59"
       Case Else
            msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
            txt_Tab0_CarCheckInTime.SetFocus
            Exit Sub
End Select

On Error GoTo err_Handle
Tran_Level = 0
Tran_Level = cn.BeginTrans

Dim intDriveTimes As Integer    '����
Dim strRouteNo As String        '���u�s��

'1.���ͨ���
str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
          "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
tmp_Rs.Close

'2.���͸��u�s��
str_SQL = "Select Isnull(Max(Cast(Right(Route_No,3) as integer))+1,1) as RouteSN " & _
          "From TRP01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'S'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
strRouteNo = "S" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
tmp_Rs.Close

'3.Insert into TRP01T ���u�s���D��
'  TRP01T.EXE_CONFIRM = '0' �s�إ߸��u�s���A�|���^�ǹL exe
str_SQL = "Insert into TRP01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
          strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
          txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'3-1.��s�`�p����
txt_Tab0_srcTotal_Case.Text = Val(txt_Tab0_srcTotal_Case.Text) - Val(txt_Tab0_Selected_Case.Text)
txt_Tab0_srcTotal_Pallet.Text = Val(txt_Tab0_srcTotal_Pallet.Text) - Val(txt_Tab0_Selected_Pallet.Text)
txt_Tab0_srcTotal_Volumn.Text = Val(txt_Tab0_srcTotal_Volumn.Text) - Val(txt_Tab0_Selected_Volumn.Text)
txt_Tab0_srcTotal_Weight.Text = Val(txt_Tab0_srcTotal_Weight.Text) - Val(txt_Tab0_Selected_Weight.Text)

'4.insert into TRP05T �����i�X�޲z
str_SQL = "Insert into TRP05T (Route_No,Vehicle_ID_No,Drive_Times,Delivery_Date,Valid_Vehicle,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Dock_No,Expect_Time,Expect_Date) Values ('" & _
          strRouteNo & "','" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",'" & _
          Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "','1'," & _
          txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'" & _
          txt_Tab0_DockNo.Text & "','" & txt_Tab0_CarCheckInTime.Text & "','" & txt_Tab0_CarCheckInDate.Text & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�Ѩ����D�ɧ�s�����������
str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
          "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strRouteNo & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�g�� SSTab1.Tab 1 [���u�s���C��]
blTab1RouteEventEnable = False
rs_Tab1_Route.AddNew
rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
rs_Tab1_Route.Fields("�G���ƨ����s").Value = strRouteNo
rs_Tab1_Route.Fields("�X�����").Value = txt_Tab0_TRPDate.Text
rs_Tab1_Route.Fields("���P���X").Value = txt_Tab0_DeliveryCarNo.Text
rs_Tab1_Route.Fields("����").Value = intDriveTimes
rs_Tab1_Route.Fields("�r�p�H").Value = txt_Tab0_DeliveryDriver.Text
rs_Tab1_Route.Fields("�c��").Value = txt_Tab0_Selected_Case.Text
rs_Tab1_Route.Fields("�O��").Value = txt_Tab0_Selected_Pallet.Text
rs_Tab1_Route.Fields("���n").Value = txt_Tab0_Selected_Volumn.Text
rs_Tab1_Route.Fields("���q").Value = txt_Tab0_Selected_Weight.Text
rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = txt_Tab0_DockNo.Text
rs_Tab1_Route.Fields("�w�p������").Value = txt_Tab0_CarCheckInDate.Text
rs_Tab1_Route.Fields("�w�p����ɶ�").Value = txt_Tab0_CarCheckInTime.Text
rs_Tab1_Route.Fields("����").Value = txt_Tab0_DeliveryCarTypeCode.Text
rs_Tab1_Route.Fields("�ƨ���").Value = User_id
rs_Tab1_Route.Update
blTab1RouteEventEnable = True

'5.update TRP01T & TRP05T [�����ƨ����s & �����ƨ������ި�]
'  �g�� SSTab1.Tab 1 [�G���ƨ����u�s�����ݤ������ƨ����u�s��]
blTab0SelectedRouteEventEnable = False
rs_Tab1_RouteDC.Filter = adFilterNone
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
   'UPDATE TRP01T
   str_SQL = "Update TRP01T Set C_ROUTE_NO = '" & strRouteNo & "',C_VEHICLE_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "',C_DRIVE_TIMES = " & intDriveTimes & " " & _
             "Where Route_No = '" & rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Update TRP05T Set C_ROUTE_NO = '" & strRouteNo & "',C_VEHICLE_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "',C_DRIVE_TIMES = " & intDriveTimes & " " & _
             "Where Route_No = '" & rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value & "' "
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�g�� SSTab1.Tab 1 [�G���ƨ����s���ݤ������ƨ����u�s���C��]
   rs_Tab1_RouteDC.AddNew
   rs_Tab1_RouteDC.Fields("�s��").Value = rs_Tab1_RouteDC.RecordCount
   rs_Tab1_RouteDC.Fields("�G���ƨ����s").Value = strRouteNo
   rs_Tab1_RouteDC.Fields("�����ƨ����s").Value = rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value
   rs_Tab1_RouteDC.Fields("�X�����").Value = rs_Tab0_SelectedRoute.Fields("�X�����").Value
   rs_Tab1_RouteDC.Fields("���P���X").Value = rs_Tab0_SelectedRoute.Fields("���P���X").Value
   rs_Tab1_RouteDC.Fields("����").Value = rs_Tab0_SelectedRoute.Fields("����").Value
   rs_Tab1_RouteDC.Fields("�r�p�H").Value = rs_Tab0_SelectedRoute.Fields("�r�p�H").Value
   rs_Tab1_RouteDC.Fields("�c��").Value = rs_Tab0_SelectedRoute.Fields("�c��").Value
   rs_Tab1_RouteDC.Fields("�O��").Value = rs_Tab0_SelectedRoute.Fields("�O��").Value
   rs_Tab1_RouteDC.Fields("���n").Value = rs_Tab0_SelectedRoute.Fields("���n").Value
   rs_Tab1_RouteDC.Fields("���q").Value = rs_Tab0_SelectedRoute.Fields("���q").Value
   rs_Tab1_RouteDC.Fields("����").Value = rs_Tab0_SelectedRoute.Fields("����").Value
   rs_Tab1_RouteDC.Update
   rs_Tab0_SelectedRoute.MoveNext
Loop


'��f�l��APP�R���P�g�J
'cn.Execute "update apporderdate set status = 'C',editdate = getdate() where c_route_no = '" & strRouteNo & "' ", RowsAffect, adExecuteNoRecords
cn.Execute "delete apporderdate where receipt_no in (select t2.receipt_no from trp02t t2 join trp01t t1 on t1.route_no = t2.route_no and t1.c_route_no = '" & strRouteNo & "') ", RowsAffect, adExecuteNoRecords

str_SQL = "insert into apporderdate(wh,C_Route_no,C_VEHICLE_ID_NO,Priority,Receipt_no,OrderGroup,Storerkey,Arrive_date,Company,Status) " & _
    "select WH = 'GYDC' ,C_Route_no = t1.C_route_no " & _
    ",C_VEHICLE_ID_NO = isnull(t1.C_VEHICLE_ID_NO,t2.VEHICLE_ID_NO) " & _
    ",Priority = t2.Priority " & _
    ",Receipt_no = t2.receipt_no " & _
    ",OrderGroup = t1m.address " & _
    ",Storerkey = rtrim(t16m.storerkey) + '_' + t16m.short_name " & _
    ",Arrive_date = convert(char(8),t2.arrive_date,112) " & _
    ",Company = t1m.short_name " & _
    ",Status = '0' " & _
    "from trp02t t2 join trp01t t1 on t1.route_no = t2.route_no and c_route_no is not null and t1.C_ROUTE_NO = '" & strRouteNo & "' " & _
    "join trp01m t1m on t1m.storerkey = t2.storerkey and t2.consigneekey = t1m.consigneekey " & _
    "join trp16m t16m on t16m.storerkey = t2.storerkey "
    
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

If dg_Tab1_Route.SelBookmarks.Count > 0 Then
   dg_Tab1_Route.SelBookmarks.Remove 0
End If
dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
rs_Tab1_RouteDC.Filter = " �G���ƨ����s = '" & rs_Tab1_Route.Fields("�G���ƨ����s").Value & "'"

blTab0SelectedRouteEventEnable = True

'5.�M�� [�w����������ƨ����u�s���C��]
blTab0SelectedRouteEventEnable = False
'�ƨ��@�~�G�w����������ƨ����u�s���C�� DBGrid �榡�]�w-ReSet
Call CreateRS_Tab0_SelectedRoute
'���s�p��w��������ƨ����u�s���G�c�ơA�O�ơA���n�A���q + �s�����s����
Call Calculate_SelectedRoute
blTab0SelectedRouteEventEnable = True

'6.�M���ƨ��@�~����
txt_Tab0_DockNo.Text = ""               '�X�Y�Ȧs
txt_Tab0_CarCheckInDate.Text = ""       '�����w�p����ɶ�
txt_Tab0_CarCheckInTime.Text = ""       '�����w�p����ɶ�
txt_Tab0_TRPDate.Text = ""              '�X�����
txt_Tab0_DeliveryCarNo.Text = ""        '���P���X
txt_Tab0_DeliveryCompany.Text = ""      '�B�餽�q
txt_Tab0_DeliveryDriver.Text = ""       '�r�p�H
txt_Tab0_DeliveryPhone.Text = ""        '�q��
txt_Tab0_DeliveryCarType.Text = ""      '����
txt_Tab0_DeliveryCarTypeCode.Text = ""  '���إN�X

SSTab1.Tab = 1
DoEvents: DoEvents

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
   rs_Tab1_Route.Filter = "�G���ƨ����s='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_RouteDC.Filter = "�G���ƨ����s='" & strRouteNo & "'"
   If Not rs_Tab1_RouteDC.EOF Then
      Do While Not rs_Tab1_RouteDC.EOF
         rs_Tab1_RouteDC.Delete
         rs_Tab1_RouteDC.MoveFirst
      Loop
   End If
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
      
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�G���ƨ�-�إߤG���ƨ����s", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_ImportRoute_Click()
'�ƨ��@�~ >> �פJ�����ƨ����u�s��
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_TRP01T.DataSource = Nothing
Set rs_TRP01T = Nothing
fra_ExtraQuery.Visible = False
strSourceFilter = adFilterNone
DoEvents

'���w��������ƨ����s�̡G�߰� user �O�_�n�M��
If rs_Tab0_SelectedRoute.RecordCount <> 0 Then
   msg_text = "[�w����������ƨ����u�s��] �O�_�i��M��"
   If MsgBox(msg_text, vbYesNo + vbInformation + vbDefaultButton2, msg_title) = vbYes Then
      '�֨��@�~�G�w����������ƨ����u�s���C�� DBGrid �榡�]�w
      Call CreateRS_Tab0_SelectedRoute
      '�M�����G�֭p����������ƨ����s�G�p�p�k 0
      txt_Tab0_Selected_Case.Text = ""
      txt_Tab0_Selected_Pallet.Text = ""
      txt_Tab0_Selected_Volumn.Text = ""
      txt_Tab0_Selected_Weight.Text = ""
   End If
End If

'�����ƨ����s�G����p�p�G�k�s
dbSelectedCount = 0
dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""

Dim str_SQL2 As String
'���^�����ƨ����u�s��
str_SQL = "Select " & _
        "' ' as '��' " & _
        ",T1.ROUTE_NO as �����ƨ����s " & _
        ",Isnull(Rtrim(T5.Dock_No),'') as �X�Y " & _
        ",Round(T1.Case_cnt,2) as �c�� " & _
        ",Round(T1.Pallet_Qty,2) as �O�� " & _
        ",Round(T1.Volumn_Weight,2) as ���n " & _
        ",Round(T1.Weight,2) as ���q " & _
        ",Rtrim(T5.VEHICLE_ID_NO) as ���P���X " & _
        ",T5.Drive_Times as ���� " & _
        ",Rtrim(Isnull(T5.Driver,'')) as �r�p�H " & _
        ",Convert(varchar , T1.Delivery_Date,112) as �X����� " & _
        ",Rtrim(Isnull(a1.Vehicle_Type,'')) as ���� " & _
        ",Case T1.EXE_Confirm When '0' Then '�s�ظ��s' When '1' Then '�]�w�^��' When '2' Then '�w�^��' When '9' Then '�w���z�f' else '�������A' End  AS EXE�^�� " & _
        ",cast(' ' as char(300)) as �Ȥ� " & _
        ",Rtrim(Isnull(T1.AddWho,'')) as �ƨ��� " & _
        "From TRP01T T1 " & _
        "inner join TRP05T T5  on T1.ROUTE_NO=T5.ROUTE_NO and sdnstatus = 0 " & _
        "inner join TRP09M A1 on A1.Vehicle_ID_No = T5.Vehicle_ID_No " & _
        "Where Left(T1.ROUTE_NO,1) <> 'S'  and T1.ROUTE_NO <> 'D' and rtrim(isnull(T1.C_ROUTE_NO,''))='' and T5.Valid_Vehicle = '1'  and 1=1 "
        
Dim str_Where As String, intloop As Integer
str_Where = ""

'�ƨ����
If Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   str_Where = "and Convert(varchar,T1.AddDate,112) Between '" & txt_FPlanDate_Start.Text & "' and '" & txt_FPlanDate_End.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) = 0 Then
   str_Where = "and Convert(varchar,T1.AddDate,112) = '" & txt_FPlanDate_Start.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) = 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   str_Where = "and  Convert(varchar,T1.AddDate,112) = '" & txt_FPlanDate_End.Text & "' "
End If

'�X�����
If Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   str_Where = str_Where & "and Convert(varchar , T1.Delivery_Date,112) Between '" & txt_FDeliveryDate_Start.Text & "' and '" & txt_FDeliveryDate_End.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) = 0 Then
   str_Where = "and Convert(varchar , T1.Delivery_Date,112) = '" & txt_FDeliveryDate_Start.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) = 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   str_Where = "and Convert(varchar , T1.Delivery_Date,112) = '" & txt_FDeliveryDate_End.Text & "' "
End If

'�ƨ��H���z��
If chk_AddWho.Value = vbChecked Then str_Where = str_Where & "and Rtrim(Isnull(T1.AddWho,'')) = '" & User_id & "' "

'str_Where = str_Where & "and EXE�^�� <> '�w�^��' "

str_SQL = str_SQL & str_Where & "Order by T1.ROUTE_NO "
'str_SQL2 = str_SQL2 & str_Where & " "
          
strSourceOrderBy = " �����ƨ����s "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧣����ƨ����u�s�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_TRP01T)
tmp_Rs.Close

With dg_TRP01T
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_TRP01T.MoveFirst
blTRP01TEventEnable = False
Set dg_TRP01T.DataSource = rs_TRP01T
With dg_TRP01T
    .RowHeight = 250
    .Columns(0).Width = 500         '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 350         '����ѧO���
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1200        '�����ƨ����u�s��
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 500        '�����ƨ����u�Ȧs�X�Y
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 800         '�c��
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 800         '�O��
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 800         '���n
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800         '���q
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 900         '���P���X
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 500         '����
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 700         '�r�p�H
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1000       '�X�����
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 500       '����
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 900       'EXE�^��
    .Columns(13).Alignment = dbgLeft
End With

'��ܫȤ�W��
rs_TRP01T.MoveFirst
Do While Not rs_TRP01T.EOF
    
    str_SQL = "select distinct t1m.short_name as short_name from trp02t t2 join trp01m t1m on t2.consigneekey = t1m.consigneekey and t2.storerkey = t1m.storerkey and t2.route_no = '" & rs_TRP01T("�����ƨ����s") & "' order by short_name"
    
    tmp_Rs.Open str_SQL, cn
    If Not tmp_Rs.EOF Then
        tmp_Rs.MoveFirst
        Do While Not tmp_Rs.EOF
            rs_TRP01T("�Ȥ�") = Trim(rs_TRP01T("�Ȥ�")) & tmp_Rs("short_name") & ","
        tmp_Rs.MoveNext
        Loop
    End If
    tmp_Rs.Close
    
    '�����ݶi��G���ƨ����@���ƨ����u�s���G�`�p��ƭ�
    txt_Tab0_srcTotal_Case.Text = Val(txt_Tab0_srcTotal_Case.Text) + Val(rs_TRP01T.Fields("�c��").Value)
    txt_Tab0_srcTotal_Pallet.Text = Val(txt_Tab0_srcTotal_Pallet.Text) + Val(rs_TRP01T.Fields("�O��").Value)
    txt_Tab0_srcTotal_Volumn.Text = Val(txt_Tab0_srcTotal_Volumn.Text) + Val(rs_TRP01T.Fields("���n").Value)
    txt_Tab0_srcTotal_Weight.Text = Val(txt_Tab0_srcTotal_Weight.Text) + Val(rs_TRP01T.Fields("���q").Value)
    
rs_TRP01T.MoveNext
Loop

rs_TRP01T.MoveFirst

blTRP01TEventEnable = True


'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'cn.CommandTimeout = 0   '�L��������
'tmp_Rs.Open str_SQL2, cn, adOpenForwardOnly, adLockReadOnly
'If Not tmp_Rs.EOF Then
'   txt_Tab0_srcTotal_Case.Text = tmp_Rs.Fields("�`�c��").Value
'   txt_Tab0_srcTotal_Pallet.Text = tmp_Rs.Fields("�`�O��").Value
'   txt_Tab0_srcTotal_Volumn.Text = tmp_Rs.Fields("�`���n").Value
'   txt_Tab0_srcTotal_Weight.Text = tmp_Rs.Fields("�`���q").Value
'End If
'tmp_Rs.Close
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�G���ƨ�-�����ƨ����s�פJ", Me.Caption, "cmd_Tab0_ImportRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Remove_Click()
'�G���ƨ��@�~ >> �� �w��������ƨ����s����
If rs_TRP01T Is Nothing Then Exit Sub
If rs_Tab0_SelectedRoute Is Nothing Then Exit Sub
'�w��������ƨ����s�Y�L�ϥտ���GDisable �w��������ʧ@�A����~�R
If dg_Tab0_SelectedRoute.SelBookmarks.Count = 0 Then Exit Sub

blTab0SelectedRouteEventEnable = False

'�������������ƨ����u�s��
Dim strRouteNo As String
strRouteNo = rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value
   
'�N���R���� [�w��������ƨ����u�s��] �[�J [�ݤG���ƨ��������ƨ����u�s��]
Call SelectedRoute_Removeto_TRP01T(strRouteNo)
'���s���� [�ݤG���ƨ����u�s��] �� [�s��] ����
Call ReSet_TRP01T_SeqNo

'�R���ϥտ���������ƨ����s�G�w��������ƨ����s����
rs_Tab0_SelectedRoute.Delete
If Not rs_Tab0_SelectedRoute.EOF Then rs_Tab0_SelectedRoute.MoveFirst
If dg_Tab0_SelectedRoute.SelBookmarks.Count > 0 Then dg_Tab0_SelectedRoute.SelBookmarks.Remove 0
'���s�p��w��������ƨ����s�G�c�ơA�O�ơA���n�A���q + �s�����s����
Call Calculate_SelectedRoute
blTab0SelectedRouteEventEnable = True

'�٭� [�z��] �P [�Ƨ�] ���]�w��
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = strSourceOrderBy
blTRP01TEventEnable = True

'���s�p�� [�ݱƨ��@���ƨ����s] �`�p
Call ReCaculate_FirstRouteSum

End Sub

Private Sub cmd_Tab0_SelectCar_Click()
'DC���s�֨� >> �q�����
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
'�����ƨ������s�G���

'�����ƨ����s�G����p�p�G�k�s
dbSelectedCount = 0
dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""

'�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj

'�z��w�����
rs_TRP01T.Filter = "��='V'"
If Not rs_TRP01T.EOF Then
   dg_Tab0_SelectedRoute.Visible = False
   blTab0SelectedRouteEventEnable = False
   Do While Not rs_TRP01T.EOF
      '�P�_�O�_�w�g����L
      rs_Tab0_SelectedRoute.Filter = adFilterNone
      rs_Tab0_SelectedRoute.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
      rs_Tab0_SelectedRoute.Filter = "�����ƨ����s = '" & rs_TRP01T.Fields("�����ƨ����s").Value & "'"
      If rs_Tab0_SelectedRoute.EOF Then
         '�s�W����������ƨ����u�s��
         rs_Tab0_SelectedRoute.AddNew
         rs_Tab0_SelectedRoute.Fields("�s��").Value = 999
         rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value = rs_TRP01T.Fields("�����ƨ����s").Value
         rs_Tab0_SelectedRoute.Fields("�c��").Value = rs_TRP01T.Fields("�c��").Value
         rs_Tab0_SelectedRoute.Fields("�O��").Value = rs_TRP01T.Fields("�O��").Value
         rs_Tab0_SelectedRoute.Fields("���n").Value = rs_TRP01T.Fields("���n").Value
         rs_Tab0_SelectedRoute.Fields("���q").Value = rs_TRP01T.Fields("���q").Value
         rs_Tab0_SelectedRoute.Fields("���P���X").Value = rs_TRP01T.Fields("���P���X").Value
         rs_Tab0_SelectedRoute.Fields("����").Value = rs_TRP01T.Fields("����").Value
         rs_Tab0_SelectedRoute.Fields("�r�p�H").Value = rs_TRP01T.Fields("�r�p�H").Value
         rs_Tab0_SelectedRoute.Fields("����").Value = rs_TRP01T.Fields("����").Value
         rs_Tab0_SelectedRoute.Fields("�X�����").Value = rs_TRP01T.Fields("�X�����").Value
         rs_Tab0_SelectedRoute.Update
      Else
         '��s��蠟�����ƨ����
         rs_Tab0_SelectedRoute.Fields("�s��").Value = 999
         rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value = rs_TRP01T.Fields("�����ƨ����s").Value
         rs_Tab0_SelectedRoute.Fields("�c��").Value = rs_TRP01T.Fields("�c��").Value
         rs_Tab0_SelectedRoute.Fields("�O��").Value = rs_TRP01T.Fields("�O��").Value
         rs_Tab0_SelectedRoute.Fields("���n").Value = rs_TRP01T.Fields("���n").Value
         rs_Tab0_SelectedRoute.Fields("���q").Value = rs_TRP01T.Fields("���q").Value
         rs_Tab0_SelectedRoute.Fields("���P���X").Value = rs_TRP01T.Fields("���P���X").Value
         rs_Tab0_SelectedRoute.Fields("����").Value = rs_TRP01T.Fields("����").Value
         rs_Tab0_SelectedRoute.Fields("�r�p�H").Value = rs_TRP01T.Fields("�r�p�H").Value
         rs_Tab0_SelectedRoute.Fields("����").Value = rs_TRP01T.Fields("����").Value
         rs_Tab0_SelectedRoute.Fields("�X�����").Value = rs_TRP01T.Fields("�X�����").Value
      End If
      rs_TRP01T.MoveNext
   Loop
   '���s�� [�w��������ƨ����s] ���� [�s��] �P������Ʋέp�G�c�ơA�O�ơA���n�A���q
   Call Calculate_SelectedRoute
   dg_Tab0_SelectedRoute.Visible = True
   blTab0SelectedRouteEventEnable = True
   
   '[�ݿ�������ƨ����s] ���A�R���w����������ƨ����u�s��
   rs_TRP01T.MoveFirst
   Do While Not rs_TRP01T.EOF
      rs_TRP01T.Delete
      rs_TRP01T.MoveFirst
   Loop
   
End If
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone

rs_TRP01T.Sort = strSourceOrderBy '�M�αƧ�
'�����ϥտ�����A
If dg_TRP01T.SelBookmarks.Count > 0 Then
   dg_TRP01T.SelBookmarks.Remove 0
End If

'�M�����s���q�����
Set dg_Tab0_Orders.DataSource = Nothing
Set rs_Tab0_Orders = Nothing

blTRP01TEventEnable = True

'���s�p�� [�ݱƨ��@���ƨ����s] �`�p
Call ReCaculate_FirstRouteSum


End Sub

Private Sub cmd_Tab0_SelectedCancel_All_Click()
'�G���ƨ� >> X�ݿ��������

'�����ƨ����s�G����p�p�G�k�s
dbSelectedCount = 0
dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""

'�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj

'�z��w�����
rs_TRP01T.Filter = "��='V'"
If Not rs_TRP01T.EOF Then
   Do While Not rs_TRP01T.EOF
      rs_TRP01T.Fields("��").Value = " "
      rs_TRP01T.MoveNext
   Loop
End If
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then
   rs_TRP01T.Filter = adFilterNone
End If
rs_TRP01T.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
'�����ϥտ�����A
If dg_TRP01T.SelBookmarks.Count > 0 Then
   dg_TRP01T.SelBookmarks.Remove 0
End If
blTRP01TEventEnable = True

End Sub

Private Sub cmd_Tab0_SelectedCancel_Click()
'�G���ƨ� >> X�ݿ����
If rs_TRP01T Is Nothing Then Exit Sub
'�ݿ�������ƨ����s�Y�L�ϥտ���GDisable �ݿ�����A����~�R
If dg_TRP01T.SelBookmarks.Count = 0 Then Exit Sub

If Trim(rs_TRP01T.Fields(1).Value) = "V" Then
   dbSelectedCount = dbSelectedCount - 1
   rs_TRP01T.Fields(1).Value = " "
   '�ݿ�w�����ƨ����s�G����p�p��s
   If dbSelectedCount <> 0 Then
      dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP01T.Fields("�c��").Value
      dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP01T.Fields("�O��").Value
      dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP01T.Fields("���n").Value
      dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP01T.Fields("���q").Value
   Else
      dbsrcSelected_Case = 0
      dbsrcSelected_Pallet = 0
      dbsrcSelected_Volumn = 0
      dbsrcSelected_Weight = 0
   End If
   txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
   txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
   '�����ϥտ�����A
   If dg_TRP01T.SelBookmarks.Count > 0 Then
      dg_TRP01T.SelBookmarks.Remove 0
   End If
End If
'�M�� [�z��] �P [�Ƨ�] �]�w��
blTRP01TEventEnable = False
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = strSourceOrderBy
blTRP01TEventEnable = True
End Sub

Private Sub cmd_Tab0_SelectedRemove_All_Click()
'�G���ƨ� >> �� �w��������ƨ����s����-����
If rs_TRP01T Is Nothing Then Exit Sub
If rs_Tab0_SelectedRoute Is Nothing Then Exit Sub
If rs_Tab0_SelectedRoute.RecordCount = 0 Then Exit Sub

blTab0SelectedRouteEventEnable = False

'�������������ƨ����u�s��
Dim strRouteNo As String
'�v���g�^ [�����ƨ����s TRP01T]
rs_Tab0_SelectedRoute.Filter = adFilterNone
rs_Tab0_SelectedRoute.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
rs_Tab0_SelectedRoute.MoveFirst
Do While Not rs_Tab0_SelectedRoute.EOF
   strRouteNo = rs_Tab0_SelectedRoute.Fields("�����ƨ����s").Value
   '�N���R���� [�w��������ƨ����s] �[�J [�����ƨ����s]
   Call SelectedRoute_Removeto_TRP01T(strRouteNo)
   rs_Tab0_SelectedRoute.MoveNext
Loop
   
'���s���� [�����ƨ����s] �� [�s��] ����
Call ReSet_TRP01T_SeqNo

'�ƨ��@�~�G�w����������ƨ����s�C�� DBGrid �榡�]�w-ReSet
Call CreateRS_Tab0_SelectedRoute

'���s�p��w��������ƨ����s�G�c�ơA�O�ơA���n�A���q + �s�����s����
Call Calculate_SelectedRoute
blTab0SelectedRouteEventEnable = True

'�M�� [�z��] �P [�Ƨ�] �]�w��
blTRP01TEventEnable = False
If strSourceFilter <> "0" Then rs_TRP01T.Filter = strSourceFilter
If rs_TRP01T.EOF Then rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = strSourceOrderBy
blTRP01TEventEnable = True

'���s�p�� [�ݱƨ��@���ƨ����s] �`�p
Call ReCaculate_FirstRouteSum

End Sub

Private Sub cmd_Tab0_srcRouteQuery_Click()
'�G���ƨ��@�~ >> �����ƨ����s�j�M
If rs_TRP01T Is Nothing Then Exit Sub
If rs_TRP01T.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_TRP01T"

If ShowForm_RS_FilterAndSort(rs_TRP01T, "�����ƨ����u�s��", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = 2
End Sub

Private Sub cmd_Tab0_srcRouteReset_Click()
'�����z��Ƨ�
'�����z�����A���]�ƧǨ̾�
 blTRP01TEventEnable = False
 rs_TRP01T.Filter = adFilterNone
 rs_TRP01T.Sort = strSourceOrderBy  '�M�αƧǡA�@���ƧǸ��Ѥp�ܤj
 blTRP01TEventEnable = True

'���s�p�� [�ݱƨ��@���ƨ����s] �`�p
Call ReCaculate_FirstRouteSum

End Sub

Private Sub cmd_Tab1_RouteNoDelete_Click()
'�G���ƨ����s�C�� >> �G���ƨ����u�s���R��
If rs_Tab1_Route.RecordCount = 0 Then Exit Sub
If dg_Tab1_Route.SelBookmarks.Count = 0 Then Exit Sub

Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
strDeleteRouteNo = Trim(rs_Tab1_Route.Fields("�G���ƨ����s").Value)
strCarno = Trim(rs_Tab1_Route.Fields("���P���X").Value)
dbDriveTimes = Trim(rs_Tab1_Route.Fields("����").Value)

'���R�������s�G�O�_�w�X���T�{
Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "Select c_Route_No From SDN01T Where c_Route_No = '" & strDeleteRouteNo & "'"
'Terry 20191127 �אּ�ˬd�X�����A
str_SQL = "Select Route_No From TRP05T Where Route_No = '" & strDeleteRouteNo & "' and sdnstatus = '1' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
    tmp_Rs.Close
    msg_text = "�`�N�G�����u�s���w�X���T�{�A�L�k�R��! "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If

msg_text = "�T�{�R���G���ƨ����u�s���G" & strDeleteRouteNo
If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)

'���ұ��R�������s�A�ƨ��̬O�_�����ɵn�J���ϥΪ�
str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "��Ʋ��`�G�䤣����R�����G���ƨ����u�s��"
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
str_SQL = "Select Convert(varchar(8),Vehicle_Check_in,112) as Checkin,Convert(varchar(8),Vehicle_Check_out,112) as Checkout From TRP05T Where Route_No = '" & strDeleteRouteNo & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("Checkin").Value <> "" Or tmp_Rs.Fields("CheckOut").Value <> "" Then
   tmp_Rs.Close
   msg_text = "��Ʋ��`�G�����u�s���w���� [��������] �� [��������]�A���R�������s�A�вM�������i�X������A�i��R��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
tmp_Rs.Close

Screen.MousePointer = vbHourglass
blTab1RouteEventEnable = False
Tran_Level = 0
Tran_Level = cn.BeginTrans

''APP�R�q��
'cn.Execute "delete apporderdate where receipt_no in (select t2.receipt_no from trp02t t2 join trp01t t1 on t1.route_no = t2.route_no and t1.c_route_no = '" & strDeleteRouteNo & "') ", RowsAffect, adExecuteNoRecords

'�R���֨����u�s��

'(1).�N TRP05T �������ƨ����s�� [�G���ƨ����u�s��] �M��
str_SQL = "Update TRP05T Set C_Route_No = null,C_Vehicle_ID_No = null,C_Drive_Times = null Where C_Route_No = '" & strDeleteRouteNo & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(2).�R�� TRP05T �G���ƨ����s
str_SQL = "Delete From TRP05T Where Route_No = '" & strDeleteRouteNo & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(3).�N TRP01T �� �����ƨ����s �� [�G���ƨ����u�s��] �M��
str_SQL = "Update TRP01T Set C_Route_No = null,C_Vehicle_ID_No = null,C_Drive_Times = null Where C_Route_No = '" & strDeleteRouteNo & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(4).�R�� TRP01T ���G���ƨ����u�s��
str_SQL = "Delete From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'(5).�R���d�ߵ��G [�G���ƨ����ݤ������ƨ����s] ���ӵ����u�s��--rs_Tab1_RouteDC
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_RouteDC.Filter = "�G���ƨ����s='" & strDeleteRouteNo & "'"
   If Not rs_Tab1_RouteDC.EOF Then
      Do While Not rs_Tab1_RouteDC.EOF
         rs_Tab1_RouteDC.Delete
         rs_Tab1_RouteDC.MoveFirst
      Loop
   End If
   rs_Tab1_RouteDC.Filter = adFilterNone
   rs_Tab1_RouteDC.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj

'(6).�R���d�ߵ��G [�G���ƨ����s] ���ӵ����u�s��--rs_Tab1_Route
rs_Tab1_Route.Delete
If Not rs_Tab1_Route.EOF Then rs_Tab1_Route.MoveFirst

blTab1RouteEventEnable = True

cn.CommitTrans
Tran_Level = 0
Screen.MousePointer = vbDefault


On Error GoTo err_Handle2
    
    Dim HttpClient As Object
    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/DeleteRouteNoByWareHouse?Route_NO=" & strDeleteRouteNo & "&WareHouse=DYDC_BEST", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    
    Exit Sub

err_Handle2:
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If

   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�G���ƨ����s�C��-�G���ƨ����s�R��", Me.Caption, "cmd_Tab1_RouteNoDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_RouteNoQuery_Click()
'�G���ƨ����s�C�� >> �G���ƨ����u�s���d��
If Len(Trim(txt_Tab1_RouteNo.Text)) = 0 Then Exit Sub

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle

'�]�w�G���ƨ����s�C��
blTab1RouteEventEnable = False
Call CreateRS_Tab1_Route
blTab1RouteEventEnable = True
'�]�w�G���ƨ����ݨ̦��ƨ����s�C��
Call CreateRS_Tab1_RouteDC

str_SQL = "Select �G���ƨ����s,�X�����,���P���X,����,�r�p�H,�c��,�O��,���q,���n,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,����,�ƨ��� " & _
          "From DCRouteMerge_RouteData Where �G���ƨ����s like '%" & txt_Tab1_RouteNo.Text & "%' and left(�G���ƨ����s,1) = 'S' order by �G���ƨ����s"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧤G���ƨ����s���(TRP01T)"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
blTab1RouteEventEnable = False
Do While Not tmp_Rs.EOF
   rs_Tab1_Route.AddNew
   rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
   rs_Tab1_Route.Fields("�G���ƨ����s").Value = tmp_Rs.Fields("�G���ƨ����s").Value
   rs_Tab1_Route.Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
   rs_Tab1_Route.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
   rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
   rs_Tab1_Route.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
   rs_Tab1_Route.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
   rs_Tab1_Route.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
   rs_Tab1_Route.Fields("���n").Value = tmp_Rs.Fields("���n").Value
   rs_Tab1_Route.Fields("���q").Value = tmp_Rs.Fields("���q").Value
   rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = tmp_Rs.Fields("�X�Y�Ȧs").Value
   rs_Tab1_Route.Fields("�w�p������").Value = tmp_Rs.Fields("�w�p������").Value
   rs_Tab1_Route.Fields("�w�p����ɶ�").Value = tmp_Rs.Fields("�w�p����ɶ�").Value
   rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
   rs_Tab1_Route.Fields("�ƨ���").Value = tmp_Rs.Fields("�ƨ���").Value
   rs_Tab1_Route.Update
   tmp_Rs.MoveNext
Loop
rs_Tab1_Route.MoveFirst
blTab1RouteEventEnable = True
tmp_Rs.Close
'TRP05T
str_SQL = "Select �G���ƨ����s,�����ƨ����s,�X�����,���P���X,����,�r�p�H,�c��,�O��,���q,���n,���� " & _
          "From DCRouteMerge_RouteDCData " & _
           "Where �G���ƨ����s like '%" & txt_Tab1_RouteNo.Text & "%' Order by �G���ƨ����s,�����ƨ����s"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w�G���ƨ����u�s�����(TRP05T)"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Do While Not tmp_Rs.EOF
   rs_Tab1_RouteDC.AddNew
   rs_Tab1_RouteDC.Fields("�s��").Value = rs_Tab1_RouteDC.RecordCount
   rs_Tab1_RouteDC.Fields("�G���ƨ����s").Value = tmp_Rs.Fields("�G���ƨ����s").Value
   rs_Tab1_RouteDC.Fields("�����ƨ����s").Value = tmp_Rs.Fields("�����ƨ����s").Value
   rs_Tab1_RouteDC.Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
   rs_Tab1_RouteDC.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
   rs_Tab1_RouteDC.Fields("����").Value = tmp_Rs.Fields("����").Value
   rs_Tab1_RouteDC.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
   rs_Tab1_RouteDC.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
   rs_Tab1_RouteDC.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
   rs_Tab1_RouteDC.Fields("���q").Value = tmp_Rs.Fields("���q").Value
   rs_Tab1_RouteDC.Fields("���n").Value = tmp_Rs.Fields("���n").Value
   rs_Tab1_RouteDC.Fields("����").Value = tmp_Rs.Fields("����").Value
   rs_Tab1_RouteDC.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab1_RouteDC.MoveFirst

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�G���ƨ����s�C��-�G���ƨ����s�d��", Me.Caption, "cmd_Tab1_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub dg_Tab0_SelectedRoute_HeadClick(ByVal ColIndex As Integer)
'�H�ƹ��I�� [�w��������ƨ����u�s��] dg_Tab0_SelectedRoute �����D�ϡG�Ƨ������
Dim OrderFieldName As String
If TypeName(rs_Tab0_SelectedRoute) <> "Nothing" Then
   OrderFieldName = "[" & dg_Tab0_SelectedRoute.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_Tab0_SelectedRoute.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_Tab0_SelectedRoute.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

Private Sub dg_Tab0_SelectedRoute_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�ƨ��@�~ >> �w��������ƨ����u�s�� DBGrid
If blTab0SelectedRouteEventEnable Then
   With dg_Tab0_SelectedRoute
        '�ϥ���ܿ������ƦC
        If Not rs_Tab0_SelectedRoute.EOF Then
           dg_Tab0_SelectedRoute.SelBookmarks.Add rs_Tab0_SelectedRoute.Bookmark
        End If
   End With
End If
End Sub

Private Sub dg_Tab1_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�G���ƨ����u�s���C��G�����
If blTab1RouteEventEnable Then
   If Not rs_Tab1_Route.EOF Then
      dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
      rs_Tab1_RouteDC.Filter = " �G���ƨ����s = '" & rs_Tab1_Route.Fields("�G���ƨ����s").Value & "'"
   End If
End If
End Sub

Private Sub dg_TRP01T_HeadClick(ByVal ColIndex As Integer)
'�H�ƹ��I�� [�����ƨ����s] dg_TRP01T �����D�ϡG�Ƨ������
Dim OrderFieldName As String
If TypeName(rs_TRP01T) <> "Nothing" Then
   '�קK���� [���] ���ʧ@
   blTRP01TEventEnable = False
   OrderFieldName = "[" & dg_TRP01T.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_TRP01T.Sort = OrderFieldName & " DESC "
      strSourceOrderBy = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_TRP01T.Sort = OrderFieldName & " ASC "
      strSourceOrderBy = OrderFieldName & " ASC "
   End If
   blTRP01TEventEnable = True
End If
End Sub

Private Sub dg_TRP01T_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�G���ƨ� >> �����ƨ����u�s���C�� DBGrid
If blTRP01TEventEnable Then
   With dg_TRP01T
        '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
        If Trim(rs_TRP01T.Fields(1).Value) = "" Then
           dbSelectedCount = dbSelectedCount + 1
           rs_TRP01T.Fields(1).Value = "V"
           '����p�p��s
           dbsrcSelected_Case = dbsrcSelected_Case + rs_TRP01T.Fields("�c��").Value
           dbsrcSelected_Pallet = dbsrcSelected_Pallet + rs_TRP01T.Fields("�O��").Value
           dbsrcSelected_Volumn = dbsrcSelected_Volumn + rs_TRP01T.Fields("���n").Value
           dbsrcSelected_Weight = dbsrcSelected_Weight + rs_TRP01T.Fields("���q").Value
           txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
           txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
        Else
           dbSelectedCount = dbSelectedCount - 1
           rs_TRP01T.Fields(1).Value = " "
           '����p�p��s
           If dbSelectedCount <> 0 Then
              dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP01T.Fields("�c��").Value
              dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP01T.Fields("�O��").Value
              dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP01T.Fields("���n").Value
              dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP01T.Fields("���q").Value
           Else
              dbsrcSelected_Case = 0
              dbsrcSelected_Pallet = 0
              dbsrcSelected_Volumn = 0
              dbsrcSelected_Weight = 0
           End If
           txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
           txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
        End If
        '��ܿ�������s����
        
        '�ϥ���ܿ������ƦC
        If Not rs_TRP01T.EOF Then
           dg_TRP01T.SelBookmarks.Add rs_TRP01T.Bookmark
        End If
        '��ܸ��s���q����
        Call Display_SelectOrdersData(rs_TRP01T.Fields("�����ƨ����s").Value)
   End With
End If
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�G���ƨ��@�~"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'�ƨ��@�~�G�w��������u�s���C�� DBGrid �榡�]�w
Call CreateRS_Tab0_SelectedRoute

'�G���ƨ����ͤ��s���s�C��GDBGrid �榡�]�w
Call CreateRS_Tab1_Route
'�Q�G���ƨ��������ƨ����s�C��GDBGrid �榡�]�w
Call CreateRS_Tab1_RouteDC
'���u�s���������q����
Call CreateRS_Tab0_Orders
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�d�I��Ӫ����L����ƥ�
'�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������������
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
End If
End Sub

Private Sub Form_Resize()
On Error GoTo err_Handle

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
'   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Height = Me.ScaleHeight - 120
   SSTab1.Width = Me.ScaleWidth - 120
   
   fam_SelectedOrders.Width = SSTab1.Width - 120
   fam_SrcRoute.Width = fam_SelectedOrders.Width
   dg_Tab0_SelectedRoute.Width = fam_SelectedOrders.Width - 1600
   dg_TRP01T.Width = fam_SrcRoute.Width - 1600
   dg_Tab0_Orders.Width = dg_TRP01T.Width
   
   fam_SrcRoute.Height = SSTab1.Height - fam_RouteData.Top - fam_RouteData.Height - fam_SelectedOrders.Height
   dg_Tab0_Orders.Height = fam_SrcRoute.Height - fam_SelectedSum.Height - dg_TRP01T.Height + 900

   cmd_Tab0_SelectedRemove_All.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   Shape3.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   cmd_Tab0_srcRouteQuery.Left = Shape3.Left + 60
   cmd_Tab0_srcRouteReset.Left = Shape3.Left + 60

   Label1(11).Left = dg_TRP01T.Left + dg_TRP01T.Width + 120
   Label1(10).Left = Label1(11).Left
   Label1(9).Left = Label1(11).Left
   Label1(8).Left = Label1(11).Left
   txt_Tab0_srcTotal_Case.Left = Label1(11).Left
   txt_Tab0_srcTotal_Pallet.Left = Label1(11).Left
   txt_Tab0_srcTotal_Volumn.Left = Label1(11).Left
   txt_Tab0_srcTotal_Weight.Left = Label1(11).Left

   fam_Tab1_Query.Left = fam_Tab1_Query.Left - (dbsrcFormWidth - Me.ScaleWidth)
   fam_Tab1_Delete.Left = fam_Tab1_Delete.Left - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab1_Route.Width = dg_Tab1_Route.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab1_RouteDC.Height = dg_Tab1_RouteDC.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Tab1_RouteDC.Width = dg_Tab1_RouteDC.Width - (dbsrcFormWidth - Me.ScaleWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = Me.ScaleHeight - 120
   SSTab1.Width = Me.ScaleWidth - 120
   
   fam_SelectedOrders.Width = SSTab1.Width - 120
   fam_SrcRoute.Width = fam_SelectedOrders.Width
   dg_Tab0_SelectedRoute.Width = fam_SelectedOrders.Width - 1600
   dg_TRP01T.Width = fam_SrcRoute.Width - 1600
   dg_Tab0_Orders.Width = dg_TRP01T.Width
   
   fam_SrcRoute.Height = SSTab1.Height - fam_RouteData.Top - fam_RouteData.Height - fam_SelectedOrders.Height
   dg_Tab0_Orders.Height = fam_SrcRoute.Height - fam_SelectedSum.Height - dg_TRP01T.Height - 120
   
   cmd_Tab0_SelectedRemove_All.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   Shape3.Left = dg_Tab0_SelectedRoute.Left + dg_Tab0_SelectedRoute.Width + 120
   cmd_Tab0_srcRouteQuery.Left = Shape3.Left + 60
   cmd_Tab0_srcRouteReset.Left = Shape3.Left + 60

   Label1(11).Left = dg_TRP01T.Left + dg_TRP01T.Width + 120
   Label1(10).Left = Label1(11).Left
   Label1(9).Left = Label1(11).Left
   Label1(8).Left = Label1(11).Left
   txt_Tab0_srcTotal_Case.Left = Label1(11).Left
   txt_Tab0_srcTotal_Pallet.Left = Label1(11).Left
   txt_Tab0_srcTotal_Volumn.Left = Label1(11).Left
   txt_Tab0_srcTotal_Weight.Left = Label1(11).Left
   
   fam_Tab1_Query.Left = fam_Tab1_Query.Left + (Me.ScaleWidth - dbsrcFormWidth)
   fam_Tab1_Delete.Left = fam_Tab1_Delete.Left + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab1_Route.Width = dg_Tab1_Route.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab1_RouteDC.Height = dg_Tab1_RouteDC.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Tab1_RouteDC.Width = dg_Tab1_RouteDC.Width + (Me.ScaleWidth - dbsrcFormWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
End If

Exit Sub
err_Handle:
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_OP_DCRouteMerge = Nothing
End Sub

Private Sub CreateRS_Tab0_SelectedRoute()
'�ƨ��@�~�G�w����������ƨ����s�C��
Call ReDim_Recordset(rs_Tab0_SelectedRoute)
With rs_Tab0_SelectedRoute
     .Fields.Append "�s��", adDouble
     .Fields.Append "�����ƨ����s", adVarChar, 10
     .Fields.Append "�c��", adDouble
     .Fields.Append "�O��", adDouble
     .Fields.Append "���n", adDouble
     .Fields.Append "���q", adDouble
     .Fields.Append "���P���X", adVarChar, 20
     .Fields.Append "����", adDouble
     .Fields.Append "�r�p�H", adVarChar, 30
     .Fields.Append "�X�����", adVarChar, 12
     .Fields.Append "����", adVarChar, 10
     .Fields.Append "EXE�^��", adVarChar, 20
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
Set dg_Tab0_SelectedRoute.DataSource = rs_Tab0_SelectedRoute
'�]�w������
With dg_Tab0_SelectedRoute
    .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
    .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
    .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
    .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
    .Columns(0).Width = 500         '�s��
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1200        '�����ƨ����u�s��
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 800         '�c��
    .Columns(2).Alignment = dbgRight
    .Columns(3).Width = 800         '�O��
    .Columns(3).Alignment = dbgRight
    .Columns(4).Width = 800         '���n
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 800         '���q
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 900         '���P���X
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 500         '����
    .Columns(7).Alignment = dbgCenter
    .Columns(8).Width = 800         '�r�p�H
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 1000        '�X�����
    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 500       '����
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 800        'EXE�^��
    .Columns(11).Alignment = dbgLeft
End With
End Sub

Private Sub Calculate_SelectedRoute()
'�@�~���e�G
'1.�w��w��������ƥX���s�C��A�̣����ƨ����s���s���� [�s��] ����
'2.�p��w��������ƨ����s���֭p���
Dim dbSeqNo As Double
dbSeqNo = 0
txt_Tab0_Selected_Case.Text = ""
txt_Tab0_Selected_Pallet.Text = ""
txt_Tab0_Selected_Volumn.Text = ""
txt_Tab0_Selected_Weight.Text = ""

rs_Tab0_SelectedRoute.Filter = adFilterNone
rs_Tab0_SelectedRoute.Sort = "�����ƨ����s asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
If Not rs_Tab0_SelectedRoute.EOF Then
   rs_Tab0_SelectedRoute.MoveFirst
Else
   '�M�X�z�����A���L��ƪ̡A���� SubProgram ����
   Exit Sub
End If
Do While Not rs_Tab0_SelectedRoute.EOF
   dbSeqNo = dbSeqNo + 1
   rs_Tab0_SelectedRoute.Fields("�s��").Value = dbSeqNo
   txt_Tab0_Selected_Case.Text = Val(txt_Tab0_Selected_Case.Text) + rs_Tab0_SelectedRoute.Fields("�c��").Value
   txt_Tab0_Selected_Pallet.Text = Val(txt_Tab0_Selected_Pallet.Text) + rs_Tab0_SelectedRoute.Fields("�O��").Value
   txt_Tab0_Selected_Volumn.Text = Val(txt_Tab0_Selected_Volumn.Text) + rs_Tab0_SelectedRoute.Fields("���n").Value
   txt_Tab0_Selected_Weight.Text = Val(txt_Tab0_Selected_Weight.Text) + rs_Tab0_SelectedRoute.Fields("���q").Value
   rs_Tab0_SelectedRoute.MoveNext
Loop
rs_Tab0_SelectedRoute.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
If Not rs_Tab0_SelectedRoute.EOF Then rs_Tab0_SelectedRoute.MoveFirst
End Sub

Private Sub SelectedRoute_Removeto_TRP01T(ByVal strRouteNo As String)
'�N���w�� [�����ƨ����s] �[�J [�����ƨ����s�C��]
blTRP01TEventEnable = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj

rs_TRP01T.Filter = "�����ƨ����s = '" & strRouteNo & "'"
If Not rs_TRP01T.EOF Then
   '�����ƨ����u�s���w�s�b���ܡA���i��s�W�A�]����s
   rs_TRP01T.Filter = adFilterNone
   rs_TRP01T.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   blTRP01TEventEnable = True
   Exit Sub
End If

'���^�����ƨ����u�s��
str_SQL = "Select ' ' as '��',�����ƨ����s,�X�Y,�c��,�O��,���n,���q,���P���X,����,�r�p�H,�X�����,����,EXE�^�� " & _
          "From DCRouteMerge_DCRouteData Where �����ƨ����s = '" & strRouteNo & "' Order by �����ƨ����s "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX��w�������ƨ����u�s����ƥi�H����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   blTRP01TEventEnable = True
   Exit Sub
End If

rs_TRP01T.AddNew
rs_TRP01T.Fields("�s��").Value = rs_TRP01T.RecordCount
rs_TRP01T.Fields("�����ƨ����s").Value = tmp_Rs.Fields("�����ƨ����s").Value
rs_TRP01T.Fields("�X�Y").Value = tmp_Rs.Fields("�X�Y").Value
rs_TRP01T.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
rs_TRP01T.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
rs_TRP01T.Fields("���n").Value = tmp_Rs.Fields("���n").Value
rs_TRP01T.Fields("���q").Value = tmp_Rs.Fields("���q").Value
rs_TRP01T.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
rs_TRP01T.Fields("����").Value = tmp_Rs.Fields("����").Value
rs_TRP01T.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
rs_TRP01T.Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
rs_TRP01T.Fields("����").Value = tmp_Rs.Fields("����").Value
rs_TRP01T.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
rs_TRP01T.Update
tmp_Rs.Close

rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "�����ƨ����s asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
If Not rs_TRP01T.EOF Then rs_TRP01T.MoveFirst
blTRP01TEventEnable = True
End Sub

Private Sub ReSet_TRP01T_SeqNo()
'���s���� [�����ƨ����u�s��] �� [�s��] ����
blTRP01TEventEnable = False
dg_TRP01T.Visible = False
rs_TRP01T.Filter = adFilterNone
rs_TRP01T.Sort = "�����ƨ����s asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
If Not rs_TRP01T.EOF Then rs_TRP01T.MoveFirst

Dim dbSeqNo As Double
dbSeqNo = 0
Do While Not rs_TRP01T.EOF
   dbSeqNo = dbSeqNo + 1
   rs_TRP01T.Fields("�s��").Value = dbSeqNo
   rs_TRP01T.MoveNext
Loop
If rs_TRP01T.RecordCount > 0 Then rs_TRP01T.MoveFirst
dg_TRP01T.Visible = True
blTRP01TEventEnable = True
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
       Case "�X�����"
            txt_Tab0_TRPDate.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�w�p������"
            txt_Tab0_CarCheckInDate.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�ƨ����.�_"
            txt_FPlanDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�ƨ����.��"
            txt_FPlanDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�X�����.�_"
            txt_FDeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�X�����.��"
            txt_FDeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_FPlanDate_Start_Click()
'�G���ƨ� >> �פJ�@���ƨ����s >> �ƨ�����G�_
If Trim(txt_FPlanDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_Start.Text, 4) & "/" & Mid(txt_FPlanDate_Start.Text, 5, 2) & "/" & Right(txt_FPlanDate_Start.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FPlanDate_Start.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FPlanDate_Start.Top + txt_FPlanDate_Start.Height
mvDate.Tag = "�ƨ����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_Start_KeyPress(KeyAscii As Integer)
'�G���ƨ� >> �פJ�@���ƨ����s >> �ƨ�����G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanDate_End.SelStart = 0: txt_FPlanDate_End.SelLength = Len(txt_FPlanDate_End.Text)
         txt_FPlanDate_End.SetFocus
End Select
End Sub

Private Sub txt_FPlanDate_End_Click()
'�G���ƨ� >> �פJ�@���ƨ����s >> �ƨ�����G��
If Trim(txt_FPlanDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_End.Text, 4) & "/" & Mid(txt_FPlanDate_End.Text, 5, 2) & "/" & Right(txt_FPlanDate_End.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FPlanDate_End.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FPlanDate_End.Top + txt_FPlanDate_End.Height
mvDate.Tag = "�ƨ����.��"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_End_KeyPress(KeyAscii As Integer)
'�G���ƨ� >> �פJ�@���ƨ����s >> �ƨ�����G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_Start.SelStart = 0: txt_FDeliveryDate_Start.SelLength = Len(txt_FDeliveryDate_Start.Text)
         txt_FDeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_Start_Click()
'�G���ƨ� >> �פJ�@���ƨ����s >> �X������G�_
If Trim(txt_FDeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_Start.Text, 4) & "/" & Mid(txt_FDeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FDeliveryDate_Start.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FDeliveryDate_Start.Top + txt_FDeliveryDate_Start.Height
mvDate.Tag = "�X�����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�G���ƨ� >> �פJ�@���ƨ����s >> �X������G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_End.SelStart = 0: txt_FDeliveryDate_End.SelLength = Len(txt_FDeliveryDate_End.Text)
         txt_FDeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_End_Click()
'�G���ƨ� >> �פJ�@���ƨ����s >> �X������G��
If Trim(txt_FDeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_End.Text, 4) & "/" & Mid(txt_FDeliveryDate_End.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = fra_ExtraQuery.Left + txt_FDeliveryDate_End.Left
mvDate.Top = fra_ExtraQuery.Top + txt_FDeliveryDate_End.Top + txt_FDeliveryDate_End.Height
mvDate.Tag = "�X�����.��"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_End_KeyPress(KeyAscii As Integer)
'�G���ƨ� >> �פJ�@���ƨ����s >> �X������G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         cmd_Tab0_ImportRoute.SetFocus
End Select
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
'�G���ƨ� >> �w�p����ɶ�
Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_KeyPress(KeyAscii As Integer)
'�G���ƨ� >> ���P���X
Select Case KeyAscii
       Case 97 To 122   '�ഫ���j�g�r��
            KeyAscii = KeyAscii - 32
End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_LostFocus()
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
'�G���ƨ� >> �X�Y�Ȧs
Select Case KeyAscii
       Case 97 To 122   '�ഫ���j�g�r��
            KeyAscii = KeyAscii - 32
       Case vbKeyReturn
            KeyAscii = 0
            txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
            txt_Tab0_CarCheckInDate.SetFocus
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
'�G���ƨ� > [�X�����] ��Ʈ榡�Gyyyymmdd
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
Public Sub frm_OP_DCRouteMerge_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
'��椽�ΰƵ{���A�� frm_RS_FilterAndSort ���I�s
'�ǤJ�ȡGstrCode      �ʧ@�ѧO�X
'                     [FILTER] �ۭq�z��    [SORT] �Ƨ�
'        strReturn    �z�� or �Ƨ� ���]�w�r��

Select Case strCode
       Case "FILTER"  '�ۭq�z��
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP01T"   '�����ƨ����u�s�����
                        blTRP01TEventEnable = False
                        rs_TRP01T.Filter = adFilterNone
                        rs_TRP01T.Filter = strReturn
                        strSourceFilter = strReturn
                        If rs_TRP01T.RecordCount = 0 Then
                           msg_text = "��p���A�䤣��ŦX���󪺣����ƨ����u�s����"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_TRP01T.Filter = adFilterNone
                           strSourceFilter = adFilterNone
                           rs_TRP01T.Sort = strSourceOrderBy  '�M�αƧǡA�@���ƧǸ��Ѥp�ܤj
                           blTRP01TEventEnable = True
                           Exit Sub
                        End If
                        blTRP01TEventEnable = True
                        '���s�p�� [�ݱƨ��@���ƨ����s] �`�p
                        Call ReCaculate_FirstRouteSum

            End Select
       Case "SORT"    '�Ƨ�
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP01T"   '�����ƨ����u�s�����
                        blTRP01TEventEnable = False
                        rs_TRP01T.Sort = strReturn
                        strSourceOrderBy = strReturn
                        blTRP01TEventEnable = True
            End Select
End Select
End Sub

Private Sub CreateRS_Tab1_Route()
'�ƨ��@�~�G�G���ƨ����ͤ����u�s���C��
Call ReDim_Recordset(rs_Tab1_Route)
With rs_Tab1_Route
     .Fields.Append "�s��", adVarChar, 10
     .Fields.Append "�G���ƨ����s", adVarChar, 10
     .Fields.Append "�X�����", adVarChar, 8
     .Fields.Append "���P���X", adVarChar, 10
     .Fields.Append "����", adDouble
     .Fields.Append "�r�p�H", adVarChar, 20
     .Fields.Append "�c��", adDouble
     .Fields.Append "�O��", adDouble
     .Fields.Append "���n", adDouble
     .Fields.Append "���q", adDouble
     .Fields.Append "�X�Y�Ȧs", adVarChar, 10
     .Fields.Append "�w�p������", adVarChar, 8
     .Fields.Append "�w�p����ɶ�", adVarChar, 4
     .Fields.Append "����", adVarChar, 10
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
    .Columns(1).Width = 1200        '�G���ƨ����u�s��
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
    .Columns(10).Width = 900        '�X�Y�Ȧs
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1200       '�w�p����������
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1200       '�w�p��������ɶ�
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 500       '����
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1300       '�ƨ���
    .Columns(14).Alignment = dbgLeft
End With
End Sub

Private Sub CreateRS_Tab1_RouteDC()
'�ƨ��@�~�G�w�s���G���ƨ����s�� �����ƨ� ���s�C��
Call ReDim_Recordset(rs_Tab1_RouteDC)
With rs_Tab1_RouteDC
     .Fields.Append "�s��", adVarChar, 10
     .Fields.Append "�G���ƨ����s", adVarChar, 10
     .Fields.Append "�����ƨ����s", adVarChar, 10
     .Fields.Append "�X�����", adVarChar, 8
     .Fields.Append "���P���X", adVarChar, 10
     .Fields.Append "����", adDouble
     .Fields.Append "�r�p�H", adVarChar, 20
     .Fields.Append "�c��", adDouble
     .Fields.Append "�O��", adDouble
     .Fields.Append "���n", adDouble
     .Fields.Append "���q", adDouble
     .Fields.Append "����", adVarChar, 10
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
Set dg_Tab1_RouteDC.DataSource = rs_Tab1_RouteDC
'�]�w������
With dg_Tab1_RouteDC
    .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
    .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
    .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
    .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
    .Columns(0).Width = 500         '�s��
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1200        '�G���ƨ����s
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1200        '�����ƨ����s
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 900         '�X�����
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 850         '���P���X
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 500         '����
    .Columns(5).Alignment = dbgCenter
    .Columns(6).Width = 900         '�r�p�H
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 700         '�c��
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700         '�O��
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700         '���n
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 700        '���q
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 500       '����
    .Columns(11).Alignment = dbgLeft
End With
End Sub

Private Sub txt_Tab1_RouteNo_KeyPress(KeyAscii As Integer)
'�֨����u�s���C�� >> �֨����u�s��
Select Case KeyAscii
     Case 97 To 122   '�ഫ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          cmd_Tab1_RouteNoQuery.SetFocus
End Select
End Sub

Private Sub Display_SelectOrdersData(ByVal strRouteNo As String)
'��ܶǤJ�����s�������q��
Set dg_Tab0_Orders.DataSource = Nothing
Set rs_Tab0_Orders = Nothing
Call CreateRS_Tab0_Orders

'str_SQL = "Select ���u�s��,�e�f��,�q��s��,ZIP,Area,�Ȥ�W��,�c��,�O��,���n,���q,�q��Ƶ�,Receipt_No,EXE�^�� " & _
'          "From DCRouteMerge_RouteOrders " & _
'           "Where ���u�s�� like '" & strRouteNo & "%' Order by Receipt_No"
           
str_SQL = "Select  Rtrim(a1.Route_No) as ���u�s�� " & _
        ", Convert(varchar,a1.Arrive_Date,112) as �e�f�� " & _
        ", Rtrim(a1.Receipt_No) + '(' + Rtrim(a1.StorerKey)+'-'+Rtrim(a1.Extern)+')' as �q��s�� " & _
        ", Rtrim(a2.ZIP) as ZIP , Rtrim(Isnull(a2.Area_Code,'')) as Area , Rtrim(a2.Full_Name) as �Ȥ�W�� " & _
        ", Round(a1.Case_cnt,2) as �c�� ,  Round(a1.Pallet_Qty,2) as �O�� " & _
        ", Round(a1.Volumn_Weight,2) as ���n " & _
        ", Round(a1.Weight,2) as ���q " & _
        ",�q��Ƶ� = rtrim(a1.description) " & _
        ", Rtrim(a1.Receipt_No) as Receipt_No " & _
        ", Case a1.EXE_Confirm When '0' Then '�s�ظ��s' When '1' Then '�]�w�^��' When '2' Then '�w�^��' When '9' Then '�w���z�f' else '�������A' End  AS EXE�^�� " & _
        "From TRP02T a1(nolock) inner join TRP01M a2(nolock) on a2.ConsigneeKey = a1.ConsigneeKey and a2.storerkey = a1.storerkey " & _
        "where Rtrim(a1.Route_No) = '" & strRouteNo & "' order by Rtrim(a1.Receipt_No) "

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
Do While Not tmp_Rs.EOF
   rs_Tab0_Orders.AddNew
   rs_Tab0_Orders.Fields("�s��").Value = rs_Tab0_Orders.RecordCount
   rs_Tab0_Orders.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
   rs_Tab0_Orders.Fields("�e�f��").Value = tmp_Rs.Fields("�e�f��").Value
   rs_Tab0_Orders.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
   rs_Tab0_Orders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
   rs_Tab0_Orders.Fields("Area").Value = tmp_Rs.Fields("Area").Value
   rs_Tab0_Orders.Fields("�Ȥ�W��").Value = tmp_Rs.Fields("�Ȥ�W��").Value
   rs_Tab0_Orders.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
   rs_Tab0_Orders.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
   rs_Tab0_Orders.Fields("���n").Value = tmp_Rs.Fields("���n").Value
   rs_Tab0_Orders.Fields("���q").Value = tmp_Rs.Fields("���q").Value
   rs_Tab0_Orders.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value & ""
   rs_Tab0_Orders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
   rs_Tab0_Orders.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
   rs_Tab0_Orders.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
rs_Tab0_Orders.MoveFirst

End Sub

Private Sub CreateRS_Tab0_Orders()
'�ƨ��@�~�G�w�s�����s���q��C��
Call ReDim_Recordset(rs_Tab0_Orders)
With rs_Tab0_Orders
     .Fields.Append "�s��", adVarChar, 10
     .Fields.Append "���u�s��", adVarChar, 10
     .Fields.Append "�e�f��", adVarChar, 20
     .Fields.Append "�q��s��", adVarChar, 60
     .Fields.Append "ZIP", adVarChar, 60
     .Fields.Append "Area", adVarChar, 60
     .Fields.Append "�Ȥ�W��", adVarChar, 120
     .Fields.Append "�c��", adDouble
     .Fields.Append "�O��", adDouble
     .Fields.Append "���n", adDouble
     .Fields.Append "���q", adDouble
     .Fields.Append "�q��Ƶ�", adVarChar, 300
     .Fields.Append "Receipt_No", adVarChar, 60
     .Fields.Append "EXE�^��", adVarChar, 20
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
Set dg_Tab0_Orders.DataSource = rs_Tab0_Orders
'�]�w������
With dg_Tab0_Orders
    .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
    .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
    .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
    .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
    .Columns(0).Width = 500         '�s��
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1050        '���u�s��
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900         '�e�f��
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 2150        '�q��s��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 400         'ZIP
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 400         'Area �B�e�ϰ�N�X
    .Columns(5).Alignment = dbgCenter
    .Columns(6).Width = 1500        '�Ȥ�W��
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 700         '�c��
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700         '�O��
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700         '���n
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 700         '���q
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 2500       '�q��Ƶ�
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1100       'Receipt_No
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 1100       'EXE�^��
    .Columns(13).Alignment = dbgLeft
End With
End Sub

Private Sub ReCaculate_FirstRouteSum()
'�p�� [�ݱƨ��@���ƨ����s] �`�p
'���ĥΦA���^���έp�A�]���|���M�� [�z�����] �����D
txt_Tab0_srcTotal_Case.Text = ""
txt_Tab0_srcTotal_Pallet.Text = ""
txt_Tab0_srcTotal_Volumn.Text = ""
txt_Tab0_srcTotal_Weight.Text = ""

If rs_TRP01T Is Nothing Then Exit Sub
If rs_TRP01T.RecordCount = 0 Then Exit Sub

Dim dbTotalCase As Double
Dim dbTotalPallet As Double
Dim dbTotalWeight As Double
Dim dbTotalVolumn As Double
dbTotalCase = 0: dbTotalPallet = 0: dbTotalVolumn = 0: dbTotalWeight = 0
blTRP01TEventEnable = False
dg_TRP01T.Visible = False
rs_TRP01T.MoveFirst
Do While Not rs_TRP01T.EOF
   dbTotalCase = dbTotalCase + rs_TRP01T.Fields("�c��").Value
   dbTotalPallet = dbTotalPallet + rs_TRP01T.Fields("�O��").Value
   dbTotalVolumn = dbTotalVolumn + rs_TRP01T.Fields("���n").Value
   dbTotalWeight = dbTotalWeight + rs_TRP01T.Fields("���q").Value
   rs_TRP01T.MoveNext
Loop
rs_TRP01T.MoveFirst

txt_Tab0_srcTotal_Case.Text = dbTotalCase
txt_Tab0_srcTotal_Pallet.Text = dbTotalPallet
txt_Tab0_srcTotal_Volumn.Text = dbTotalVolumn
txt_Tab0_srcTotal_Weight.Text = dbTotalWeight
dg_TRP01T.Visible = True
blTRP01TEventEnable = True
End Sub
