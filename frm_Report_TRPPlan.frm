VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_Report_TRPPlan 
   Caption         =   " �ƨ��@�~����"
   ClientHeight    =   7140
   ClientLeft      =   135
   ClientTop       =   1020
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3600
      TabIndex        =   87
      Top             =   4200
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
      StartOfWeek     =   122355713
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7800
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   13758
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "VLL �˸�"
      TabPicture(0)   =   "frm_Report_TRPPlan.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CmnDialog"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fam_Tab0_Header"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dg_Tab0_VLL"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "�����˸��J�`��"
      TabPicture(1)   =   "frm_Report_TRPPlan.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fam_Tab1_Header"
      Tab(1).Control(1)=   "dg_Tab1_VLLSum"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "�q���`��"
      TabPicture(2)   =   "frm_Report_TRPPlan.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fam_Tab2"
      Tab(2).Control(1)=   "dg_Tab2_OrdersSum"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "�z�f�˸��]�֪�"
      TabPicture(3)   =   "frm_Report_TRPPlan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fam_Tab3"
      Tab(3).Control(1)=   "dg_Tab3_PickLoadCheck"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "��B�����u���`��"
      TabPicture(4)   =   "frm_Report_TRPPlan.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTab2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "�ƨ��@����"
      TabPicture(5)   =   "frm_Report_TRPPlan.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dg_Tab5_PlanList"
      Tab(5).Control(1)=   "fam_Tab5_Header"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "�̪O���@by���s"
      TabPicture(6)   =   "frm_Report_TRPPlan.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "dg_Tab6_PlanList"
      Tab(6).Control(1)=   "fam_Tab6_Header"
      Tab(6).ControlCount=   2
      Begin VB.Frame fam_Tab6_Header 
         Height          =   1500
         Left            =   -74880
         TabIndex        =   137
         Top             =   720
         Width           =   10185
         Begin VB.TextBox txt_Tab6_route_End 
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
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   144
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab6_route_Start 
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
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   143
            Top             =   720
            Width           =   1365
         End
         Begin VB.CommandButton cmd_Tab6_SaveToExcel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   7725
            Picture         =   "frm_Report_TRPPlan.frx":00C4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   142
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab6_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��Ƭd��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   6570
            Picture         =   "frm_Report_TRPPlan.frx":0C86
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   141
            Top             =   240
            Width           =   1065
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
            Height          =   990
            Index           =   7
            Left            =   8880
            Picture         =   "frm_Report_TRPPlan.frx":1550
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   140
            Top             =   210
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab6_DeliveryDate_End 
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
            Left            =   2700
            MaxLength       =   8
            TabIndex        =   139
            Top             =   270
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab6_DeliveryDate_Start 
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
            Left            =   1140
            MaxLength       =   8
            TabIndex        =   138
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label Label1 
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
            Index           =   37
            Left            =   105
            TabIndex        =   149
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Index           =   38
            Left            =   2550
            TabIndex        =   148
            Top             =   840
            Width           =   240
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00008080&
            BorderWidth     =   2
            Height          =   1260
            Index           =   2
            Left            =   6480
            Top             =   120
            Width           =   3570
         End
         Begin VB.Label Label1 
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
            Index           =   40
            Left            =   2430
            TabIndex        =   147
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   41
            Left            =   120
            TabIndex        =   146
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȭd30�� �B�w�X���T�{"
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
            Height          =   195
            Index           =   42
            Left            =   4080
            TabIndex        =   145
            Top             =   360
            Width           =   2160
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab5_PlanList 
         Height          =   4890
         Left            =   -74805
         TabIndex        =   102
         Top             =   2280
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   8625
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
      Begin VB.Frame fam_Tab5_Header 
         Height          =   1500
         Left            =   -74805
         TabIndex        =   88
         Top             =   720
         Width           =   12105
         Begin VB.CommandButton cmd_Tab5_PrintReport1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���f��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   8520
            Picture         =   "frm_Report_TRPPlan.frx":1992
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   134
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab5_SaveToExcel_NEW 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel NEW"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   6840
            Picture         =   "frm_Report_TRPPlan.frx":1C9C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   129
            ToolTipText     =   "�ѩ�]�t�B�O�պ�A�ФŤ@���d�߹L�h�ѼƸ��"
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab5_route_End 
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
            Left            =   2835
            MaxLength       =   10
            TabIndex        =   113
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab5_route_Start 
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
            Left            =   1155
            MaxLength       =   10
            TabIndex        =   112
            Top             =   1080
            Width           =   1365
         End
         Begin VB.ComboBox cmb_Tab5_AreaCode 
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
            ForeColor       =   &H00800080&
            Height          =   315
            Left            =   1170
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   100
            Top             =   225
            Width           =   3960
         End
         Begin VB.CommandButton cmd_Tab5_ReSet 
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
            Height          =   360
            Left            =   5085
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   99
            Top             =   195
            Width           =   765
         End
         Begin VB.CommandButton cmd_Tab5_SaveToExcel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   7365
            Picture         =   "frm_Report_TRPPlan.frx":285E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   98
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab5_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��Ƭd��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   6210
            Picture         =   "frm_Report_TRPPlan.frx":3420
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   97
            Top             =   240
            Width           =   1065
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
            Height          =   990
            Index           =   6
            Left            =   10860
            Picture         =   "frm_Report_TRPPlan.frx":3CEA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   96
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab5_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�ƨ��ި��C�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   9705
            Picture         =   "frm_Report_TRPPlan.frx":412C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   95
            Top             =   225
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab5_DeliveryDate_End 
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
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   91
            Top             =   630
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab5_DeliveryDate_Start 
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
            Left            =   1155
            MaxLength       =   8
            TabIndex        =   90
            Top             =   615
            Width           =   1245
         End
         Begin VB.CheckBox chk_Tab5_PreView 
            Caption         =   "�w���C�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   4680
            TabIndex        =   89
            Top             =   1080
            Value           =   1  '�֨�
            Width           =   1425
         End
         Begin VB.Label Label1 
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
            Index           =   28
            Left            =   120
            TabIndex        =   115
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Index           =   27
            Left            =   2565
            TabIndex        =   114
            Top             =   1080
            Width           =   240
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00008080&
            BorderWidth     =   2
            Height          =   1140
            Index           =   1
            Left            =   6135
            Top             =   165
            Width           =   5865
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϽX"
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
            TabIndex        =   101
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Left            =   2445
            TabIndex        =   94
            Top             =   690
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   24
            Left            =   135
            TabIndex        =   93
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȭd7��"
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
            Height          =   195
            Index           =   17
            Left            =   4200
            TabIndex        =   92
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame fam_Tab3 
         Height          =   2100
         Left            =   -74805
         TabIndex        =   66
         Top             =   840
         Width           =   11070
         Begin VB.TextBox txt_Tab3_route_Start 
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
            Left            =   1875
            MaxLength       =   10
            TabIndex        =   121
            Top             =   1680
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab3_route_End 
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
            Left            =   3555
            MaxLength       =   10
            TabIndex        =   120
            Top             =   1680
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_Start 
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
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   104
            Top             =   1230
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_End 
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
            Left            =   3495
            MaxLength       =   8
            TabIndex        =   103
            Top             =   1230
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab3_UploadMinute_End 
            Alignment       =   1  '�a�k���
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
            Left            =   3450
            MaxLength       =   2
            TabIndex        =   79
            Top             =   870
            Width           =   375
         End
         Begin VB.TextBox txt_Tab3_UploadMinute_Start 
            Alignment       =   1  '�a�k���
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
            Left            =   3450
            MaxLength       =   2
            TabIndex        =   78
            Top             =   525
            Width           =   375
         End
         Begin VB.CheckBox chk_Tab3_PreView 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   8160
            TabIndex        =   77
            Top             =   1320
            Value           =   1  '�֨�
            Width           =   1155
         End
         Begin VB.TextBox txt_Tab3_UploadHour_End 
            Alignment       =   1  '�a�k���
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
            Left            =   3075
            MaxLength       =   2
            TabIndex        =   76
            Top             =   870
            Width           =   375
         End
         Begin VB.TextBox txt_Tab3_UploadHour_Start 
            Alignment       =   1  '�a�k���
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
            Left            =   3075
            MaxLength       =   2
            TabIndex        =   75
            Top             =   525
            Width           =   375
         End
         Begin VB.CommandButton cmd_Tab3_PrintReport 
            BackColor       =   &H00C0FFC0&
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
            Left            =   8685
            Picture         =   "frm_Report_TRPPlan.frx":4436
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   74
            Top             =   240
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Tab3_Reset 
            BackColor       =   &H00FFC0FF&
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
            Height          =   360
            Left            =   5235
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   73
            Top             =   570
            Width           =   795
         End
         Begin VB.TextBox txt_Tab3_UploadDate_Start 
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
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   72
            Top             =   525
            Width           =   1185
         End
         Begin VB.TextBox txt_Tab3_UploadDate_End 
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
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   71
            Top             =   870
            Width           =   1185
         End
         Begin VB.ComboBox cmb_Tab3_AreaCode 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            ItemData        =   "frm_Report_TRPPlan.frx":4740
            Left            =   1110
            List            =   "frm_Report_TRPPlan.frx":4742
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   70
            Top             =   165
            Width           =   4500
         End
         Begin VB.CommandButton cmd_Tab3_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��Ƭd��"
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
            Left            =   6420
            Picture         =   "frm_Report_TRPPlan.frx":4744
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   69
            Top             =   255
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "��  �}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Index           =   3
            Left            =   9795
            Picture         =   "frm_Report_TRPPlan.frx":500E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   68
            Top             =   255
            Width           =   1005
         End
         Begin VB.CommandButton cmd_Tab3_SaveToExcel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   7515
            Picture         =   "frm_Report_TRPPlan.frx":5450
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   67
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            Index           =   32
            Left            =   3285
            TabIndex        =   123
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   31
            Left            =   900
            TabIndex        =   122
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Index           =   7
            Left            =   900
            TabIndex        =   106
            Top             =   1290
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Left            =   3210
            TabIndex        =   105
            Top             =   1290
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "mm�G0 ~ 59"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   12
            Left            =   4065
            TabIndex        =   86
            Top             =   825
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "hh�G0 ~ 23"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   14
            Left            =   4155
            TabIndex        =   85
            Top             =   615
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "yyyymmdd   hh   mm"
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
            Height          =   195
            Index           =   15
            Left            =   6000
            TabIndex        =   84
            Top             =   1320
            Width           =   1860
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��Ʈ榡�G"
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
            Height          =   195
            Index           =   16
            Left            =   4920
            TabIndex        =   83
            Top             =   1320
            Width           =   1050
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
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   4
            Left            =   1605
            TabIndex        =   82
            Top             =   915
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϰ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   81
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�^�Ǥ���G�_"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   2
            Left            =   405
            TabIndex        =   80
            Top             =   600
            Width           =   1440
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00008080&
            BorderWidth     =   2
            Height          =   1080
            Index           =   0
            Left            =   6300
            Top             =   150
            Width           =   4620
         End
      End
      Begin VB.Frame fam_Tab2 
         Height          =   1320
         Left            =   -74850
         TabIndex        =   30
         Top             =   720
         Width           =   11145
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
            Height          =   990
            Index           =   2
            Left            =   9930
            Picture         =   "frm_Report_TRPPlan.frx":6012
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   40
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab2_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��Ƭd��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   6420
            Picture         =   "frm_Report_TRPPlan.frx":6454
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   39
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab2_SaveToExcel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   7560
            Picture         =   "frm_Report_TRPPlan.frx":6D1E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   38
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab2_ReSet 
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
            Height          =   375
            Left            =   4035
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   37
            Top             =   885
            Width           =   630
         End
         Begin VB.CheckBox chk_Tab2_PreView 
            Caption         =   "�w���C�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   1695
            TabIndex        =   36
            Top             =   960
            Value           =   1  '�֨�
            Width           =   1425
         End
         Begin VB.TextBox txt_Tab2_DeliveryDate_End 
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
            Left            =   2730
            MaxLength       =   8
            TabIndex        =   35
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab2_DeliveryDate_Start 
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
            Left            =   1125
            MaxLength       =   8
            TabIndex        =   34
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab2_RouteNo_End 
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
            Left            =   3075
            MaxLength       =   10
            TabIndex        =   33
            Top             =   525
            Width           =   1605
         End
         Begin VB.TextBox txt_Tab2_RouteNo_Start 
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
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   32
            Top             =   525
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab2_PrintReport 
            BackColor       =   &H00C0FFC0&
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
            Height          =   990
            Left            =   8760
            Picture         =   "frm_Report_TRPPlan.frx":78E0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   31
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            Index           =   11
            Left            =   2445
            TabIndex        =   45
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   10
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Left            =   2790
            TabIndex        =   43
            Top             =   585
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   8
            Left            =   120
            TabIndex        =   42
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����榡�Gyyyymmdd"
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
            Height          =   195
            Index           =   5
            Left            =   4020
            TabIndex        =   41
            Top             =   225
            Width           =   2010
         End
      End
      Begin VB.Frame fam_Tab1_Header 
         Height          =   1530
         Left            =   -74850
         TabIndex        =   3
         Top             =   840
         Width           =   11145
         Begin VB.TextBox txt_Tab1_route_Start 
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
            Left            =   1155
            MaxLength       =   10
            TabIndex        =   117
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab1_route_End 
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
            Left            =   2835
            MaxLength       =   10
            TabIndex        =   116
            Top             =   1080
            Width           =   1365
         End
         Begin VB.CheckBox chk_Tab1_PreView 
            Caption         =   "�w���C�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   4680
            TabIndex        =   28
            Top             =   1080
            Value           =   1  '�֨�
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Tab1_PrintReport 
            BackColor       =   &H00C0FFC0&
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
            Height          =   990
            Left            =   8745
            Picture         =   "frm_Report_TRPPlan.frx":7BEA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   27
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab1_Reset 
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
            Height          =   360
            Left            =   5100
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   15
            Top             =   180
            Width           =   765
         End
         Begin VB.ComboBox cmb_Tab1_AreaCode 
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
            ForeColor       =   &H00800080&
            Height          =   315
            Left            =   1155
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   13
            Top             =   210
            Width           =   3960
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
            Height          =   990
            Index           =   1
            Left            =   9975
            Picture         =   "frm_Report_TRPPlan.frx":7EF4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   8
            Top             =   180
            Width           =   1065
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate_Start 
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
            Left            =   1155
            MaxLength       =   8
            TabIndex        =   7
            Top             =   615
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate_End 
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
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   6
            Top             =   630
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab1_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��Ƭd��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   6315
            Picture         =   "frm_Report_TRPPlan.frx":8336
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab1_SaveToExcel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   7545
            Picture         =   "frm_Report_TRPPlan.frx":8C00
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   4
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            Index           =   30
            Left            =   2565
            TabIndex        =   119
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   29
            Left            =   120
            TabIndex        =   118
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϽX"
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
            Left            =   135
            TabIndex        =   12
            Top             =   255
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����榡�Gyyyymmdd"
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
            Height          =   195
            Index           =   3
            Left            =   4200
            TabIndex        =   11
            Top             =   600
            Width           =   2010
         End
         Begin VB.Label Label1 
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
            Index           =   2
            Left            =   135
            TabIndex        =   10
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Left            =   2445
            TabIndex        =   9
            Top             =   690
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_VLL 
         Height          =   4545
         Left            =   150
         TabIndex        =   2
         Top             =   2160
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8017
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
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
      Begin VB.Frame fam_Tab0_Header 
         Height          =   1320
         Left            =   150
         TabIndex        =   1
         Top             =   720
         Width           =   13425
         Begin VB.CheckBox chkVllPallet 
            Caption         =   "�u�L�̪O�ި��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   4680
            TabIndex        =   152
            Top             =   600
            Value           =   1  '�֨�
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkDetail 
            Caption         =   "���ӦC�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   5520
            TabIndex        =   151
            Top             =   960
            Width           =   1185
         End
         Begin VB.CheckBox chkLFAShipList 
            Caption         =   "�t�Q�ץX�f��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   136
            Top             =   960
            Width           =   1545
         End
         Begin VB.CheckBox chkKAOShipList 
            Caption         =   "�t����X�f��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   135
            Top             =   240
            Width           =   1545
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
            Height          =   990
            Index           =   0
            Left            =   12240
            Picture         =   "frm_Report_TRPPlan.frx":97C2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   133
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_PrintReport 
            BackColor       =   &H00C0FFC0&
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
            Height          =   990
            Left            =   11040
            Picture         =   "frm_Report_TRPPlan.frx":9C04
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   132
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_SaveToExcel1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   9840
            Picture         =   "frm_Report_TRPPlan.frx":9F0E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   131
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��Ƭd��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   8640
            Picture         =   "frm_Report_TRPPlan.frx":AAD0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   130
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkNSLShipList 
            Caption         =   "�t�Ȱ��X�f��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   127
            Top             =   720
            Width           =   1545
         End
         Begin VB.CheckBox chkTHLShipList 
            Caption         =   "�t�ʨƥX�f��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   126
            Top             =   480
            Width           =   1545
         End
         Begin VB.CommandButton cmdLTHL01ShipDate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "THL�X�f���"
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
            Left            =   8640
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   125
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkSUMDetail 
            Caption         =   "���Ӷ��`�C�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   3840
            TabIndex        =   124
            Top             =   1000
            Width           =   1665
         End
         Begin VB.TextBox txt_Tab0_RouteNo_Start 
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
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   21
            Top             =   555
            Width           =   1605
         End
         Begin VB.TextBox txt_Tab0_RouteNo_End 
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
            MaxLength       =   10
            TabIndex        =   20
            Top             =   555
            Width           =   1605
         End
         Begin VB.TextBox txt_Tab0_DeliveryDate_Start 
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
            Left            =   1125
            MaxLength       =   8
            TabIndex        =   19
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab0_DeliveryDate_End 
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
            Left            =   2730
            MaxLength       =   8
            TabIndex        =   18
            Top             =   180
            Width           =   1245
         End
         Begin VB.CheckBox chk_Tab0_PreView 
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
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   2625
            TabIndex        =   17
            Top             =   1000
            Value           =   1  '�֨�
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Tab0_ReSet 
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
            Height          =   375
            Left            =   5040
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   16
            Top             =   120
            Width           =   630
         End
         Begin VB.CheckBox chk_Tab0_PrintedRoute 
            Caption         =   "�t�w�C�L�L�����u�s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Left            =   150
            TabIndex        =   22
            Top             =   1000
            Width           =   2625
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȭd7��"
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
            Left            =   4080
            TabIndex        =   128
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label1 
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
            Index           =   21
            Left            =   120
            TabIndex        =   26
            Top             =   615
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Index           =   20
            Left            =   2790
            TabIndex        =   25
            Top             =   615
            Width           =   240
         End
         Begin VB.Label Label1 
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
            TabIndex        =   24
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Index           =   18
            Left            =   2445
            TabIndex        =   23
            Top             =   240
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_VLLSum 
         Height          =   4950
         Left            =   -74850
         TabIndex        =   14
         Top             =   2400
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
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
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dg_Tab2_OrdersSum 
         Height          =   5250
         Left            =   -74850
         TabIndex        =   29
         Top             =   2040
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   9260
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   0
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
      Begin MSDataGridLib.DataGrid dg_Tab3_PickLoadCheck 
         Height          =   4365
         Left            =   -74805
         TabIndex        =   46
         Top             =   3000
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7699
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   6570
         Left            =   -74880
         TabIndex        =   47
         Top             =   720
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   11589
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "��ƿz��"
         TabPicture(0)   =   "frm_Report_TRPPlan.frx":B39A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(4)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label2(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label2(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(13)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(22)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label1(23)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dg_Tab4_RouteList"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmd_Tab4_Query_RouteDetail"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmd_Tab4_QueryBySRouteNo"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmd_Exit(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt_Tab4_SecondRouteNo"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt_Tab4_DeliveryDate_Start"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt_Tab4_DeliveryDate_End"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chk_Tab4_Selected"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "�C�L���"
         TabPicture(1)   =   "frm_Report_TRPPlan.frx":B3B6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmd_Tab4_PrintReport"
         Tab(1).Control(1)=   "chk_Tab4_PreView"
         Tab(1).Control(2)=   "dg_Tab4_OrderDetail"
         Tab(1).ControlCount=   3
         Begin VB.CheckBox chk_Tab4_Selected 
            Caption         =   "�d�ߵ��G����"
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
            Height          =   195
            Left            =   225
            TabIndex        =   111
            Top             =   1080
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab4_DeliveryDate_End 
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
            Left            =   2790
            MaxLength       =   8
            TabIndex        =   108
            Top             =   1395
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab4_DeliveryDate_Start 
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
            Left            =   1215
            MaxLength       =   8
            TabIndex        =   107
            Top             =   1395
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab4_SecondRouteNo 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   240
            MaxLength       =   10
            TabIndex        =   55
            Top             =   645
            Width           =   1650
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
            Height          =   945
            Index           =   5
            Left            =   9825
            Picture         =   "frm_Report_TRPPlan.frx":B3D2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   54
            Top             =   735
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab4_QueryBySRouteNo 
            BackColor       =   &H008080FF&
            Caption         =   "���s�z��"
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
            Left            =   2370
            Picture         =   "frm_Report_TRPPlan.frx":B814
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   53
            Top             =   405
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab4_PrintReport 
            BackColor       =   &H00C0FFC0&
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
            Height          =   930
            Left            =   -68535
            Picture         =   "frm_Report_TRPPlan.frx":C0DE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   52
            Top             =   555
            Width           =   2100
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
            Height          =   870
            Index           =   4
            Left            =   -65040
            Picture         =   "frm_Report_TRPPlan.frx":C3E8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   51
            Top             =   645
            Width           =   1065
         End
         Begin VB.CheckBox chk_Tab4_PreView 
            Caption         =   "�w���C�L"
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
            Height          =   240
            Left            =   -71070
            TabIndex        =   50
            Top             =   1170
            Value           =   1  '�֨�
            Width           =   1380
         End
         Begin VB.CommandButton cmd_Tab4_Query_RouteDetail 
            BackColor       =   &H00FF8080&
            Caption         =   "���u���`"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   8565
            Picture         =   "frm_Report_TRPPlan.frx":C82A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   49
            Top             =   735
            Width           =   1035
         End
         Begin MSDataGridLib.DataGrid dg_Tab4_OrderDetail 
            Height          =   4740
            Left            =   -74850
            TabIndex        =   48
            Top             =   1650
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   8361
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab4_RouteList 
            Height          =   4530
            Left            =   120
            TabIndex        =   56
            Top             =   1845
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   7990
            _Version        =   393216
            BackColor       =   -2147483624
            Cols            =   11
            TextStyleFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   11
         End
         Begin VB.Label Label1 
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
            Index           =   23
            Left            =   2535
            TabIndex        =   110
            Top             =   1455
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   22
            Left            =   225
            TabIndex        =   109
            Top             =   1455
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ƨ����u�s��"
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
            Height          =   240
            Index           =   13
            Left            =   225
            TabIndex        =   65
            Top             =   345
            Width           =   1530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ާ@�B�J�G"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   195
            Index           =   3
            Left            =   4140
            TabIndex        =   64
            Top             =   555
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����w���w�ѹp�g�L���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   195
            Index           =   0
            Left            =   -74400
            TabIndex        =   63
            Top             =   1050
            Width           =   2310
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "Fujitsu 16 ADV �C�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   180
            Index           =   0
            Left            =   -73845
            TabIndex        =   62
            Top             =   1365
            Width           =   1755
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "A4 ���L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   0
            Left            =   -73200
            TabIndex        =   61
            Top             =   585
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "����榡�G"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   225
            Index           =   3
            Left            =   -74415
            TabIndex        =   60
            Top             =   585
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "1. ��J [�ƨ����u�s��]�A���� [���s�z��]"
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
            Index           =   1
            Left            =   4425
            TabIndex        =   59
            Top             =   840
            Width           =   3795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "2. �T�{���C�L�����u�s�����"
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
            Index           =   2
            Left            =   4425
            TabIndex        =   58
            Top             =   1125
            Width           =   2745
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "3. ���� [���u���`]�A���X�C�L���"
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
            Index           =   4
            Left            =   4425
            TabIndex        =   57
            Top             =   1410
            Width           =   3165
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab6_PlanList 
         Height          =   4410
         Left            =   -74880
         TabIndex        =   150
         Top             =   2280
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   7779
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
   End
End
Attribute VB_Name = "frm_Report_TRPPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private blVLLReportEventEnable As Boolean   'VLL�˸�
Private arAreaCode() As String
Private rs_Tab0_VLL As ADODB.Recordset           'VLL �˸���G�ݿ���s�M��
Private rs_Tab0_VLLSum As ADODB.Recordset        'VLL �˸��`��
Private rs_Tab0_VLLDetail As ADODB.Recordset     'VLL �˸����Ӫ�
Private rs_Tab0_VLLSUMDetail As ADODB.Recordset     'VLL �˸����Ӫ�
Private rs_Tab0_VLLOrder As ADODB.Recordset      'VLL �X�f��
Private rs_Tab1_VLLSum As ADODB.Recordset        '�����˸��J�`��
Private rs_Tab2_OrdersSum As ADODB.Recordset     '�q���`��
Private rs_Tab3_PickLoadCheck As ADODB.Recordset '�z�f�˸��]�֪�
Private rs_Tab4_OrderDetail As ADODB.Recordset   '��B�����u���`��G�q��W��
Private rs_Tab5_PlanList As ADODB.Recordset      '�ƨ��@����
Private rs_Tab5_TRPPlanList As ADODB.Recordset   '�ƨ��@����_�s
Private rs_Tab6_PlanList As ADODB.Recordset
Private str_SQL_Excel As String
Private strAccessDBFileName_FullPath As String
Private MSAccessAP As access.Application
Private rs_Access As ADODB.Recordset
Private rs_Access1 As ADODB.Recordset            '��B�����u���`��
Private rs_Access2 As ADODB.Recordset            '��B�����u���`��A
Private rs_Tab0_VLLDetailxSection As ADODB.Recordset

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Set rs_Tab0_VLL = Nothing
Set rs_Tab0_VLLSum = Nothing
Set rs_Tab0_VLLDetail = Nothing
Set rs_Tab0_VLLSUMDetail = Nothing
Set rs_Tab0_VLLOrder = Nothing
Set rs_Tab1_VLLSum = Nothing
Set rs_Tab2_OrdersSum = Nothing
Set rs_Tab3_PickLoadCheck = Nothing
Set rs_Tab4_OrderDetail = Nothing
Set rs_Tab5_PlanList = Nothing
Set rs_Tab5_TRPPlanList = Nothing
Set rs_Access1 = Nothing
Set rs_Access2 = Nothing
Set rs_Tab0_VLLDetailxSection = Nothing

Unload Me
End Sub

Private Sub cmd_Tab0_SaveToExcel1_Click()

Recordset2Excel "VLL�˸�", rs_Tab0_VLL
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd_Tab2_SaveToExcel_Click()
'�q���`�� >> �� EXCEL
Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel �ɮצW��
CmnDialog.DialogTitle = "��s Excel ��"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "�q���`��_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel�ɮ�(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '�� [�}������] ��ܤ�����A���U [����] �s
   msg_text = "��� [����] ���s�A������ Excel ���ۦ�s��"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab2_OrdersSum) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab2_OrdersSum.MoveFirst
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q���`��-�� EXCEL", Me.Caption, "cmd_Tab2_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_PrintReport_Click()
'�z�f�˸��]�֪� >> ����C�L
If rs_Tab3_PickLoadCheck Is Nothing Then Exit Sub
If rs_Tab3_PickLoadCheck.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. ��Ƽg�X Access ��Ʈw >> �q���`��
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From �z�f�˸��]�֪�"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "�z�f�˸��]�֪�", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab3_PickLoadCheck.MoveFirst
Do While Not rs_Tab3_PickLoadCheck.EOF
   rs_Access.AddNew
   rs_Access.Fields("�Ǹ�").Value = rs_Tab3_PickLoadCheck.Fields("�s��").Value
   rs_Access.Fields("�ϰ�").Value = rs_Tab3_PickLoadCheck.Fields("�B�e�ϰ�").Value
   rs_Access.Fields("�W�Ǥ��").Value = rs_Tab3_PickLoadCheck.Fields("�^�Ǥ��").Value
   rs_Access.Fields("���u�s��").Value = rs_Tab3_PickLoadCheck.Fields("���u�s��").Value
   rs_Access.Fields("�X�����").Value = rs_Tab3_PickLoadCheck.Fields("�X�����").Value
   rs_Access.Fields("�q��i��").Value = rs_Tab3_PickLoadCheck.Fields("�q���").Value
   rs_Access.Fields("�e�f�I").Value = rs_Tab3_PickLoadCheck.Fields("�e�f�I").Value
   rs_Access.Fields("�Ȥ�²��").Value = rs_Tab3_PickLoadCheck.Fields("�Ȥ�²��").Value
   rs_Access.Fields("�c��").Value = rs_Tab3_PickLoadCheck.Fields("�c��").Value
   rs_Access.Fields("�O��").Value = rs_Tab3_PickLoadCheck.Fields("�O��").Value
   rs_Access.Fields("���q").Value = rs_Tab3_PickLoadCheck.Fields("���q").Value
   rs_Access.Fields("���n").Value = rs_Tab3_PickLoadCheck.Fields("���n").Value
   rs_Access.Fields("�f�B���q").Value = rs_Tab3_PickLoadCheck.Fields("�f�B���q").Value
   rs_Access.Fields("����").Value = rs_Tab3_PickLoadCheck.Fields("����").Value
   rs_Access.Fields("����").Value = rs_Tab3_PickLoadCheck.Fields("����").Value
   rs_Access.Fields("�w�p����ɶ�").Value = rs_Tab3_PickLoadCheck.Fields("�w�p����ɶ�").Value
   rs_Access.Fields("�X�Y").Value = rs_Tab3_PickLoadCheck.Fields("�X�Y").Value
   rs_Access.Update
   rs_Tab3_PickLoadCheck.MoveNext
Loop
rs_Tab3_PickLoadCheck.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access �C�L����
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab3_PreView.Value = vbChecked Then
   '�w���C�L
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "�z�f�˸��]�֪�", acViewPreview
Else
   '�����C�L�ܦL���
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "�z�f�˸��]�֪�", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cnAccess.RollbackTrans
      Tran_Level = 0
   End If
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�z�f�˸��]�֪�-�C�L", Me.Caption, "cmd_Tab3_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab0_PrintReport_Click()
'����C�L
If rs_Tab0_VLL Is Nothing Then Exit Sub
If rs_Tab0_VLL.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle
Dim strPrintDate As String       '�C�L�ɶ�
Dim strUserName As String        '�C�L��
Dim strRouteNo As String         '���u�s��
Dim iLoop As Double
Dim strTmp As String, strRoute_No As String '�Ҧ����s�P���s�Ȧs���
Dim i As Integer, strCompany As String

blVLLReportEventEnable = False
Dim strSelectedRouteNo As String    '��������u�s��

strSelectedRouteNo = ""
rs_Tab0_VLL.MoveFirst
Do While Not rs_Tab0_VLL.EOF
   If Len(Trim(rs_Tab0_VLL.Fields(1).Value)) > 0 Then
      If strSelectedRouteNo = "" Then
         strSelectedRouteNo = "'" & rs_Tab0_VLL.Fields("���u�s��").Value & "'"
      Else
         strSelectedRouteNo = strSelectedRouteNo & ",'" & rs_Tab0_VLL.Fields("���u�s��").Value & "'"
      End If
   End If
   rs_Tab0_VLL.MoveNext
Loop

blVLLReportEventEnable = True
If strSelectedRouteNo = "" Then
   msg_text = "��ƿ��~�G��������C�L�� [���u�s��]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
cmd_Tab0_PrintReport.Enabled = False    'daniel

Screen.MousePointer = 11

'�@�BVLL �˸��`��
str_SQL = "Select Distinct ' ' as '��',�X�����,���u�s��,���P���X,����,�r�p�H,�B�餽�q,�f�D�渹,�Ȥ�s��,�Ȥ�W��," & _
          "   �e�f�a�},�q��Ƶ�,�c��,�Ӽ�,�O��,���n,���q,�q���,�w�p������,�w�p����ɶ�,�X�Y�Ȧs,�C�L����," & _
          "   �C�L�ɶ�,���w�������O,�ƨ��� as LoginUserID,����," & _
          "   '                          ' as ���Ƶ��O,Receipt_No,�G���ƨ����s,�q������,��� " & _
          "From Report_VLL Where �G���ƨ����s IN (" & strSelectedRouteNo & ") or ���u�s�� in (" & strSelectedRouteNo & ") order by �G���ƨ����s "
          
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ���ƩΥX������w�W�L�d�߭���!!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   blVLLReportEventEnable = True
   cmd_Tab0_PrintReport.Enabled = True     'daniel
   Screen.MousePointer = 0
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLSum)
tmp_Rs.Close

'1.1 ����h���A���ƭp��
Dim strExtern As String
rs_Tab0_VLLSum.Filter = " ���� > 1 "
rs_Tab0_VLLSum.Sort = " �f�D�渹 desc "
If Not rs_Tab0_VLLSum.EOF Then
   Do While Not rs_Tab0_VLLSum.EOF
   
      If rs_Tab0_VLLSum.Fields("�f�D�渹").Value <> strExtern Then
         '��s���ƭp�������
         str_SQL = "exec VLL_Extern_CarCount '" & rs_Tab0_VLLSum.Fields("�f�D�渹").Value & "' "
         cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
         
         strExtern = rs_Tab0_VLLSum.Fields("�f�D�渹").Value
      End If
      
      '���o���ƭp�������
      str_SQL = "Select Rtrim(isnull(Car_Notes,' ')) as ���Ƶ��O From TRP02T Where Receipt_No = '" & rs_Tab0_VLLSum.Fields("Receipt_No").Value & "'"
      tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
      If Not tmp_Rs.EOF Then
         rs_Tab0_VLLSum.Fields("���Ƶ��O").Value = tmp_Rs.Fields("���Ƶ��O").Value
      End If
      tmp_Rs.Close
      rs_Tab0_VLLSum.MoveNext
   Loop
End If

rs_Tab0_VLLSum.Filter = adFilterNone
rs_Tab0_VLLSum.Sort = " �G���ƨ����s desc "

'1-2. ���o DB Server �ɶ�
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select Convert(varchar,GetDate(),111) + ' ' + convert(varchar,GetDate(),108) as '�C�L�ɶ�' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
strPrintDate = tmp_Rs.Fields("�C�L�ɶ�").Value
tmp_Rs.Close

'1-3. ��s TRP01T �����
strRouteNo = ""
rs_Tab0_VLLSum.MoveFirst
Do While Not rs_Tab0_VLLSum.EOF
   If strRouteNo <> rs_Tab0_VLLSum.Fields("�G���ƨ����s").Value Then
      rs_Tab0_VLLSum.Fields("�C�L����").Value = rs_Tab0_VLLSum.Fields("�C�L����").Value + 1
      rs_Tab0_VLLSum.Fields("�C�L�ɶ�").Value = strPrintDate
      
      '�H ���u�s�� ����ƳB�z�̾�
      strRouteNo = rs_Tab0_VLLSum.Fields("�G���ƨ����s").Value
      str_SQL = "Update TRP01T Set VLListCount = " & rs_Tab0_VLLSum.Fields("�C�L����").Value & ",VLListPrintDate = '" & strPrintDate & "' " & _
                "Where Route_No = '" & strRouteNo & "' or C_Route_No = '" & strRouteNo & "'"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0_VLLSum.MoveNext
Loop
rs_Tab0_VLLSum.MoveFirst

If chkSUMDetail.Value = vbChecked Then

        '�G�BVLL �[�`�˸����Ӫ�
'        str_SQL = "Select ���u�s��,�X�����,�w�p������,�w�p����ɶ�,�X�Y�Ȧs,���P���X,����,�r�p�H,�B�餽�q," & _
'                  "  �ƨ���,�f�D�渹,TMS�渹,�ܧO,�N��,�f��,�~�W,�X�f�c��=isnull(�X�f�c��,0),�X�f�Ӽ�=isnull(�X�f�Ӽ�,0),�z�f���q,�z�f���n,�C�L����,�C�L�ɶ�,�G���ƨ����s,�ƨ��c��,�s�y��,����� " & _
'                  "From Report_VLLSUMDetail Where �G���ƨ����s IN (" & strSelectedRouteNo & ") or ���u�s�� in (" & strSelectedRouteNo & ")"
                    
         str_SQL = "Select ���u�s�� = a1.Route_No,�X����� = Case When Isnull(a1.C_Route_No,'') = '' Then Convert(varchar,a1.Delivery_Date,112) else Convert(varchar,t01t2.Delivery_Date,112) End " & _
                    ",�w�p������ = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Date,'')) else Rtrim(t05t2.Expect_Date) End " & _
                    ",�w�p����ɶ� = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Time,'')) else Rtrim(t05t2.Expect_Time) End " & _
                    ",�X�Y�Ȧs = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Dock_No,'')) else Rtrim(t05t2.Dock_No) End " & _
                    ",���P���X = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Vehicle_ID_No) else Rtrim(t05t2.Vehicle_ID_No) End " & _
                    ",���� = Case When Isnull(a1.C_Route_No,'') = '' Then Round(Cast(a2.Drive_Times as float),2) else Round(Cast(t05t2.Drive_Times as float),2) End " & _
                    ",�r�p�H = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Driver) else Rtrim(t05t2.Driver) End " & _
                    ",�B�餽�q = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(c2.C_Name,'')) else Rtrim(Isnull(t08m2.C_Name,'')) End " & _
                    ",�ƨ��� = Case When Isnull(a1.C_Route_No,'') = '' Then Isnull(Rtrim(a1.AddWho),'') else Rtrim(t01t2.AddWho) End " & _
                    ",�f�D�渹 = ' ',TMS�渹 = ' ',�ܧO = rtrim(l.lottable06),�N�� = substring(sp.skugroup,7,1),�a�q = case when t02t.priority = 'C' then '' else rtrim(loc.sectionkey) end,�f�� = Rtrim(t03t.Product_No),�~�W = Isnull(Rtrim(sp.Descr),'') " & _
                    ",�X�f�c�� = case when sp.casecnt = 0 then 0 else Isnull(floor(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) /(sp.Casecnt)),0) end " & _
                    ",�X�f�Ӽ� = case when sp.casecnt = 0 then sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) else cast(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) as int) % cast(sp.Casecnt as int) end " & _
                    ",�z�f���q = Isnull(Round((sp.STDGrossWGT * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0),�z�f���n = Isnull(Round((sp.STDCUBE * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0) " & _
                    ",�C�L���� = Isnull(a1.VLListCount,0),�C�L�ɶ� = Isnull((Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()),111) + ' ' + Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()) , 108)),'') " & _
                    ",�G���ƨ����s = Case When Isnull(a1.C_Route_No,'') = '' Then a1.Route_No Else a1.C_Route_No End,�ƨ��c�� = 0,�s�y�� = ' ',����� = ' ' " & _
                    "From TRP01T a1 inner join trp02t t02t on t02t.Route_No = a1.Route_No join TRP03T t03t on t03t.receipt_No = t02t.receipt_No and a1.Route_No <> 'D' inner join TRP05T a2 on a2.Route_No = a1.Route_No " & _
                    "inner join TRP09M c1 on c1.Vehicle_ID_No = a2.Vehicle_ID_No inner join gv_SKUxpack sp on sp.StorerKey = t03t.StorerKey and sp.SKU = t03t.Product_No " & _
                    "Left outer join TRP01T t01t2 on t01t2.Route_No = a1.C_Route_No " & _
                    "Left outer join TRP05T t05t2 on t05t2.Route_No = a1.C_Route_No Left outer join TRP09M t09m2 on t09m2.Vehicle_ID_No = t05t2.Vehicle_ID_No " & _
                    "Left outer join TRP08M t08m2 on t08m2.Company_Code = t09m2.TRP_Company_Code Left outer join TRP08M c2 on c2.Company_Code = c1.TRP_Company_Code " & _
                    "Left join " & strWMSDB & "..orders o (nolock) on o.updatesource = t03t.receipt_no Left join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey and od.externlineno = t03t.seq_no  " & _
                    "Left join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = od.orderkey and od.orderlinenumber  = p.orderlinenumber Left join " & strWMSDB & "..lotattribute l (nolock) on p.lot = l.lot and p.sku = l.sku Left join " & strWMSDB & "..loc loc (nolock) on p.loc = loc.loc " & _
                    "where a1.C_Route_No IN (" & strSelectedRouteNo & ") or a1.Route_No in (" & strSelectedRouteNo & ") " & _
                    "Group by t02t.priority,a1.Route_No , loc.sectionkey,a1.Delivery_Date , a2.Expect_Date , a2.Expect_Time , a2.Dock_No , a2.Vehicle_ID_No,a2.Drive_Times , a2.Driver , c2.C_Name , a1.C_Route_No , a1.AddWho,t03t.Product_No,sp.Descr,a1.VLListCount , a1.VLListPrintDate ,sp.CaseCnt,sp.STDGROSSWGT,sp.STDCUBE,t01t2.Delivery_Date , t05t2.Expect_Date , t05t2.Expect_Time , t05t2.Dock_No , t05t2.Vehicle_ID_No , t05t2.Driver , " & _
                    "t08m2.C_Name , t01t2.AddWho , t05t2.Drive_Times,l.lottable06,substring(sp.skugroup,7,1) "

        Call DB_CheckConnectStatus
        Call Confirm_Recordset_Closed(tmp_Rs)
        cn.CommandTimeout = 600
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
           tmp_Rs.Close
           cmd_Tab0_PrintReport.Enabled = True
           msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧭q����Ӹ��"
           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
           Screen.MousePointer = vbDefault
           blVLLReportEventEnable = True
           Exit Sub
        End If
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLSUMDetail)
        tmp_Rs.Close
End If
If chkDetail.Value = vbChecked Then

        '�G�BVLL �˸����Ӫ�
'        str_SQL = "Select ���u�s��,�X�����,�w�p������,�w�p����ɶ�,�X�Y�Ȧs,���P���X,����,�r�p�H,�B�餽�q," & _
'                  "  �ƨ���,�f�D�渹,TMS�渹,�ܧO,�N��,�f��,�~�W,�X�f�c��=isnull(�X�f�c��,0),�X�f�Ӽ�=isnull(�X�f�Ӽ�,0),�z�f���q,�z�f���n,�C�L����,�C�L�ɶ�,�G���ƨ����s,�ƨ��c��,�s�y��,����� " & _
'                  "From Report_VLLDetail Where �G���ƨ����s IN (" & strSelectedRouteNo & ") or ���u�s�� in (" & strSelectedRouteNo & ")"
        
        str_SQL = "Select ���u�s�� = a1.Route_No,�X����� = Case When Isnull(a1.C_Route_No,'') = '' Then Convert(varchar,a1.Delivery_Date,112) else Convert(varchar,t01t2.Delivery_Date,112) End " & _
                    ",�w�p������ = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Date,'')) else Rtrim(t05t2.Expect_Date) End " & _
                    ",�w�p����ɶ� = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Expect_Time,'')) else Rtrim(t05t2.Expect_Time) End " & _
                    ",�X�Y�Ȧs = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(a2.Dock_No,'')) else Rtrim(t05t2.Dock_No) End " & _
                    ",���P���X = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Vehicle_ID_No) else Rtrim(t05t2.Vehicle_ID_No) End " & _
                    ",���� = Case When Isnull(a1.C_Route_No,'') = '' Then Round(Cast(a2.Drive_Times as float),2) else Round(Cast(t05t2.Drive_Times as float),2) End " & _
                    ",�r�p�H = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(a2.Driver) else Rtrim(t05t2.Driver) End " & _
                    ",�B�餽�q = Case When Isnull(a1.C_Route_No,'') = '' Then Rtrim(Isnull(c2.C_Name,'')) else Rtrim(Isnull(t08m2.C_Name,'')) End " & _
                    ",�ƨ��� = Case When Isnull(a1.C_Route_No,'') = '' Then Isnull(Rtrim(a1.AddWho),'') else Rtrim(t01t2.AddWho) End " & _
                    ",�f�D�渹 = t03t.extern,TMS�渹 = t03t.receipt_no,�ܧO = rtrim(l.lottable06),�N�� = substring(sp.skugroup,7,1),�a�q = case when t02t.priority = 'C' then '' else rtrim(loc.sectionkey) end,�f�� = Rtrim(t03t.Product_No),�~�W = Isnull(Rtrim(sp.Descr),'') " & _
                    ",�X�f�c�� = case when sp.casecnt = 0 then 0 else Isnull(floor(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) /(sp.Casecnt)),0) end " & _
                    ",�X�f�Ӽ� = case when sp.casecnt = 0 then sum(isnull(p.qty,0)) else cast(sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0)) as int) % cast(sp.Casecnt as int) end " & _
                    ",�z�f���q = Isnull(Round((sp.STDGrossWGT * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0),�z�f���n = Isnull(Round((sp.STDCUBE * sum(isnull(case when t02t.priority = 'C' then t03t.order_qty else p.qty end,0))),2),0) " & _
                    ",�C�L���� = Isnull(a1.VLListCount,0),�C�L�ɶ� = Isnull((Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()),111) + ' ' + Convert(varchar,Isnull(a1.VLListPrintDate,Getdate()) , 108)),'') " & _
                    ",�G���ƨ����s = Case When Isnull(a1.C_Route_No,'') = '' Then a1.Route_No Else a1.C_Route_No End,�ƨ��c�� = 0,�s�y�� = isnull(convert(char(10),l.lottable04,111),''),����� = isnull(convert(char(10),l.lottable05,111),'') " & _
                    "From TRP01T a1 inner join trp02t t02t on t02t.Route_No = a1.Route_No join TRP03T t03t on t03t.receipt_No = t02t.receipt_No and a1.Route_No <> 'D' inner join TRP05T a2 on a2.Route_No = a1.Route_No " & _
                    "inner join TRP09M c1 on c1.Vehicle_ID_No = a2.Vehicle_ID_No inner join gv_SKUxpack sp on sp.StorerKey = t03t.StorerKey and sp.SKU = t03t.Product_No " & _
                    "Left outer join TRP01T t01t2 on t01t2.Route_No = a1.C_Route_No " & _
                    "Left outer join TRP05T t05t2 on t05t2.Route_No = a1.C_Route_No Left outer join TRP09M t09m2 on t09m2.Vehicle_ID_No = t05t2.Vehicle_ID_No " & _
                    "Left outer join TRP08M t08m2 on t08m2.Company_Code = t09m2.TRP_Company_Code Left outer join TRP08M c2 on c2.Company_Code = c1.TRP_Company_Code " & _
                    "Left join " & strWMSDB & "..orders o (nolock) on o.updatesource = t03t.receipt_no Left join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey and od.externlineno = t03t.seq_no  " & _
                    "Left join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = od.orderkey and od.orderlinenumber  = p.orderlinenumber Left join " & strWMSDB & "..lotattribute l (nolock) on p.lot = l.lot and p.sku = l.sku Left join " & strWMSDB & "..loc loc (nolock) on p.loc = loc.loc " & _
                    "where a1.C_Route_No IN (" & strSelectedRouteNo & ") or a1.Route_No in (" & strSelectedRouteNo & ") " & _
                    "Group by t02t.priority,a1.Route_No , a1.Delivery_Date , a2.Expect_Date , a2.Expect_Time , a2.Dock_No , a2.Vehicle_ID_No,a2.Drive_Times , a2.Driver , c2.C_Name , a1.C_Route_No , a1.AddWho , t03t.Product_No ,sp.Descr,a1.VLListCount , a1.VLListPrintDate ,sp.CaseCnt,sp.STDGROSSWGT,sp.STDCUBE,t01t2.Delivery_Date , t05t2.Expect_Date , t05t2.Expect_Time , t05t2.Dock_No , t05t2.Vehicle_ID_No , t05t2.Driver , " & _
                    "t08m2.C_Name , t01t2.AddWho , t05t2.Drive_Times,l.lottable04,l.lottable05,t03t.extern,t03t.receipt_no,l.lottable06,substring(sp.skugroup,7,1),loc.sectionkey "
        
        Call DB_CheckConnectStatus
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
           tmp_Rs.Close
           cmd_Tab0_PrintReport.Enabled = True
           msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧭q����Ӹ��"
           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
           Screen.MousePointer = vbDefault
           blVLLReportEventEnable = True
           cmd_Tab0_PrintReport.Enabled = True
           Screen.MousePointer = 0
           Exit Sub
        End If
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLDetail)
        tmp_Rs.Close
        
End If
        
        '4. ��Ƽg�X Access ��Ʈw >> VLL�W�f��
        Call AccessDB_Connect
        Tran_Level = 0
        Tran_Level = cnAccess.BeginTrans
        str_SQL = "Delete From VLL�˸��`��"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL�˸��`��", cnAccess, adOpenStatic, adLockOptimistic
            
        '���t�m���
        Dim rsTmp As New ADODB.Recordset, strSectionKey As String, lngSectionCS As Long, lngSectionEA As Long, strSectionQty As String
        
        str_SQL = "select distinct sectionkey ,o.updatesource " & _
        "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = o.orderkey " & _
        "join " & strWMSDB & "..loc l (nolock) on l.loc = p.loc " & _
        "left join trp02t t2 on t2.receipt_no = o.updatesource " & _
        "left join trp01t t1 on t2.route_no = t1.route_no " & _
        "where isnull(t1.c_route_no,t1.route_no) in (" & strSelectedRouteNo & ") " & _
        "order by sectionkey ,o.updatesource "
        
        tmp_Rs.Open str_SQL, cn
        Call Replication_Recordset(tmp_Rs, rsTmp)
        tmp_Rs.Close
        
        '���a�q�ƶq
        str_SQL = "select route_no = isnull(t1.c_route_no,t1.route_no),sectionkey " & _
                ",CS = case when s.casecnt = 0 then 0 else Isnull(floor(sum(isnull(p.qty,0)) /(s.Casecnt)),0) end " & _
                ",EA = case when s.casecnt = 0 then sum(isnull(p.qty,0)) else cast(sum(isnull(p.qty,0)) as int) % cast(s.Casecnt as int) end " & _
                "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = o.orderkey " & _
                "join " & strWMSDB & "..loc l (nolock) on l.loc = p.loc " & _
                "join gv_skuxpack s on s.storerkey = p.storerkey and s.sku = p.sku " & _
                "left join trp02t t2 on t2.receipt_no = o.updatesource " & _
                "left join trp01t t1 on t2.route_no = t1.route_no " & _
                "where isnull(t1.c_route_no,t1.route_no) in (" & strSelectedRouteNo & ") " & _
                "group by sectionkey ,isnull(t1.c_route_no,t1.route_no),p.orderkey,p.orderlinenumber,s.Casecnt " & _
                "union all select '          ','                                                                                                         ',0,0 " & _
                "order by isnull(t1.c_route_no,t1.route_no),sectionkey  "
        tmp_Rs.Open str_SQL, cn
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLDetailxSection)
        tmp_Rs.Close
        
        rs_Tab0_VLLSum.MoveFirst
        Do While Not rs_Tab0_VLLSum.EOF
        
            '�έp���sx�a�q�X�f�c�ƭӼ�
            rs_Tab0_VLLDetailxSection.MoveFirst
            strRouteNo = rs_Tab0_VLLSum("�G���ƨ����s")
            strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey")
            strSectionQty = "": lngSectionCS = 0: lngSectionEA = 0
            
            '�����s�L�t�f�h�U�@�Ӹ��s
            rs_Tab0_VLLDetailxSection.Filter = "(route_no = '" & rs_Tab0_VLLSum("�G���ƨ����s") & "')"
            If rs_Tab0_VLLDetailxSection.EOF Then rs_Tab0_VLLDetailxSection.Filter = "": GoTo nestRoute
            strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey")
    
            Do While Not rs_Tab0_VLLDetailxSection.EOF
            
                If strRouteNo = rs_Tab0_VLLDetailxSection("route_no") Then '�P���s
                    If strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey") Then '�P�@�a�q
                        lngSectionCS = lngSectionCS + rs_Tab0_VLLDetailxSection("cs")
                        lngSectionEA = lngSectionEA + rs_Tab0_VLLDetailxSection("ea")
                    Else
                        strSectionQty = strSectionQty & RTrim(strSectionKey) & " �@ " & lngSectionCS & "CS / " & lngSectionEA & "EA;"
                        strSectionKey = rs_Tab0_VLLDetailxSection("sectionkey")
                        lngSectionCS = rs_Tab0_VLLDetailxSection("cs")
                        lngSectionEA = rs_Tab0_VLLDetailxSection("ea")
                    End If
                End If
                rs_Tab0_VLLDetailxSection.MoveNext
            Loop
            strSectionQty = strSectionQty & RTrim(strSectionKey) & " �@ " & lngSectionCS & "CS / " & lngSectionEA & "EA ;"
            
            rs_Tab0_VLLDetailxSection.Filter = ""
nestRoute:
        
           rs_Access.AddNew
           rs_Access.Fields("�Ǹ�").Value = rs_Tab0_VLLSum.Fields("�s��").Value
           rs_Access.Fields("���u�s��").Value = rs_Tab0_VLLSum.Fields("���u�s��").Value
           rs_Access.Fields("�X�����").Value = rs_Tab0_VLLSum.Fields("�X�����").Value
           rs_Access.Fields("����").Value = rs_Tab0_VLLSum.Fields("���P���X").Value
           rs_Access.Fields("�q��").Value = rs_Tab0_VLLSum.Fields("�r�p�H").Value
           rs_Access.Fields("�f�B��").Value = rs_Tab0_VLLSum.Fields("�B�餽�q").Value
           rs_Access.Fields("�q��s��").Value = rs_Tab0_VLLSum.Fields("�f�D�渹").Value
           rs_Access.Fields("Receipt_no").Value = rs_Tab0_VLLSum.Fields("Receipt_no").Value
           rs_Access.Fields("�Ȥ�s��").Value = rs_Tab0_VLLSum.Fields("�Ȥ�s��").Value
           rs_Access.Fields("�Ȥ�W��").Value = rs_Tab0_VLLSum.Fields("�Ȥ�W��").Value
           rs_Access.Fields("�e�f�a�}").Value = rs_Tab0_VLLSum.Fields("�e�f�a�}").Value
           rs_Access.Fields("�e�f�Ƶ�").Value = rs_Tab0_VLLSum.Fields("�q��Ƶ�").Value
           rs_Access.Fields("�e�f�Ƶ�").Value = rs_Tab0_VLLSum.Fields("�q��Ƶ�").Value
           rs_Access.Fields("�X�f�O��").Value = rs_Tab0_VLLSum.Fields("�O��").Value
           rs_Access.Fields("�X�f�c��").Value = rs_Tab0_VLLSum.Fields("�c��").Value
           rs_Access.Fields("�X�f�Ӽ�").Value = rs_Tab0_VLLSum.Fields("�Ӽ�").Value
           rs_Access.Fields("���n").Value = rs_Tab0_VLLSum.Fields("���n").Value
           rs_Access.Fields("���q").Value = rs_Tab0_VLLSum.Fields("���q").Value
           rs_Access.Fields("�C�L����").Value = rs_Tab0_VLLSum.Fields("�C�L����").Value + 1 'daniel<�Ĥ@���C�L����1>
           rs_Access.Fields("�C�L�ɶ�").Value = rs_Tab0_VLLSum.Fields("�C�L�ɶ�").Value
           rs_Access.Fields("���w�������O").Value = rs_Tab0_VLLSum.Fields("���w�������O").Value
           rs_Access.Fields("LoginUserID").Value = rs_Tab0_VLLSum.Fields("LoginUserID").Value
           rs_Access.Fields("�w�p������").Value = rs_Tab0_VLLSum.Fields("�w�p������").Value
           rs_Access.Fields("�w�p����ɶ�").Value = rs_Tab0_VLLSum.Fields("�w�p����ɶ�").Value
           rs_Access.Fields("�X�Y�Ȧs").Value = rs_Tab0_VLLSum.Fields("�X�Y�Ȧs").Value
           rs_Access.Fields("���Ƶ��O").Value = rs_Tab0_VLLSum.Fields("���Ƶ��O").Value
           rs_Access.Fields("�G���ƨ����s").Value = rs_Tab0_VLLSum.Fields("�G���ƨ����s").Value
           rs_Access.Fields("�a�q�ƶq").Value = IIf(Len(RTrim(strSectionQty)) = 0, "���t�m", strSectionQty)

            rsTmp.Filter = "(updatesource = '" & rs_Tab0_VLLSum.Fields("Receipt_no") & "')"
        
            strSectionKey = ""
        
            If rsTmp.EOF Then
                rs_Access.Fields("�q������") = "���t�m"
            Else
                Do While Not rsTmp.EOF
                    strSectionKey = strSectionKey & RTrim(rsTmp("sectionkey")) & ";"
                    rsTmp.MoveNext
                Loop
                rs_Access.Fields("�q������") = strSectionKey
            End If
            
            rsTmp.Filter = ""
           
           rs_Access.Fields("���").Value = rs_Tab0_VLLSum.Fields("���").Value
           
           '���Ҧ����s
           If strTmp <> rs_Tab0_VLLSum.Fields("�G���ƨ����s") Then
               strTmp = rs_Tab0_VLLSum.Fields("�G���ƨ����s")
               
               str_SQL = "select distinct route_no from trp01t where isnull(c_route_no,route_no) = '" & rs_Tab0_VLLSum.Fields("�G���ƨ����s") & "'and left(route_no,1) <> 'S' order by route_no "
               Call Confirm_Recordset_Closed(tmp_Rs)
               tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
               strRoute_No = ""
               tmp_Rs.MoveFirst
               Do While Not tmp_Rs.EOF
    
                strRoute_No = strRoute_No & RTrim(tmp_Rs("route_no")) & "; "
                tmp_Rs.MoveNext
    
               Loop
               tmp_Rs.Close
           
           End If
           
           rs_Access.Fields("���s��").Value = strRoute_No & ""
           rs_Access.Update
           rs_Tab0_VLLSum.MoveNext
        Loop
        rs_Tab0_VLLDetailxSection.Close
        rs_Tab0_VLLSum.MoveFirst
        rs_Access.Close

'VLL�˸����Ӫ�
If chkDetail = vbChecked Then
    str_SQL = "Delete From VLL�˸����Ӫ�"
    cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
    Call ReDim_Recordset(rs_Access)
    rs_Access.Open "VLL�˸����Ӫ�", cnAccess, adOpenStatic, adLockOptimistic
    rs_Tab0_VLLDetail.MoveFirst
    Do While Not rs_Tab0_VLLDetail.EOF
       rs_Access.AddNew
       For iLoop = 0 To rs_Tab0_VLLDetail.Fields.Count - 1
           rs_Access.Fields(iLoop).Value = rs_Tab0_VLLDetail.Fields(iLoop).Value
       Next iLoop
       rs_Access.Update
       rs_Tab0_VLLDetail.MoveNext
    Loop
rs_Tab0_VLLDetail.MoveFirst
End If

If chkSUMDetail = vbChecked Then
    str_SQL = "Delete From VLL�˸����Ӫ�"
    cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
    Call ReDim_Recordset(rs_Access)
    rs_Access.Open "VLL�˸����Ӫ�", cnAccess, adOpenStatic, adLockOptimistic
    rs_Tab0_VLLSUMDetail.MoveFirst
    Do While Not rs_Tab0_VLLSUMDetail.EOF
       rs_Access.AddNew
       For iLoop = 0 To rs_Tab0_VLLSUMDetail.Fields.Count - 1
           rs_Access.Fields(iLoop).Value = rs_Tab0_VLLSUMDetail.Fields(iLoop).Value
       Next iLoop
       rs_Access.Update
       rs_Tab0_VLLSUMDetail.MoveNext
    Loop
rs_Tab0_VLLSUMDetail.MoveFirst
End If


'VLL�X�f��
Dim blnPrintVLLorder As Boolean
blnPrintVLLorder = True

str_SQL = "Select * From gv_Report_VLLOrder where �ըƹF�渹 in (select trp02t.receipt_no from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & "))"

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then '�L��ƮɵL���C�L
        blnPrintVLLorder = False
Else
        Call Replication_Recordset(tmp_Rs, rs_Tab0_VLLOrder)
        tmp_Rs.Close
        str_SQL = "Delete From VLL�X�f��"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL�X�f��", cnAccess, adOpenStatic, adLockOptimistic
        With rs_Tab0_VLLOrder
            .MoveFirst
            Do While Not .EOF
            
            If .Fields("�f�D") = "LPSI01" And chkTHLShipList = 0 Then GoTo NextRow
            If .Fields("�f�D") = "LKAO01" And chkKAOShipList = 0 Then GoTo NextRow
            If .Fields("�f�D") = "LABT01" And chkNSLShipList = 0 Then GoTo NextRow
            If .Fields("�f�D") = "LLFA01" And chkLFAShipList = 0 Then GoTo NextRow
            
'            If .Fields("�f�D") = "LNSL01" Then
'                If chkNSLShipList = 0 And Left(.Fields("�f�D�渹"), 1) = "8" Then GoTo NextRow
'            End If
                   rs_Access.AddNew
                   rs_Access.Fields("�s��").Value = .Fields("�s��").Value
                   rs_Access.Fields("�f�D�W��").Value = .Fields("�f�D�W��").Value
                   rs_Access.Fields("���u�s��").Value = .Fields("���u�s��").Value
                   rs_Access.Fields("�X�����").Value = .Fields("�X�����").Value
                   rs_Access.Fields("�ݨD���").Value = .Fields("�ݨD���").Value
                   rs_Access.Fields("�f�D�渹").Value = .Fields("�f�D�渹").Value
                   rs_Access.Fields("�ըƹF�渹").Value = .Fields("�ըƹF�渹").Value
                   rs_Access.Fields("���ʽs��").Value = .Fields("���ʽs��").Value & ""
                   rs_Access.Fields("�Ȥ�W��").Value = .Fields("�Ȥ�W��").Value
                   rs_Access.Fields("�Ȥ�a�}").Value = .Fields("�Ȥ�a�}").Value
                   rs_Access.Fields("�q��").Value = .Fields("�q��").Value
                   rs_Access.Fields("�Ȥ�ݨD").Value = .Fields("�Ȥ�ݨD").Value
                   rs_Access.Fields("�Ƶ�").Value = .Fields("�Ƶ�").Value
                   rs_Access.Fields("�r�p").Value = .Fields("�r�p").Value
                   rs_Access.Fields("����").Value = .Fields("����").Value
                   rs_Access.Fields("����").Value = .Fields("����").Value
                   rs_Access.Fields("�f��").Value = .Fields("�f��").Value
                   rs_Access.Fields("�~�W").Value = .Fields("�~�W").Value
                   rs_Access.Fields("�c��").Value = .Fields("�X�f�c��").Value
                   rs_Access.Fields("�j�]��").Value = .Fields("�j�]��").Value
                   rs_Access.Fields("�Ӽ�").Value = .Fields("�X�f�Ӽ�").Value
                   rs_Access.Fields("�p�]��").Value = .Fields("�p�]��").Value
                   rs_Access.Fields("�`�Ӽ�").Value = .Fields("�`�Ӽ�").Value
                   rs_Access.Fields("�ܧO").Value = .Fields("�ܧO").Value
                   rs_Access.Fields("�G���ƨ����s").Value = .Fields("�G���ƨ����s").Value
                   rs_Access.Fields("���").Value = .Fields("���").Value
                '   rs_Access.Fields("�s�y��").Value = .Fields("�s�y��").Value
                '   rs_Access.Fields("�����").Value = .Fields("�����").Value
                    rs_Access.Fields("USER").Value = User_Name
                   rs_Access.Update
              i = i + 1
               
NextRow:
               .MoveNext
            Loop
        
        .MoveFirst
        End With
End If

'LCHF�X�f�� add by Terry 20180724
Dim blnPrintLCHForder As Boolean
blnPrintLCHForder = True

str_SQL = "Select * From Xv_Report_VLLOrder_LCHF01 where �ըƹF�渹 in (select trp02t.receipt_no from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & "))"

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then '�L��ƮɵL���C�L
        blnPrintLCHForder = False
Else
        str_SQL = "Delete From VLL�X�f��_LCHF01"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL�X�f��_LCHF01", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            Do While Not .EOF
                   rs_Access.AddNew
                   'rs_Access.Fields("�s��").Value = .Fields("�s��").Value
                   rs_Access.Fields("�f�D�W��").Value = .Fields("�f�D�W��").Value
                   rs_Access.Fields("���u�s��").Value = .Fields("���u�s��").Value
                   rs_Access.Fields("�X�����").Value = .Fields("�X�����").Value
                   rs_Access.Fields("�ݨD���").Value = .Fields("�ݨD���").Value
                   rs_Access.Fields("�f�D�渹").Value = .Fields("�f�D�渹").Value
                   rs_Access.Fields("�ըƹF�渹").Value = .Fields("�ըƹF�渹").Value
                   rs_Access.Fields("���ʽs��").Value = .Fields("���ʽs��").Value & ""
                   rs_Access.Fields("�Ȥ�W��").Value = .Fields("�Ȥ�W��").Value
                   rs_Access.Fields("�Ȥ�a�}").Value = .Fields("�Ȥ�a�}").Value
                   rs_Access.Fields("�q��").Value = .Fields("�q��").Value
                   rs_Access.Fields("�Ȥ�ݨD").Value = .Fields("�Ȥ�ݨD").Value
                   rs_Access.Fields("�Ƶ�").Value = .Fields("�Ƶ�").Value
                   rs_Access.Fields("�r�p").Value = .Fields("�r�p").Value
                   rs_Access.Fields("����").Value = .Fields("����").Value
                   rs_Access.Fields("����").Value = .Fields("����").Value
                   rs_Access.Fields("�f��").Value = .Fields("�f��").Value
                   rs_Access.Fields("�~�W").Value = .Fields("�~�W").Value & " Exp." & .Fields("�����").Value
                   rs_Access.Fields("�c��").Value = .Fields("�X�f�c��").Value
                   rs_Access.Fields("�j�]��").Value = .Fields("�j�]��").Value
                   rs_Access.Fields("�Ӽ�").Value = .Fields("�X�f�Ӽ�").Value
                   rs_Access.Fields("�p�]��").Value = .Fields("�p�]��").Value
                   rs_Access.Fields("�`�Ӽ�").Value = .Fields("�`�Ӽ�").Value
                   rs_Access.Fields("�ܧO").Value = .Fields("�ܧO").Value
                   rs_Access.Fields("�G���ƨ����s").Value = .Fields("�G���ƨ����s").Value
                   rs_Access.Fields("���").Value = .Fields("���").Value
                   rs_Access.Fields("USER").Value = User_Name
                   rs_Access.Fields("�����Ƶ�").Value = .Fields("�����Ƶ�").Value
                   rs_Access.Update
              i = i + 1
               
               .MoveNext
            Loop
        
        .MoveFirst
        End With
End If

'LNVA�X�f�� add by Terry 20190422
Dim blnPrintLNVAorder As Boolean
blnPrintLNVAorder = True

'str_SQL = "Select * From Xv_Report_VLLOrder_LNVA01 where �ըƹF�渹 in (select trp02t.receipt_no from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & "))"
str_SQL = "delete from codelkup where listname = 'VLLReport_LNVA01'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "insert into CODELKUP (LISTNAME,Code,Description,AddDate,AddWho,EditDate,EditWho) select 'VLLReport_LNVA01',trp02t.receipt_no,'LNVA01�X�f��',GETDATE(),'" & User_Name & "',getdate(),'" & User_Name & "' from trp02t trp02t join trp01t trp01t on trp01t.route_no = trp02t.route_no where trp01t.c_route_no in (" & strSelectedRouteNo & ") or trp02t.route_no in (" & strSelectedRouteNo & ")"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "exec Xs_VLLReport_LNVA01"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then '�L��ƮɵL���C�L
        blnPrintLNVAorder = False
Else
        str_SQL = "Delete From VLL�X�f��_LNVA01"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VLL�X�f��_LNVA01", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            Do While Not .EOF
                   rs_Access.AddNew
                   'rs_Access.Fields("�s��").Value = .Fields("�s��").Value
                   rs_Access.Fields("�f�D�W��").Value = .Fields("�f�D�W��").Value
                   rs_Access.Fields("���u�s��").Value = .Fields("���u�s��").Value
                   rs_Access.Fields("�X�����").Value = .Fields("�X�����").Value
                   rs_Access.Fields("�ݨD���").Value = .Fields("�ݨD���").Value
                   rs_Access.Fields("�f�D�渹").Value = .Fields("�f�D�渹").Value
                   rs_Access.Fields("�ըƹF�渹").Value = .Fields("�ըƹF�渹").Value
                   rs_Access.Fields("���ʽs��").Value = .Fields("���ʽs��").Value & ""
                   rs_Access.Fields("�Ȥ�W��").Value = .Fields("�Ȥ�W��").Value
                   rs_Access.Fields("�Ȥ�a�}").Value = .Fields("�Ȥ�a�}").Value
                   rs_Access.Fields("�q��").Value = .Fields("�q��").Value
                   rs_Access.Fields("�Ȥ�ݨD").Value = .Fields("�Ȥ�ݨD").Value
                   rs_Access.Fields("�Ƶ�").Value = .Fields("�Ƶ�").Value
                   rs_Access.Fields("�r�p").Value = .Fields("�r�p").Value
                   rs_Access.Fields("����").Value = .Fields("����").Value
                   rs_Access.Fields("����").Value = .Fields("����").Value
                   rs_Access.Fields("�f��").Value = .Fields("�f��").Value
                   rs_Access.Fields("�~�W").Value = .Fields("�~�W").Value & "  " & .Fields("�帹").Value
                   rs_Access.Fields("�c��").Value = .Fields("�X�f�c��").Value
                   rs_Access.Fields("�j�]��").Value = .Fields("�j�]��").Value
                   rs_Access.Fields("�Ӽ�").Value = .Fields("�X�f�Ӽ�").Value
                   rs_Access.Fields("�p�]��").Value = .Fields("�p�]��").Value
                   rs_Access.Fields("�`�Ӽ�").Value = .Fields("�`�Ӽ�").Value
                   rs_Access.Fields("�ܧO").Value = .Fields("�ܧO").Value
                   rs_Access.Fields("�G���ƨ����s").Value = .Fields("�G���ƨ����s").Value
                   rs_Access.Fields("���").Value = .Fields("���").Value
                    rs_Access.Fields("USER").Value = User_Name
                   rs_Access.Update
              i = i + 1

               .MoveNext
            Loop

        .MoveFirst
        End With
End If

'VTL�X�f��
Dim blnPrintVTLorder As Boolean: blnPrintVTLorder = True
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from gv_Report_VTLOrder Where �G���ƨ����s IN (" & strSelectedRouteNo & ") or ���u�s�� in (" & strSelectedRouteNo & ") "
tmp_Rs.Open str_SQL, cn
If tmp_Rs.EOF Then '�L��ƮɵL���C�L

        blnPrintVTLorder = False

Else
        str_SQL = "Delete From VTL�X�f��"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "VTL�X�f��", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            
            Do While Not .EOF
            
                '�P�_���L�S���q�O
                If UCase(Left(.Fields("�ӹB�ӥN��"), 1)) = "P" Then
                strCompany = "����"
                ElseIf UCase(Left(.Fields("�X�f�渹�X"), 1)) = "V" Then strCompany = "���L�S"
                ElseIf UCase(Left(.Fields("�X�f�渹�X"), 1)) = "C" Then strCompany = "���y"
                ElseIf UCase(Left(.Fields("�X�f�渹�X"), 1)) = "E" Then strCompany = "���o"
                ElseIf UCase(Left(.Fields("���~�N��"), 1)) = "O" Then strCompany = "�u�A�K"
                Else
                strCompany = "���L�S"
                End If
            
               rs_Access.AddNew
               rs_Access.Fields("�X�f�渹�X").Value = .Fields("�X�f�渹�X").Value
               rs_Access.Fields("TMS�渹").Value = .Fields("TMS�渹").Value
               rs_Access.Fields("�ƥX���").Value = .Fields("�ƥX���").Value
               rs_Access.Fields("���u�s��").Value = .Fields("���u�s��").Value
               rs_Access.Fields("�G���ƨ����s").Value = .Fields("�G���ƨ����s").Value
               rs_Access.Fields("�b�ګȤ�N��").Value = .Fields("�b�ګȤ�N��").Value
               rs_Access.Fields("�b�ګȤ�").Value = .Fields("�b�ګȤ�").Value
               rs_Access.Fields("�e�f�Ȥ�N��").Value = .Fields("�e�f�Ȥ�N��").Value
               rs_Access.Fields("�e�f�Ȥ�").Value = .Fields("�e�f�Ȥ�").Value
               rs_Access.Fields("�̪O�ϥ�").Value = .Fields("�̪O�ϥ�").Value
               rs_Access.Fields("�e�f�a�}").Value = .Fields("�e�f�a�}").Value & ""
               rs_Access.Fields("�q��").Value = .Fields("�q��").Value
               rs_Access.Fields("�ӹB�ӥN��").Value = .Fields("�ӹB�ӥN��").Value
               rs_Access.Fields("�ӹB�ӦW��").Value = .Fields("�ӹB�ӦW��").Value
               rs_Access.Fields("����").Value = .Fields("����").Value
               rs_Access.Fields("����").Value = .Fields("����").Value
               rs_Access.Fields("����").Value = .Fields("����").Value
               rs_Access.Fields("��]").Value = .Fields("��]").Value
               rs_Access.Fields("���~�N��").Value = .Fields("���~�N��").Value
               rs_Access.Fields("���~�W��").Value = .Fields("���~�W��").Value
               rs_Access.Fields("����").Value = .Fields("����").Value
               rs_Access.Fields("����").Value = .Fields("����").Value
'               rs_Access.Fields("�c").Value = .Fields("�c").Value
'               rs_Access.Fields("��").Value = .Fields("��").Value
'               rs_Access.Fields("�`����").Value = .Fields("�`����").Value
               rs_Access.Fields("�Ƶ�").Value = .Fields("�Ƶ�").Value
               rs_Access.Fields("USER").Value = User_Name
               rs_Access.Fields("���q�O").Value = strCompany
               rs_Access.Update
               .MoveNext
            Loop
        
        .MoveFirst
        End With
End If

'YFY�X�f��
Dim blnPrintYFYorder As Boolean: blnPrintYFYorder = True
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "select * from ev_Report_YFYOrder Where ���u�s�� in (" & strSelectedRouteNo & ") or �G�����u�s�� in (" & strSelectedRouteNo & ")"

tmp_Rs.Open str_SQL, cn
If tmp_Rs.EOF Then '�L��ƮɵL���C�L

        blnPrintYFYorder = False

Else
        str_SQL = "Delete From YFY�X�f��"
        cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Call ReDim_Recordset(rs_Access)
        rs_Access.Open "YFY�X�f��", cnAccess, adOpenStatic, adLockOptimistic
        With tmp_Rs
            .MoveFirst
            Do While Not .EOF
               rs_Access.AddNew
               rs_Access.Fields("TMS�渹").Value = .Fields("TMS�渹").Value
               rs_Access.Fields("�q�渹�X").Value = .Fields("�q�渹�X").Value
               'rs_Access.Fields("�q��Ӷ�").Value = .Fields("�q��Ӷ�").Value
               rs_Access.Fields("�f�D�W��").Value = .Fields("�f�D�W��").Value
               rs_Access.Fields("�Ȥ�W��").Value = .Fields("�Ȥ�W��").Value
               rs_Access.Fields("�Ƶ�").Value = .Fields("�Ƶ�").Value
               rs_Access.Fields("�a�}").Value = .Fields("�a�}").Value
               rs_Access.Fields("�p���H").Value = .Fields("�p���H").Value
               rs_Access.Fields("�C�L���").Value = .Fields("�C�L���").Value
               rs_Access.Fields("���u�s��").Value = .Fields("���u�s��").Value
               'rs_Access.Fields("�G�����u�s��").Value = .Fields("�G�����u�s��").Value
               rs_Access.Fields("����").Value = .Fields("����").Value
               rs_Access.Fields("�~��").Value = .Fields("�~��").Value
               rs_Access.Fields("�~�W").Value = .Fields("�~�W").Value
               rs_Access.Fields("�ƶq").Value = .Fields("�ƶq").Value
               rs_Access.Fields("���n").Value = .Fields("���n").Value
               rs_Access.Fields("�Ȥ�q�渹�X").Value = .Fields("�Ȥ�q�渹�X").Value
               rs_Access.Fields("���ʳ渹").Value = .Fields("���ʳ渹").Value
               rs_Access.Fields("�X�f��").Value = .Fields("�X�f��").Value
               rs_Access.Update
               .MoveNext
            Loop

        .MoveFirst
        End With
End If

cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'5. call Access �C�L����
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab0_PreView.Value = vbChecked Then
   '�w���C�L
   MSAccessAP.Visible = True
   If chkVllPallet = vbChecked Then
      MSAccessAP.DoCmd.OpenReport "VLL�˸��`��_²��", acViewPreview
   Else
      MSAccessAP.DoCmd.OpenReport "VLL�˸��`��", acViewPreview
   End If
      
   MSAccessAP.DoCmd.OpenReport "��������", acViewPreview
   
   If chkSUMDetail.Value = vbChecked Then
    MSAccessAP.DoCmd.OpenReport "VLL�˸����Ӷ��`��", acViewPreview
   End If
   
   If chkDetail.Value = vbChecked Then
    MSAccessAP.DoCmd.OpenReport "VLL�˸����Ӫ�", acViewPreview
   End If
   
   If blnPrintVLLorder = True And i > 0 Then MSAccessAP.DoCmd.OpenReport "VLL�X�f��", acViewPreview
   If blnPrintLCHForder = True Then MSAccessAP.DoCmd.OpenReport "VLL�X�f��_LCHF01", acViewPreview
   If blnPrintLNVAorder = True Then MSAccessAP.DoCmd.OpenReport "VLL�X�f��_LNVA01", acViewPreview
   If blnPrintVTLorder = True Then MSAccessAP.DoCmd.OpenReport "VTL�X�f��", acViewPreview
   If blnPrintYFYorder = True Then MSAccessAP.DoCmd.OpenReport "YFY�X�f��", acViewPreview
   MSAccessAP.DoCmd.Maximize
Else
   '�����C�L�ܦL���
   MSAccessAP.Visible = False
   If chkVllPallet = vbChecked Then
        MSAccessAP.DoCmd.OpenReport "VLL�˸��`��_²��", acViewNormal
   Else
        MSAccessAP.DoCmd.OpenReport "VLL�˸��`��", acViewNormal
   End If
   MSAccessAP.DoCmd.OpenReport "��������", acViewPreview
   
   If chkSUMDetail.Value = vbChecked Then
        MSAccessAP.DoCmd.OpenReport "VLL�˸����Ӷ��`��", acViewNormal
   End If
   
   If chkDetail.Value = vbChecked Then
        MSAccessAP.DoCmd.OpenReport "VLL�˸����Ӫ�", acViewNormal
   End If
   
   If blnPrintVLLorder = True And i > 0 Then MSAccessAP.DoCmd.OpenReport "VLL�X�f��", acViewNormal
   If blnPrintLCHForder = True Then MSAccessAP.DoCmd.OpenReport "VLL�X�f��_LCHF01", acViewNormal
   If blnPrintLNVAorder = True Then MSAccessAP.DoCmd.OpenReport "VLL�X�f��_LNVA01", acViewNormal
   If blnPrintVTLorder = True Then MSAccessAP.DoCmd.OpenReport "VTL�X�f��", acViewNormal
   If blnPrintYFYorder = True Then MSAccessAP.DoCmd.OpenReport "YFY�X�f��", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If

cmd_Tab0_PrintReport.Enabled = True     'daniel
Screen.MousePointer = 0
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   
   Call Unload_RunLogForm
   cmd_Tab0_PrintReport.Enabled = True  'daniel
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL�W�f��-�C�L", Me.Caption, "cmd_Tab0_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab0_Query_Click()
On Error GoTo err_Handle

'�^�Ǵz�f�q
str_SQL = "select o.route " & _
        ",o.storerkey " & _
        ",o.orderkey " & _
        ",o.updatesource " & _
        ",o.Externorderkey " & _
        ",ExternLineno = case when o.storerkey = 'LLFA01' and o.IncoTerm <> '' then od.orderlinenumber else od.ExternLineno end " & _
        ",od.sku " & _
        ",shippedqty = (od.shippedqty + od.qtyallocated + od.qtypicked) " & _
        ",od.editdate " & _
        ",o.status " & _
        "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey and o.yfystatus = '0' and convert(char(8),o.deliverydate,112) > convert(char(8),getdate()-7,112) " & _
        "where (od.shippedqty + od.qtyallocated + od.qtypicked) > 0 " & _
        "and len(rtrim(isnull(o.updatesource,''))) > 9 and o.updatesource in (select t2.receipt_no from trp02t t2 where t2.receipt_no = o.updatesource and t2.exe_confirm = 2) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'�L���
If Not tmp_Rs.EOF Then

    tmp_Rs.MoveFirst
    Tran_Level = cn.BeginTrans
    Do While Not tmp_Rs.EOF
    
            str_SQL = "UPDATE TRP03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03W set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�g�J����
'            Call WriteLog(Err.Number & Chr(9) & "�z�f�ƶq�T�{" & Chr(9) & "WMS: " & tmp_Rs("orderkey") & ",TMS: " & tmp_Rs("route") & "," & tmp_Rs("storerkey") & "," & tmp_Rs("updatesource") & "," & RTrim(tmp_Rs("Externorderkey")) & "," & tmp_Rs("Externlineno") & "," & tmp_Rs("sku") & "," & tmp_Rs("shippedqty") & "," & User_id)
            
            '��sYFYstatus�^�Ǫ��A
            If Trim(tmp_Rs("status")) = "9" And Trim(tmp_Rs("storerkey")) <> "LTKK01" Then
                str_SQL = "UPDATE " & strWMSDB & "..Orders set YFYstatus = '1' ,TrafficCop = null where orderkey = '" & tmp_Rs("orderkey") & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
        
        tmp_Rs.MoveNext
    Loop
    
    '�����X�f�q=�q��q mark by Gemini @20150805 4 SHIP_QTY ���TRP02T Trigger �g�J
'            str_SQL = "UPDATE TRP03T set TRP03T.SHIP_QTY=TRP03T.order_qty from trp02t join trp03t on trp02t.receipt_no = trp03t.receipt_no where trp02t.priority = 'C' and ship_qty = 0 "
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
    cn.CommitTrans: Tran_Level = 0
End If

tmp_Rs.Close

'��LF���t�m�q add by Eric
Call LLFA01Ship2TMS

'VLL�W�f�� >> �d��
Set dg_Tab0_VLL.DataSource = Nothing
Set rs_Tab0_VLL = Nothing
blVLLReportEventEnable = False  '�Ը�

Screen.MousePointer = vbHourglass

'str_SQL = "Select ' ' as '��',���u�s�� ,�X�����,�C�L����,�C�L�ɶ�,���P���X,����,�r�p�H,�B�餽�q,�ƨ��c��,�z�f�c�� " & _
'          "From Report_VLL_RouteList "
str_SQL = "Select ' ' as '��',t01t.Route_No as ���u�s�� , Convert(varchar(8),t01t.Delivery_Date,112) as �X����� , Isnull(t01t.VLListCount,0) as �C�L���� , " & _
        "Isnull((Convert(varchar,Isnull(t01t.VLListPrintDate,Getdate()),111) + ' ' + Convert(varchar,Isnull(t01t.VLListPrintDate,Getdate()) , 108)),'') as �C�L�ɶ� , " & _
        "Rtrim(t05t.Vehicle_ID_No) as ���P���X , t05t.Drive_Times as ���� , Rtrim(t05t.Driver) as �r�p�H , " & _
        "Rtrim(Isnull(t08m.C_Name,'')) as �B�餽�q, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(order_qty),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(order_qty),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as �ƨ��Ӽ�, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(SHIP_QTY),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(SHIP_QTY),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as �z�f�Ӽ�, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(WEIGHT),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(WEIGHT),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as �ƨ����q, " & _
        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(VOLUMN_WEIGHT),0),0) from trp03t where Route_No=t01t.Route_No ) else (select isnull(round(sum(VOLUMN_WEIGHT),0),0) from trp03t_S where Route_No_S=t01t.Route_No ) end as �ƨ����n " & _
        "From TRP01T t01t " & _
        "inner join TRP05T t05t on t05t.Route_No = t01t.Route_No and convert(char(8),t01t.delivery_date,112) > convert(char(8),getdate()-7,112) " & _
        "inner join TRP09M t09m on t09m.Vehicle_ID_No = t05t.Vehicle_ID_No " & _
        "Left outer join TRP08M t08m on t08m.Company_Code = t09m.TRP_Company_Code "
        
'�z�f���n���q
'        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(t3.ship_qty * sp.stdgrosswgt),0),0) from trp03t t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where t3.Route_No=t01t.Route_No ) else (select isnull(round(sum(t3.ship_qty * sp.stdgrosswgt),0),0) from trp03t_S t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where Route_No_S=t01t.Route_No ) end as �ƨ����q, " & _
'        "case when left(t01t.Route_No,1)='F' then  (select isnull(round(sum(t3.ship_qty * sp.stdcube),0),0) from trp03t t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where t3.Route_No=t01t.Route_No ) else (select isnull(round(sum(t3.ship_qty * sp.stdcube),0),0) from trp03t_S t3 join gv_skuxpack sp on sp.sku = t3.product_no and t3.storerkey = sp.storerkey where Route_No_S=t01t.Route_No ) end as �ƨ����n " & _

Dim strWhere As String, strTmp As String
strWhere = "Where t01t.Route_No <> 'D'"

'�X�����
strTmp = ""
If Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),t01t.Delivery_Date,112) between '" & txt_Tab0_DeliveryDate_Start.Text & "' and '" & txt_Tab0_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(varchar(8),t01t.Delivery_Date,112) = '" & txt_Tab0_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) = 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),t01t.Delivery_Date,112) = '" & txt_Tab0_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'���u�s��
strTmp = ""
If Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   strTmp = " t01t.Route_No between '" & txt_Tab0_RouteNo_Start.Text & "' and '" & txt_Tab0_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) = 0 Then
   strTmp = " t01t.Route_No = '" & txt_Tab0_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) = 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   strTmp = " t01t.Route_No = '" & txt_Tab0_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'�t�w�C�L�L�� Wave
strTmp = ""
If chk_Tab0_PrintedRoute.Value = vbUnchecked Then
   strTmp = " Isnull(t01t.VLListCount,0) = 0 "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp & " and t01t.C_ROUTE_NO is null"
   Else
      strWhere = strWhere & strTmp & " and t01t.C_ROUTE_NO is null"
   End If
End If

'�u��ܤjroute
If Len(strWhere) > 0 Then
   strWhere = strWhere & " and t01t.C_ROUTE_NO is null"
Else
   strWhere = "t01t.C_ROUTE_NO is null"
End If
If strWhere <> "" Then
   str_SQL = str_SQL & strWhere
Else
   msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by ���u�s�� "
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
cn.CommandTimeout = 300

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   blVLLReportEventEnable = True
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0_VLL)
tmp_Rs.Close

With dg_Tab0_VLL
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab0_VLL.MoveFirst
Set dg_Tab0_VLL.DataSource = rs_Tab0_VLL

SetDataGridColWidth "VLL�˸�", dg_Tab0_VLL

With dg_Tab0_VLL
    .ColumnHeaders = True         '���D�����
    .RowHeight = 300

End With
rs_Tab0_VLL.MoveFirst
rs_Tab0_VLL.Filter = adFilterNone
rs_Tab0_VLL.Sort = " �s�� "
rs_Tab0_VLL.MoveFirst
blVLLReportEventEnable = True
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL�W�f��-�d��", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Reset_Click()
'VLL�W�f�� >> �M��
txt_Tab0_DeliveryDate_Start.Text = "": txt_Tab0_DeliveryDate_End.Text = ""
txt_Tab0_RouteNo_Start.Text = "": txt_Tab0_RouteNo_End.Text = ""
chk_Tab0_PrintedRoute.Value = vbUnchecked
Set dg_Tab0_VLL.DataSource = Nothing
Set rs_Tab0_VLL = Nothing
End Sub

Private Sub cmd_Tab0_SaveToExcel_Click()
'VLL�W�f�� >> �� EXCEL
blVLLReportEventEnable = False

Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel �ɮצW��
CmnDialog.DialogTitle = "��s Excel ��"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "VLL�W�f��_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel�ɮ�(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '�� [�}������] ��ܤ�����A���U [����] �s
   msg_text = "��� [����] ���s�A������ Excel ���ۦ�s��"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab0_VLL) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab0_VLL.MoveFirst
Exit Sub

err_Handle:
   blVLLReportEventEnable = True
   Dim tmpString As String
   Screen.MousePointer = vbDefault
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL�W�f��-�� EXCEL", Me.Caption, "cmd_Tab0_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab1_PrintReport_Click()
'�����˸����`�� >> ����C�L
If rs_Tab1_VLLSum Is Nothing Then Exit Sub
If rs_Tab1_VLLSum.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. ��Ƽg�X Access ��Ʈw >> �����˸����`��
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From �����˸����`��"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "�����˸����`��", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab1_VLLSum.MoveFirst
Do While Not rs_Tab1_VLLSum.EOF
   rs_Access.AddNew
   rs_Access.Fields("SerialNo").Value = rs_Tab1_VLLSum.Fields("�s��").Value
   rs_Access.Fields("���u�s��").Value = rs_Tab1_VLLSum.Fields("���u�s��").Value
   rs_Access.Fields("����").Value = rs_Tab1_VLLSum.Fields("���P���X").Value
   rs_Access.Fields("����").Value = rs_Tab1_VLLSum.Fields("����").Value
   rs_Access.Fields("�X�����").Value = rs_Tab1_VLLSum.Fields("�X�����").Value
   rs_Access.Fields("�B�餽�q²��").Value = rs_Tab1_VLLSum.Fields("�B�餽�q²��").Value
   rs_Access.Fields("�S��ݨD1").Value = rs_Tab1_VLLSum.Fields("�S��ݨD2").Value
   rs_Access.Fields("�S��ݨD2").Value = rs_Tab1_VLLSum.Fields("�S��ݨD1").Value
   rs_Access.Fields("�f�D").Value = rs_Tab1_VLLSum.Fields("�f�D").Value
   rs_Access.Fields("�Ȥ�q��s��").Value = rs_Tab1_VLLSum.Fields("�f�D�渹").Value
   rs_Access.Fields("�Ȥ�s��").Value = rs_Tab1_VLLSum.Fields("�Ȥ�s��").Value
   rs_Access.Fields("�Ȥ�W��").Value = rs_Tab1_VLLSum.Fields("�Ȥ�W��").Value
   rs_Access.Fields("�l���ϸ�").Value = rs_Tab1_VLLSum.Fields("zip").Value
   rs_Access.Fields("�e�f�a�}").Value = rs_Tab1_VLLSum.Fields("�e�f�a�}").Value
   rs_Access.Fields("�c��").Value = rs_Tab1_VLLSum.Fields("�c��").Value
   rs_Access.Fields("�Ӽ�").Value = rs_Tab1_VLLSum.Fields("�Ӽ�").Value
   rs_Access.Fields("���n").Value = rs_Tab1_VLLSum.Fields("���n").Value
   rs_Access.Fields("���q").Value = rs_Tab1_VLLSum.Fields("���q").Value
   rs_Access.Fields("�O��").Value = rs_Tab1_VLLSum.Fields("�O��").Value
   rs_Access.Fields("�G���ƨ����s").Value = rs_Tab1_VLLSum.Fields("�G���ƨ����s").Value
   rs_Access.Fields("�X�Y�Ȧs").Value = rs_Tab1_VLLSum.Fields("�X�Y�Ȧs").Value
   rs_Access.Fields("�w�p�������ɶ�").Value = rs_Tab1_VLLSum.Fields("�w�p�������ɶ�").Value
   rs_Access.Fields("�q������").Value = rs_Tab1_VLLSum.Fields("�q������").Value
   rs_Access.Fields("�Ȥ�ݨD").Value = rs_Tab1_VLLSum.Fields("�Ȥ�ݨD").Value
   rs_Access.Fields("�q��Ƶ�").Value = rs_Tab1_VLLSum.Fields("�q��Ƶ�").Value
   rs_Access.Update
   rs_Tab1_VLLSum.MoveNext
Loop
rs_Tab1_VLLSum.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access �C�L����
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab1_PreView.Value = vbChecked Then
   '�w���C�L
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "�����˸����`��", acViewPreview
Else
   '�����C�L�ܦL���
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "�����˸����`��", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Call Unload_RunLogForm
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--�����˸����`��C�L", Me.Caption, "cmd_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab1_Query_Click()
'�����˸����`�� >> �d��
Set dg_Tab1_VLLSum.DataSource = Nothing
Set rs_Tab1_VLLSum = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select ���P���X,���u�s��,����,�X�����,�B�餽�q²��,�S��ݨD1,�S��ݨD2,�q��s��,�f�D,�f�D�渹," & _
          "   �Ȥ�s��,�Ȥ�W��,ZIP,�e�f�a�},�c��,�Ӽ�,�O��,���n,���q,�B�e�ϰ�,�G���ƨ����s,�X�Y�Ȧs,�w�p�������ɶ�,�q������,�Ȥ�ݨD,�q��Ƶ�  " & _
          "From Report_LoadingSummary "

Dim strWhere As String, strTmp As String
strWhere = ""
'�q����
strTmp = ""
If Len(txt_Tab1_DeliveryDate_Start.Text) > 0 And Len(txt_Tab1_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� between '" & txt_Tab1_DeliveryDate_Start.Text & "' and '" & txt_Tab1_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab1_DeliveryDate_Start.Text) > 0 And Len(txt_Tab1_DeliveryDate_End.Text) = 0 Then
   strTmp = " �X����� = '" & txt_Tab1_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab1_DeliveryDate_Start.Text) = 0 And Len(txt_Tab1_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� = '" & txt_Tab1_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'���u�s��
strTmp = ""
If Len(txt_Tab1_route_Start.Text) > 0 And Len(txt_Tab1_route_End.Text) > 0 Then
   strTmp = " ���u�s�� between '" & txt_Tab1_route_Start.Text & "' and '" & txt_Tab1_route_End.Text & "' "
ElseIf Len(txt_Tab1_route_Start.Text) > 0 And Len(txt_Tab1_route_End.Text) = 0 Then
   strTmp = " ���u�s�� = '" & txt_Tab1_route_Start.Text & "' "
ElseIf Len(txt_Tab1_route_Start.Text) = 0 And Len(txt_Tab1_route_End.Text) > 0 Then
   strTmp = " ���u�s�� = '" & txt_Tab1_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'�B�e�ϰ�
strTmp = ""
If cmb_Tab1_AreaCode.ListIndex <> -1 Then
   strTmp = " �B�e�ϰ�N�X = '" & arAreaCode(cmb_Tab1_AreaCode.ListIndex) & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by ���P���X,���u�s��,����,�X����� "
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab1_VLLSum)
tmp_Rs.Close

With dg_Tab1_VLLSum
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab1_VLLSum.MoveFirst
Set dg_Tab1_VLLSum.DataSource = rs_Tab1_VLLSum

With dg_Tab1_VLLSum
    .ColumnHeaders = True          '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '���P���X
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1100       '���u�s��
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 500        '����
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 900        '�X�����
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1200       '�B�餽�q²��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1500       '�S��ݨD1
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1500       '�S��ݨD2
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1100       '�q��s��
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 500        '�f�D
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 900       '�f�D�渹
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1100      '�Ȥ�s��
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 2500      '�Ȥ�W��
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 500       'ZIP
    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 2400      '�e�f�a�}
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 800       '�c��
    .Columns(15).Alignment = dbgRight
    .Columns(16).Width = 800       '�O��
    .Columns(16).Alignment = dbgRight
    .Columns(17).Width = 800       '���n
    .Columns(17).Alignment = dbgRight
    .Columns(18).Width = 800       '���q
    .Columns(18).Alignment = dbgRight
    .Columns(19).Width = 3400      '�B�e�ϰ�
    .Columns(19).Alignment = dbgLeft
    .Columns(20).Width = 1300      '�G���ƨ����s
    .Columns(20).Alignment = dbgLeft
    .Columns(21).Width = 1300      '�X�Y�Ȧs
    .Columns(21).Alignment = dbgLeft
    .Columns(22).Width = 1300      '�w�p�������ɶ�
    .Columns(22).Alignment = dbgLeft
End With
rs_Tab1_VLLSum.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�����˸����`��-�d��", Me.Caption, "cmd_Tab1_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Reset_Click()
'�����˸����`�� >> �M��
cmb_Tab1_AreaCode.ListIndex = -1
txt_Tab1_DeliveryDate_Start.Text = ""
txt_Tab1_DeliveryDate_End.Text = ""
txt_Tab1_route_Start.Text = ""
txt_Tab1_route_End.Text = ""

Set dg_Tab1_VLLSum.DataSource = Nothing
Set rs_Tab1_VLLSum = Nothing
End Sub

Private Sub cmd_Tab1_SaveToExcel_Click()
'�����˸����`�� >> �� EXCEL
Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel �ɮצW��
CmnDialog.DialogTitle = "��s Excel ��"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "�����˸����`��_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel�ɮ�(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '�� [�}������] ��ܤ�����A���U [����] �s
   msg_text = "��� [����] ���s�A������ Excel ���ۦ�s��"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab1_VLLSum) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab1_VLLSum.MoveFirst
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�����˸����`��-�� EXCEL", Me.Caption, "cmd_Tab1_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_PrintReport_Click()
'�q���`�� >> ����C�L
If rs_Tab2_OrdersSum Is Nothing Then Exit Sub
If rs_Tab2_OrdersSum.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. ��Ƽg�X Access ��Ʈw >> �q���`��
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From �q���`��"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords

Call ReDim_Recordset(rs_Access)
rs_Access.Open "�q���`��", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab2_OrdersSum.MoveFirst

Do While Not rs_Tab2_OrdersSum.EOF

'    '�ˬd�Y�{�O�_��ƽT�{
'    If rs_Tab2_OrdersSum.Fields("�f�D�s��").Value = "LTHL01" And Len(RTrim(rs_Tab2_OrdersSum.Fields("��ƽT�{�ɶ�").Value)) = 0 Then
'        MsgBox "�Y�{�|�����T�{���!", 64, "�q���`��C�L����"
'        cnAccess.RollbackTrans
'        rs_Access.Close
'        Set rs_Access = Nothing
'        Exit Sub
'    End If
    
   rs_Access.AddNew
   rs_Access.Fields("�Ǹ�").Value = rs_Tab2_OrdersSum.Fields("�s��").Value
   rs_Access.Fields("���u�s��").Value = rs_Tab2_OrdersSum.Fields("���u�s��").Value
   rs_Access.Fields("�X�f���").Value = rs_Tab2_OrdersSum.Fields("��f��").Value
   rs_Access.Fields("����").Value = rs_Tab2_OrdersSum.Fields("���P���X").Value
   rs_Access.Fields("�q��").Value = rs_Tab2_OrdersSum.Fields("�r�p�H").Value
   rs_Access.Fields("����").Value = rs_Tab2_OrdersSum.Fields("����").Value
   rs_Access.Fields("�X�����").Value = rs_Tab2_OrdersSum.Fields("�X�����").Value
   rs_Access.Fields("�f�B��").Value = rs_Tab2_OrdersSum.Fields("�B�餽�q").Value
   rs_Access.Fields("�f�D�渹").Value = rs_Tab2_OrdersSum.Fields("�f�D�渹").Value
   rs_Access.Fields("�Ȥ�s��").Value = rs_Tab2_OrdersSum.Fields("�Ȥ�s��").Value
   rs_Access.Fields("�Ȥ�W��").Value = rs_Tab2_OrdersSum.Fields("�Ȥ�W��").Value
   rs_Access.Fields("�l���ϸ�").Value = rs_Tab2_OrdersSum.Fields("zip").Value
   rs_Access.Fields("�e�f�a�}").Value = rs_Tab2_OrdersSum.Fields("�e�f�a�}").Value
   rs_Access.Fields("�e�f�Ƶ�").Value = rs_Tab2_OrdersSum.Fields("�q��Ƶ�").Value
   rs_Access.Fields("���w�������O").Value = rs_Tab2_OrdersSum.Fields("���O").Value
   rs_Access.Fields("�f�DPO").Value = rs_Tab2_OrdersSum.Fields("�f�DPO").Value
   rs_Access.Fields("�q������").Value = rs_Tab2_OrdersSum.Fields("�q������").Value
   rs_Access.Fields("�N��").Value = rs_Tab2_OrdersSum.Fields("�N��").Value
   rs_Access.Fields("�Ȧs��").Value = rs_Tab2_OrdersSum.Fields("�Ȧs��").Value
   rs_Access.Fields("�Ȥ�ݨD").Value = rs_Tab2_OrdersSum.Fields("�Ȥ�ݨD").Value
   rs_Access.Fields("�c��").Value = rs_Tab2_OrdersSum.Fields("�c��").Value
   rs_Access.Fields("�Ӽ�").Value = rs_Tab2_OrdersSum.Fields("�Ӽ�").Value
   rs_Access.Update
   rs_Tab2_OrdersSum.MoveNext
Loop
rs_Tab2_OrdersSum.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access �C�L����
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab2_PreView.Value = vbChecked Then
   '�w���C�L
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "�q���`��", acViewPreview
Else
   '�����C�L�ܦL���
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "�q���`��", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cnAccess.RollbackTrans
      Tran_Level = 0
   End If
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q���`��-�C�L", Me.Caption, "cmd_Tab2_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab2_Query_Click()
'�q���`�� >> �d��
Set dg_Tab2_OrdersSum.DataSource = Nothing
Set rs_Tab2_OrdersSum = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select �X�����,��f��,���u�s��,���P���X,����,�r�p�H,�B�餽�q,�f�D�s��,�f�D�渹,�Ȥ�s��,�Ȥ�W��,�e�f�a�},�q��Ƶ�,�q���,'          ' as ���O,ZIP,�f�DPO,�q������, �N��, �Ȧs��, �Ȥ�ݨD, �c��, �Ӽ� ,��ƽT�{�ɶ� " & _
          "From Report_OrdersSum "

Dim strWhere As String, strTmp As String
strWhere = ""
'�X�����
strTmp = ""
If Len(txt_Tab2_DeliveryDate_Start.Text) > 0 And Len(txt_Tab2_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� between '" & txt_Tab2_DeliveryDate_Start.Text & "' and '" & txt_Tab2_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab2_DeliveryDate_Start.Text) > 0 And Len(txt_Tab2_DeliveryDate_End.Text) = 0 Then
   strTmp = " �X����� = '" & txt_Tab2_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab2_DeliveryDate_Start.Text) = 0 And Len(txt_Tab2_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� = '" & txt_Tab2_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'���u�s��
strTmp = ""
If Len(txt_Tab2_RouteNo_Start.Text) > 0 And Len(txt_Tab2_RouteNo_End.Text) > 0 Then
   strTmp = " ���u�s�� between '" & txt_Tab2_RouteNo_Start.Text & "' and '" & txt_Tab2_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab2_RouteNo_Start.Text) > 0 And Len(txt_Tab2_RouteNo_End.Text) = 0 Then
   strTmp = " ���u�s�� = '" & txt_Tab2_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab2_RouteNo_Start.Text) = 0 And Len(txt_Tab2_RouteNo_End.Text) > 0 Then
   strTmp = " ���u�s�� = '" & txt_Tab2_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by ���u�s��,�f�D�渹 "
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
Call Replication_Recordset(tmp_Rs, rs_Tab2_OrdersSum)
tmp_Rs.Close

With dg_Tab0_VLL
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab2_OrdersSum.MoveFirst
Set dg_Tab2_OrdersSum.DataSource = rs_Tab2_OrdersSum

With dg_Tab2_OrdersSum
    .ColumnHeaders = True         '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500       '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000      '�X�����
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000      '��f��
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1100      '���u�s��
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 900       '���P���X
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 500       '����
    .Columns(5).Alignment = dbgCenter
    .Columns(6).Width = 1000      '�r�p�H
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1500      '�B�餽�q
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 900       '�f�D�渹
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 1200      '�Ȥ�s��
    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 2000     '�Ȥ�W��
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 2000     '�e�f�a�}
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1200     '�Ȥ�Ƶ�
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 900      '�q���
    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 1400      '���w�������O
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 500      'ZIP
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1200      '�f�DPO
    .Columns(16).Alignment = dbgLeft
End With
'��s [���O] ���  ==> �q�榳 [���w�����] �X�f�A�g�J [***]
rs_Tab2_OrdersSum.MoveFirst
Do While Not rs_Tab2_OrdersSum.EOF
   str_SQL = "Select Count(LotTable05) as 'CNT' From OrderDetail Where ExternOrderKey = '" & rs_Tab2_OrdersSum.Fields("�f�D�渹").Value & "' and LotTable05 is not null "
   tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
   If tmp_Rs.Fields("CNT").Value = 0 Then
      rs_Tab2_OrdersSum.Fields("���O").Value = ""
   Else
      rs_Tab2_OrdersSum.Fields("���O").Value = "���w�����"
   End If
   tmp_Rs.Close
   rs_Tab2_OrdersSum.MoveNext
Loop
rs_Tab2_OrdersSum.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q���`��-�d��", Me.Caption, "cmd_Tab2_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Reset_Click()
'�q���`�� >> �M��
txt_Tab2_DeliveryDate_Start.Text = "": txt_Tab2_DeliveryDate_End.Text = ""
txt_Tab2_RouteNo_Start.Text = "": txt_Tab2_RouteNo_End.Text = ""
Set dg_Tab2_OrdersSum.DataSource = Nothing
Set rs_Tab2_OrdersSum = Nothing

End Sub

Private Sub cmd_Tab3_Query_Click()
'�z�f�˸��]�֪� >> �d��
Set dg_Tab3_PickLoadCheck.DataSource = Nothing
Set rs_Tab3_PickLoadCheck = Nothing

txt_Tab3_UploadDate_Start.Text = Trim(txt_Tab3_UploadDate_Start.Text)
txt_Tab3_UploadHour_Start.Text = Format(Val(txt_Tab3_UploadHour_Start.Text), "00")
txt_Tab3_UploadMinute_Start.Text = Format(Val(txt_Tab3_UploadMinute_Start.Text), "00")
txt_Tab3_UploadDate_End.Text = Trim(txt_Tab3_UploadDate_End.Text)
txt_Tab3_UploadHour_End.Text = Format(Val(txt_Tab3_UploadHour_End.Text), "00")
txt_Tab3_UploadMinute_End.Text = Format(Val(txt_Tab3_UploadMinute_End.Text), "00")

'If Len(txt_Tab3_UploadDate_Start.Text) = 0 Or Len(txt_Tab3_UploadDate_End.Text) = 0 Then
'   msg_text = "������ҡG�п�J [�^�Ǥ��] "
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'ElseIf Val(txt_Tab3_UploadHour_Start.Text) = 24 Or Val(txt_Tab3_UploadHour_End.Text) = 24 Then
'   'Wave �إ߮ɶ��d�򤣱��� 24�G00
'   msg_text = "������ҡG[�^�Ǥ��] ��ƿ��~�A" & vbCrLf & "" & vbCrLf & _
'              "�i��������ƽd��Gyyyymmdd 00�G00 �� yyyymmdd 23�G59"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   If Val(txt_Tab3_UploadHour_Start.Text) = 24 Then
'      txt_Tab3_UploadHour_Start.SelStart = 0: txt_Tab3_UploadHour_Start.SelLength = Len(txt_Tab3_UploadHour_Start.Text)
'      txt_Tab3_UploadHour_Start.SetFocus
'   Else
'      txt_Tab3_UploadHour_End.SelStart = 0: txt_Tab3_UploadHour_End.SelLength = Len(txt_Tab3_UploadHour_End.Text)
'      txt_Tab3_UploadHour_End.SetFocus
'   End If
'   Exit Sub
'End If
'txt_Tab3_UploadMinute_Start.Text = Trim(txt_Tab3_UploadMinute_Start.Text)
'txt_Tab3_UploadMinute_End.Text = Trim(txt_Tab3_UploadMinute_End.Text)
'If Len(txt_Tab3_UploadMinute_Start.Text) = 0 Or Len(txt_Tab3_UploadMinute_Start.Text) = 0 Then
'   msg_text = "������ҡG�п�J [�^�Ǯɶ�] "
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'End If

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select �B�e�ϰ�,���u�s��,�q���,�e�f�I,�Ȥ�²��,�c��,�O��,���q,���n,�f�B���q,����,����,�^�Ǥ��,�X�����,�w�p����ɶ�,�X�Y " & _
          "From Report_PickLoadCheck "

Dim tmpString1 As String, tmpString2 As String
Dim strWhere As String, strTmp As String
strWhere = ""
'�B�e�ϰ�
strTmp = ""
If cmb_Tab3_AreaCode.ListIndex <> -1 Then
   strTmp = " �B�e�ϰ�X = '" & arAreaCode(cmb_Tab3_AreaCode.ListIndex) & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'���w [�^�Ǥ��]
If txt_Tab3_UploadDate_Start.Text <> "" And txt_Tab3_UploadDate_End.Text <> "" Then
   strTmp = ""
   tmpString1 = Mid(txt_Tab3_UploadDate_Start.Text, 1, 4) & "-" & Mid(txt_Tab3_UploadDate_Start.Text, 5, 2) & "-" & Mid(txt_Tab3_UploadDate_Start.Text, 7, 2) & " " & _
                txt_Tab3_UploadHour_Start.Text & ":" & txt_Tab3_UploadMinute_Start.Text & ":00"
   tmpString2 = Mid(txt_Tab3_UploadDate_End.Text, 1, 4) & "-" & Mid(txt_Tab3_UploadDate_End.Text, 5, 2) & "-" & Mid(txt_Tab3_UploadDate_End.Text, 7, 2) & " " & _
                txt_Tab3_UploadHour_End.Text & ":" & txt_Tab3_UploadMinute_End.Text & ":00"
   strTmp = "UploadDate between convert(datetime,'" & tmpString1 & "',120) and convert(datetime,'" & tmpString2 & "',120) "
   If Len(strTmp) > 0 Then
      If Len(strWhere) = 0 Then
         strWhere = strTmp
      Else
         strWhere = strWhere & " and " & strTmp
      End If
   End If
End If
'���u�s��
strTmp = ""
If Len(txt_Tab3_route_Start.Text) > 0 And Len(txt_Tab3_route_End.Text) > 0 Then
   strTmp = " ���u�s�� between '" & txt_Tab3_route_Start.Text & "' and '" & txt_Tab3_route_End.Text & "' "
ElseIf Len(txt_Tab3_route_Start.Text) > 0 And Len(txt_Tab3_route_End.Text) = 0 Then
   strTmp = " ���u�s�� = '" & txt_Tab3_route_Start.Text & "' "
ElseIf Len(txt_Tab3_route_Start.Text) = 0 And Len(txt_Tab3_route_End.Text) > 0 Then
   strTmp = " ���u�s�� = '" & txt_Tab3_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
'�X�����
strTmp = ""
If Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� between '" & txt_Tab3_DeliveryDate_Start.Text & "' and '" & txt_Tab3_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) = 0 Then
   strTmp = " �X����� = '" & txt_Tab3_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) = 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� = '" & txt_Tab3_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If



If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " order by �B�e�ϰ�,���u�s�� "
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
Call Replication_Recordset(tmp_Rs, rs_Tab3_PickLoadCheck)
tmp_Rs.Close

With dg_Tab3_PickLoadCheck
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab3_PickLoadCheck.MoveFirst
Set dg_Tab3_PickLoadCheck.DataSource = rs_Tab3_PickLoadCheck

With dg_Tab3_PickLoadCheck
    .ColumnHeaders = True         '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500       '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 2500      '�B�e�ϰ�
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000      '���u�s��
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 700      '�q���
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 700       '�e�f�I
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1000       '�Ȥ�²��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 700      '�c��
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 700      '�O��
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700       '���q
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700      '���n
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 1000     '�f�B���q
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1000     '����
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 500     '����
    .Columns(12).Alignment = dbgCenter
    .Columns(13).Width = 850     '�^�Ǥ��
    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 850     '�X�����
    .Columns(14).Alignment = dbgCenter
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�z�f�˸��]��-�d��", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_SaveToExcel_Click()
'�z�f�˸��]�֪� >> �� EXCEL
Dim ExcelTitle As String
Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           'Excel �ɮצW��
CmnDialog.DialogTitle = "��s Excel ��"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "�z�f�˸��]�֪�_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "Excel�ɮ�(*.xls)|*.xls"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '�� [�}������] ��ܤ�����A���U [����] �s
   msg_text = "��� [����] ���s�A������ Excel ���ۦ�s��"
   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
   strTranFileName = ""
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
If SaveTo_ExcelFile(strTranFileName, rs_Tab3_PickLoadCheck) = 1 Then
   Screen.MousePointer = vbDefault
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
Else
   Screen.MousePointer = vbDefault
   If Len(strTranFileName) > 0 Then
      msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
rs_Tab3_PickLoadCheck.MoveFirst
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�z�f�˸��]�֪�-�� EXCEL", Me.Caption, "cmd_Tab3_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_PrintReport_Click()
'��B�����u���`�� >> �C�L��� >> ����C�L
If rs_Tab4_OrderDetail Is Nothing Then Exit Sub
If rs_Tab4_OrderDetail.RecordCount = 0 Then Exit Sub
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass

'1. ��Ƽg�X Access ��Ʈw
Call AccessDB_Connect

' Wave ���q������X
str_SQL = "Delete From ��B�����u���`��_OrderList"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access1)

Dim i As Double
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
rs_Access1.Open "��B�����u���`��_OrderList", cnAccess, adOpenStatic, adLockOptimistic
With dg_Tab4_RouteList


   For i = 1 To .Rows - 2
       .Row = i
       .Col = 1   '�C�L���O
       If Len(Trim(.Text)) > 0 Then
          rs_Access1.AddNew
          .Col = 3
          rs_Access1.Fields("�G���ƨ����u�s��").Value = txt_Tab4_SecondRouteNo.Text
          .Col = 2
          rs_Access1.Fields("�����ƨ����u�s��").Value = Trim(.Text)
          .Col = 7
          rs_Access1.Fields("�f�D�渹").Value = Trim(.Text)
          .Col = 8
          rs_Access1.Fields("�q����").Value = Trim(.Text)
          .Col = 9
          rs_Access1.Fields("�e�f���").Value = Trim(.Text)
          .Col = 10
          rs_Access1.Fields("�Ȥ�s��").Value = Trim(.Text)
          .Col = 11
          rs_Access1.Fields("�Ȥ�W��").Value = Trim(.Text)
          .Col = 5
          rs_Access1.Fields("���P���X").Value = Trim(.Text)
          .Col = 6
          rs_Access1.Fields("����").Value = Trim(.Text)
          rs_Access1.Update
       End If
   Next i
End With
cnAccess.CommitTrans

'�G���ƨ����u�����u���`�����X
str_SQL = "Delete From ��B�����u���`��_RouteList"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access2)
cnAccess.BeginTrans
rs_Access2.Open "��B�����u���`��_RouteList", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab4_OrderDetail.MoveFirst
Do While Not rs_Tab4_OrderDetail.EOF
   rs_Access2.AddNew
   rs_Access2.Fields("�Ǹ�").Value = rs_Tab4_OrderDetail.Fields("�s��").Value
   rs_Access2.Fields("�G���ƨ����s").Value = rs_Tab4_OrderDetail.Fields("�G�����s").Value
   rs_Access2.Fields("�f��").Value = rs_Tab4_OrderDetail.Fields("�f��").Value
   rs_Access2.Fields("����~�W").Value = rs_Tab4_OrderDetail.Fields("�~�W").Value
   rs_Access2.Fields("���w�����_�Х�").Value = rs_Tab4_OrderDetail.Fields("���O").Value
   rs_Access2.Fields("���w�����_���").Value = rs_Tab4_OrderDetail.Fields("���O���e").Value
   rs_Access2.Fields("�q��q_CaseQty").Value = rs_Tab4_OrderDetail.Fields("�q��c��").Value
   rs_Access2.Fields("�z�f�q_CaseQty").Value = rs_Tab4_OrderDetail.Fields("�z�f�c��").Value
   rs_Access2.Fields("�O��").Value = rs_Tab4_OrderDetail.Fields("�z�f�O��").Value
   rs_Access2.Fields("���n").Value = rs_Tab4_OrderDetail.Fields("�z�f���n").Value
   rs_Access2.Fields("���q").Value = rs_Tab4_OrderDetail.Fields("�z�f���q").Value
   rs_Access2.Fields("�G���ƨ�����").Value = rs_Tab4_OrderDetail.Fields("�G���ƨ�����").Value
   rs_Access2.Fields("�G���ƨ�����").Value = rs_Tab4_OrderDetail.Fields("�G���ƨ�����").Value
   rs_Access2.Update
   rs_Tab4_OrderDetail.MoveNext
Loop
rs_Tab4_OrderDetail.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access �C�L����
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab4_PreView.Value = vbChecked Then
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "��B�����u���`��_RouteList", acViewPreview
Else
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "��B�����u���`��_RouteList", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
chk_Tab4_PreView.Value = vbUnchecked
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cnAccess.RollbackTrans
      Tran_Level = 0
   End If
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--����C�L", Me.Caption, "cmd_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_QueryBySRouteNo_Click()
'��B�����u���`�� >> ��ƿz�� >> ���s�z��

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass

'��B�����u���`��G�G���ƨ����s�P�@���ƨ����s�B�q������C��
Call SetGridFormat_Tab4_RouteList

str_SQL = "Select �����ƨ����s,�G���ƨ����s,�X�����,���P���X,����,�f�D�渹,�q����,�e�f���,�Ȥ�s��,�Ȥ�W�� " & _
          "From Report_DCRouteSumSrc "
          

Dim strWhere As String, strTmp As String
strWhere = ""
'���u�s��
If txt_Tab4_SecondRouteNo.Text <> "" Then
   strTmp = " (�G���ƨ����s = '" & txt_Tab4_SecondRouteNo.Text & "' or �����ƨ����s = '" & txt_Tab4_SecondRouteNo.Text & "') "
   If Len(strTmp) > 0 Then
      If Len(strWhere) > 0 Then
         strWhere = strWhere & " and " & strTmp
      Else
         strWhere = strWhere & strTmp
      End If
   End If
End If
'�X�����
strTmp = ""
If Len(txt_Tab4_DeliveryDate_Start.Text) > 0 And Len(txt_Tab4_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� between '" & txt_Tab4_DeliveryDate_Start.Text & "' and '" & txt_Tab4_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab4_DeliveryDate_Start.Text) > 0 And Len(txt_Tab4_DeliveryDate_End.Text) = 0 Then
   strTmp = " �X����� = '" & txt_Tab4_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab4_DeliveryDate_Start.Text) = 0 And Len(txt_Tab4_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� = '" & txt_Tab4_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If
If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "�п�J�A���d�߱���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
str_SQL = str_SQL & " Order by �����ƨ����s,���P���X,�f�D�渹"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�q���Ƭd�ߵ��G�G�L�������q����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   txt_Tab4_SecondRouteNo.SelStart = 0: txt_Tab4_SecondRouteNo.SelLength = Len(txt_Tab4_SecondRouteNo.Text): txt_Tab4_SecondRouteNo.SetFocus
   Exit Sub
End If

Dim iLoop As Double
iLoop = 0
dg_Tab4_RouteList.Visible = False
Do While Not tmp_Rs.EOF
   iLoop = iLoop + 1
   With dg_Tab4_RouteList
        If iLoop + 1 >= .Rows Then .Rows = .Rows + 1
            .Row = iLoop
            .Col = 0: .Text = iLoop
            If chk_Tab4_Selected.Value = vbChecked Then
               .Col = 1: .Text = "��"
            Else
               .Col = 1: .Text = " " '"��"
            End If
            .Col = 2: .Text = tmp_Rs.Fields("�����ƨ����s").Value
            .Col = 3: .Text = tmp_Rs.Fields("�G���ƨ����s").Value
            .Col = 4: .Text = tmp_Rs.Fields("�X�����").Value
            .Col = 5: .Text = tmp_Rs.Fields("���P���X").Value
            .Col = 6: .Text = tmp_Rs.Fields("����").Value
            .Col = 7: .Text = tmp_Rs.Fields("�f�D�渹").Value
            .Col = 8: .Text = tmp_Rs.Fields("�q����").Value
            .Col = 9: .Text = tmp_Rs.Fields("�e�f���").Value
            .Col = 10: .Text = tmp_Rs.Fields("�Ȥ�s��").Value
            .Col = 11: .Text = tmp_Rs.Fields("�Ȥ�W��").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
dg_Tab4_RouteList.Visible = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "��B�����u���`��-��ƿz��-���s�z��", Me.Caption, "cmd_Tab4_QueryBySRouteNo_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_Query_RouteDetail_Click()
'��B�����u���`�� >> ��ƿz�� >> ���s�z��
Dim strOrderkey As String, i As Double
Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
strOrderkey = ""
With dg_Tab4_RouteList
     For i = 1 To .Rows - 2   'OrderList Grid �û��|�O�d�@�C���ť�
         .Row = i: .Col = 1   '�O�_����G�n��X��
         If Len(Trim(.Text)) > 0 Then
            .Col = 7   '�f�D�渹 ���
            If Len(strOrderkey) > 0 Then
               strOrderkey = strOrderkey & ",'" & RTrim(.Text) & "'"
            Else
               strOrderkey = "'" & RTrim(.Text) & "'"
            End If
         End If
     Next i
End With
If Len(strOrderkey) = 0 Then
   msg_text = "���u���`��ƿz��@�~���~�T���G�S������n���q��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

SSTab2.Tab = 1
'str_SQL = "Select �G�����s,�f��,�~�W,���O,���O���e,sum(�q��c��) as �q��c��,sum(�z�f�c��) as �z�f�c��,sum(�˳f�O��) as �z�f�O��," & _
'          "      sum(�˳f���q) as �z�f���q,sum(�˳f���n) as �z�f���n,�G���ƨ�����,�G���ƨ����� " & _
'          "From  Report_DCRouteSum " & _
'          "Where �f�D�渹 in (" & strOrderKey & ") and ���u�s�� = '" & txt_Tab4_SecondRouteNo.Text & "' " & _
'          "Group by �G�����s,�f��,�~�W,���O,���O���e,�G���ƨ�����,�G���ƨ����� Order by �G�����s,�f��,�~�W,���O "
'daniel_2004100
str_SQL = "Select �G�����s,�f��,�~�W,���O,���O���e,sum(�q��c��) as �q��c��,sum(�z�f�c��) as �z�f�c��,sum(�˳f�O��) as �z�f�O��," & _
          "      sum(�˳f���q) as �z�f���q,sum(�˳f���n) as �z�f���n,�G���ƨ�����,�G���ƨ�����,�Ȥ�s�� " & _
          "From  Report_DCRouteSum " & _
          "Where �f�D�渹 in (" & strOrderkey & ")  " & _
          "Group by �G�����s,�f��,�~�W,���O,���O���e,�G���ƨ�����,�G���ƨ�����,�Ȥ�s�� Order by �G�����s,�f��,�~�W,���O "
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Call Replication_Recordset(tmp_Rs, rs_Tab4_OrderDetail)
With dg_Tab4_OrderDetail
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 230                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab4_OrderDetail.MoveFirst
Set dg_Tab4_OrderDetail.DataSource = rs_Tab4_OrderDetail
With dg_Tab4_OrderDetail
    .Columns(0).Width = 500        '�s��
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1100       '�G�����s
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900        '�f��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '����~�W
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500        '�q����w�������O��
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 900        '�q����w�����--���
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '�q��c��
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800        '�z�f�c��
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 800        '�z�f�O��
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 800        '�z�f���q
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800        '�z�f���n
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 1300        '�G���ƨ����P���X
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1300        '�G���ƨ�����
    .Columns(12).Alignment = dbgLeft
End With
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "��B�����u���`��-��ƿz��-���u���`", Me.Caption, "cmd_Tab4_Query_RouteDetail_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5_PrintReport_Click()
'�ƨ��@���� >> ����C�L
If rs_Tab5_PlanList Is Nothing Then Exit Sub
If rs_Tab5_PlanList.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle

'1. ��Ƽg�X Access ��Ʈw >> �����˸����`��
Dim iLoop As Double
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From �ƨ��@����"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "�ƨ��@����", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab5_PlanList.MoveFirst

Do While Not rs_Tab5_PlanList.EOF
   rs_Access.AddNew
   For iLoop = 0 To rs_Tab5_PlanList.Fields.Count - 1
       rs_Access.Fields(iLoop).Value = rs_Tab5_PlanList.Fields(iLoop).Value
   Next iLoop
   rs_Access.Update
   rs_Tab5_PlanList.MoveNext
Loop

rs_Tab5_PlanList.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access �C�L����
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
'MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab5_PreView.Value = vbChecked Then
   '�w���C�L
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "�ƨ�������", acViewPreview
Else
   '�����C�L�ܦL���
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "�ƨ�������", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Call Unload_RunLogForm
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@����-�C�L", Me.Caption, "cmd_Tab5_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab5_PrintReport1_Click()
'�ƨ��@���� >> ���f��C�L
Dim strTmp As String, strShortName As String, strRoute_No As String, strOrderkey As String

On Error GoTo err_Handle
str_SQL = "Select * From Report_TRPPlanList1 where 1 = 1 "
    
'�q����
If Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = "and �X����� between '" & txt_Tab5_DeliveryDate_Start.Text & "' and '" & txt_Tab5_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) = 0 Then
   strTmp = "and �X����� = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) = 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = "and �X����� = '" & txt_Tab5_DeliveryDate_End.Text & "' "
End If

str_SQL = str_SQL & strTmp

'���u�s��
strTmp = ""
If Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = "and Rtrim(���u�s��) between '" & txt_Tab5_route_Start.Text & "' and '" & txt_Tab5_route_End.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) = 0 Then
   strTmp = "and Rtrim(���u�s��) = '" & txt_Tab5_route_Start.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) = 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = "and Rtrim(���u�s��) = '" & txt_Tab5_route_End.Text & "' "
End If

str_SQL = str_SQL & strTmp

'�B�e�ϰ�
strTmp = ""
If cmb_Tab5_AreaCode.ListIndex <> -1 Then
   strTmp = "and �ϽX = '" & mySplit(cmb_Tab5_AreaCode, " ", 0) & "' "
End If

str_SQL = str_SQL & strTmp & " order by �X�����,���u�s��,�~�� "

Screen.MousePointer = 11

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim rs_Tab5_PlanList1 As New ADODB.Recordset

Call Replication_Recordset(tmp_Rs, rs_Tab5_PlanList1)
tmp_Rs.Close

'1. ��Ƽg�X Access ��Ʈw >> �����˸����`��
Dim iLoop As Double
Call AccessDB_Connect
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From ���f��"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "���f��", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab5_PlanList1.MoveFirst

Do While Not rs_Tab5_PlanList1.EOF
   
   rs_Access.AddNew
   For iLoop = 0 To rs_Tab5_PlanList1.Fields.Count - 1
       rs_Access.Fields(iLoop).Value = rs_Tab5_PlanList1.Fields(iLoop).Value
   Next iLoop
   
   If strRoute_No <> rs_Tab5_PlanList1("���u�s��") Then
   strRoute_No = rs_Tab5_PlanList1("���u�s��")
   '���Ȥ�W��
    str_SQL = "select distinct isnull(t1m.Short_Name,'') as Short_Name , m2t.route_no " & _
                "from trp02t m2t join TRP01M t1m on t1m.ConsigneeKey = m2t.ConsigneeKey and t1m.storerkey = m2t.storerkey " & _
                "where m2t.route_no = '" & strRoute_No & "' order by t1m.Short_Name desc "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    strShortName = ""
    tmp_Rs.MoveFirst
    Do While Not tmp_Rs.EOF
        strShortName = strShortName & RTrim(tmp_Rs("Short_Name")) & ";"
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    
    '���q�渹�X
    str_SQL = "select distinct RTRIM(Extern) as Orderkey from trp02t where route_no = '" & strRoute_No & "' GROUP BY Extern order by RTRIM(Extern) "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    strOrderkey = ""
    tmp_Rs.MoveFirst
    
    Do While Not tmp_Rs.EOF
        strOrderkey = strOrderkey & RTrim(tmp_Rs("Orderkey")) & ";"
    tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    
    End If

    rs_Access.Fields("�Ȥ�²��") = strShortName
    rs_Access.Fields("�q�渹�X") = strOrderkey
   
'   rs_Access.Fields("�Ȥ�²��").Value = .Fields("�s��").Value
'   rs_Access.Fields("�q�渹�X").Value = .Fields("�s��").Value

   rs_Access.Update
   rs_Tab5_PlanList1.MoveNext
Loop

cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)
Set rs_Tab5_PlanList1 = Nothing

'2. call Access �C�L����
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
'MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab5_PreView.Value = vbChecked Then
   '�w���C�L
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "���f��", acViewPreview
Else
   '�����C�L�ܦL���
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "���f��", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If

Exit Sub

err_Handle:
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
      Tran_Level = 0
   End If
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd_Tab5_Query_Click()
'�ƨ��@���� >> �d��
Set dg_Tab5_PlanList.DataSource = Nothing
Set rs_Tab5_PlanList = Nothing

Screen.MousePointer = 11
On Error GoTo err_Handle
'
str_SQL = "Select �X�����,�ϰ�,�B�e�ϰ�,�Ȧs��,�����O,�f�B���q,���P���X,����,�@��h��,�r�p�H," & _
          "       �i�����q,�i�����n,���u�s��,�B�e�I��,�B�e�c��,�B�e�Ӽ�,�B�e�O��,�B�e���q,�B�e���n,�f�B���q�N�X,�Ƶ�,�w�p�������ɶ�,�f�D�W��,�Ȥ�²��,����,�~��,�[�u " & _
          "From Report_TRPPlanList "
          
Dim strWhere As String, strTmp As String
strWhere = ""
'�q����
strTmp = ""
If Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� between '" & txt_Tab5_DeliveryDate_Start.Text & "' and '" & txt_Tab5_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) = 0 Then
   strTmp = " �X����� = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) = 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
   strTmp = " �X����� = '" & txt_Tab5_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'���u�s��
strTmp = ""
If Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = " Rtrim(���u�s��) between '" & txt_Tab5_route_Start.Text & "' and '" & txt_Tab5_route_End.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) = 0 Then
   strTmp = " Rtrim(���u�s��) = '" & txt_Tab5_route_Start.Text & "' "
ElseIf Len(txt_Tab5_route_Start.Text) = 0 And Len(txt_Tab5_route_End.Text) > 0 Then
   strTmp = " Rtrim(���u�s��) = '" & txt_Tab5_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'�B�e�ϰ�
strTmp = ""
If cmb_Tab5_AreaCode.ListIndex <> -1 Then
   strTmp = " �ϰ� = '" & arAreaCode(cmb_Tab5_AreaCode.ListIndex) & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If strWhere <> "" Then
   str_SQL = str_SQL & " Where " & strWhere
Else
   msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

str_SQL = str_SQL & " order by �X�����,���u�s�� "
'str_SQL = str_SQL & " order by �X�����,�Ƶ�,���u�s��,���P���X "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab5_PlanList)
tmp_Rs.Close

With dg_Tab5_PlanList
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_Tab5_PlanList.MoveFirst
Set dg_Tab5_PlanList.DataSource = rs_Tab5_PlanList
dg_Tab5_PlanList.Visible = False

With dg_Tab5_PlanList
    .ColumnHeaders = True          '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�X�����
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 500        '�ϰ�
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1500       '�B�e�ϰ�
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 800        '�Ȧs��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1500       '�����O
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1500       '�f�B���q
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1000       '���P���X
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 500        '����
    .Columns(8).Alignment = dbgCenter
    .Columns(9).Width = 800        '�@��h��
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 1000       '�r�p�H
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 800        '�i�����q
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800       '�i�����n
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 1100      '���u�s��
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 800       '�B�e�I��
    .Columns(14).Alignment = dbgCenter
    .Columns(15).Width = 800       '�B�e�c��
    .Columns(15).Alignment = dbgRight
    .Columns(16).Width = 800       '�B�e�O��
    .Columns(16).Alignment = dbgRight
    .Columns(17).Width = 800       '�B�e���q
    .Columns(17).Alignment = dbgRight
    .Columns(18).Width = 800       '�B�e���n
    .Columns(18).Alignment = dbgRight
    .Columns(19).Width = 1200       '�f�B���q�N�X
    .Columns(19).Alignment = dbgLeft
    .Columns(20).Width = 1200       '�Ƶ�(�G���ƨ����u�s��)
    .Columns(20).Alignment = dbgLeft
    .Columns(21).Width = 1200       '�w�p�������ɶ�
    .Columns(21).Alignment = dbgCenter
    .Columns(22).Width = 3000       '�Ȥ�²��
    .Columns(22).Alignment = dbgLeft
End With
rs_Tab5_PlanList.MoveFirst

'���Ҧ��Ȥ�W��
Dim strShort_name As String
Call Confirm_Recordset_Closed(tmp_Rs)

Do While Not rs_Tab5_PlanList.EOF
    '���Ȥ�W��
    str_SQL = "select distinct isnull(t1m.Short_Name,'') as Short_Name , m2t.route_no " & _
              "from trp02t m2t join TRP01M t1m on t1m.ConsigneeKey = m2t.ConsigneeKey and t1m.storerkey = m2t.storerkey " & _
              "where m2t.route_no = '" & rs_Tab5_PlanList("���u�s��") & "' order by t1m.Short_Name desc "

    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    strShort_name = ""
    tmp_Rs.MoveFirst
    
    Do While Not tmp_Rs.EOF

    strShort_name = strShort_name & RTrim(tmp_Rs("Short_Name")) & ";"

    tmp_Rs.MoveNext

    Loop
    tmp_Rs.Close

    rs_Tab5_PlanList("�Ȥ�²��") = strShort_name
        
    '���t�m���
    str_SQL = "select Facility = sum(case when sectionkey = 'FACILITY' then 1 else 0 end) , Wild = sum(case when sectionkey <> 'FACILITY' then 1 else 0 end) , Repacking = sum(len(rtrim(isnull(od.updatesource,'')))) " & _
                "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..orderdetail od on od.orderkey = o.orderkey " & _
                "join " & strWMSDB & "..pickdetail p (nolock) on p.orderkey = od.orderkey and od.orderlinenumber = p.orderlinenumber " & _
                "join " & strWMSDB & "..loc l (nolock) on l.loc = p.loc " & _
                "where o.route = '" & rs_Tab5_PlanList("���u�s��") & "' "
    
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn
    
    If Not tmp_Rs.EOF Then
    
        If tmp_Rs("facility") > 0 Then rs_Tab5_PlanList("����") = "V"
        If tmp_Rs("wild") > 0 Then rs_Tab5_PlanList("�~��") = "V"
        If tmp_Rs("repacking") > 0 Then rs_Tab5_PlanList("�[�u") = "V"
    
    End If
    
    tmp_Rs.Close

rs_Tab5_PlanList.MoveNext
Loop
rs_Tab5_PlanList.MoveFirst
dg_Tab5_PlanList.Visible = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@����-�d��", Me.Caption, "cmd_Tab5_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5_ReSet_Click()
'�ƨ��@���� >> �M��
cmb_Tab5_AreaCode.ListIndex = -1
txt_Tab5_DeliveryDate_Start.Text = ""
txt_Tab5_DeliveryDate_End.Text = ""
txt_Tab5_route_Start.Text = ""
txt_Tab5_route_End.Text = ""
Set dg_Tab5_PlanList.DataSource = Nothing
Set rs_Tab5_PlanList = Nothing

End Sub

Private Sub cmd_Tab5_SaveToExcel_Click()
'�ƨ��@���� >> �� EXCEL

    If rs_Tab5_PlanList Is Nothing Then Exit Sub
    rs_Tab5_PlanList.MoveFirst
    On Error GoTo err_Handle
    
    Screen.MousePointer = 11
    
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
'    MyXlsApp.Sheets("Sheet1").Name = "�ƨ��@����"
    MyXlsApp.ActiveSheet.Name = "�ƨ��@����"
    i = 1
    'tr_SQL = "Select �X�����,�ϰ�,�B�e�ϰ�,�Ȧs��,�����O,�f�B���q,���P���X,����,�@��h��,�r�p�H,�i�����q,�i�����n,���u�s��,�B�e�I��,�B�e�c��,�B�e�O��,�B�e���q,�B�e���n,�f�B���q�N�X,�Ƶ�,�w�p�������ɶ�,�Ȥ�²�� "
    MyXlsApp.Cells(i, 1).Value = "�s��"
    MyXlsApp.Cells(i, 2).Value = "�X�����"
    MyXlsApp.Cells(i, 3).Value = "�ϰ�"
    MyXlsApp.Cells(i, 4).Value = "�Ȧs��"
    MyXlsApp.Cells(i, 5).Value = "�����O"
    MyXlsApp.Cells(i, 6).Value = "���P���X"
    MyXlsApp.Cells(i, 7).Value = "����"
    MyXlsApp.Cells(i, 8).Value = "�r�p�H"
    MyXlsApp.Cells(i, 9).Value = "���u�s��"
    MyXlsApp.Cells(i, 10).Value = "�B�e�c��"
    MyXlsApp.Cells(i, 11).Value = "�B�e�Ӽ�"
    MyXlsApp.Cells(i, 12).Value = "�B�e���q"
    MyXlsApp.Cells(i, 13).Value = "�B�e���n"
    MyXlsApp.Cells(i, 14).Value = "�Ƶ�"
    MyXlsApp.Cells(i, 15).Value = "�ɶ�"
    MyXlsApp.Cells(i, 16).Value = "�f�D�W��"
    MyXlsApp.Cells(i, 17).Value = "�Ȥ�²��"
    MyXlsApp.Cells(i, 18).Value = "����"
    MyXlsApp.Cells(i, 19).Value = "�~��"
    MyXlsApp.Cells(i, 20).Value = "�[�u"
    MyXlsApp.Cells(i, 21).Value = "�l�ܮɶ�"
    MyXlsApp.Cells(i, 22).Value = "�T�{"
    MyXlsApp.Cells(i, 23).Value = "�ɥX"
    MyXlsApp.Cells(i, 24).Value = "�^��"
    MyXlsApp.Cells(i, 25).Value = "�j�O"
    i = i + 1
    j = i
    rs_Tab5_PlanList.MoveFirst
    '���,����,�渹,�Z�O,�ɥX,�٤J
    Do While Not rs_Tab5_PlanList.EOF
        If i > 2 Then
            If MyXlsApp.Cells(i - 1, 6).Value <> rs_Tab5_PlanList.Fields(7) Then
                '�������P,�j�@��b�g�Jexcel
                MyXlsApp.Cells(i, 10).Value = "=SUM(J" & CStr(j) & ":J" & CStr(i - 1) & ")"  '�B�e�c��
                MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")"  '�B�e�Ӽ�
                MyXlsApp.Cells(i, 13).Value = "=SUM(M" & CStr(j) & ":M" & CStr(i - 1) & ")" '�B�e���q
                MyXlsApp.Cells(i, 12).Value = "=SUM(L" & CStr(j) & ":L" & CStr(i - 1) & ")" '�B�e���n
                i = i + 2
                j = i
            End If
        End If
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab5_PlanList.Fields(1)) '�X�����
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rs_Tab5_PlanList.Fields(2) '�ϰ�
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rs_Tab5_PlanList.Fields(4) '�Ȧs��
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rs_Tab5_PlanList.Fields(5) '�����O
        MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 6).Value = rs_Tab5_PlanList.Fields(7) '����
        MyXlsApp.Cells(i, 7).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 7).Value = rs_Tab5_PlanList.Fields(8) '����
        'MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 8).Value = rs_Tab5_PlanList.Fields(10) '�r�p�H
        MyXlsApp.Cells(i, 9).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 9).Value = rs_Tab5_PlanList.Fields(13)    '���u�s��
        MyXlsApp.Cells(i, 10).Value = rs_Tab5_PlanList.Fields(15)    '�B�e�c��
        MyXlsApp.Cells(i, 11).Value = rs_Tab5_PlanList.Fields(16)    '�B�e�Ӽ�
        MyXlsApp.Cells(i, 13).Value = rs_Tab5_PlanList.Fields(19)   '�B�e���q
        MyXlsApp.Cells(i, 12).Value = rs_Tab5_PlanList.Fields(18)   '�B�e���n
        MyXlsApp.Cells(i, 14).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 14).Value = rs_Tab5_PlanList.Fields(21)   '�Ƶ�
        MyXlsApp.Cells(i, 15).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 15).Value = Mid(rs_Tab5_PlanList.Fields(22), 10, 4)   '�ɶ�
        MyXlsApp.Cells(i, 16).Value = rs_Tab5_PlanList.Fields(23)   '�f�D�W��
        MyXlsApp.Cells(i, 17).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 17).Value = rs_Tab5_PlanList.Fields(24)   '�Ȥ�²��
        MyXlsApp.Cells(i, 18).Value = rs_Tab5_PlanList.Fields(25)   '����
        MyXlsApp.Cells(i, 19).Value = rs_Tab5_PlanList.Fields(26)   '�~��
        MyXlsApp.Cells(i, 20).Value = rs_Tab5_PlanList.Fields(27)   '�[�u
        rs_Tab5_PlanList.MoveNext
        i = i + 1
    Loop
    '�p��c�ƭӼƭ��q���n
    MyXlsApp.Cells(i, 10).Value = "=SUM(J" & CStr(j) & ":J" & CStr(i - 1) & ")"  '�B�e�c��
    MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")"  '�B�e�Ӽ�
    MyXlsApp.Cells(i, 13).Value = "=SUM(M" & CStr(j) & ":M" & CStr(i - 1) & ")" '�B�e���q
    MyXlsApp.Cells(i, 12).Value = "=SUM(L" & CStr(j) & ":L" & CStr(i - 1) & ")" '�B�e���n
    i = i + 1
    '�̾A��e
    MyXlsApp.Columns("A:X").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w,���q�M���n
    MyXlsApp.Columns("L:M").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:X" & i - 1).Select
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
    
    '���y���
    str_SQL = "select VEHICLE_ID_NO,DRIVER,DRIVER_PHONE from TRP09M"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.MoveFirst
        MyXlsApp.Sheets.Add
'        MyXlsApp.Sheets("Sheet2").Select
'        MyXlsApp.Sheets("Sheet2").Name = "���y���"
        MyXlsApp.ActiveSheet.Name = "���y���"
        i = 1
        MyXlsApp.Cells(i, 1).Value = "����"
        MyXlsApp.Cells(i, 2).Value = "�q��"
        MyXlsApp.Cells(i, 3).Value = "�q��"
        i = i + 1
        Do While Not tmp_Rs.EOF
            MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
            MyXlsApp.Cells(i, 1).Value = Trim(tmp_Rs.Fields(0))
            MyXlsApp.Cells(i, 2).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 2).Value = Trim(tmp_Rs.Fields(1))
            MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 3).Value = Trim(tmp_Rs.Fields(2))
            tmp_Rs.MoveNext
            i = i + 1
        Loop
        '�q������
        MyXlsApp.Sheets("�ƨ��@����").Select
        MyXlsApp.Range("H2").Select
        MyXlsApp.ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-2])=0,"""",VLOOKUP(RC[-2],���y���!C[-7]:C[-5],2,FALSE))"
    End If
    tmp_Rs.Close
    
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@����-�� EXCEL", Me.Caption, "cmd_Tab5_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5_SaveToExcel_NEW_Click()
'�ƨ��@���� >> �� EXCEL NEW

    If rs_Tab5_PlanList Is Nothing Then Exit Sub
    rs_Tab5_PlanList.MoveFirst
    On Error GoTo err_Handle
    Screen.MousePointer = 11
    
    Dim strWhere As String, strTmp As String, lngAR As Long, lngAP As Long, lngSorting As Long
    
    str_SQL = "select * from gv_TRPPlanLst where 1 = 1 and "
    
    strWhere = ""
    '�q����
    strTmp = ""
    If Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
       strTmp = " �X����� between '" & txt_Tab5_DeliveryDate_Start.Text & "' and '" & txt_Tab5_DeliveryDate_End.Text & "' "
    ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) > 0 And Len(txt_Tab5_DeliveryDate_End.Text) = 0 Then
       strTmp = " �X����� = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
    ElseIf Len(txt_Tab5_DeliveryDate_Start.Text) = 0 And Len(txt_Tab5_DeliveryDate_End.Text) > 0 Then
       strTmp = " �X����� = '" & txt_Tab5_DeliveryDate_End.Text & "' "
    End If
    
    If Len(strTmp) > 0 Then
       If Len(strWhere) > 0 Then
          strWhere = strWhere & " and " & strTmp
       Else
          strWhere = strWhere & strTmp
       End If
    End If
    
    '���u�s��
    strTmp = ""
    If Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) > 0 Then
       strTmp = " Rtrim(���u�s��) between '" & txt_Tab5_route_Start.Text & "' and '" & txt_Tab5_route_End.Text & "' "
    ElseIf Len(txt_Tab5_route_Start.Text) > 0 And Len(txt_Tab5_route_End.Text) = 0 Then
       strTmp = " Rtrim(���u�s��) = '" & txt_Tab5_route_Start.Text & "' "
    ElseIf Len(txt_Tab5_route_Start.Text) = 0 And Len(txt_Tab5_route_End.Text) > 0 Then
       strTmp = " Rtrim(���u�s��) = '" & txt_Tab5_route_End.Text & "' "
    End If
    
    If Len(strTmp) > 0 Then
       If Len(strWhere) > 0 Then
          strWhere = strWhere & " and " & strTmp
       Else
          strWhere = strWhere & strTmp
       End If
    End If
        
    str_SQL = str_SQL & strWhere & " order by �X�����,left(�ϰ�,1),�Ƶ�,�f�D"
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    
    Call Replication_Recordset(tmp_Rs, rs_Tab5_TRPPlanList)
    tmp_Rs.Close

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
'    MyXlsApp.Sheets("Sheet1").Name = "�ƨ��@����"
    MyXlsApp.ActiveSheet.Name = "�ƨ��@����"
    i = 1
    'tr_SQL = "Select �X�����,�ϰ�,�B�e�ϰ�,�f�B���q,���P���X,����,�@��h��,�r�p�H,�i�����q,�i�����n,���u�s��,�B�e�I��,�B�e�c��,�B�e�O��,�B�e���q,�B�e���n,�f�B���q�N�X,�Ƶ�,�w�p�������ɶ�,�Ȥ�²�� "
    MyXlsApp.Cells(i, 1).Value = "�s��"
    MyXlsApp.Cells(i, 2).Value = "�X�����"
    MyXlsApp.Cells(i, 3).Value = "�ϰ�"
    MyXlsApp.Cells(i, 4).Value = "���P���X"
    MyXlsApp.Cells(i, 5).Value = "����"
    MyXlsApp.Cells(i, 6).Value = "�r�p�H"
    MyXlsApp.Cells(i, 7).Value = "���u�s��"
    MyXlsApp.Cells(i, 8).Value = "�f�D"
    MyXlsApp.Cells(i, 9).Value = "����"
    MyXlsApp.Cells(i, 10).Value = "���I"
    MyXlsApp.Cells(i, 11).Value = "½�O�z�f"
    MyXlsApp.Cells(i, 12).Value = "�B�e���"
    MyXlsApp.Cells(i, 13).Value = "�B�e�c��"
    MyXlsApp.Cells(i, 14).Value = "�B�e���q"
    MyXlsApp.Cells(i, 15).Value = "�B�e���n"
    MyXlsApp.Cells(i, 16).Value = "�Ƶ�"
    MyXlsApp.Cells(i, 17).Value = "�ɶ�"
    MyXlsApp.Cells(i, 18).Value = "�Ȥ�²��"
    MyXlsApp.Cells(i, 19).Value = "�l�ܮɶ�"
    MyXlsApp.Cells(i, 20).Value = "�T�{"
    MyXlsApp.Cells(i, 21).Value = "�ɥX"
    MyXlsApp.Cells(i, 22).Value = "�^��"
    MyXlsApp.Cells(i, 23).Value = "�j�O"
    i = i + 1
    j = i
    
    rs_Tab5_TRPPlanList.MoveFirst
    '���,����,�渹,�Z�O,�ɥX,�٤J
    Do While Not rs_Tab5_TRPPlanList.EOF
        If i > 2 Then
            If RTrim(MyXlsApp.Cells(i - 1, 4).Value) & RTrim(MyXlsApp.Cells(i - 1, 15).Value) <> RTrim(rs_Tab5_TRPPlanList.Fields(3)) & RTrim(rs_Tab5_TRPPlanList.Fields(12)) Then
                '�������P,�j�@��A�g�Jexcel
                MyXlsApp.Cells(i, 9).Value = "=SUM(i" & CStr(j) & ":i" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 10).Value = "=SUM(j" & CStr(j) & ":j" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 11).Value = "=SUM(k" & CStr(j) & ":k" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 12).Value = "=SUM(l" & CStr(j) & ":l" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 13).Value = "=SUM(m" & CStr(j) & ":m" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 14).Value = "=SUM(n" & CStr(j) & ":n" & CStr(i - 1) & ")"
                MyXlsApp.Cells(i, 15).Value = "=SUM(o" & CStr(j) & ":o" & CStr(i - 1) & ")"
                i = i + 2
                j = i
            End If
        End If
        
        '���Ҧ��Ȥ�W��
        Dim strShort_name As String
        Call Confirm_Recordset_Closed(tmp_Rs)
               
        str_SQL = "select distinct isnull(t1m.Short_Name,'') as Short_Name , m2t.route_no " & _
                    "from trp02t m2t join TRP01M t1m on t1m.ConsigneeKey = m2t.ConsigneeKey and m2t.storerkey = t1m.storerkey " & _
                    "join trp16m t16m on t16m.storerkey = t1m.storerkey and t16m.short_name = '" & rs_Tab5_TRPPlanList("�f�D") & "' " & _
                    "where m2t.route_no = '" & rs_Tab5_TRPPlanList("���u�s��") & "' order by t1m.Short_Name desc "
    
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        strShort_name = ""
        tmp_Rs.MoveFirst
        
        Do While Not tmp_Rs.EOF
    
            strShort_name = strShort_name & RTrim(tmp_Rs("Short_Name")) & ";"
    
        tmp_Rs.MoveNext
    
        Loop
        tmp_Rs.Close
        
        '�B�O�w���p��
        If rs_Tab5_TRPPlanList("�X�����") > Format(Now - 7, "YYYYMMDD") Then cn.Execute "exec gs_precost '" & IIf(Len(Trim(rs_Tab5_TRPPlanList("�Ƶ�"))) = 0, rs_Tab5_TRPPlanList("���u�s��"), rs_Tab5_TRPPlanList("�Ƶ�")) & "','" & rs_Tab5_TRPPlanList("�f�D") & "' ", RowsAffect, adExecuteNoRecords
        
        '�����u�s�������I���
        lngAR = 0: lngAP = 0: lngSorting = 0
        
        Call Confirm_Recordset_Closed(tmp_Rs)
               
        str_SQL = "select ar=sum(t2.receivable),ap = sum(t2.payable) ,sorting = isnull((select sum((palletqty * 50 )+ ((case when storer = 'LTHL01' then 45 else 40 end) * sortingqty/1000)) from gt_LoadSorting where route_no = t2.route_no and storer = t2.storerkey ),0) from trp02t t2 join trp16m t16m on t16m.storerkey = t2.storerkey and t16m.short_name = '" & rs_Tab5_TRPPlanList("�f�D") & "' where t2.route_no = '" & rs_Tab5_TRPPlanList("���u�s��") & "' group by t2.route_no ,t2.storerkey"
    
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
'        lngAR = tmp_rs("ar")������������B
        lngAP = tmp_Rs("ap")
        lngSorting = tmp_Rs("sorting")
        
        tmp_Rs.Close
        
        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab5_TRPPlanList.Fields(1))
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rs_Tab5_TRPPlanList.Fields(2)
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rs_Tab5_TRPPlanList.Fields(3)
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rs_Tab5_TRPPlanList.Fields(4)
        'MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 6).Value = rs_Tab5_TRPPlanList.Fields(5)
        MyXlsApp.Cells(i, 7).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 7).Value = rs_Tab5_TRPPlanList.Fields(6)
        MyXlsApp.Cells(i, 8).Value = rs_Tab5_TRPPlanList.Fields(7)
        MyXlsApp.Cells(i, 9).Value = lngAR
        MyXlsApp.Cells(i, 10).Value = lngAP
        MyXlsApp.Cells(i, 11).Value = lngSorting
        MyXlsApp.Cells(i, 12).Value = rs_Tab5_TRPPlanList.Fields(8)
        MyXlsApp.Cells(i, 13).Value = rs_Tab5_TRPPlanList.Fields(9)
        MyXlsApp.Cells(i, 14).Value = rs_Tab5_TRPPlanList.Fields(10)
        MyXlsApp.Cells(i, 15).Value = rs_Tab5_TRPPlanList.Fields(11)
        MyXlsApp.Cells(i, 16).Value = rs_Tab5_TRPPlanList.Fields(12)
        MyXlsApp.Cells(i, 17).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 17).Value = Mid(rs_Tab5_TRPPlanList.Fields(13), 10, 4)
        MyXlsApp.Cells(i, 18).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 18).Value = strShort_name
        rs_Tab5_TRPPlanList.MoveNext
        i = i + 1
    Loop
    
    MyXlsApp.Cells(i, 9).Value = "=SUM(I" & CStr(j) & ":I" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 10).Value = "=SUM(j" & CStr(j) & ":j" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 11).Value = "=SUM(k" & CStr(j) & ":k" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 12).Value = "=SUM(l" & CStr(j) & ":l" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 13).Value = "=SUM(m" & CStr(j) & ":m" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 14).Value = "=SUM(n" & CStr(j) & ":n" & CStr(i - 1) & ")"
    MyXlsApp.Cells(i, 15).Value = "=SUM(o" & CStr(j) & ":o" & CStr(i - 1) & ")"
    i = i + 1
    '�̾A��e
    MyXlsApp.Columns("A:W").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w
    MyXlsApp.Columns("m:n").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:W" & i - 1).Select
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
    
    '���y���
    str_SQL = "select VEHICLE_ID_NO,DRIVER,DRIVER_PHONE from TRP09M"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.MoveFirst
        MyXlsApp.Sheets.Add
'        MyXlsApp.Sheets("Sheet2").Select
'        MyXlsApp.Sheets("Sheet2").Name = "���y���"
        MyXlsApp.ActiveSheet.Name = "���y���"
        i = 1
        MyXlsApp.Cells(i, 1).Value = "����"
        MyXlsApp.Cells(i, 2).Value = "�q��"
        MyXlsApp.Cells(i, 3).Value = "�q��"
        i = i + 1
        Do While Not tmp_Rs.EOF
            MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
            MyXlsApp.Cells(i, 1).Value = Trim(tmp_Rs.Fields(0))
            MyXlsApp.Cells(i, 2).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 2).Value = Trim(tmp_Rs.Fields(1))
            MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
            MyXlsApp.Cells(i, 3).Value = Trim(tmp_Rs.Fields(2))
            tmp_Rs.MoveNext
            i = i + 1
        Loop
        '�q������
        MyXlsApp.Sheets("�ƨ��@����").Select
        MyXlsApp.Range("F2").Select
        MyXlsApp.ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-2])=0,"""",VLOOKUP(RC[-2],���y���!C[-5]:C[-3],2,FALSE))"
    End If
    tmp_Rs.Close
    
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@����-�� EXCEL", Me.Caption, "cmd_Tab5_SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6_Query_Click()
'�̪O���@by���s >> �d��
' add by Terry 20180518
Set dg_Tab6_PlanList.DataSource = Nothing
Set rs_Tab6_PlanList = Nothing


Screen.MousePointer = 11
On Error GoTo err_Handle

str_SQL = "select �X����� = Convert(VarChar, s01t.Delivery_Date, 112),�f�B���q = Rtrim(Isnull(t08m.C_Name,'')),�f�B���q�N�X = t08m.COMPANY_CODE,�Ȥ�²�� = ISNULL(RTRIM(t01m.short_name),'') " & _
          ",���P���X = Rtrim(Isnull(s01t.C_VEHICLE_ID_NO,'')),�r�p�H = Isnull(rtrim(s01t.Driver),''),���u�s�� = Rtrim(Isnull(s02t.ROUTE_NO,'')),�G�����s = Rtrim(Isnull(s02t.C_ROUTE_NO,'')) " & _
          ",�B�e�c�� = sum(case when p.casecnt = 0 then 0 else floor((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end) /p.casecnt) end) " & _
          ",�B�e�Ӽ� = sum(case when p.casecnt = 0 then (case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end) else cast((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end) as int)%cast(p.casecnt as int) end) " & _
          ",�B�e�O�� = sum(case when p.pallet = 0 then 0 else round((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end)/p.pallet,2) end) " & _
          ",�B�e���q = sum(round((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end)*s.STDGROSSWGT,2)) " & _
          ",�B�e���n = sum(round((case when s03t.ship_qty = 0 then s03t.order_qty else s03t.ship_qty end)*s.stdcube,2)) " & _
          ",�̪O���@ = s01t.PalletDefend " & _
          "From SDN01T s01t join SDN02T s02t on s01t.C_Route_No = s02t.C_ROUTE_NO " & _
          "join SDN03T s03t on s03t.receipt_no = s02t.receipt_no " & _
          "join " & strWMSDB & "..sku s on s03t.PRODUCT_NO = s.sku " & _
          "join " & strWMSDB & "..pack p on s.packkey = p.packkey " & _
          "join TRP01M t01m on s02t.CONSIGNEEKEY = t01m.CONSIGNEEKEY " & _
          "Left join TRP05T t05t on t05t.Route_No = s02t.Route_No " & _
          "Left join TRP08M t08m on t08m.Company_Code = t05t.TRP_Company_Code " & _
          "Where s01t.Delivery_Date > getdate() - 30 "


Dim tmpString1 As String, tmpString2 As String
Dim strWhere As String, strTmp As String
strWhere = ""



'�q����
strTmp = ""
If Len(txt_Tab6_DeliveryDate_Start.Text) > 0 And Len(txt_Tab6_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(VarChar, s01t.Delivery_Date, 112) between '" & txt_Tab6_DeliveryDate_Start.Text & "' and '" & txt_Tab6_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab6_DeliveryDate_Start.Text) > 0 And Len(txt_Tab6_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(VarChar, s01t.Delivery_Date, 112) = '" & txt_Tab5_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab6_DeliveryDate_Start.Text) = 0 And Len(txt_Tab6_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(VarChar, s01t.Delivery_Date, 112) = '" & txt_Tab6_DeliveryDate_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

'���u�s��
strTmp = ""
If Len(txt_Tab6_route_Start.Text) > 0 And Len(txt_Tab6_route_End.Text) > 0 Then
   strTmp = " Rtrim(s02t.route_no) between '" & txt_Tab6_route_Start.Text & "' and '" & txt_Tab6_route_End.Text & "' "
ElseIf Len(txt_Tab6_route_Start.Text) > 0 And Len(txt_Tab6_route_End.Text) = 0 Then
   strTmp = " Rtrim(s02t.route_no) = '" & txt_Tab5_route_Start.Text & "' "
ElseIf Len(txt_Tab6_route_Start.Text) = 0 And Len(txt_Tab6_route_End.Text) > 0 Then
   strTmp = " Rtrim(s02t.route_no) = '" & txt_Tab6_route_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If


If strWhere <> "" Then
   str_SQL = str_SQL & " and " & strWhere
End If

str_SQL = str_SQL & " group by s01t.Delivery_Date,t08m.C_Name,s01t.C_VEHICLE_ID_NO,s01t.Driver,s02t.ROUTE_NO,t08m.COMPANY_CODE,s02t.C_ROUTE_NO,t01m.SHORT_NAME,s01t.PalletDefend order by �X�����,���u�s�� "
str_SQL_Excel = str_SQL
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab6_PlanList)
tmp_Rs.Close

With dg_Tab6_PlanList
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_Tab6_PlanList.MoveFirst
Set dg_Tab6_PlanList.DataSource = rs_Tab6_PlanList
dg_Tab6_PlanList.Visible = False

With dg_Tab6_PlanList
    .ColumnHeaders = True          '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�X�����
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1500       '�f�B���q
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1200       '�f�B���q�N�X
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 3000       '�Ȥ�²��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1000       '���P���X
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1000       '�r�p�H
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1100      '���u�s��
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1200       '�Ƶ�(�G���ƨ����u�s��)
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 800       '�B�e�c��
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800       '�B�e�Ӽ�
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 800       '�B�e�O��
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800       '�B�e���q
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 800       '�B�e���n
    .Columns(13).Alignment = dbgRight
    .Columns(14).Width = 800       '�̪O���@
    .Columns(14).Alignment = dbgCenter
End With
rs_Tab6_PlanList.MoveFirst

dg_Tab6_PlanList.Visible = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�̪O���@by���s-�d��", Me.Caption, "cmd_Tab6_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmd_Tab6_SaveToExcel_Click()
On Error GoTo err_Handle
If rs_Tab6_PlanList Is Nothing Then Exit Sub
If rs_Tab6_PlanList.RecordCount = 0 Then Exit Sub

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL_Excel, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'��Excel
Call Recordset2Excel("PalletDefend", rsTmp)

Set MyXlsApp = Nothing
rsTmp.Close: Set rsTmp = Nothing
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdLTHL01ShipDate_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = adUseClient
str_SQL = "select * from gv_LTHL01ShipData where 1 = 1 "

'�X�����
Dim strTmp As String
strTmp = ""
If Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = "and �X����� between '" & txt_Tab0_DeliveryDate_Start.Text & "' and '" & txt_Tab0_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) = 0 Then
   strTmp = "and �X����� = '" & txt_Tab0_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) = 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = "and �X����� = '" & txt_Tab0_DeliveryDate_End.Text & "' "
End If

str_SQL = str_SQL & strTmp

rsTmp.Open str_SQL, cn
If rsTmp.EOF Then MsgBox "", 64, cmdLTHL01ShipDate.Caption: Screen.MousePointer = 0: Exit Sub

'���r��
'If Dir("C:\LTHL01\�X�f�^��", vbDirectory) = "" Then MkDirs "C:\LTHL01\�X�f�^��"
Open "C:\ShipDate.txt" For Output As #1

rsTmp.Sort = "�q�渹�X"

rsTmp.MoveFirst
Do While Not rsTmp.EOF
    Print #1, rsTmp("�q�渹�X"); rsTmp("���u�s��"); rsTmp("TMS�渹")
    rsTmp.MoveNext
Loop

'�����ɮ�
Close

MsgBox "�@��X " & rsTmp.RecordCount & "���q��A��r�ɦs��C:\ShipDate.txt", 64, "�X�f�����X"
Screen.MousePointer = 0

Exit Sub
err_Handle:
Close
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub dg_Tab0_VLL_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dg_Tab0_VLL
If Len(dg.Columns(ColIndex).DataField) = 0 Then Exit Sub
SaveSetting App.title, "VLL�˸�" & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dg_Tab0_VLL_HeadClick(ByVal ColIndex As Integer)
'VLL �˸�����
'�H�ƹ��I�� dg_Tab0_VLL �����D��
Dim OrderFieldName As String
If TypeName(rs_Tab0_VLL) <> "Nothing" Then
   OrderFieldName = "[" & dg_Tab0_VLL.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_Tab0_VLL.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_Tab0_VLL.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

Private Sub dg_Tab0_VLL_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If blVLLReportEventEnable Then
   
   With dg_Tab0_VLL
        '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
        If Trim(rs_Tab0_VLL.Fields(1).Value) = "" Then
           'If rs_Tab0_VLL("�z�f�Ӽ�") = 0 Then MsgBox "�z�f�q��0�L�k���!", 64, "�`�N": Exit Sub
           rs_Tab0_VLL.Fields(1).Value = "V"
           dg_Tab0_VLL.SelBookmarks.Add rs_Tab0_VLL.Bookmark
           If rs_Tab0_VLL("�ƨ��Ӽ�") <> rs_Tab0_VLL("�z�f�Ӽ�") Then MsgBox "�ƨ��ӼƤ�����z�f�ӼƩδz�f�q��0�A�нT�{�t�m�z�f�q�O�_�����I", 16, "�`�N"
        Else
           rs_Tab0_VLL.Fields(1).Value = " "
           If dg_Tab0_VLL.SelBookmarks.Count <> 0 Then dg_Tab0_VLL.SelBookmarks.Remove 0
           
        End If
   End With
End If
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�ƨ��t�Χ@�~����"
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
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

Dim tmp_cnt As Double
'���X�Ҧ��B�e�ϰ�N�X TRP03M
cmb_Tab1_AreaCode.Clear: cmb_Tab3_AreaCode.Clear: cmb_Tab5_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_Tab1_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab3_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab5_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If
cmb_Tab1_AreaCode.ListIndex = -1
cmb_Tab3_AreaCode.ListIndex = -1
cmb_Tab5_AreaCode.ListIndex = -1
tmp_Rs.Close

cmd_Exit(0).Picture = BaseObject.cmdExit.Picture

SSTab1.Tab = 0

'��B�����u���`��G�G���ƨ����s�P�@���ƨ����s�B�q������C��
Call SetGridFormat_Tab4_RouteList
End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   fam_Tab0_Header.Left = fam_Tab0_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab0_VLL.Width = dg_Tab0_VLL.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab0_VLL.Height = dg_Tab0_VLL.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab1_Header.Left = fam_Tab1_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab1_VLLSum.Width = dg_Tab1_VLLSum.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab1_VLLSum.Height = dg_Tab1_VLLSum.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab2.Left = fam_Tab2.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab2_OrdersSum.Width = dg_Tab2_OrdersSum.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab2_OrdersSum.Height = dg_Tab2_OrdersSum.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab3.Left = fam_Tab3.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab3_PickLoadCheck.Width = dg_Tab3_PickLoadCheck.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab3_PickLoadCheck.Height = dg_Tab3_PickLoadCheck.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   SSTab2.Left = SSTab2.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   SSTab2.Top = SSTab2.Top - ((dbsrcFormHeight - Me.ScaleHeight) / 2)
   
   fam_Tab5_Header.Left = fam_Tab5_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab5_PlanList.Width = dg_Tab5_PlanList.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab5_PlanList.Height = dg_Tab5_PlanList.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab6_Header.Left = fam_Tab6_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab6_PlanList.Width = dg_Tab6_PlanList.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab6_PlanList.Height = dg_Tab6_PlanList.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   fam_Tab0_Header.Left = fam_Tab0_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab0_VLL.Width = dg_Tab0_VLL.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab0_VLL.Height = dg_Tab0_VLL.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab1_Header.Left = fam_Tab1_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab1_VLLSum.Width = dg_Tab1_VLLSum.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab1_VLLSum.Height = dg_Tab1_VLLSum.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab2.Left = fam_Tab2.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab2_OrdersSum.Width = dg_Tab2_OrdersSum.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab2_OrdersSum.Height = dg_Tab2_OrdersSum.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab3.Left = fam_Tab3.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab3_PickLoadCheck.Width = dg_Tab3_PickLoadCheck.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab3_PickLoadCheck.Height = dg_Tab3_PickLoadCheck.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   SSTab2.Left = SSTab2.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   SSTab2.Top = SSTab2.Top + ((Me.ScaleHeight - dbsrcFormHeight) / 2)
   
   fam_Tab5_Header.Left = fam_Tab5_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab5_PlanList.Width = dg_Tab5_PlanList.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab5_PlanList.Height = dg_Tab5_PlanList.Height + (Me.ScaleHeight - dbsrcFormHeight)

   fam_Tab6_Header.Left = fam_Tab6_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab6_PlanList.Width = dg_Tab6_PlanList.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab6_PlanList.Height = dg_Tab6_PlanList.Height - (dbsrcFormHeight - Me.ScaleHeight)

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
Set frm_Report_TRPPlan = Nothing

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
    Case "VLL�˸���.�X�����.�_"
         txt_Tab0_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "VLL�˸���.�X�����.��"
         txt_Tab0_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�����˸����`��.�X�����.�_"
         txt_Tab1_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�����˸����`��.�X�����.��"
         txt_Tab1_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�q���`��.�X�����.�_"
         txt_Tab2_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�q���`��.�X�����.��"
         txt_Tab2_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�z�f�˸��]�֪�.�^�Ǥ��.�_"
         txt_Tab3_UploadDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�z�f�˸��]�֪�.�^�Ǥ��.��"
         txt_Tab3_UploadDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�z�f�˸��]�֪�.�X�����.�_"
         txt_Tab3_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�z�f�˸��]�֪�.�X�����.��"
         txt_Tab3_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "��B�����u���`��.�X�����.�_"
         txt_Tab4_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "��B�����u���`��.�X�����.��"
         txt_Tab4_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�ƨ��@����.�X�����.�_"
         txt_Tab5_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�ƨ��@����.�X�����.��"
         txt_Tab5_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�̪O���@by���s.�^�Ǥ��.�_"
         txt_Tab6_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�̪O���@by���s.�^�Ǥ��.��"
         txt_Tab6_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case Else
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    mvDate.Visible = False
End Sub

Private Sub txt_Tab0_DeliveryDate_End_Click()
'VLL�˸� >> �X����� >> ��
If Trim(txt_Tab0_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab0_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab0_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "VLL�˸���.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab0_Header.Top + txt_Tab0_DeliveryDate_End.Top + txt_Tab0_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Header.Left + txt_Tab0_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab0_DeliveryDate_Start_Click()
'VLL�˸��� >> �X����� >> �_
If Trim(txt_Tab0_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab0_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab0_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "VLL�˸���.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab0_Header.Top + txt_Tab0_DeliveryDate_Start.Top + txt_Tab0_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Header.Left + txt_Tab0_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab0_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'VLL�˸��� >> �X����� >> �_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab0_DeliveryDate_Start.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_Start.Text) = 1 Then
                msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_DeliveryDate_Start.SelStart = 0: txt_Tab0_DeliveryDate_Start.SelLength = Len(txt_Tab0_DeliveryDate_Start.Text): txt_Tab0_DeliveryDate_Start.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab0_DeliveryDate_End.SelStart = 0
          txt_Tab0_DeliveryDate_End.SelLength = Len(txt_Tab0_DeliveryDate_End.Text)
          txt_Tab0_DeliveryDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab0_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'VLL�˸��� >> �X����� >> ��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab0_DeliveryDate_End.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_End.Text) = 1 Then
                msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_DeliveryDate_End.SelStart = 0: txt_Tab0_DeliveryDate_End.SelLength = Len(txt_Tab0_DeliveryDate_End.Text): txt_Tab0_DeliveryDate_End.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab0_RouteNo_Start.SelStart = 0: txt_Tab0_RouteNo_Start.SelLength = Len(txt_Tab0_RouteNo_Start.Text)
          txt_Tab0_RouteNo_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab0_RouteNo_End_KeyPress(KeyAscii As Integer)
'VLL�W�f�� >> ���u�s�� >> ��
   Select Case KeyAscii
     Case 97 To 122   '�p�g�r���אּ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab0_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab0_RouteNo_Start_KeyPress(KeyAscii As Integer)
'VLL�W�f�� >> ���u�s�� >> �_
   Select Case KeyAscii
     Case 97 To 122   '�p�g�r���אּ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab0_RouteNo_End.SelStart = 0: txt_Tab0_RouteNo_End.SelLength = Len(txt_Tab0_RouteNo_End.Text)
          txt_Tab0_RouteNo_End.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_DeliveryDate_End_Click()
'�����˸����`�� >> �X����� >> ��
If Trim(txt_Tab1_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�����˸����`��.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_End.Top + txt_Tab1_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DeliveryDate_Start_Click()
'�����˸����`�� >> �X����� >> �_
If Trim(txt_Tab1_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�����˸����`��.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_Start.Top + txt_Tab1_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab2_DeliveryDate_End_Click()
'�q���`�� >> �X����� >> ��
If Trim(txt_Tab2_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab2_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab2_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab2_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�q���`��.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab2.Top + txt_Tab2_DeliveryDate_End.Top + txt_Tab2_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab2.Left + txt_Tab2_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab2_DeliveryDate_Start_Click()
'�q���`�� >> �X����� >> �_
If Trim(txt_Tab2_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab2_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab2_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab2_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�q���`��.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab2.Top + txt_Tab2_DeliveryDate_Start.Top + txt_Tab2_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab2.Left + txt_Tab2_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab2_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�q���`�� >> �e�f��� >> �_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab2_DeliveryDate_Start.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_Start.Text) = 1 Then
                msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab2_DeliveryDate_Start.SelStart = 0: txt_Tab2_DeliveryDate_Start.SelLength = Len(txt_Tab2_DeliveryDate_Start.Text): txt_Tab2_DeliveryDate_Start.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab2_DeliveryDate_End.SelStart = 0
          txt_Tab2_DeliveryDate_End.SelLength = Len(txt_Tab2_DeliveryDate_End.Text)
          txt_Tab2_DeliveryDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab2_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'�q���`�� >> �X����� >> ��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab2_DeliveryDate_End.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab2_DeliveryDate_End.Text) = 1 Then
                msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab2_DeliveryDate_End.SelStart = 0: txt_Tab2_DeliveryDate_End.SelLength = Len(txt_Tab2_DeliveryDate_End.Text): txt_Tab2_DeliveryDate_End.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab2_RouteNo_Start.SelStart = 0: txt_Tab2_RouteNo_Start.SelLength = Len(txt_Tab2_RouteNo_Start.Text)
          txt_Tab2_RouteNo_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab2_RouteNo_End_KeyPress(KeyAscii As Integer)
'�q���`�� >> ���u�s�� >> ��
   Select Case KeyAscii
     Case 97 To 122   '�p�g�r���אּ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab2_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab2_RouteNo_Start_KeyPress(KeyAscii As Integer)
'�q���`�� >> ���u�s�� >> �_
   Select Case KeyAscii
     Case 97 To 122   '�p�g�r���אּ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab2_RouteNo_End.SelStart = 0: txt_Tab2_RouteNo_End.SelLength = Len(txt_Tab2_RouteNo_End.Text)
          txt_Tab2_RouteNo_End.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�˸����`�� >> �X����� >> �_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab1_DeliveryDate_Start.SelStart = 0: txt_Tab1_DeliveryDate_Start.SelLength = Len(txt_Tab1_DeliveryDate_Start.Text): txt_Tab1_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab1_DeliveryDate_End.SelStart = 0
             txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text)
             txt_Tab1_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab1_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'�˸����`�� >> �X����� >> ��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab1_DeliveryDate_End.SelStart = 0: txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text): txt_Tab1_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab1_Query.SetFocus
          End If
   End Select
End Sub

Private Sub txt_Tab3_UploadDate_End_Click()
'�z�f�˸��]�֪� >> �^�Ǥ�� >> ����G��
If Trim(txt_Tab3_UploadDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_UploadDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_UploadDate_End.Text, 4) & "/" & Mid(txt_Tab3_UploadDate_End.Text, 5, 2) & "/" & Right(txt_Tab3_UploadDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�z�f�˸��]�֪�.�^�Ǥ��.��"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_UploadDate_End.Top + txt_Tab3_UploadDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_UploadDate_End.Left
mvDate.Visible = True
End Sub


Private Sub txt_Tab3_UploadDate_Start_Click()
'�z�f�˸��]�֪� >> �^�Ǥ�� >> ����G�_
If Trim(txt_Tab3_UploadDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_UploadDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_UploadDate_Start.Text, 4) & "/" & Mid(txt_Tab3_UploadDate_Start.Text, 5, 2) & "/" & Right(txt_Tab3_UploadDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�z�f�˸��]�֪�.�^�Ǥ��.�_"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_UploadDate_Start.Top + txt_Tab3_UploadDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_UploadDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab3_uploadDate_Start_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �^�Ǥ�� >> ����G�_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_Start.SelStart = 0: txt_Tab3_UploadHour_Start.SelLength = Len(txt_Tab3_UploadHour_Start.Text)
          txt_Tab3_UploadHour_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab3_uploadhour_Start_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �^�Ǥ�� >> �����G�_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_Start.Text = Format(Val(txt_Tab3_UploadHour_Start.Text), "00")
          txt_Tab3_UploadMinute_Start.SelStart = 0: txt_Tab3_UploadMinute_Start.SelLength = Len(txt_Tab3_UploadMinute_Start.Text)
          txt_Tab3_UploadMinute_Start.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_uploadminute_Start_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �^�Ǥ�� >> ��ơG�_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadMinute_Start.Text = Format(Val(txt_Tab3_UploadMinute_Start.Text), "00")
          txt_Tab3_UploadDate_End.SelStart = 0: txt_Tab3_UploadDate_End.SelLength = Len(txt_Tab3_UploadDate_End.Text)
          txt_Tab3_UploadDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_uploadDate_End_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �^�Ǥ�� >> ����G��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_End.SelStart = 0: txt_Tab3_UploadHour_End.SelLength = Len(txt_Tab3_UploadHour_End.Text)
          txt_Tab3_UploadHour_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_uploadhour_End_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �^�Ǥ�� >> �����G��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadHour_End.Text = Format(Val(txt_Tab3_UploadHour_End.Text), "00")
          txt_Tab3_UploadMinute_End.SelStart = 0: txt_Tab3_UploadMinute_End.SelLength = Len(txt_Tab3_UploadMinute_End.Text)
          txt_Tab3_UploadMinute_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab3_Uploadminute_End_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �^�Ǥ�� >> ��ơG��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab3_UploadMinute_End.Text = Format(Val(txt_Tab3_UploadMinute_End.Text), "00")
          cmd_Tab3_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab3_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �X����� >> �_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_Start.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab3_DeliveryDate_Start.SelStart = 0: txt_Tab3_DeliveryDate_Start.SelLength = Len(txt_Tab3_DeliveryDate_Start.Text): txt_Tab3_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab3_DeliveryDate_End.SelStart = 0
             txt_Tab3_DeliveryDate_End.SelLength = Len(txt_Tab3_DeliveryDate_End.Text)
             txt_Tab3_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab3_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'�z�f�˸��]�֪� >> �X����� >> ��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_End.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab3_DeliveryDate_End.SelStart = 0: txt_Tab3_DeliveryDate_End.SelLength = Len(txt_Tab3_DeliveryDate_End.Text): txt_Tab3_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab3_Query.SetFocus
          End If
   End Select
End Sub

Private Sub txt_Tab3_DeliveryDate_Start_Click()
'�z�f�˸��]�֪� >> �X����� >> �_
If Trim(txt_Tab3_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�z�f�˸��]�֪�.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_DeliveryDate_Start.Top + txt_Tab3_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab3_DeliveryDate_End_Click()
'�z�f�˸��]�֪� >> �X����� >> ��
If Trim(txt_Tab3_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�z�f�˸��]�֪�.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab3.Top + txt_Tab3_DeliveryDate_End.Top + txt_Tab3_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab3.Left + txt_Tab3_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub cmd_Tab3_Reset_Click()
'�z�f�˸��]�֪� >> ���]
Set dg_Tab3_PickLoadCheck.DataSource = Nothing
Set rs_Tab3_PickLoadCheck = Nothing
cmb_Tab3_AreaCode.ListIndex = -1
txt_Tab3_UploadDate_Start.Text = ""
txt_Tab3_UploadHour_Start.Text = ""
txt_Tab3_UploadMinute_Start.Text = ""
txt_Tab3_UploadDate_End.Text = ""
txt_Tab3_UploadHour_End.Text = ""
txt_Tab3_UploadMinute_End.Text = ""
End Sub


Private Sub txt_Tab4_SecondRouteNo_KeyPress(KeyAscii As Integer)
'��B�����u���`�� >> ��ƿz�� >> �G���ƨ����u�s�� >> �_
   Select Case KeyAscii
     Case 97 To 122   '�p�g�r���אּ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab4_QueryBySRouteNo.SetFocus
   End Select
End Sub

Private Sub txt_Tab4_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'��B�����u�J�`�� >> �X����� >> �_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_Start.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab4_DeliveryDate_Start.SelStart = 0: txt_Tab4_DeliveryDate_Start.SelLength = Len(txt_Tab4_DeliveryDate_Start.Text): txt_Tab4_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab4_DeliveryDate_End.SelStart = 0
             txt_Tab4_DeliveryDate_End.SelLength = Len(txt_Tab4_DeliveryDate_End.Text)
             txt_Tab4_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab4_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'��B�����u�J�`�� >> �X����� >> ��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_End.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab4_DeliveryDate_End.SelStart = 0: txt_Tab4_DeliveryDate_End.SelLength = Len(txt_Tab4_DeliveryDate_End.Text): txt_Tab4_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab4_QueryBySRouteNo.SetFocus
          End If
   End Select
End Sub

Private Sub txt_Tab4_DeliveryDate_Start_Click()
'��B�����u�J�`�� >> �X����� >> �_
If Trim(txt_Tab4_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab4_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab4_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab4_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "��B�����u���`��.�X�����.�_"
mvDate.Top = SSTab1.Top + SSTab2.Top + txt_Tab4_DeliveryDate_Start.Top + txt_Tab4_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + SSTab2.Left + txt_Tab4_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab4_DeliveryDate_End_Click()
'��B�����u�J�`�� >> �X����� >> ��
If Trim(txt_Tab4_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab4_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab4_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab4_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab4_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "��B�����u���`��.�X�����.��"
mvDate.Top = SSTab1.Top + SSTab2.Top + txt_Tab4_DeliveryDate_End.Top + txt_Tab4_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + SSTab2.Left + txt_Tab4_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub SetGridFormat_Tab4_RouteList()
'�]�w ��B�����u���`�� �� [�����ƨ����s���] Grid �榡
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Tab4_RouteList.Visible = False
With dg_Tab4_RouteList
     .Rows = 2: .Cols = 12
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
     .ColWidth(0) = 500    '�Ǹ�
     .ColWidth(1) = 500    '�C�L�P�_
     .ColWidth(2) = 1300   '�����ƨ����s
     .ColWidth(3) = 1300   '�G���ƨ����s
     .ColWidth(4) = 1000   '�X�����
     .ColWidth(5) = 900    '���P���X
     .ColWidth(6) = 500    '����
     .ColWidth(7) = 1000   '�f�D�渹
     .ColWidth(8) = 1000   '�q����
     .ColWidth(9) = 1000   '�e�f���
     .ColWidth(10) = 1200   '�Ȥ�s��
     .ColWidth(11) = 2600   '�Ȥ�W��
     
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "�Ǹ�"
     .Col = 1: .Text = "�C�L"
     .Col = 2: .Text = "�����ƨ����s"
     .Col = 3: .Text = "�G���ƨ����s"
     .Col = 4: .Text = "�X�����"
     .Col = 5: .Text = "���P���X"
     .Col = 6: .Text = "����"
     .Col = 7: .Text = "�f�D�渹"
     .Col = 8: .Text = "�q����"
     .Col = 9: .Text = "�e�f���"
     .Col = 10: .Text = "�Ȥ�s��"
     .Col = 11: .Text = "�Ȥ�W��"
     
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignCenterCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignCenterCenter
     .ColAlignment(7) = flexAlignCenterCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .ColAlignment(10) = flexAlignCenterCenter
     .ColAlignment(11) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Tab4_RouteList.Visible = True
End Sub

Private Sub DG_TAB4_ROUTELIST_Click()
'Wave ���ݭq����
'�I�@���G����A�I�ĤG���G�������
Dim i As Double
With dg_Tab4_RouteList
     .Col = 5   'Exceed�f�D�渹
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 1
     If Len(Trim(.Text)) = 0 Then
        .Text = "��"
     Else
        .Text = ""
     End If
     .Col = 0
'     For i = 0 To .Cols - 1
'         .ColSel = i
'     Next i
End With
End Sub

Private Sub txt_Tab5_DeliveryDate_End_Click()
'�ƨ��@���� >> �X����� >> ��
If Trim(txt_Tab5_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab5_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab5_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab5_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�ƨ��@����.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab5_Header.Top + txt_Tab1_DeliveryDate_End.Top + txt_Tab5_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab5_Header.Left + txt_Tab5_DeliveryDate_End.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab5_DeliveryDate_Start_Click()
'�ƨ��@���� >> �X����� >> �_
If Trim(txt_Tab5_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab5_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab5_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab5_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�ƨ��@����.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab5_Header.Top + txt_Tab1_DeliveryDate_Start.Top + txt_Tab5_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab5_Header.Left + txt_Tab5_DeliveryDate_Start.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab5_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�ƨ��@���� >> �X����� >> �_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_Start.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab5_DeliveryDate_Start.SelStart = 0: txt_Tab5_DeliveryDate_Start.SelLength = Len(txt_Tab5_DeliveryDate_Start.Text): txt_Tab5_DeliveryDate_Start.SetFocus
             Exit Sub
          Else
             txt_Tab5_DeliveryDate_End.SelStart = 0
             txt_Tab5_DeliveryDate_End.SelLength = Len(txt_Tab5_DeliveryDate_End.Text)
             txt_Tab5_DeliveryDate_End.SetFocus
          End If
   End Select
End Sub
Private Sub txt_Tab5_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'�ƨ��@���� >> �X����� >> ��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Fun_ChkDateFormat(txt_Tab5_DeliveryDate_End.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_Tab5_DeliveryDate_End.SelStart = 0: txt_Tab5_DeliveryDate_End.SelLength = Len(txt_Tab5_DeliveryDate_End.Text): txt_Tab5_DeliveryDate_End.SetFocus
             Exit Sub
          Else
             cmd_Tab5_Query.SetFocus
          End If
   End Select
End Sub


Private Sub LLFA01Ship2TMS()
On Error GoTo err_Handle
'�^�Ǵz�f�q

str_SQL = "select o.route " & _
        ",o.storerkey " & _
        ",o.orderkey " & _
        ",o.updatesource " & _
        ",o.Externorderkey " & _
        ",ExternLineno = case when o.storerkey = 'LLFA01' then od.orderlinenumber else od.ExternLineno end " & _
        ",od.sku " & _
        ",shippedqty = (od.shippedqty + od.qtyallocated + od.qtypicked) " & _
        ",od.editdate " & _
        "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey  and convert(char(8),o.deliverydate,112) > convert(char(8),getdate()-7,112) " & _
        "where (od.shippedqty + od.qtyallocated + od.qtypicked) > 0 " & _
        "and len(rtrim(isnull(o.updatesource,''))) > 9 and o.updatesource in (select distinct receipt_no from trp03t where storerkey = 'LLFA01' and ship_qty = 0) "


Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'�L���
If Not tmp_Rs.EOF Then

    tmp_Rs.MoveFirst
    Tran_Level = cn.BeginTrans
    Do While Not tmp_Rs.EOF
    
            str_SQL = "UPDATE TRP03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "UPDATE SDN03W set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
                     "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and  SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
                     "and receipt_no ='" & tmp_Rs("updatesource") & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�g�J����
'            Call WriteLog(Err.Number & Chr(9) & "�z�f�ƶq�T�{" & Chr(9) & "WMS: " & tmp_Rs("orderkey") & ",TMS: " & tmp_Rs("route") & "," & tmp_Rs("storerkey") & "," & tmp_Rs("updatesource") & "," & RTrim(tmp_Rs("Externorderkey")) & "," & tmp_Rs("Externlineno") & "," & tmp_Rs("sku") & "," & tmp_Rs("shippedqty") & "," & User_id)
            
'            '��sYFYstatus�^�Ǫ��A
'            str_SQL = "UPDATE " & strWMSDB & "..Orders set YFYstatus = '1' ,TrafficCop = null where orderkey = '" & tmp_Rs("orderkey") & "'"
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        tmp_Rs.MoveNext
    Loop
    
    '�����X�f�q=�q��q
            str_SQL = "UPDATE TRP03T set TRP03T.SHIP_QTY=TRP03T.order_qty from trp02t join trp03t on trp02t.receipt_no = trp03t.receipt_no where trp02t.priority = 'C' and ship_qty = 0 "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
    cn.CommitTrans: Tran_Level = 0
End If

tmp_Rs.Close

Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-VLL�W�f��-�d��", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub
Private Sub txt_Tab6_DeliveryDate_End_Click()
'�̪O���@by���s >> �^�Ǥ�� >> ����G��
If Trim(txt_Tab6_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab6_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab6_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab6_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab6_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�̪O���@by���s.�^�Ǥ��.��"
mvDate.Top = SSTab1.Top + fam_Tab6_Header.Top + txt_Tab6_DeliveryDate_End.Top + txt_Tab6_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + txt_Tab6_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab6_DeliveryDate_Start_Click()
'�̪O���@by���s >> �^�Ǥ�� >> ����G�_
If Trim(txt_Tab6_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab6_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab6_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab6_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab6_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�̪O���@by���s.�^�Ǥ��.�_"
mvDate.Top = SSTab1.Top + fam_Tab6_Header.Top + txt_Tab6_DeliveryDate_Start.Top + txt_Tab6_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + txt_Tab6_DeliveryDate_Start.Left
mvDate.Visible = True
End Sub
