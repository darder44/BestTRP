VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Report_Other 
   Caption         =   "��L�ƨ�����"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11460
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   2520
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3240
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
      StartOfWeek     =   47644673
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7080
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12488
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "��L�X�f��"
      TabPicture(0)   =   "frm_Report_Other.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fam_Tab0_Header"
      Tab(0).Control(1)=   "dg_Tab0_VLL"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "��L�ƨ��@����"
      TabPicture(1)   =   "frm_Report_Other.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "dg_Tab1_VLLSum"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fam_Tab1_Header"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frm_Report_Other.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "dg_DivideSku"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   1530
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   11145
         Begin VB.TextBox Text4 
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
            Left            =   7275
            MaxLength       =   10
            TabIndex        =   57
            Top             =   1200
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox Text3 
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
            Left            =   8955
            MaxLength       =   10
            TabIndex        =   56
            Top             =   1200
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CheckBox Check1 
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
            TabIndex        =   55
            Top             =   1080
            Value           =   1  '�֨�
            Width           =   1425
         End
         Begin VB.CommandButton Command4 
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
            Picture         =   "frm_Report_Other.frx":0054
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   54
            Top             =   195
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton Command3 
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
            TabIndex        =   53
            Top             =   180
            Width           =   765
         End
         Begin VB.ComboBox cmb_Storerkey 
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
            TabIndex        =   52
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
            Index           =   2
            Left            =   9975
            Picture         =   "frm_Report_Other.frx":035E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   51
            Top             =   180
            Width           =   1065
         End
         Begin VB.TextBox DateS 
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
            TabIndex        =   50
            Top             =   615
            Width           =   1245
         End
         Begin VB.TextBox DateE 
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
            TabIndex        =   49
            Top             =   600
            Width           =   1245
         End
         Begin VB.CommandButton Command2 
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
            Left            =   6360
            Picture         =   "frm_Report_Other.frx":07A0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   48
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
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
            Picture         =   "frm_Report_Other.frx":106A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   47
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label sumlab 
            BeginProperty Font 
               Name            =   "�з���"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   67
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label sumlab 
            Caption         =   "�`�Ӽ�:"
            BeginProperty Font 
               Name            =   "�з���"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   1080
            Width           =   1215
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
            Index           =   14
            Left            =   8685
            TabIndex        =   63
            Top             =   1200
            Visible         =   0   'False
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
            Index           =   13
            Left            =   6240
            TabIndex        =   62
            Top             =   1200
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Lab_Storerkey 
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
            Left            =   135
            TabIndex        =   61
            Top             =   255
            Width           =   480
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
            Index           =   11
            Left            =   4200
            TabIndex        =   60
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
            Index           =   10
            Left            =   135
            TabIndex        =   59
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
            Index           =   9
            Left            =   2445
            TabIndex        =   58
            Top             =   690
            Width           =   240
         End
      End
      Begin VB.Frame fam_Tab0_Header 
         Height          =   1665
         Left            =   -74880
         TabIndex        =   35
         Top             =   660
         Width           =   11145
         Begin VB.TextBox txtExternOrderkeyS 
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
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1260
            Width           =   1605
         End
         Begin VB.TextBox txtExternOrderkeyE 
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
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1260
            Width           =   1605
         End
         Begin VB.TextBox txtOrderkeyS 
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
            TabIndex        =   4
            Top             =   900
            Width           =   1605
         End
         Begin VB.TextBox txtOrderkeyE 
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
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   5
            Top             =   900
            Width           =   1605
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
            Left            =   9615
            Picture         =   "frm_Report_Other.frx":1C2C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   13
            Top             =   195
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
            Left            =   6105
            Picture         =   "frm_Report_Other.frx":206E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   10
            Top             =   195
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_SaveToExcel 
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
            Left            =   7260
            Picture         =   "frm_Report_Other.frx":2938
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   11
            Top             =   195
            Width           =   1065
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
            Left            =   5055
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   8
            Top             =   525
            Width           =   630
         End
         Begin VB.CheckBox chk_Tab0_PreView 
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
            Left            =   4785
            TabIndex        =   9
            Top             =   1305
            Value           =   1  '�֨�
            Width           =   1425
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
            TabIndex        =   1
            Top             =   180
            Width           =   1245
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
            TabIndex        =   0
            Top             =   180
            Width           =   1245
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
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   3
            Top             =   555
            Width           =   1605
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
            TabIndex        =   2
            Top             =   555
            Width           =   1605
         End
         Begin VB.CommandButton cmd_Tab0_PrintReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�X�f��"
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
            Left            =   8445
            Picture         =   "frm_Report_Other.frx":34FA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   12
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��s��"
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
            TabIndex        =   44
            Top             =   1320
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
            Index           =   7
            Left            =   2790
            TabIndex        =   43
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "TMS�渹"
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
            TabIndex        =   42
            Top             =   960
            Width           =   975
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
            Index           =   5
            Left            =   2790
            TabIndex        =   41
            Top             =   960
            Width           =   240
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
            TabIndex        =   40
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
            Index           =   19
            Left            =   120
            TabIndex        =   39
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
            Index           =   20
            Left            =   2790
            TabIndex        =   38
            Top             =   615
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
            Index           =   21
            Left            =   120
            TabIndex        =   37
            Top             =   615
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
            Index           =   0
            Left            =   4020
            TabIndex        =   36
            Top             =   225
            Width           =   2010
         End
      End
      Begin VB.Frame fam_Tab1_Header 
         Height          =   1530
         Left            =   120
         TabIndex        =   28
         Top             =   660
         Width           =   11145
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
            Picture         =   "frm_Report_Other.frx":3804
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   23
            Top             =   195
            Width           =   1065
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
            Picture         =   "frm_Report_Other.frx":43C6
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   22
            Top             =   210
            Width           =   1065
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
            TabIndex        =   17
            Top             =   630
            Width           =   1245
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
            TabIndex        =   16
            Top             =   615
            Width           =   1245
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
            Picture         =   "frm_Report_Other.frx":4C90
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   25
            Top             =   180
            Width           =   1065
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
            TabIndex        =   15
            Top             =   210
            Width           =   3960
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
            TabIndex        =   20
            Top             =   180
            Width           =   765
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
            Picture         =   "frm_Report_Other.frx":50D2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
            Top             =   195
            Width           =   1065
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
            TabIndex        =   21
            Top             =   1080
            Value           =   1  '�֨�
            Width           =   1425
         End
         Begin VB.TextBox txt_Tab1_RouteNo_End 
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
            TabIndex        =   19
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab1_RouteNo_Start 
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
            TabIndex        =   18
            Top             =   1080
            Width           =   1365
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
            TabIndex        =   34
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
            Index           =   2
            Left            =   135
            TabIndex        =   33
            Top             =   660
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
            TabIndex        =   32
            Top             =   600
            Width           =   2010
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
            TabIndex        =   31
            Top             =   255
            Width           =   960
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
            TabIndex        =   30
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
            Index           =   30
            Left            =   2565
            TabIndex        =   29
            Top             =   1080
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_VLL 
         Height          =   4545
         Left            =   -74850
         TabIndex        =   14
         Top             =   2445
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8017
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         HeadLines       =   2
         RowHeight       =   20
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
      Begin MSDataGridLib.DataGrid dg_Tab1_VLLSum 
         Height          =   4710
         Left            =   150
         TabIndex        =   26
         Top             =   2280
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8308
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   120
         Top             =   -480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dg_DivideSku 
         Height          =   4950
         Left            =   -74850
         TabIndex        =   64
         Top             =   1980
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   0   'False
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
   End
   Begin VB.Label Label3 
      Caption         =   "�`�Ӽ�:"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frm_Report_Other"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private strAccessDBFileName_FullPath As String
Private MSAccessAP As access.Application

Private arAreaCode() As String
Private rsMain0 As ADODB.Recordset
Private rsMain1 As ADODB.Recordset
Private rsMain2 As ADODB.Recordset
Private rsMain2_1 As ADODB.Recordset
Private rsMain2_2 As ADODB.Recordset



Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_Tab0_PrintReport_Click()
Dim i As Integer, Tran_Level, j As Integer, strTmp As String

'����C�L
If rsMain0 Is Nothing Then MsgBox "�L��ƥi�ѦC�L�I", vbOKOnly + vbInformation, "����C�L": Exit Sub: chk_Tab0_PreView = 0

On Error GoTo err_Handle

'��Ƽg�J Access ��Ʈw
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From �h�fñ����"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Dim rs_Access As New ADODB.Recordset
rs_Access.Open "�h�fñ����", cnAccess, adOpenStatic, adLockOptimistic
rsMain0.MoveFirst
Do While Not rsMain0.EOF

   rs_Access.AddNew
   For i = 0 To rsMain0.Fields.Count - 3
   
'    If i = 15 Then'����C��15�����
'         If strTmp <> rsMain0("�h�f�渹") Then
'             j = 0: strTmp = rsMain0("�h�f�渹")
'         End If
'             j = j + 1
'             rs_Access.Fields(i).Value = j
'    Else
    
         
         rs_Access.Fields(i).Value = RTrim(rsMain0.Fields(i).Value)
'    End If
   Next i
   rs_Access.Update
   rsMain0.MoveNext
Loop
rsMain0.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'��s�C�L����
'str_SQL = "Update Ort01T Set VLListCount = " & rs_Tab0_VLLSum.Fields("�C�L����").Value & ",VLListPrintDate = '" & strPrintDate & "' " & _
'          "Where Route_No = '" & strRouteNo & "' or C_Route_No = '" & strRouteNo & "'"
'cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab0_PreView = 1 Then
   '�w���C�L
    MSAccessAP.Visible = True
    MSAccessAP.DoCmd.OpenReport "�h�fñ����", acViewPreview
    MSAccessAP.DoCmd.Maximize
   
Else
   '�����C�L�ܦL���
    MSAccessAP.Visible = False
    MSAccessAP.DoCmd.OpenReport "�h�fñ����", acViewNormal
    MSAccessAP.CloseCurrentDatabase
    MSAccessAP.Quit: Set MSAccessAP = Nothing
End If
'chk_Tab0_PreView = 0
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
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�h�fñ����-�C�L", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Query_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dg_Tab0_VLL.DataSource = Nothing
Dim chcOrderby As String, chcDeliveryDate As String, chcRoute As String, chcOrderkey As String, chcExternOrderkey As String
Dim i As Integer

str_SQL = "select �q�����O = case o2t.priority when 'RC' then '���f�J�w��' when 'A2B' then '���f�t�e��' else case when o2t.storerkey = 'LTKK01' and substring(o2t.extern,3,2) = '12' then '�h�f��(���f)' else '�h�f��' end end " & _
        ", �f�D�W�� =  (select rtrim(t16.c_name) from trp16m t16 where t16.storerkey = o2t.storerkey ) " & _
        ", ���u�s�� = o2t.route_no , �ѦҸ��s = rtrim(o.ContainerType) " & _
        ", �X����� = convert(char(8) , o1t.delivery_date , 112) " & _
        ", ���f��� = convert(char(8) , o2t.arrive_date , 112) " & _
        ", ���� = rtrim(o2t.vehicle_id_no) , �r�p = rtrim(t9m.driver) " & _
        ", TMS�渹 = o2t.receipt_no , �f�D�渹 = rtrim(o2t.extern) " & _
        ", �Ȥ�q�渹�X = rtrim(o.customerorderkey) " & _
        ", �Ȥ�W�� = rtrim(t1m.short_name) , �Ȥ�a�} = rtrim(t1m.address) ,�q�� = rtrim(t1m.phone), �Ȥ�ݨD = t1m.notes " & _
        ", ��f�Ȥ� = case when o2t.priority in ('R','RC') then '�f�e�G' + rtrim(o.facility) when len(rtrim(o.b_company)) > 0 then '�f�e�G' + rtrim(t1ma.short_name) + '-'+ rtrim(t1ma.address) + ' ' + rtrim(t1ma.phone) else '' end " & _
        ", ���� = rtrim(o3t.seq_no) , �f�� = Rtrim(o3t.Product_No)  " & _
        ", �~�W = rtrim(sp.descr) " & _
        ", �c�� =isnull(case when sp.casecnt = 0 then 0 else floor(o3t.order_qty/sp.Casecnt) end ,0) ,�j�]�� = isnull(rtrim(sp.busr3),'�c') " & _
        ", �Ӽ� =isnull(case when sp.casecnt = 0 then o3t.order_qty else cast(o3t.order_qty as int)%cast(sp.Casecnt as int) end ,0) , �p�]�� = isnull(rtrim(sp.busr1),'��') " & _
        ", �Ƶ� = case when len(cast(o.notes as varchar(1000))) > 0 or len(cast(od.notes as varchar(1000))) > 0 then cast(isnull(o.notes,'') as varchar(1000)) + '_' + cast(isnull(od.notes,'') as varchar(1000)) else ' ' end , �`�Ӽ�= o3t.order_qty " & _
        ", �ƨ��� = Case When Isnull(o1t.C_Route_No,'') = '' Then Isnull(Rtrim(o1t.AddWho),'') else Rtrim(o1t.AddWho) End , �`���n= o3t.order_qty * sp.stdcube , �`���q= o3t.order_qty * sp.stdgrosswgt " & _
        "from ort01t o1t join ort02t o2t on o1t.route_no = o2t.route_no " & _
        "join ort03t o3t on o3t.receipt_no = o2t.receipt_no " & _
        "join orders o on o.orderkey = o2t.receipt_no " & _
        "left join trp01m t1m on o2t.consigneekey = t1m.consigneekey and t1m.storerkey = o2t.storerkey " & _
        "left join trp01m t1ma on o.b_company = t1ma.consigneekey and t1ma.storerkey = o.storerkey  " & _
        "left join trp09m t9m on t9m.vehicle_id_no = o2t.vehicle_id_no " & _
        "join orderdetail od on od.orderkey = o.orderkey and od.orderlinenumber = o3t.seq_no  " & _
        "join gv_skuxpack sp on sp.sku = od.sku and sp.storerkey = od.storerkey " & _
        "where left(o2t.route_no,1) = 'R' "

chcOrderby = " "

'�X�����
chcDeliveryDate = ""
If Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   chcDeliveryDate = "and convert(char(8) , o1t.delivery_date , 112) between '" & txt_Tab0_DeliveryDate_Start.Text & "' and '" & txt_Tab0_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) = 0 Then
   chcDeliveryDate = "and convert(char(8) , o1t.delivery_date , 112) = '" & txt_Tab0_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) = 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   chcDeliveryDate = "and convert(char(8) , o1t.delivery_date , 112) = '" & txt_Tab0_DeliveryDate_End.Text & "' "
End If

'���u�s��
chcRoute = ""
If Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   chcRoute = "and o2t.route_no between '" & txt_Tab0_RouteNo_Start.Text & "' and '" & txt_Tab0_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) = 0 Then
   chcRoute = "and o2t.route_no = '" & txt_Tab0_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) = 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   chcRoute = "and o2t.route_no = '" & txt_Tab0_RouteNo_End.Text & "' "
End If

'�q�渹�X
chcOrderkey = ""
If Len(txtOrderkeyS.Text) > 0 And Len(txtOrderkeyE.Text) > 0 Then
   chcOrderkey = "and o2t.receipt_no between '" & txtOrderkeyS.Text & "' and '" & txtOrderkeyE.Text & "' "
ElseIf Len(txtOrderkeyS.Text) > 0 And Len(txtOrderkeyE.Text) = 0 Then
   chcOrderkey = "and o2t.receipt_no = '" & txtOrderkeyS.Text & "' "
ElseIf Len(txtOrderkeyS.Text) = 0 And Len(txtOrderkeyE.Text) > 0 Then
   chcOrderkey = "and o2t.receipt_no = '" & txtOrderkeyE.Text & "' "
End If

'�Ȥ�渹
chcExternOrderkey = ""
If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chcExternOrderkey = "and o2t.extern between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
   chcExternOrderkey = "and o2t.extern = '" & txtExternOrderkeyS.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chcExternOrderkey = "and o2t.extern = '" & txtExternOrderkeyE.Text & "' "
End If

'�զX�r��
str_SQL = str_SQL & chcDeliveryDate & chcRoute & chcOrderkey & chcExternOrderkey & chcOrderby

Set rsMain0 = New ADODB.Recordset
rsMain0.CursorLocation = 3
cn.CommandTimeout = 0
rsMain0.Open str_SQL, cn ', adOpenForwardOnly, adLockPessimistic
If rsMain0.EOF Then MsgBox "�d�L��ơI", 64, Me.Caption: Screen.MousePointer = 0: Exit Sub
Set dg_Tab0_VLL.DataSource = rsMain0
 
With dg_Tab0_VLL
    .ColumnHeaders = True        '���D�����
    .RowHeight = 300
'    .Columns(11).Alignment = dbgRight
'    .Columns(12).Alignment = dbgRight
'    .Columns(13).Alignment = dbgRight
'    .Columns(14).Alignment = dbgRight
'    .Columns(15).Alignment = dbgCenter
'    .Columns(18).Alignment = dbgRight
'    .Columns(19).Alignment = dbgRight

End With

rsMain0.Sort = " ���u�s�� , TMS�渹 , ���� "
SetDataGridColWidth Me.Caption, dg_Tab0_VLL
Screen.MousePointer = 0

Exit Sub
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�h�fñ����-�d��", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Reset_Click()
'VLL�W�f�� >> �M��
txt_Tab0_DeliveryDate_Start.Text = "": txt_Tab0_DeliveryDate_End.Text = ""
txtOrderkeyS.Text = "": txtOrderkeyE.Text = ""
txtExternOrderkeyS.Text = "": txtExternOrderkeyE.Text = ""
txt_Tab0_RouteNo_Start.Text = "": txt_Tab0_RouteNo_End.Text = ""
Set dg_Tab0_VLL.DataSource = Nothing
Set rsMain0 = Nothing
End Sub

Private Sub cmd_Tab0_SaveToExcel_Click()
'��ƱƧ�
Recordset2Excel "��L�ƨ�����", rsMain0

'..�b���s��EXCEL
With MyXlsApp
   
End With

Set MyXlsApp = Nothing

End Sub

Private Sub cmd_Tab1_PrintReport_Click()
Dim i As Integer, j As Integer, strTmp As String

'����C�L
If rsMain1 Is Nothing Then MsgBox "�L��ƥi�ѦC�L�I", vbOKOnly + vbInformation, "����C�L": Exit Sub

On Error GoTo err_Handle

'��Ƽg�X Access ��Ʈw
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From " & "�h�f�ƨ��@����"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Dim rs_Access As New ADODB.Recordset
rs_Access.Open "�h�f�ƨ��@����", cnAccess, adOpenStatic, adLockOptimistic
rsMain1.MoveFirst
Do While Not rsMain1.EOF
   
   rs_Access.AddNew
   For i = 0 To rsMain1.Fields.Count - 1
       rs_Access.Fields(i).Value = rsMain1.Fields(i).Value
   Next i
   rs_Access.Update
   rsMain1.MoveNext
Loop
rsMain1.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

Dim MSAccessAP As New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (App.Path & "\" & App.title & ".mdb")

'[����C�L] �R�O�s -- �Q�� Access ����
If chk_Tab1_PreView = 1 Then
   '�w���C�L
    MSAccessAP.Visible = True
    MSAccessAP.DoCmd.OpenReport "�h�f�ƨ��@����", acViewPreview
    MSAccessAP.DoCmd.Maximize
   
Else
   '�����C�L�ܦL���
    MSAccessAP.Visible = False
    MSAccessAP.DoCmd.OpenReport "�h�f�ƨ��@����", acViewNormal
    MSAccessAP.CloseCurrentDatabase
    MSAccessAP.Quit: Set MSAccessAP = Nothing
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
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�h�f�ƨ��@����-�C�L", Me.Caption, "cmd_Tab1_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Query_Click()
'�ƨ��@���� >> �d��
Set dg_Tab1_VLLSum.DataSource = Nothing
Set rsMain1 = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "Select �X�����,����,�ϰ�,���u�s��,[�G������(A)],[�G���r�p(A)],[�@������(B)],[�@���r�p(B)] " & _
          ",�B�e�O��,�B�e�c��,�B�e�Ӽ�,�B�e���q,�B�e���n,[���f�Ȥ�(A)],[��f�Ȥ�(B)],�Ƶ� " & _
          " From Report_ORTPlanList "

    
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
If Len(txt_Tab1_RouteNo_Start.Text) > 0 And Len(txt_Tab1_RouteNo_End.Text) > 0 Then
   strTmp = " Rtrim(���u�s��) between '" & txt_Tab1_RouteNo_Start.Text & "' and '" & txt_Tab1_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab1_RouteNo_Start.Text) > 0 And Len(txt_Tab1_RouteNo_End.Text) = 0 Then
   strTmp = " Rtrim(���u�s��) = '" & txt_Tab1_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab1_RouteNo_Start.Text) = 0 And Len(txt_Tab1_RouteNo_End.Text) > 0 Then
   strTmp = " Rtrim(���u�s��) = '" & txt_Tab1_RouteNo_End.Text & "' "
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
   strTmp = " �ϰ� = '" & arAreaCode(cmb_Tab1_AreaCode.ListIndex) & "' "
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
str_SQL = str_SQL & " order by �X�����,[�@������(B)],����,�ϰ� "
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rsMain1)
tmp_Rs.Close

With dg_Tab1_VLLSum
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
     rsMain1.MoveFirst
     Set .DataSource = rsMain1

    .ColumnHeaders = True          '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�X�����
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 500        '����
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 500        '�ϰ�
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 1100      '���u�s��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1200       '�G�����P���X
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1200       '�G���r�p�H
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1200       '�@�����P���X
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1200       '�@���r�p�H
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 800       '�B�e�O��
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800       '�B�e�c��
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 800       '�B�e�Ӽ�
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 800       '�B�e���q
    .Columns(12).Alignment = dbgRight
    .Columns(13).Width = 800       '�B�e���n
    .Columns(13).Alignment = dbgRight
    .Columns(14).Width = 2000       '�Ȥ�²��(A)
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 2000       '�Ȥ�²��(B)
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1200       '�Ƶ�(�G���ƨ����u�s��)
    .Columns(16).Alignment = dbgLeft
End With
rsMain1.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�h�f�ƨ��@����-�d��", Me.Caption, "cmd_Tab1_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Reset_Click()
'�h�f�ƨ��@���� >> �M��
cmb_Tab1_AreaCode.ListIndex = -1
txt_Tab1_DeliveryDate_Start.Text = ""
txt_Tab1_DeliveryDate_End.Text = ""
txt_Tab1_RouteNo_Start.Text = ""
txt_Tab1_RouteNo_End.Text = ""
chk_Tab1_PreView = 0
Set dg_Tab1_VLLSum.DataSource = Nothing
Set rsMain1 = Nothing
End Sub

Private Sub cmd_Tab1_SaveToExcel_Click()


'�ƨ��@���� >> �� EXCEL

    If rsMain1 Is Nothing Then Exit Sub
    rsMain1.MoveFirst
    On Error GoTo err_Handle

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
'    MyXlsApp.Cells(i, 1).Value = "�s��"
'    MyXlsApp.Cells(i, 2).Value = "�X�����"
'    MyXlsApp.Cells(i, 3).Value = "�ϰ�"
'    MyXlsApp.Cells(i, 4).Value = "���P���X"
'    MyXlsApp.Cells(i, 5).Value = "����"
'    MyXlsApp.Cells(i, 6).Value = "�r�p�H"
'    MyXlsApp.Cells(i, 7).Value = "���u�s��"
'    MyXlsApp.Cells(i, 8).Value = "�B�e�c��"
'    MyXlsApp.Cells(i, 9).Value = "�B�e�Ӽ�"
'    MyXlsApp.Cells(i, 10).Value = "�B�e���q"
'    MyXlsApp.Cells(i, 11).Value = "�B�e���n"
'    MyXlsApp.Cells(i, 12).Value = "�Ƶ�"
'    MyXlsApp.Cells(i, 13).Value = "�ɶ�"
'    MyXlsApp.Cells(i, 14).Value = "�Ȥ�²��"
'    MyXlsApp.Cells(i, 15).Value = "�l�ܮɶ�"
'    MyXlsApp.Cells(i, 16).Value = "�T�{"
'    MyXlsApp.Cells(i, 17).Value = "�ɥX"
'    MyXlsApp.Cells(i, 18).Value = "�^��"
'    MyXlsApp.Cells(i, 19).Value = "�j�O"
    

    MyXlsApp.Cells(i, 1).Value = "�s��"
    MyXlsApp.Cells(i, 2).Value = "�X�����"
    MyXlsApp.Cells(i, 3).Value = "����"
    MyXlsApp.Cells(i, 4).Value = "�ϰ�"
    MyXlsApp.Cells(i, 5).Value = "���u�s��"
    MyXlsApp.Cells(i, 6).Value = "�G������(A)"
    MyXlsApp.Cells(i, 7).Value = "�G���r�p(A)"
    MyXlsApp.Cells(i, 8).Value = "�@������(B)"
    MyXlsApp.Cells(i, 9).Value = "�@���r�p(B)"
    MyXlsApp.Cells(i, 10).Value = "�B�e�O��"
    MyXlsApp.Cells(i, 11).Value = "�B�e�c��"
    MyXlsApp.Cells(i, 12).Value = "�B�e�Ӽ�"
    MyXlsApp.Cells(i, 13).Value = "�B�e���q"
    MyXlsApp.Cells(i, 14).Value = "�B�e���n"
    MyXlsApp.Cells(i, 15).Value = "���f�Ȥ�(A)"
    MyXlsApp.Cells(i, 16).Value = "��f�Ȥ�(B)"
    MyXlsApp.Cells(i, 17).Value = "�Ƶ�"
    MyXlsApp.Cells(i, 18).Value = "�l�ܮɶ�"
    MyXlsApp.Cells(i, 19).Value = "�ɥX"
    MyXlsApp.Cells(i, 20).Value = "�^��"
    
    
    
    i = i + 1
    j = i
    rsMain1.MoveFirst
    '���,����,�渹,�Z�O,�ɥX,�٤J
    Do While Not rsMain1.EOF
        If i > 2 Then
            If MyXlsApp.Cells(i - 1, 4).Value <> rsMain1.Fields(8) Then
                '�������P,�j�@��b�g�Jexcel
                MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")"  '�c��
                MyXlsApp.Cells(i, 12).Value = "=SUM(L" & CStr(j) & ":L" & CStr(i - 1) & ")"  '�Ӽ�
                MyXlsApp.Cells(i, 13).Value = "=SUM(M" & CStr(j) & ":M" & CStr(i - 1) & ")" '�B�e���q
                MyXlsApp.Cells(i, 14).Value = "=SUM(N" & CStr(j) & ":N" & CStr(i - 1) & ")" '���n
                i = i + 2
                j = i
            End If
        End If

        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
        MyXlsApp.Cells(i, 2).Value = Trim(rsMain1.Fields(1)) '�X��
        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 3).Value = rsMain1.Fields(2) '����
        MyXlsApp.Cells(i, 4).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 4).Value = rsMain1.Fields(3) '�ϰ�
        MyXlsApp.Cells(i, 5).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 5).Value = rsMain1.Fields(4) '���s
        MyXlsApp.Cells(i, 6).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 6).Value = rsMain1.Fields(5) '�G������
        MyXlsApp.Cells(i, 7).Value = rsMain1.Fields(6) '�G���r�p
        MyXlsApp.Cells(i, 8).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 8).Value = rsMain1.Fields(7) '�@������
        MyXlsApp.Cells(i, 9).Value = rsMain1.Fields(8) '�@���r�p
        MyXlsApp.Cells(i, 10).Value = rsMain1.Fields(9) '�B�e�O��
        MyXlsApp.Cells(i, 11).Value = rsMain1.Fields(10) '�B�e�c��
        MyXlsApp.Cells(i, 12).Value = rsMain1.Fields(11) '�B�e�Ӽ�
        MyXlsApp.Cells(i, 13).Value = rsMain1.Fields(12) '�B�e���q
        MyXlsApp.Cells(i, 14).Value = rsMain1.Fields(13) '�B�e���n
        MyXlsApp.Cells(i, 15).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 15).Value = rsMain1.Fields(14) '�Ȥ�²��A
        MyXlsApp.Cells(i, 16).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 16).Value = rsMain1.Fields(15) '�Ȥ�²��B
        MyXlsApp.Cells(i, 17).NumberFormatLocal = "@"
        MyXlsApp.Cells(i, 17).Value = rsMain1.Fields(16) '�Ƶ�
        rsMain1.MoveNext
        i = i + 1
    Loop
                MyXlsApp.Cells(i, 11).Value = "=SUM(K" & CStr(j) & ":K" & CStr(i - 1) & ")"  '�c��
                MyXlsApp.Cells(i, 12).Value = "=SUM(L" & CStr(j) & ":L" & CStr(i - 1) & ")"  '�Ӽ�
                MyXlsApp.Cells(i, 13).Value = "=SUM(M" & CStr(j) & ":M" & CStr(i - 1) & ")" '�B�e���q
                MyXlsApp.Cells(i, 14).Value = "=SUM(N" & CStr(j) & ":N" & CStr(i - 1) & ")" '���n
    i = i + 1
    '�̾A��e
    MyXlsApp.Columns("A:T").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w
    MyXlsApp.Columns("K:N").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:T" & i - 1).Select
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
'        '�q������
'        MyXlsApp.Sheets("�ƨ��@����").Select
'        MyXlsApp.Range("I2").Select
'        MyXlsApp.ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-2])=0,"""",VLOOKUP(RC[-2],���y���!C[-5]:C[-3],2,FALSE))"
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

Private Sub Command1_Click()
'
''��ƱƧ�
'Recordset2Excel "���f���`��", rsMain2
'
''..�b���s��EXCEL
'With MyXlsApp
'
''    .Range("s:t").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
''    '�ƥ��ɮ�
''    If Dir("C:\LTKK01\�t�e���`", vbDirectory) = "" Then MkDirs "C:\LTKK01\�t�e���`"
''    .ActiveWorkbook.SaveAs "C:\LTKK01\�t�e���`\�t�e���`_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'
'End With
'
'Set MyXlsApp = Nothing
'Exit Sub

Dim SaveToExcel As Boolean
'�`��C�L
If rsMain2 Is Nothing Then Exit Sub
    If rsMain2.RecordCount = 0 Then Exit Sub
    Dim ExcelTitle As String
    Call DocStoreDirectory(strDocPath)
    Dim strTranFileName As String           'Excel �ɮצW��
    CmnDialog.DialogTitle = "��s���f���`�� Excel ��"
    CmnDialog.InitDir = "c:\my documents"
    CmnDialog.FileName = cmb_Storerkey.Text & "���f���`��_" & Format(Now, "YYYYMMDDHHNNSS")
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
    SaveToExcel = False
    On Error GoTo err_Handle
    Screen.MousePointer = vbHourglass
    If SaveTo_ExcelFile_OTHER(strTranFileName, rsMain2, "���f���`��", 1) = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
    Else
       Screen.MousePointer = vbDefault
    End If
    rsMain2.MoveFirst
    SaveToExcel = True
    
    '���f��c��

If rsMain2_1 Is Nothing Then Exit Sub
    If rsMain2_1.RecordCount = 0 Then Exit Sub
    strTranFileName = Replace(CmnDialog.FileName, "�`��", "�c��")
    SaveToExcel = False
    Screen.MousePointer = vbHourglass
    If SaveTo_ExcelFile(strTranFileName, rsMain2_1, "���f���`��_�c", 1) = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
    Else
       Screen.MousePointer = vbDefault
    End If
    rsMain2_1.MoveFirst
    SaveToExcel = True

    '���f��Ӽ�
If rsMain2_2 Is Nothing Then Exit Sub
    If rsMain2_2.RecordCount = 0 Then Exit Sub
   strTranFileName = Replace(CmnDialog.FileName, "�`��", "�Ӽ�")
    SaveToExcel = False
    Screen.MousePointer = vbHourglass
    If SaveTo_ExcelFile(strTranFileName, rsMain2_2, "���f���`��_��", 1) = 1 Then
       Screen.MousePointer = vbDefault
       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
    Else
       Screen.MousePointer = vbDefault
       If Len(strTranFileName) > 0 Then
          strTranFileName = Replace(CmnDialog.FileName, "�`��", "")
          msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
          MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       End If
    End If
    rsMain2_2.MoveFirst
    SaveToExcel = True
    Exit Sub

err_Handle:
   Dim tmpString As String
   SaveToExcel = True
   Screen.MousePointer = vbDefault
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--�� EXCEL", Me.Caption, "cmd_Tab3SavetoExcel_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub Command2_Click()
sumlab(1).Caption = ""
'�ƨ��@���� >> �d��
Set dg_DivideSku.DataSource = Nothing
Set rsMain2 = Nothing
Dim Dob_sube As Double, Int_section As Integer, Int_loc As Integer, Str_Custname As String, Int_Cloc As Integer, Str_Loc As String
Dob_sube = 0
Int_section = 65
Int_loc = 0
Int_Cloc = 0
Str_Custname = ""
Str_Loc = ""
'Dim Door_Array
'Door_Array = Array("", "RA01", "RA02", "RA03", "RA04", "RA05", "RA06", "RA07", "RA08", "RA09", "RA10", _
'                    "RB01", "RB02", "RB03", "RB04", "RB05", "RB06", "RB07", "RB08", "RB09", "RB10", "RB11", "RB12", "RB13", "RB14", _
'                    "RC01", "RC02", "RC03", "RC04", "RC05", "RC06", "RC07", "RC08", "RC09", "RC10", "RC11", "RC12", "RC13", "RC14", _
'                    "RD01", "RD02", "RD03", "RD04", "RD05", "RD06", "RD07", "RD08", "RD09", "RD10", "RD11", "RD12", "RD13", "RD14", _
'                    "RE01", "RE02", "RE03", "RE04", "RE05", "RE06", "RE07", "RE08", "RE09", "RE10", "RE11", "RE12", "RE13", "RE14", _
'                    "RF01", "RF02", "RF03", "RF04", "RF05", "RF06", "RF07", "RF08", "RF09", "RF10", _
'                    "AB03-1A", "AB04-1A", "AB05-1A", "AB06-1A", "AB07-1A""AB08-1A", "AB09-1A", "AB10-1A", "AB11-1A", "AB12-1A", "AB13-1A", _
'                    "AB14-1A", "AB15-1A", "AB16-1A", "AB17-1A", "AB18-1A""AB19-1A", "AB20-1A", "AB21-1A", "AB22-1A", "AB23-1A", "AB24-1A", _
'                    "AB25-1A", "AB26-1A", "AB27-1A", "AB28-1A", "AB29-1A""AB30-1A", "AB31-1A", "AB32-1A", "AB33-1A", "AB34-1A", _
'                    "N1", "N2", "N3", "N4", "N5", "N6", "N7", "N8", "N9", "N10", "N11", "N12", "N13", "N14", "N15", "N16", "N17", "N18", "N19", "N20", _
'                    "N21", "N22", "N23", "N24", "N25", "N26", "N27", "N28", "N29", "N30", "N31", "N32", "N33", "N34", "N35", "N36", "N37", "N38", "N39", "N40")

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
str_SQL = "exec Es_DivideSku 'LKAO01','" & DateS.Text & "','" & DateE.Text & "'"

    
Dim strWhere As String, strTmp As String

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧤��f����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rsMain2)
tmp_Rs.Close

str_SQL = "select �f�D,�X����,�Ȧs�X�Y,�Ȥ�,�~�W,�~��,���X,�c�J��,�c��,'�T�{1' = �T�{,'�T�{2' = �T�{,'�T�{3' = �T�{ from ##DivideSku where �c�� > 0 order by  �X����,�~�W,�Ȥ�,�Ȧs�X�Y"
'���f��c
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF Then
'   Screen.MousePointer = vbDefault
'   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧤��f��c���"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
'   Exit Sub
'End If
Call Replication_Recordset(tmp_Rs, rsMain2_1)
tmp_Rs.Close

'���f���
str_SQL = "select �f�D,�X����,�Ȧs�X�Y,�Ȥ�,�~�W,�~��,���X,�c�J��,�Ӽ�,'�T�{1'=�T�{,'�T�{2'=�T�{,'�T�{3' = �T�{ from ##DivideSku where �Ӽ� > 0  order by  �X����,�~�W,�Ȥ�,�Ȧs�X�Y"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF Then
'   Screen.MousePointer = vbDefault
'   msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧤��f��Ӹ��"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
'   Exit Sub
'End If
Call Replication_Recordset(tmp_Rs, rsMain2_2)
tmp_Rs.Close


str_SQL = "select �`�ƶq = sum(�`�ƶq) from ##DivideSku"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
sumlab(1).Caption = tmp_Rs.Fields("�`�ƶq")
tmp_Rs.Close
'
'rsMain2.MoveFirst
'Do While Not rsMain2.EOF
''�۰ʥH45�����@�ӼȦs�Ϧ۰ʽs��
'    If Str_Custname <> rsMain2.Fields("�Ȥ�W��") Then
'        Str_Custname = Trim(rsMain2.Fields("�Ȥ�W��"))
'        Dob_sube = Val(rsMain2.Fields("�`���n"))
'        Int_Cloc = Round(Dob_sube / 45 + 0.5)
'        Int_loc = Int_loc + Int_Cloc '�L����i��
'
''        If Int_loc >= 106 Then
''            rsMain2.Fields("�Ȧs�X�Y").Value = "N"
''            GoTo Exitif
''        End If
'
'        If Int_Cloc > 1 Then
'            '2�ӼȦs�ϡA���϶�
'            rsMain2.Fields("�Ȧs�X�Y").Value = Door_Array(Int_loc - Int_Cloc + 1) & "~" & Door_Array(Int_loc)
'            Str_Loc = Door_Array(Int_loc - Int_Cloc + 1) & "~" & Door_Array(Int_loc)
'        Else
'            '�@�ӼȦs��
'            rsMain2.Fields("�Ȧs�X�Y").Value = Door_Array(Int_loc)
'            Str_Loc = Door_Array(Int_loc)
'        End If
'
'    Else
''        If Int_loc >= 106 Then
''            rsMain2.Fields("�Ȧs�X�Y").Value = "N"
''            GoTo Exitif
''        End If
'        '�P�@�ӫȤ᩵��W�@�ӼȦs��
'         rsMain2.Fields("�Ȧs�X�Y").Value = Str_Loc
'    End If
'Exitif:
'    rsMain2.MoveNext
'Loop


With dg_DivideSku
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
     rsMain2.MoveFirst
     Set .DataSource = rsMain2

    .ColumnHeaders = True          '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1200       '�f�D
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000        '�X����
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1000       '�~��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 2000       '�~�W
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 4000       '���X
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '�Ȥ�
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1600        '�Ȧs�X�Y
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 700       '�c�J��
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 600        '�c��
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 750        '�Ӽ�
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 1100        '�`�ƶq
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 1100       '�T�{
    .Columns(12).Alignment = dbgLeft


End With
rsMain2.MoveFirst

'If Int_loc >= 106 Then
'            msg_text = "�W�X�Ȧs���x��(N�}�Y��)�A����X�ɡA��ʭק�"
'            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'End If

DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-��L�ƨ��@����-�d��", Me.Caption, "cmd_Tab1_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub DateE_Click()
'�h�f�ƨ��@���� >> �X����� >> �_
If Trim(DateE.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(DateE.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(DateE.Text, 4) & "/" & Mid(DateE.Text, 5, 2) & "/" & Right(DateE.Text, 2))
   End If
End If
mvDate.Tag = "���f��.�X�����.��"
mvDate.Top = SSTab1.Top + DateE.Top + DateE.Top + DateE.Height
mvDate.Left = SSTab1.Left + Frame1.Left + DateE.Left
mvDate.Visible = True
End Sub


Private Sub DateS_Click()
'�h�f�ƨ��@���� >> �X����� >> �_
If Trim(DateS.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(DateS.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(DateS.Text, 4) & "/" & Mid(DateS.Text, 5, 2) & "/" & Right(DateS.Text, 2))
   End If
End If
mvDate.Tag = "���f��.�X�����.�_"
mvDate.Top = SSTab1.Top + DateS.Top + DateS.Top + DateS.Height
mvDate.Left = SSTab1.Left + Frame1.Left + DateS.Left
mvDate.Visible = True

End Sub


Private Sub dg_Tab0_VLL_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dg_Tab0_VLL
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dg_Tab0_VLL_HeadClick(ByVal ColIndex As Integer)
'�h�fñ����
'�H�ƹ��I�� dg_Tab0_VLL �����D��
Dim OrderFieldName As String
If TypeName(rsMain0) <> "Nothing" Then
   OrderFieldName = "[" & dg_Tab0_VLL.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rsMain0.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rsMain0.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

'Private Sub dg_Tab1_VLLSum_HeadClick(ByVal ColIndex As Integer)
''�h�f�ƨ��@�� ��
''�H�ƹ��I�� dg_Tab0_VLL �����D��
'Dim OrderFieldName As String
'If TypeName(rsMain1) <> "Nothing" Then
'   OrderFieldName = "[" & dg_Tab1_VLLSum.Columns(ColIndex).Caption & "]"
'   If strOrder = "ASC" Then
'      strOrder = "DESC"
'      rsMain1.Sort = OrderFieldName & " DESC "
'   Else
'      strOrder = "ASC"
'      rsMain1.Sort = OrderFieldName & " ASC "
'   End If
'End If
'
'End Sub

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
cmb_Tab1_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_Tab1_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If

cmb_Storerkey.AddItem "LKAO01"
cmb_Storerkey.Text = "LKAO01"


cmb_Tab1_AreaCode.ListIndex = -1
tmp_Rs.Close
txt_Tab0_DeliveryDate_Start.Text = Format(Now, "yyyymmdd")
DateS.Text = Format(Now, "yyyymmdd")
DateE.Text = Format(Now, "yyyymmdd")
SSTab1.Tab = 0

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
   dg_DivideSku.Width = dg_DivideSku.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_DivideSku.Height = dg_DivideSku.Height - (dbsrcFormHeight - Me.ScaleHeight)
  
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
   dg_DivideSku.Width = dg_DivideSku.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_DivideSku.Height = dg_DivideSku.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
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
Set rsMain0 = Nothing
Set rsMain1 = Nothing
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
    Case "VLL�˸���.�X�����.�_"
         txt_Tab0_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "VLL�˸���.�X�����.��"
         txt_Tab0_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�h�f�ƨ��@����.�X�����.�_"
         txt_Tab1_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�h�f�ƨ��@����.�X�����.��"
         txt_Tab1_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "���f��.�X�����.�_"
        DateS.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "���f��.�X�����.��"
        DateE.Text = Format(mvDate.Value, "YYYYMMDD")
    Case Else
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call Form_KeyDown(27, 0)
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
'�h�f�ƨ��@���� >> �X����� >> ��
If Trim(txt_Tab1_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�h�f�ƨ��@����.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_End.Top + txt_Tab1_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DeliveryDate_Start_Click()
'�h�f�ƨ��@���� >> �X����� >> �_
If Trim(txt_Tab1_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�h�f�ƨ��@����.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_Start.Top + txt_Tab1_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_Start.Left
mvDate.Visible = True

End Sub

Private Sub txt_Tab1_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�h�f�ƨ��@���� >> �X����� >> �_
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab1_DeliveryDate_Start.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
                msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab1_DeliveryDate_Start.SelStart = 0: txt_Tab1_DeliveryDate_Start.SelLength = Len(txt_Tab1_DeliveryDate_Start.Text): txt_Tab1_DeliveryDate_Start.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab1_DeliveryDate_End.SelStart = 0
          txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text)
          txt_Tab1_DeliveryDate_End.SetFocus
   End Select
End Sub
Private Sub txt_Tab1_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'�h�f�ƨ��@���� >> �X����� >> ��
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          KeyAscii = 0
          If Trim(txt_Tab1_DeliveryDate_End.Text) <> "" Then
             If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
                msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab1_DeliveryDate_End.SelStart = 0: txt_Tab1_DeliveryDate_End.SelLength = Len(txt_Tab1_DeliveryDate_End.Text): txt_Tab1_DeliveryDate_End.SetFocus
                Exit Sub
             End If
          End If
          txt_Tab1_RouteNo_Start.SelStart = 0: txt_Tab1_RouteNo_Start.SelLength = Len(txt_Tab1_RouteNo_Start.Text)
          txt_Tab1_RouteNo_Start.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_RouteNo_End_KeyPress(KeyAscii As Integer)
'�h�f�ƨ��@���� >> ���u�s�� >> ��
   Select Case KeyAscii
     Case 97 To 122   '�p�g�r���אּ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Tab1_Query.SetFocus
   End Select
End Sub

Private Sub txt_Tab1_RouteNo_Start_KeyPress(KeyAscii As Integer)
'�h�f�ƨ��@���� >> ���u�s�� >> �_
   Select Case KeyAscii
     Case 97 To 122   '�p�g�r���אּ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_Tab1_RouteNo_End.SelStart = 0: txt_Tab1_RouteNo_End.SelLength = Len(txt_Tab1_RouteNo_End.Text)
          txt_Tab1_RouteNo_End.SetFocus
   End Select
End Sub

Public Function SaveTo_ExcelFile_OTHER(ByVal strFileName As String, ByRef in_rs As ADODB.Recordset, _
                Optional ByVal title As String, Optional ByVal OrientSelect As Integer) As Integer
'��s Excel ��
Dim excelAP As Excel.Application
Dim tmp_col As Double, tmp_row As Double
Dim tmp_letter As String, tmp_RangNo As String, tmpI As Integer
Dim Dob_CS As Double, Dob_EA As Double, Dob_Total As Double, Str_Sku As String
Dim bl_first As Boolean
bl_first = True
Str_Sku = ""
Dob_CS = 0
Dob_EA = 0
Dob_Total = 0

'SaveTo_ExcelFile = 1
If TypeName(in_rs) = "Nothing" Then
   funRtn_msg = "���ɿ��~�G�L��X���"
   Exit Function
ElseIf in_rs.RecordCount = 0 Then
   funRtn_msg = "���ɿ��~�G�L��X���"
   Exit Function
End If

'�]�w���檬�A Form ���
fgTransferToExcel = True
Load frm_WaitWindows
frm_WaitWindows.Tag = "Transfertoexcel"
frm_WaitWindows.ZOrder
frm_WaitWindows.Refresh
DoEvents: DoEvents

On Error GoTo err_Handle
Set excelAP = New Excel.Application
excelAP.Visible = False
excelAP.Workbooks.Add
DoEvents

'���ͲĤ@��G�H���W�ٷ���D�C
in_rs.MoveFirst
tmp_row = 1
For tmp_col = 0 To in_rs.Fields.Count - 1
    tmp_letter = Chr(65 + tmp_col)      ' A �� ascii code
    If Asc(tmp_letter) > 90 Then        ' > Z �h�ܦ� AA �_�l
       tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
    End If
        tmp_RangNo = tmp_letter & (tmp_row)
        excelAP.Range(tmp_RangNo) = in_rs.Fields(tmp_col).Name

Next tmp_col
excelAP.Range("A1", tmp_RangNo).Select
excelAP.Selection.Font.Name = "�s�ө���"
excelAP.Selection.Font.FontStyle = "����"

'�]�w�G�󭶪�������D�C�L
With excelAP.ActiveSheet.PageSetup
     .PrintTitleRows = "$1:$1"
End With

'�ۼg��Ʀ� Excel File
tmp_row = tmp_row + 1
Do While Not in_rs.EOF
    DoEvents
    '�P�_�ϥΪ̬O�_�������ɧ@�~
    If fgTransferToExcel = False Then
       err.Raise vbObjectError + 513, "Excel ���ɧ@�~", "�ϥΪ̭n�D���� Excel ���ɧ@�~�A���ɧ@�~������"
    End If
            
         If Trim(in_rs.Fields("�Ȧs�X�Y").Value) <> Str_Sku Then
            '�h�@��[�`�ƶq�A�ñN�ƶq���m�A���s�]�w�X�Y
            If bl_first = False Then
                For tmp_col = 0 To in_rs.Fields.Count - 1
                    tmp_letter = Chr(65 + tmp_col)      ' A �� ascii code
                    If Asc(tmp_letter) > 90 Then        ' > Z �h�ܦ� AA �_�l
                       tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                    End If
                        tmp_RangNo = tmp_letter & (tmp_row)
                        excelAP.Range(tmp_RangNo) = ""
                
                        With excelAP.Range(tmp_RangNo)
                            .NumberFormatLocal = "@"      '�x�s��榡 >> �Ʀr >> ���O = ��r
                            '.Font.Name = "�s�ө���"       '�x�s��榡 >> �r�� >> �r�� = Times New Roman
                            '.Font.FontStyle = "�з�"      '�x�s��榡 >> �r�� >> �~���˦� = �з�
                            '.Font.Size = 12               '�x�s��榡 >> �r�� >> �j�p = 12
                            .Font.Name = "�s�ө���"
                            .Font.FontStyle = "����"
                            .Interior.Color = RGB(173, 255, 47)
                        End With
                        
                        If Left(tmp_RangNo, 1) = "I" Then excelAP.Range(tmp_RangNo) = "�`�ơG"
                        If Left(tmp_RangNo, 1) = "J" Then excelAP.Range(tmp_RangNo) = Dob_CS
                        If Left(tmp_RangNo, 1) = "K" Then excelAP.Range(tmp_RangNo) = Dob_EA
                        If Left(tmp_RangNo, 1) = "L" Then excelAP.Range(tmp_RangNo) = Dob_Total
    
                Next tmp_col
            End If
            
            Str_Sku = Trim(in_rs.Fields("�Ȧs�X�Y").Value)
            Dob_CS = Trim(in_rs.Fields("�c��").Value): Dob_EA = Trim(in_rs.Fields("�Ӽ�").Value): Dob_Total = Trim(in_rs.Fields("�`�ƶq").Value)
            
            If bl_first = False Then
                tmp_row = tmp_row + 1
            End If
            bl_first = False
        Else
            Dob_CS = Dob_CS + Trim(in_rs.Fields("�c��").Value)
            Dob_EA = Dob_EA + Trim(in_rs.Fields("�Ӽ�").Value)
            Dob_Total = Dob_Total + Trim(in_rs.Fields("�`�ƶq").Value)
        End If
        
    For tmp_col = 0 To in_rs.Fields.Count - 1
        tmp_letter = Chr(65 + tmp_col)      ' A �� ascii code
        If Asc(tmp_letter) > 90 Then        ' > Z �h�ܦ� AA �_�l
           tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
        End If
        tmp_RangNo = tmp_letter & (tmp_row)
        '�]�w�榡
        
        With excelAP.Range(tmp_RangNo)
            .NumberFormatLocal = "@"      '�x�s��榡 >> �Ʀr >> ���O = ��r
            .Font.Name = "�s�ө���"       '�x�s��榡 >> �r�� >> �r�� = Times New Roman
            .Font.FontStyle = "�з�"      '�x�s��榡 >> �r�� >> �~���˦� = �з�
            '.Font.Size = 12               '�x�s��榡 >> �r�� >> �j�p = 12
        End With
        excelAP.Range(tmp_RangNo) = Trim(in_rs.Fields(tmp_col).Value)
    Next tmp_col
    in_rs.MoveNext
    tmp_row = tmp_row + 1
Loop

'�̫�@��
                For tmp_col = 0 To in_rs.Fields.Count - 1
                    tmp_letter = Chr(65 + tmp_col)      ' A �� ascii code
                    If Asc(tmp_letter) > 90 Then        ' > Z �h�ܦ� AA �_�l
                       tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                    End If
                        tmp_RangNo = tmp_letter & (tmp_row)
                
                        With excelAP.Range(tmp_RangNo)
                            .NumberFormatLocal = "@"      '�x�s��榡 >> �Ʀr >> ���O = ��r
                            '.Font.Name = "�s�ө���"       '�x�s��榡 >> �r�� >> �r�� = Times New Roman
                            '.Font.FontStyle = "�з�"      '�x�s��榡 >> �r�� >> �~���˦� = �з�
                            '.Font.Size = 12               '�x�s��榡 >> �r�� >> �j�p = 12
                            .Font.Name = "�s�ө���"
                            .Font.FontStyle = "����"
                            .Interior.Color = RGB(173, 255, 47)
                        End With
                
                        excelAP.Range(tmp_RangNo) = ""
                        If Left(tmp_RangNo, 1) = "I" Then excelAP.Range(tmp_RangNo) = "�`�ơG"
                        If Left(tmp_RangNo, 1) = "J" Then excelAP.Range(tmp_RangNo) = Dob_CS
                        If Left(tmp_RangNo, 1) = "K" Then excelAP.Range(tmp_RangNo) = Dob_EA
                        If Left(tmp_RangNo, 1) = "L" Then excelAP.Range(tmp_RangNo) = Dob_Total
    
                Next tmp_col
                tmp_row = tmp_row + 1
            
'���ؽu
DoEvents
'�P�_�ϥΪ̬O�_�������ɧ@�~
If fgTransferToExcel = False Then
   err.Raise vbObjectError + 513, "Excel ���ɧ@�~", "�ϥΪ̭n�D���� Excel ���ɧ@�~�A���ɧ@�~������"
End If
excelAP.Range("A1", tmp_RangNo).Select
With excelAP.Selection
     .Font.Name = "�s�ө���"
     .Font.Size = 9
     .Borders(xlEdgeLeft).LineStyle = xlContinuous
     .Borders(xlEdgeLeft).Weight = xlThin
     '.Borders(xlEdgeLeft).ColorIndex = xlAutomatic
     .Borders(xlEdgeTop).LineStyle = xlContinuous
     .Borders(xlEdgeTop).Weight = xlThin
     '.Borders(xlEdgeTop).ColorIndex = xlAutomatic
     .Borders(xlEdgeBottom).LineStyle = xlContinuous
     .Borders(xlEdgeBottom).Weight = xlThin
     '.Borders(xlEdgeBottom).ColorIndex = xlAutomatic
     .Borders(xlEdgeRight).LineStyle = xlContinuous
     .Borders(xlEdgeRight).Weight = xlThin
     '.Borders(xlEdgeRight).ColorIndex = xlAutomatic
     .Borders(xlInsideVertical).LineStyle = xlContinuous
     .Borders(xlInsideVertical).Weight = xlThin
     '.Borders(xlInsideVertical).ColorIndex = xlAutomatic
     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
     .Borders(xlInsideHorizontal).Weight = xlThin
     '.Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
End With

'�۰ʽվ���e
DoEvents
Dim str_cnStart As String, str_cnEnd As String, str_cn As String
Dim int_cn As Integer
int_cn = 1
Do   '���o���
  Select Case Mid(tmp_RangNo, int_cn, 1)
         Case 0 To 9
              Exit Do
         Case Else
              int_cn = int_cn + 1
  End Select
Loop
str_cnEnd = Mid(tmp_RangNo, 1, int_cn - 1)
str_cnStart = "A"
DoEvents
Do
   '�P�_�ϥΪ̬O�_�������ɧ@�~
   If fgTransferToExcel = False Then
      err.Raise vbObjectError + 513, "Excel ���ɧ@�~", "�ϥΪ̭n�D���� Excel ���ɧ@�~�A���ɧ@�~������"
   End If
   str_cn = str_cnStart & ":" & str_cnStart
   excelAP.Columns(str_cn).EntireColumn.AutoFit
   If str_cnStart = str_cnEnd Then
      Exit Do
   End If
   If str_cnStart = "Z" Then
      str_cnStart = "AA"
   Else
      If Len(str_cnStart) > 1 Then
         str_cnStart = "A" & Chr(Asc(Mid(str_cnStart, 2, 1)) + 1)
      Else
         str_cnStart = Chr(Asc(str_cnStart) + 1)
      End If
   End If
   DoEvents
Loop

'�O�����ɮɶ�
tmp_row = tmp_row + 1
tmp_RangNo = "A" & (tmp_row)
excelAP.Range(tmp_RangNo) = "���ɤH���G" & Get_LoginUserName
tmp_row = tmp_row + 1
tmp_RangNo = "A" & (tmp_row)
excelAP.Range(tmp_RangNo) = "�q���W�١G" & GetComputerName_rtnString
tmp_row = tmp_row + 1
tmp_RangNo = "A" & (tmp_row)
excelAP.Range(tmp_RangNo) = "���ɮɶ��G" & Format(Now, "yyyy/mm/dd hh:nn:ss")

'�ۭq����
With excelAP.ActiveSheet.PageSetup
     If Len(title) > 0 Then
        .CenterHeader = "&""�з���,����""&18" & title
     End If
     .RightFooter = "�@&""Times New Roman,�з�"" &N &""�ө���,�з�""���A&""�s�ө���,�з�""��&""Times New Roman,�з�"" &P &""�s�ө���,�з�""��"
     If OrientSelect = 1 Then
        .Orientation = xlLandscape    '��L
        .LeftMargin = excelAP.InchesToPoints(0.75)
        .RightMargin = excelAP.InchesToPoints(0.75)
        .TopMargin = excelAP.InchesToPoints(0.81)
        .BottomMargin = excelAP.InchesToPoints(0.62)
        .HeaderMargin = excelAP.InchesToPoints(0.39)
        .FooterMargin = excelAP.InchesToPoints(0.36)
     End If
End With

DoEvents
If Len(strFileName) > 0 Then
   excelAP.ActiveWorkbook.SaveAs FileName:=strFileName, FileFormat:=xlNormal
   excelAP.ActiveWindow.Close
   excelAP.Visible = False
   Set excelAP = Nothing
Else
   excelAP.Visible = True
End If
in_rs.MoveFirst

'�������檬�A Form
Unload frm_WaitWindows
Set frm_WaitWindows = Nothing
fgTransferToExcel = True

'SaveTo_ExcelFile = 0
Exit Function

err_Handle:
   fgTransferToExcel = False
   Call Release_RunningForm
   Dim tmpString As String
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   funRtn_msg = "��s excel �@�~�{�ǥ��ѡA���~�T���p�U�G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   If TypeName(excelAP) <> "Nothing" Then
      excelAP.ActiveWorkbook.Close SaveChanges:=False
      Set excelAP = Nothing
   End If
End Function

