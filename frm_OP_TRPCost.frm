VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_OP_TRPCost 
   Caption         =   "�B�O���R"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11475
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   1560
      TabIndex        =   34
      Top             =   3840
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
      StartOfWeek     =   196542465
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12488
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "�̳q���O���R"
      TabPicture(0)   =   "frm_OP_TRPCost.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dg_Tab0_Cost"
      Tab(0).Control(1)=   "fam_Tab0_Header"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "�̳f�~���O���R"
      TabPicture(1)   =   "frm_OP_TRPCost.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dg_Tab1_Cost"
      Tab(1).Control(1)=   "fam_Tab1_Header"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "UTL�˸����q�έp"
      TabPicture(2)   =   "frm_OP_TRPCost.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dg_Tab2_Cost"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "��L�B�O"
      TabPicture(3)   =   "frm_OP_TRPCost.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dg_Tab3_Cost"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "�̭q�����O���R"
      TabPicture(4)   =   "frm_OP_TRPCost.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "dgCost4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame3"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame3 
         Height          =   1410
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   11145
         Begin VB.CommandButton cmd2Excel4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
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
            Picture         =   "frm_OP_TRPCost.frx":008C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   76
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdQuery4 
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
            Picture         =   "frm_OP_TRPCost.frx":0C4E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   75
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtDeliveryDateE4 
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
            TabIndex        =   74
            Top             =   390
            Width           =   1245
         End
         Begin VB.TextBox txtDeliveryDateS4 
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
            TabIndex        =   73
            Top             =   375
            Width           =   1245
         End
         Begin VB.CommandButton cmdExit4 
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
            Index           =   4
            Left            =   9975
            Picture         =   "frm_OP_TRPCost.frx":1518
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   72
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
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
            Picture         =   "frm_OP_TRPCost.frx":195A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   71
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
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
            TabIndex        =   70
            Top             =   840
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox txtRouteE4 
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
            TabIndex        =   69
            Top             =   840
            Width           =   1365
         End
         Begin VB.TextBox txtRouteS4 
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
            TabIndex        =   68
            Top             =   840
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
            Index           =   22
            Left            =   2445
            TabIndex        =   81
            Top             =   450
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
            Index           =   17
            Left            =   135
            TabIndex        =   80
            Top             =   420
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
            Index           =   16
            Left            =   4200
            TabIndex        =   79
            Top             =   360
            Width           =   2010
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
            Index           =   15
            Left            =   120
            TabIndex        =   78
            Top             =   900
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
            Index           =   14
            Left            =   2565
            TabIndex        =   77
            Top             =   840
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1410
         Left            =   -74880
         TabIndex        =   51
         Top             =   360
         Width           =   11145
         Begin VB.TextBox txt_Tab3_RouteNo_Start 
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
            TabIndex        =   60
            Top             =   840
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab3_RouteNo_End 
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
            TabIndex        =   59
            Top             =   840
            Width           =   1365
         End
         Begin VB.CheckBox chk_Tab3_PreView 
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
            TabIndex        =   58
            Top             =   840
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.CommandButton Command6 
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
            Picture         =   "frm_OP_TRPCost.frx":1C64
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
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
            Index           =   3
            Left            =   9975
            Picture         =   "frm_OP_TRPCost.frx":1F6E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   56
            Top             =   240
            Width           =   1065
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
            Left            =   1155
            MaxLength       =   8
            TabIndex        =   55
            Top             =   375
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
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   54
            Top             =   390
            Width           =   1245
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
            Height          =   990
            Left            =   6315
            Picture         =   "frm_OP_TRPCost.frx":23B0
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   53
            Top             =   240
            Width           =   1065
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
            Height          =   990
            Left            =   7545
            Picture         =   "frm_OP_TRPCost.frx":2C7A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   52
            Top             =   240
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
            Index           =   13
            Left            =   2565
            TabIndex        =   65
            Top             =   840
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
            Index           =   12
            Left            =   120
            TabIndex        =   64
            Top             =   900
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
            Index           =   11
            Left            =   4200
            TabIndex        =   63
            Top             =   360
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
            TabIndex        =   62
            Top             =   420
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
            TabIndex        =   61
            Top             =   450
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1410
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   11145
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
            Left            =   1155
            MaxLength       =   10
            TabIndex        =   44
            Top             =   840
            Width           =   1365
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
            Left            =   2835
            MaxLength       =   10
            TabIndex        =   43
            Top             =   840
            Width           =   1365
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
            Left            =   4680
            TabIndex        =   42
            Top             =   840
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.CommandButton Command3 
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
            Picture         =   "frm_OP_TRPCost.frx":383C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
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
            Index           =   2
            Left            =   9975
            Picture         =   "frm_OP_TRPCost.frx":3B46
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   40
            Top             =   240
            Width           =   1065
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
            Left            =   1155
            MaxLength       =   8
            TabIndex        =   39
            Top             =   375
            Width           =   1245
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
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   38
            Top             =   390
            Width           =   1245
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
            Left            =   6315
            Picture         =   "frm_OP_TRPCost.frx":3F88
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   37
            Top             =   240
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
            Left            =   7545
            Picture         =   "frm_OP_TRPCost.frx":4852
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   36
            Top             =   240
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
            Index           =   8
            Left            =   2565
            TabIndex        =   49
            Top             =   840
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
            Index           =   7
            Left            =   120
            TabIndex        =   48
            Top             =   900
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
            Index           =   6
            Left            =   4200
            TabIndex        =   47
            Top             =   360
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
            Index           =   5
            Left            =   135
            TabIndex        =   46
            Top             =   420
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
            Index           =   4
            Left            =   2445
            TabIndex        =   45
            Top             =   450
            Width           =   240
         End
      End
      Begin VB.Frame fam_Tab0_Header 
         Height          =   1320
         Left            =   -74850
         TabIndex        =   17
         Top             =   330
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
            Index           =   0
            Left            =   9615
            Picture         =   "frm_OP_TRPCost.frx":5414
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   27
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
            Picture         =   "frm_OP_TRPCost.frx":5856
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   26
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
            Picture         =   "frm_OP_TRPCost.frx":6120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   25
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
            Left            =   4695
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
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
            Left            =   1080
            TabIndex        =   23
            Top             =   960
            Visible         =   0   'False
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   555
            Width           =   1605
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
            Left            =   8445
            Picture         =   "frm_OP_TRPCost.frx":6CE2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   18
            Top             =   195
            Visible         =   0   'False
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
            Index           =   18
            Left            =   2445
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   225
            Width           =   2010
         End
      End
      Begin VB.Frame fam_Tab1_Header 
         Height          =   1410
         Left            =   -74850
         TabIndex        =   1
         Top             =   360
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
            Picture         =   "frm_OP_TRPCost.frx":6FEC
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   10
            Top             =   240
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
            Picture         =   "frm_OP_TRPCost.frx":7BAE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   9
            Top             =   240
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
            TabIndex        =   8
            Top             =   390
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
            TabIndex        =   7
            Top             =   375
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
            Picture         =   "frm_OP_TRPCost.frx":8478
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   6
            Top             =   240
            Width           =   1065
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
            Picture         =   "frm_OP_TRPCost.frx":88BA
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
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
            TabIndex        =   4
            Top             =   840
            Visible         =   0   'False
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
            TabIndex        =   3
            Top             =   840
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
            TabIndex        =   2
            Top             =   840
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
            TabIndex        =   15
            Top             =   450
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
            TabIndex        =   14
            Top             =   420
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
            TabIndex        =   13
            Top             =   360
            Width           =   2010
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
            TabIndex        =   12
            Top             =   900
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
            TabIndex        =   11
            Top             =   840
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab0_Cost 
         Height          =   5265
         Left            =   -74850
         TabIndex        =   16
         Top             =   1665
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   9287
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
      Begin MSDataGridLib.DataGrid dg_Tab1_Cost 
         Height          =   5070
         Left            =   -74850
         TabIndex        =   33
         Top             =   1860
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8943
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
         Top             =   -480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dg_Tab2_Cost 
         Height          =   5070
         Left            =   -74880
         TabIndex        =   50
         Top             =   1860
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8943
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
      Begin MSDataGridLib.DataGrid dg_Tab3_Cost 
         Height          =   5070
         Left            =   -74880
         TabIndex        =   66
         Top             =   1860
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8943
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
      Begin MSDataGridLib.DataGrid dgCost4 
         Height          =   5070
         Left            =   120
         TabIndex        =   82
         Top             =   1860
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8943
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
   End
End
Attribute VB_Name = "frm_OP_TRPCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private iLoop As Double

Private rs_Tab0_Cost As ADODB.Recordset
Private rs_Tab1_Cost As ADODB.Recordset
Private rs_Tab2_Cost As ADODB.Recordset
Private rs_Tab3_Cost As ADODB.Recordset
Private rsCost4 As ADODB.Recordset
Private Sub cmd_Exit_Click(Index As Integer)
    '���}
    Unload Me
End Sub

Private Sub cmd_Tab0_Query_Click()
'�̳q���O�B�O���R >>�d��
Set dg_Tab0_Cost.DataSource = Nothing
Set rs_Tab0_Cost = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_handle

str_SQL = "select  isnull(m1.channel,'�L') as �q���O,o.Priority as ���O,d2.extern as �q�渹�X,m1.AREA_CODE as �a��,d2.CUST_NAME as �Ȥ�W��,d2.ship_wt as �q�歫�q, " & _
        "(select sum(ship_wt) from SDN02T where C_Route_No=d2.C_Route_No) as �`���q, " & _
        "(select isnull(sum(sumreceivable),0) from SDN05T where C_Route_No=d2.C_Route_No) as �`�B�O, " & _
        "round((d2.ship_wt/(select sum(ship_wt) from SDN02T where C_Route_No=d2.C_Route_No))*(select isnull(sum(sumreceivable),0) from SDN05T where C_Route_No=d2.C_Route_No),3) as ���u�B�O, " & _
        "o.ConsigneeKey as �Ȥ�s��,d2.C_ROUTE_NO as ���u�s�� " & _
        "from SDN02T d2 " & _
        "inner join orders o on o.ExternOrderKey=d2.extern " & _
        "inner  join  trp01m  m1  on  o.ConsigneeKey=m1.ConsigneeKey "
                
Dim strWhere As String, strTmp As String
strWhere = ""
'�X�����
strTmp = ""
If Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),d2.ARRIVE_DATE,112) between '" & txt_Tab0_DeliveryDate_Start.Text & "' and '" & txt_Tab0_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) > 0 And Len(txt_Tab0_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(varchar(8),d2.ARRIVE_DATE,112) = '" & txt_Tab0_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab0_DeliveryDate_Start.Text) = 0 And Len(txt_Tab0_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),d2.ARRIVE_DATE,112) = '" & txt_Tab0_DeliveryDate_End.Text & "' "
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
   strTmp = " d2.c_route_no between '" & txt_Tab0_RouteNo_Start.Text & "' and '" & txt_Tab0_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) = 0 Then
   strTmp = " d2.c_route_no = '" & txt_Tab0_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab0_RouteNo_Start.Text) = 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
   strTmp = " d2.c_route_no = '" & txt_Tab0_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If Len(strWhere) > 0 Then
   strWhere = strWhere & " and (select sum(ship_wt) from SDN02T where C_Route_No=d2.C_Route_No)>0 order by isnull(m1.channel,'�L')"
End If
If strWhere <> "" Then
    str_SQL = str_SQL & "where" & strWhere
Else
    msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If

cmd_Tab0_Query.Enabled = False

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
cn.CommandTimeout = 0
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    cmd_Tab0_Query.Enabled = True
    tmp_Rs.Close
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0_Cost)
tmp_Rs.Close

With dg_Tab0_Cost
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab0_Cost.MoveFirst
Set dg_Tab0_Cost.DataSource = rs_Tab0_Cost
With dg_Tab0_Cost
    .ColumnHeaders = True         '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500       '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 800       '�q���O
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 500       '���O
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '�q�渹�X
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500       '�a��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 2000      '�Ȥ�W��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1000      '�q�歫�q
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 1000       '�`���q
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 1000       '�`�B�O
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 1000      '���u�B�O
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 1000      '�Ȥ�s��
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1000     '���u�s��
    .Columns(11).Alignment = dbgLeft
End With
rs_Tab0_Cost.MoveFirst
rs_Tab0_Cost.Filter = adFilterNone
rs_Tab0_Cost.Sort = " �s�� "
rs_Tab0_Cost.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
cmd_Tab0_Query.Enabled = True
Exit Sub

err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�̳q���O�B�O���R >>-�d��", Me.Caption, "cmd_Tab0_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmd_Tab0_Reset_Click()
    txt_Tab0_DeliveryDate_Start = ""
    txt_Tab0_DeliveryDate_End = ""
    txt_Tab0_RouteNo_Start = ""
    txt_Tab0_RouteNo_End = ""
End Sub

Private Sub cmd_Tab0_SaveToExcel_Click()
    'ñ��d�� >> �� EXCEL
    If rs_Tab0_Cost Is Nothing Then Exit Sub
    Screen.MousePointer = 11
    rs_Tab0_Cost.MoveFirst
    On Error GoTo err_handle
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
'    MyXlsApp.Sheets("Sheet1").Name = "�̳q���O���R�B�O"
    MyXlsApp.ActiveSheet.Name = "�̳q���O���R�B�O"
'    i = 1
'    '�q���O,�q�渹�X,�q�歫�q,�`���q,�`�B�O,���u�B�O
'    MyXlsApp.Cells(i, 1).Value = "�q���O"
'    MyXlsApp.Cells(i, 2).Value = "���O"
'    MyXlsApp.Cells(i, 3).Value = "�q�渹�X"
'    MyXlsApp.Cells(i, 4).Value = "�a��"
'    MyXlsApp.Cells(i, 5).Value = "�Ȥ�W��"
'    MyXlsApp.Cells(i, 6).Value = "�q�歫�q"
'    MyXlsApp.Cells(i, 7).Value = "�`���q"
'    MyXlsApp.Cells(i, 8).Value = "�`�B�O"
'    MyXlsApp.Cells(i, 9).Value = "���u�B�O"
'    MyXlsApp.Cells(i, 10).Value = "�Ȥ�s��"
'    MyXlsApp.Cells(i, 11).Value = "���u�s��"
'    i = i + 1
'    rs_Tab0_Cost.MoveFirst
'    Do While Not rs_Tab0_Cost.EOF
'        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab0_Cost.Fields(1))
'        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab0_Cost.Fields(2))
'        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab0_Cost.Fields(3))
'        MyXlsApp.Cells(i, 4).Value = Trim(rs_Tab0_Cost.Fields(4))
'        MyXlsApp.Cells(i, 5).Value = rs_Tab0_Cost.Fields(5)
'        MyXlsApp.Cells(i, 6).Value = rs_Tab0_Cost.Fields(6)
'        MyXlsApp.Cells(i, 7).Value = rs_Tab0_Cost.Fields(7)
'        MyXlsApp.Cells(i, 8).Value = rs_Tab0_Cost.Fields(8)
'        MyXlsApp.Cells(i, 9).Value = rs_Tab0_Cost.Fields(9)
'        MyXlsApp.Cells(i, 10).Value = rs_Tab0_Cost.Fields(10)
'        MyXlsApp.Cells(i, 11).Value = rs_Tab0_Cost.Fields(10)
'        rs_Tab0_Cost.MoveNext
'        i = i + 1
'    Loop
'    i = i + 1

'���D�C
For i = 1 To rs_Tab0_Cost.Fields.Count - 1
MyXlsApp.Range("A1").Value = "����"
MyXlsApp.Range(Chr(65 + i) & "1").Value = rs_Tab0_Cost(i).Name
Next i

'��Ƽg�J
MyXlsApp.Range("A2").CopyFromRecordset rs_Tab0_Cost

    '�̾A��e
    MyXlsApp.Columns("A:L").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w
    MyXlsApp.Columns("G:J").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:L" & rs_Tab0_Cost.RecordCount + 1).Select
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
    
    '�p�p
    MyXlsApp.Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(10), Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    MyXlsApp.ActiveSheet.Outline.ShowLevels RowLevels:=2
    
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�B�O���R-�� EXCEL", Me.Caption, "cmd_Tab1_SaveToExcel_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmd_Tab1_Query_Click()
'�̳f���B�O���R >>�d��
Set dg_Tab1_Cost.DataSource = Nothing
Set rs_Tab1_Cost = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_handle

str_SQL = "select d3.EXTERN as �渹,d3.PRODUCT_NO as �f��, isnull(rtrim(s.BUSR1),'�L') as ���~�O,isnull(d3.SHIP_QTY,0)*s.STDGROSSWGT as �f�~���q, " & _
        "(select sum(ship_wt) from SDN02T where C_Route_No=d3.C_Route_No) as �`���q, " & _
        "(select isnull(sum(sumreceivable),0) from SDN05T where C_Route_No=d3.C_Route_No) as �`�B�O, " & _
        "(isnull(d3.SHIP_QTY,0)*s.STDGROSSWGT/(select sum(ship_wt) from SDN02T where C_Route_No=d3.C_Route_No))*(select isnull(sum(sumreceivable),0) from SDN05T where C_Route_No=d3.C_Route_No) as ���u�B�O " & _
        "from SDN03T d3 " & _
        "inner join SDN02T  d2 on d3.C_Route_No=d2.C_Route_No and d2.EXTERN=d3.EXTERN " & _
        "inner join sku s on s.sku=d3.PRODUCT_NO and s.storerkey = d3.storerkey "
                
Dim strWhere As String, strTmp As String
strWhere = ""
'�X�����
strTmp = ""
If Len(txt_Tab1_DeliveryDate_Start.Text) > 0 And Len(txt_Tab1_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),d2.ARRIVE_DATE,112) between '" & txt_Tab1_DeliveryDate_Start.Text & "' and '" & txt_Tab1_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab1_DeliveryDate_Start.Text) > 0 And Len(txt_Tab1_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(varchar(8),d2.ARRIVE_DATE,112) = '" & txt_Tab1_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab1_DeliveryDate_Start.Text) = 0 And Len(txt_Tab1_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),d2.ARRIVE_DATE,112) = '" & txt_Tab1_DeliveryDate_End.Text & "' "
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
   strTmp = " d2.c_route_no between '" & txt_Tab1_RouteNo_Start.Text & "' and '" & txt_Tab1_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab1_RouteNo_Start.Text) > 0 And Len(txt_Tab1_RouteNo_End.Text) = 0 Then
   strTmp = " d2.c_route_no = '" & txt_Tab1_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab1_RouteNo_Start.Text) = 0 And Len(txt_Tab1_RouteNo_End.Text) > 0 Then
   strTmp = " d2.c_route_no = '" & txt_Tab1_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If Len(strWhere) > 0 Then
   strWhere = strWhere & " and (select sum(ship_wt) from SDN02T where C_Route_No=d3.C_Route_No)>0 order by d3.PRODUCT_NO"
End If
If strWhere <> "" Then
    str_SQL = str_SQL & "where" & strWhere
Else
    msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
cn.CommandTimeout = 0
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab1_Cost)
tmp_Rs.Close

With dg_Tab1_Cost
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab1_Cost.MoveFirst
Set dg_Tab1_Cost.DataSource = rs_Tab1_Cost
With dg_Tab1_Cost
    .ColumnHeaders = True         '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500       '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�q�渹�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '�f��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '���~�O
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000      '�f�~���q
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 1000      '�`���q
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 1000       '�`�B�O
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 1500       '���u�B�O
    .Columns(7).Alignment = dbgRight
End With
rs_Tab1_Cost.MoveFirst
rs_Tab1_Cost.Filter = adFilterNone
rs_Tab1_Cost.Sort = " �s�� "
rs_Tab1_Cost.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�̳f���B�O���R >>-�d��", Me.Caption, "cmd_Tab1_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_SaveToExcel_Click()
    '�̳f���B�O���R >> �� EXCEL
    If rs_Tab1_Cost Is Nothing Then Exit Sub
    rs_Tab1_Cost.MoveFirst
    On Error GoTo err_handle
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
'    MyXlsApp.Sheets("Sheet1").Name = "�̳q���O���R�B�O"
    MyXlsApp.ActiveSheet.Name = "�̳q���O���R�B�O"
    i = 1
    '�Ǹ�'�q�渹�X'�f��'���~�O'�f�~���q '�`���q'�`�B�O '���u�B�O
    MyXlsApp.Cells(i, 1).Value = "�q�渹�X"
    MyXlsApp.Cells(i, 2).Value = "�f��"
    MyXlsApp.Cells(i, 3).Value = "���~�O"
    MyXlsApp.Cells(i, 4).Value = "�f�~���q"
    MyXlsApp.Cells(i, 5).Value = "�`���q"
    MyXlsApp.Cells(i, 6).Value = "�`�B�O"
    MyXlsApp.Cells(i, 7).Value = "���u�B�O"
    i = i + 1
    rs_Tab1_Cost.MoveFirst
    Do While Not rs_Tab1_Cost.EOF
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab1_Cost.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab1_Cost.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab1_Cost.Fields(3))
        MyXlsApp.Cells(i, 4).Value = Trim(rs_Tab1_Cost.Fields(4))
        MyXlsApp.Cells(i, 5).Value = rs_Tab1_Cost.Fields(5)
        MyXlsApp.Cells(i, 6).Value = rs_Tab1_Cost.Fields(6)
        MyXlsApp.Cells(i, 7).Value = rs_Tab1_Cost.Fields(7)
        rs_Tab1_Cost.MoveNext
        i = i + 1
    Loop
    i = i + 1
    '�̾A��e
    MyXlsApp.Columns("A:G").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w
    MyXlsApp.Columns("D:G").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:G" & i - 1).Select
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
    
    '�p�p
    MyXlsApp.Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(7), Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    MyXlsApp.ActiveSheet.Outline.ShowLevels RowLevels:=2
    
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�B�O���R-�� EXCEL", Me.Caption, "cmd_Tab1_SaveToExcel_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmd_Tab2_Query_Click()
'�̳f���B�O���R >>�d��
Set dg_Tab2_Cost.DataSource = Nothing
Set rs_Tab2_Cost = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_handle

str_SQL = "select o.ConsigneeKey,o.Priority,t2.extern as �q�渹�X,m1.AREA_CODE as �a��,m1.FULL_NAME as �Ȥ�W��,t2.WEIGHT as �q�歫�q  " & _
        "from TRP02T t2 " & _
        "inner join TRP05T t5 on t2.ROUTE_NO=t5.ROUTE_NO " & _
        "inner join orders o on o.ExternOrderKey=t2.extern " & _
        "inner  join  trp01m  m1  on  o.ConsigneeKey=m1.ConsigneeKey "
'where t5.Receiver ='�D�g��' and Convert(varchar,t2.ARRIVE_DATE,112) between '20050904' and '20051001'"
                
Dim strWhere As String, strTmp As String
strWhere = ""
'�X�����
strTmp = ""
If Len(txt_Tab2_DeliveryDate_Start.Text) > 0 And Len(txt_Tab2_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),t2.ARRIVE_DATE,112) between '" & txt_Tab2_DeliveryDate_Start.Text & "' and '" & txt_Tab2_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab2_DeliveryDate_Start.Text) > 0 And Len(txt_Tab2_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(varchar(8),t2.ARRIVE_DATE,112) = '" & txt_Tab2_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab2_DeliveryDate_Start.Text) = 0 And Len(txt_Tab2_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),t2.ARRIVE_DATE,112) = '" & txt_Tab2_DeliveryDate_End.Text & "' "
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
   strTmp = " t2.c_route_no between '" & txt_Tab2_RouteNo_Start.Text & "' and '" & txt_Tab2_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab2_RouteNo_Start.Text) > 0 And Len(txt_Tab2_RouteNo_End.Text) = 0 Then
   strTmp = " t2.c_route_no = '" & txt_Tab2_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab2_RouteNo_Start.Text) = 0 And Len(txt_Tab2_RouteNo_End.Text) > 0 Then
   strTmp = " t2.c_route_no = '" & txt_Tab2_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If Len(strWhere) > 0 Then
   strWhere = strWhere & " and t5.Receiver ='�D�g��' order by t2.extern"
End If
If strWhere <> "" Then
    str_SQL = str_SQL & " where " & strWhere
Else
    msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
cn.CommandTimeout = 0
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab2_Cost)
tmp_Rs.Close

With dg_Tab2_Cost
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab2_Cost.MoveFirst
Set dg_Tab2_Cost.DataSource = rs_Tab2_Cost
With dg_Tab2_Cost
    .ColumnHeaders = True         '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500       '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       'ConsigneeKey
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       'Priority
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '�q�渹�X
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000      '�a��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1000      '�Ȥ�W��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1000       '�q�歫�q
    .Columns(6).Alignment = dbgRight
End With
rs_Tab2_Cost.MoveFirst
rs_Tab2_Cost.Filter = adFilterNone
rs_Tab2_Cost.Sort = " �s�� "
rs_Tab2_Cost.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-UTL�˸����q >>-�d��", Me.Caption, "cmd_Tab2_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd2Excel4_Click()
Recordset2Excel "�̭q�����O���R", rsCost4
Set MyXlsApp = Nothing
End Sub

Private Sub cmdExit4_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmd_Tab2_SaveToExcel_Click()
    '�̳f���B�O���R >> �� EXCEL
    If rs_Tab2_Cost Is Nothing Then Exit Sub
    rs_Tab2_Cost.MoveFirst
    On Error GoTo err_handle
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
'    MyXlsApp.Sheets("Sheet1").Name = "UTL�˸����q"
    MyXlsApp.ActiveSheet.Name = "UTL�˸����q"
    i = 1
    '�Ǹ�'ConsigneeKey,Priority,�q�渹�X ,�a��, �Ȥ�W��, �q�歫�q
    MyXlsApp.Cells(i, 1).Value = "�Ȥ�s��"
    MyXlsApp.Cells(i, 2).Value = "���O"
    MyXlsApp.Cells(i, 3).Value = "�q�渹�X"
    MyXlsApp.Cells(i, 4).Value = "�a��"
    MyXlsApp.Cells(i, 5).Value = "�Ȥ�W��"
    MyXlsApp.Cells(i, 6).Value = "�q�歫�q"
 
    i = i + 1
    rs_Tab2_Cost.MoveFirst
    Do While Not rs_Tab2_Cost.EOF
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab2_Cost.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab2_Cost.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab2_Cost.Fields(3))
        MyXlsApp.Cells(i, 4).Value = Trim(rs_Tab2_Cost.Fields(4))
        MyXlsApp.Cells(i, 5).Value = rs_Tab2_Cost.Fields(5)
        MyXlsApp.Cells(i, 6).Value = rs_Tab2_Cost.Fields(6)
        rs_Tab2_Cost.MoveNext
        i = i + 1
    Loop
    i = i + 1
    '�̾A��e
    MyXlsApp.Columns("A:F").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w
    MyXlsApp.Columns("D:F").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:G" & i - 1).Select
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
    Exit Sub
    
err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�B�O���R-�� EXCEL", Me.Caption, "cmd_Tab2_SaveToExcel_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_Query_Click()
'�̳f���B�O���R >>�d��
Set dg_Tab3_Cost.DataSource = Nothing
Set rs_Tab3_Cost = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_handle

str_SQL = "select   d5.SDN_No,d5.SDN_Name,d5.CostKind,d5.AreaStart,d5.AreaEnd,d5.SumReceivable " & _
        "from SDN05T d5 " & _
        "inner join SDN01T d1 on d5.C_ROUTE_NO=d1.C_ROUTE_NO "
'where  Convert(varchar,d1.DELIVERY_DATE,112)  between '20050904' and '20051001'
'and left(d1.C_ROUTE_NO,2)='WD'

                
Dim strWhere As String, strTmp As String
strWhere = ""
'�X�����
strTmp = ""
If Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),d1.DELIVERY_DATE,112) between '" & txt_Tab3_DeliveryDate_Start.Text & "' and '" & txt_Tab3_DeliveryDate_End.Text & "' "
ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) = 0 Then
   strTmp = " Convert(varchar(8),d1.DELIVERY_DATE,112) = '" & txt_Tab3_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) = 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
   strTmp = " Convert(varchar(8),d1.DELIVERY_DATE,112) = '" & txt_Tab3_DeliveryDate_End.Text & "' "
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
If Len(txt_Tab3_RouteNo_Start.Text) > 0 And Len(txt_Tab3_RouteNo_End.Text) > 0 Then
   strTmp = " d1.c_route_no between '" & txt_Tab3_RouteNo_Start.Text & "' and '" & txt_Tab3_RouteNo_End.Text & "' "
ElseIf Len(txt_Tab3_RouteNo_Start.Text) > 0 And Len(txt_Tab3_RouteNo_End.Text) = 0 Then
   strTmp = " d1.c_route_no = '" & txt_Tab3_RouteNo_Start.Text & "' "
ElseIf Len(txt_Tab3_RouteNo_Start.Text) = 0 And Len(txt_Tab3_RouteNo_End.Text) > 0 Then
   strTmp = " d1.c_route_no = '" & txt_Tab3_RouteNo_End.Text & "' "
End If
If Len(strTmp) > 0 Then
   If Len(strWhere) > 0 Then
      strWhere = strWhere & " and " & strTmp
   Else
      strWhere = strWhere & strTmp
   End If
End If

If Len(strWhere) > 0 Then
   strWhere = strWhere & " and left(d1.C_ROUTE_NO,2)='WD' "
End If
If strWhere <> "" Then
    str_SQL = str_SQL & " where " & strWhere
Else
    msg_text = "����Y�p�d�߸�ƶq�A�оA�׳]�w�d�߱���"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
cn.CommandTimeout = 0
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab3_Cost)
tmp_Rs.Close

With dg_Tab3_Cost
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab3_Cost.MoveFirst
Set dg_Tab3_Cost.DataSource = rs_Tab3_Cost
'SDN_No  ,SDN_Name ,CostKind ,AreaStart ,AreaEnd , SumReceivable
With dg_Tab3_Cost
    .ColumnHeaders = True         '���D�����
    .RowHeight = 250
    .Columns(0).Width = 500       '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       'SDN_No
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       'SDN_Name
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       'CostKind
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000      'AreaStart
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1000      'AreaEnd
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1000       '�B�O
    .Columns(6).Alignment = dbgRight
End With
rs_Tab3_Cost.MoveFirst
rs_Tab3_Cost.Filter = adFilterNone
rs_Tab3_Cost.Sort = " �s�� "
rs_Tab3_Cost.MoveFirst
DoEvents: DoEvents
Screen.MousePointer = vbDefault
Exit Sub

err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-UTL�˸����q >>-�d��", Me.Caption, "cmd_Tab3_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_SaveToExcel_Click()
    '�̳f���B�O���R >> �� EXCEL
    If rs_Tab3_Cost Is Nothing Then Exit Sub
    rs_Tab3_Cost.MoveFirst
    On Error GoTo err_handle
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
'    MyXlsApp.Sheets("Sheet1").Name = "��L�B�O"
    MyXlsApp.ActiveSheet.Name = "��L�B�O"
    i = 1
    'SDN_No  ,SDN_Name ,CostKind ,AreaStart ,AreaEnd , SumReceivable
    MyXlsApp.Cells(i, 1).Value = "�q�渹�X"
    MyXlsApp.Cells(i, 2).Value = "�Ȥ�W��"
    MyXlsApp.Cells(i, 3).Value = "�д����O"
    MyXlsApp.Cells(i, 4).Value = "�_�I"
    MyXlsApp.Cells(i, 5).Value = "���I"
    MyXlsApp.Cells(i, 6).Value = "�B�O"
 
    i = i + 1
    rs_Tab3_Cost.MoveFirst
    Do While Not rs_Tab3_Cost.EOF
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab3_Cost.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab3_Cost.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab3_Cost.Fields(3))
        MyXlsApp.Cells(i, 4).Value = Trim(rs_Tab3_Cost.Fields(4))
        MyXlsApp.Cells(i, 5).Value = rs_Tab3_Cost.Fields(5)
        MyXlsApp.Cells(i, 6).Value = rs_Tab3_Cost.Fields(6)
        rs_Tab3_Cost.MoveNext
        i = i + 1
    Loop
    i = i + 1
    '�̾A��e
    MyXlsApp.Columns("A:F").Select
    MyXlsApp.Selection.Columns.AutoFit
    
    '�x�s��榡�]�w
    MyXlsApp.Columns("D:F").Select
    MyXlsApp.Selection.NumberFormatLocal = "0.00_ "
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:G" & i - 1).Select
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
    Exit Sub
    
err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�B�O���R-�� EXCEL", Me.Caption, "cmd_Tab3_SaveToExcel_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery4_Click()
'�̭q�����O���R >>�d��

If Len(txtDeliveryDateS4.Text) + Len(txtDeliveryDateS4.Text) + Len(txtRouteS4.Text) + Len(txtRouteE4.Text) = 0 Then MsgBox "�Цܤֿ�J�@�Ӭd�߱���!", vbOKOnly, Me.Caption: Exit Sub
Set dgCost4.DataSource = Nothing
Set rsCost4 = Nothing

Screen.MousePointer = vbHourglass
On Error GoTo err_handle

str_SQL = "exec gs_costxcostkind '" & txtDeliveryDateS4.Text & "','" & txtDeliveryDateE4.Text & "','" & txtRouteS4.Text & "','" & txtRouteE4.Text & "'"

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
cn.CommandTimeout = 0
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����!"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If

tmp_Rs.Sort = "�X�����,���u�s��"
Call Replication_Recordset(tmp_Rs, rsCost4)
tmp_Rs.Close

rsCost4.MoveFirst

Set dgCost4.DataSource = rsCost4
arColCenter = Array(0, 6)
With dgCost4
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
     .Columns(0).Alignment = dbgCenter
     .Columns(6).Alignment = dbgCenter
     .Columns(7).Alignment = dbgRight
     .Columns(12).Alignment = dbgRight
     .Columns(13).Alignment = dbgRight
     .Columns(14).Alignment = dbgRight
End With
SetDataGridColWidth Me.Caption, dgCost4
DoEvents: DoEvents
Screen.MousePointer = vbDefault

Exit Sub

err_handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�̭q���������R >>-�d��", Me.Caption, "cmd_Tab3_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub dgCost4_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & "dgCost4", dgCost4.Columns(ColIndex).DataField, dgCost4.Columns(ColIndex).Width
End Sub

Private Sub dgCost4_HeadClick(ByVal ColIndex As Integer)
With dgCost4

    If .Row = -1 Then Exit Sub
    If intColIndex = ColIndex Then
        rsCost4.Sort = .Columns(ColIndex).Caption & " DESC"
        .ClearSelCols
        intColIndex = 255
    
    Else
        rsCost4.Sort = .Columns(ColIndex).Caption
        .ClearSelCols
        intColIndex = ColIndex
    
    End If

End With
End Sub

Private Sub Form_Activate()
    '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
    Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
    msg_title = "�B�O���R"
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

End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   fam_Tab0_Header.Left = fam_Tab0_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab0_Cost.Width = dg_Tab0_Cost.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab0_Cost.Height = dg_Tab0_Cost.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dgCost4.Width = dgCost4.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dgCost4.Height = dgCost4.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab1_Header.Left = fam_Tab1_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Tab1_Cost.Width = dg_Tab1_Cost.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Tab1_Cost.Height = dg_Tab1_Cost.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   fam_Tab0_Header.Left = fam_Tab0_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab0_Cost.Width = dg_Tab0_Cost.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab0_Cost.Height = dg_Tab0_Cost.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab1_Header.Left = fam_Tab1_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Tab1_Cost.Width = dg_Tab1_Cost.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Tab1_Cost.Height = dg_Tab1_Cost.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dgCost4.Width = dgCost4.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dgCost4.Height = dgCost4.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
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
    Case "�̳q���O���R.�X�����.�_"
         txt_Tab0_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�̳q���O���R.�X�����.��"
         txt_Tab0_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�̳f���O���R.�X�����.�_"
         txt_Tab1_DeliveryDate_Start.Text = Format(mvDate.Value, "YYYYMMDD")
    Case "�̳f���O���R.�X�����.��"
         txt_Tab1_DeliveryDate_End.Text = Format(mvDate.Value, "YYYYMMDD")
    Case Else
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_Tab0_DeliveryDate_End_Click()
'�̳q���O���R >> �X����� >> ��
If Trim(txt_Tab0_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab0_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab0_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�̳q���O���R.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab0_Header.Top + txt_Tab0_DeliveryDate_End.Top + txt_Tab0_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Header.Left + txt_Tab0_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab0_DeliveryDate_Start_Click()
'�̳q���O���R >> �X����� >> �_
If Trim(txt_Tab0_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab0_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab0_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab0_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab0_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�̳q���O���R.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab0_Header.Top + txt_Tab0_DeliveryDate_Start.Top + txt_Tab0_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab0_Header.Left + txt_Tab0_DeliveryDate_Start.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DeliveryDate_End_Click()
'�̳f���O���R >> �X����� >> ��
If Trim(txt_Tab1_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Tag = "�̳f���O���R.�X�����.��"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_End.Top + txt_Tab1_DeliveryDate_End.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_End.Left
mvDate.Visible = True
End Sub

Private Sub txt_Tab1_DeliveryDate_Start_Click()
'�̳f���O���R >> �X����� >> �_
If Trim(txt_Tab1_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_Tab1_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_Tab1_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab1_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab1_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Tag = "�̳f���O���R.�X�����.�_"
mvDate.Top = SSTab1.Top + fam_Tab1_Header.Top + txt_Tab1_DeliveryDate_Start.Top + txt_Tab1_DeliveryDate_Start.Height
mvDate.Left = SSTab1.Left + fam_Tab1_Header.Left + txt_Tab1_DeliveryDate_Start.Left
mvDate.Visible = True
End Sub


