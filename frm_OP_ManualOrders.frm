VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_ManualOrders 
   Caption         =   "�q����@�@�~ "
   ClientHeight    =   8310
   ClientLeft      =   240
   ClientTop       =   690
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   12480
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   2295
      Left            =   0
      TabIndex        =   104
      Top             =   4560
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   3
      RowHeight       =   20
      TabAction       =   2
      AllowDelete     =   -1  'True
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
   Begin VB.CommandButton cmdDelRs 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�R��CTrl+D"
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
      Left            =   1080
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   103
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAddRs 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�s�WCTrl+A"
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
      Left            =   0
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   102
      Top             =   3960
      Width           =   1035
   End
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   9360
      TabIndex        =   20
      Top             =   3960
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
      StartOfWeek     =   97189889
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame fam_Orders 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   3075
      Left            =   0
      TabIndex        =   30
      Top             =   840
      Width           =   11520
      Begin VB.CheckBox Chk_receive 
         Caption         =   "����"
         Height          =   255
         Left            =   9720
         TabIndex        =   110
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_OtQty 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9495
         MaxLength       =   20
         TabIndex        =   108
         Top             =   1130
         Width           =   990
      End
      Begin VB.ComboBox txt_B_city 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0000
         Left            =   7500
         List            =   "frm_OP_ManualOrders.frx":0002
         TabIndex        =   106
         Top             =   1130
         Width           =   1575
      End
      Begin VB.ComboBox cmdFacility 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0004
         Left            =   6000
         List            =   "frm_OP_ManualOrders.frx":0014
         TabIndex        =   105
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txt_Description 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         TabIndex        =   100
         Top             =   840
         Width           =   8700
      End
      Begin VB.ComboBox cbo_Priority 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0041
         Left            =   960
         List            =   "frm_OP_ManualOrders.frx":0043
         TabIndex        =   79
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame frm_OP_ManualShipToOrders 
         BackColor       =   &H00004040&
         BorderStyle     =   0  '�S���ؽu
         Enabled         =   0   'False
         Height          =   1470
         Left            =   120
         TabIndex        =   60
         Top             =   1440
         Width           =   10365
         Begin VB.TextBox txt_ShipToAreaCode 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6360
            TabIndex        =   69
            Top             =   825
            Width           =   3100
         End
         Begin VB.TextBox txt_ShipToZIP 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6360
            TabIndex        =   68
            Top             =   540
            Width           =   3100
         End
         Begin VB.TextBox txt_ShipToExtraDemand2 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5250
            TabIndex        =   67
            Top             =   1125
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToExtraDemand1 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1005
            TabIndex        =   66
            Top             =   1125
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToFullName 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1005
            TabIndex        =   65
            Top             =   255
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToShortName 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6360
            TabIndex        =   64
            Top             =   240
            Width           =   3100
         End
         Begin VB.TextBox txt_ShipToAddress 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1005
            TabIndex        =   63
            Top             =   825
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToContact 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1005
            TabIndex        =   62
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txt_ShipToPhone 
            Appearance      =   0  '����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3585
            TabIndex        =   61
            Top             =   540
            Width           =   1635
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��B��f�Ȥ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   78
            Top             =   60
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�W��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   77
            Top             =   315
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�²��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   28
            Left            =   5535
            TabIndex        =   76
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�a�}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   75
            Top             =   870
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϰ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   26
            Left            =   5535
            TabIndex        =   74
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�S��ݨD"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   73
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�l���ϸ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   24
            Left            =   5535
            TabIndex        =   72
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p���H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   71
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   22
            Left            =   3150
            TabIndex        =   70
            Top             =   585
            Width           =   390
         End
      End
      Begin VB.TextBox txtShipToKey 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4680
         TabIndex        =   59
         Top             =   1140
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShipToList 
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
         Height          =   300
         Left            =   6345
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   58
         Top             =   1125
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox cmbStorerkey 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0045
         Left            =   960
         List            =   "frm_OP_ManualOrders.frx":0047
         TabIndex        =   57
         Top             =   160
         Width           =   2175
      End
      Begin VB.TextBox txt_DeliveryDate 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9150
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txt_OrderKey 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6900
         TabIndex        =   55
         Top             =   160
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Consigneelist 
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
         Height          =   300
         Left            =   2625
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   54
         Top             =   1125
         Width           =   315
      End
      Begin VB.TextBox txt_Extern 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4050
         MaxLength       =   20
         TabIndex        =   53
         Top             =   160
         Width           =   1575
      End
      Begin VB.Frame fam_ConsigneeData 
         BackColor       =   &H00004040&
         BorderStyle     =   0  '�S���ؽu
         Enabled         =   0   'False
         Height          =   1470
         Left            =   120
         TabIndex        =   34
         Top             =   1485
         Width           =   10365
         Begin VB.ComboBox cmb_ExtraDemand2 
            BackColor       =   &H8000000B&
            Height          =   300
            Left            =   5235
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   44
            Top             =   1125
            Width           =   4230
         End
         Begin VB.ComboBox cmb_ExtraDemand1 
            BackColor       =   &H8000000A&
            Height          =   300
            Left            =   1005
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   43
            Top             =   1125
            Width           =   4230
         End
         Begin VB.ComboBox cmb_ZIP 
            BackColor       =   &H8000000B&
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
            Left            =   6360
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   42
            Top             =   525
            Width           =   3100
         End
         Begin VB.TextBox txt_Phone 
            Appearance      =   0  '����
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3585
            TabIndex        =   41
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txt_Contact 
            Appearance      =   0  '����
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1005
            TabIndex        =   40
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txt_Address 
            Appearance      =   0  '����
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1005
            TabIndex        =   39
            Top             =   825
            Width           =   4215
         End
         Begin VB.TextBox txt_ShortName 
            Appearance      =   0  '����
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6360
            TabIndex        =   38
            Top             =   240
            Width           =   3100
         End
         Begin VB.TextBox txt_FullName 
            Appearance      =   0  '����
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1005
            TabIndex        =   37
            Top             =   255
            Width           =   4215
         End
         Begin VB.ComboBox cmb_AreaCode 
            BackColor       =   &H8000000B&
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
            Left            =   6360
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   36
            Top             =   840
            Width           =   3100
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
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
            Height          =   240
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   35
            Top             =   60
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   16
            Left            =   3150
            TabIndex        =   52
            Top             =   585
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p���H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   51
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�l���ϸ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   14
            Left            =   5535
            TabIndex        =   50
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�S��ݨD"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϰ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   9
            Left            =   5535
            TabIndex        =   48
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�a�}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   47
            Top             =   870
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�²��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   7
            Left            =   5535
            TabIndex        =   46
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�W��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   45
            Top             =   315
            Width           =   780
         End
      End
      Begin VB.TextBox txt_ConsigneeKey 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         TabIndex        =   33
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox txt_OrderDate 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9150
         TabIndex        =   32
         Top             =   160
         Width           =   1335
      End
      Begin VB.TextBox txtType 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3495
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   9120
         TabIndex        =   109
         Top             =   1180
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ճ����O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   6735
         TabIndex        =   107
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q��Ƶ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   101
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q�����O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   89
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��B��f�Ȥ�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   3000
         TabIndex        =   88
         Top             =   1185
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ���ʽs��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   5700
         TabIndex        =   87
         Top             =   225
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q��s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3195
         TabIndex        =   86
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   85
         Top             =   1185
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�e�f��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   8565
         TabIndex        =   84
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   8565
         TabIndex        =   83
         Top             =   225
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�f�D�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   82
         Top             =   220
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q�檬�A"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   2640
         TabIndex        =   81
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�t�e�ܧO"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   5160
         TabIndex        =   80
         Top             =   540
         Width           =   780
      End
   End
   Begin VB.Frame fam_Header 
      Height          =   870
      Left            =   0
      TabIndex        =   16
      Top             =   -75
      Width           =   11520
      Begin VB.CommandButton cmd_AddNew 
         BackColor       =   &H00FF80FF&
         Caption         =   "�s  �W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Modify 
         BackColor       =   &H00FF8080&
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Save 
         BackColor       =   &H008080FF&
         Caption         =   "�s  ��"
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
         Height          =   495
         Left            =   7560
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Delete 
         BackColor       =   &H000080FF&
         Caption         =   "�R  ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         Top             =   240
         Width           =   1200
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
         Height          =   495
         Index           =   0
         Left            =   10200
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   15
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Cancel 
         BackColor       =   &H0080FFFF&
         Caption         =   "��  ��"
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
         Height          =   495
         Left            =   8880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_OrdersQuery 
         BackColor       =   &H0080C0FF&
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
         Height          =   435
         Left            =   2640
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txt_QueryExternOrderKey 
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
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   0
         Top             =   300
         Width           =   1590
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000080&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Height          =   585
         Left            =   3555
         Top             =   195
         Width           =   7920
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00004040&
         BackStyle       =   1  '���z��
         Height          =   495
         Left            =   2610
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "TMS�渹"
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
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame fam_OrderDetail 
      BackColor       =   &H8000000B&
      Height          =   4485
      Left            =   -240
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   16920
      Begin VB.CommandButton cmd_DetailVerify 
         BackColor       =   &H00808080&
         Caption         =   "�Ӷ�����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   9840
         Picture         =   "frm_OP_ManualOrders.frx":0049
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   21
         ToolTipText     =   "�s�W"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmd_DetailCancel 
         BackColor       =   &H0080FFFF&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10605
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   12
         ToolTipText     =   "�s�W"
         Top             =   765
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailSave 
         BackColor       =   &H008080FF&
         Caption         =   "�x�s"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9765
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         ToolTipText     =   "�s�W"
         Top             =   765
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailDel 
         BackColor       =   &H000080FF&
         Caption         =   "�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10605
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   13
         ToolTipText     =   "�R��"
         Top             =   285
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailAddNew 
         BackColor       =   &H00FF80FF&
         Caption         =   "�s�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8925
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
         ToolTipText     =   "�s�W"
         Top             =   285
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailModify 
         BackColor       =   &H00FF8080&
         Caption         =   "�ק�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9765
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         ToolTipText     =   "�s�W"
         Top             =   285
         Width           =   765
      End
      Begin VB.Frame fam_DetailData 
         Appearance      =   0  '����
         BackColor       =   &H00004000&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   8700
         Begin VB.TextBox txtCasecnt 
            Alignment       =   1  '�a�k���
            BackColor       =   &H80000004&
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
            Left            =   585
            Locked          =   -1  'True
            TabIndex        =   98
            ToolTipText     =   "�C�c�J��"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox txtLot5 
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
            Left            =   7635
            MaxLength       =   8
            TabIndex        =   96
            ToolTipText     =   "���w�����"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox txtLot4 
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
            Left            =   7635
            MaxLength       =   8
            TabIndex        =   94
            ToolTipText     =   "���w�s�y��"
            Top             =   480
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtOrderCS 
            Alignment       =   1  '�a�k���
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
            Left            =   600
            TabIndex        =   92
            ToolTipText     =   "�q��c��"
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtOrderEA 
            Alignment       =   1  '�a�k���
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
            Left            =   2025
            TabIndex        =   90
            ToolTipText     =   "�q��Ӽ�"
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtLot3 
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
            Left            =   5280
            TabIndex        =   28
            ToolTipText     =   "�Ͳ��帹"
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox cboLot6 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frm_OP_ManualOrders.frx":0913
            Left            =   3480
            List            =   "frm_OP_ManualOrders.frx":0915
            TabIndex        =   26
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNotes 
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
            Left            =   2040
            TabIndex        =   8
            ToolTipText     =   "�ӫ~�Ƶ�"
            Top             =   825
            Width           =   4860
         End
         Begin VB.ComboBox cboSku 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frm_OP_ManualOrders.frx":0917
            Left            =   600
            List            =   "frm_OP_ManualOrders.frx":0919
            TabIndex        =   7
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox txt_SkuDescr 
            BackColor       =   &H80000000&
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
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   120
            Width           =   4545
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�C�c"
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
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   99
            Top             =   900
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�����"
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
            Height          =   195
            Index           =   35
            Left            =   6960
            TabIndex        =   97
            Top             =   900
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�s�y��"
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
            Height          =   195
            Index           =   36
            Left            =   6960
            TabIndex        =   95
            Top             =   540
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   93
            Top             =   540
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ӽ�"
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
            Height          =   195
            Index           =   18
            Left            =   1560
            TabIndex        =   91
            Top             =   540
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�帹"
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
            Height          =   195
            Index           =   34
            Left            =   4800
            TabIndex        =   29
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lable14 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ܧO"
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
            Height          =   195
            Index           =   17
            Left            =   3000
            TabIndex        =   27
            Top             =   555
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ƶ�"
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
            Height          =   255
            Index           =   19
            Left            =   1560
            TabIndex        =   25
            Top             =   900
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f��"
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
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�~�W"
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
            Height          =   195
            Index           =   31
            Left            =   3600
            TabIndex        =   22
            Top             =   180
            Width           =   420
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_OrderDetail 
         Height          =   2520
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   4445
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         Height          =   495
         Index           =   4
         Left            =   9720
         Top             =   720
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         Height          =   495
         Index           =   1
         Left            =   10560
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         Height          =   495
         Index           =   0
         Left            =   10560
         Top             =   720
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404080&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00004040&
         BorderWidth     =   2
         Height          =   495
         Index           =   3
         Left            =   8880
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         Height          =   495
         Index           =   2
         Left            =   9720
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_OP_ManualOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private iLoop As Double              '�j��p��

Private intsrcSKUNowRow As Double

Private arZip() As String               '�l���ϸ�
Private arZIPArea() As String           '�l���ϸ��ɳ]�w�� AreaCode
Private arAreaCode() As String          '�ϰ�N�X
Private arExtraDemand() As String       '�S��ݨD
Private rsMain As ADODB.Recordset

Private Sub cmdAddRs_Click()
If rsMain Is Nothing Then Exit Sub
If dgMain.Enabled = False Then Exit Sub
Dim lngSeq As Long

'������
If rsMain.RecordCount > 0 Then
    rsMain.MoveLast
    lngSeq = Val(rsMain("����"))
End If

'�s�W
rsMain.AddNew
rsMain("����") = Format(lngSeq + 1, "00000")
'Call dgMain_RowColChange(1, 1)
dgMain.SetFocus
dgMain.Col = 1
End Sub

Private Sub cmdDelRs_Click()
If rsMain Is Nothing Then Exit Sub
If rsMain.RecordCount = 0 Then Exit Sub
If dgMain.Enabled = False Then Exit Sub
On Error GoTo err_Handle

If MsgBox("�T�w�R��������(" & RTrim(rsMain("����")) & ")?", vbOKCancel, "�T�{�R��") <> vbOK Then: Exit Sub

'�R��
rsMain.Delete

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, Me.Caption & "_cmdDelRs_Click")

End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 200 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgmain_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then 'Ctrl+A
    Call cmdAddRs_Click
    dgMain.SetFocus: dgMain.Col = 0
End If

If KeyAscii = 4 Then 'Ctrl+D
    cmdDelRs.SetFocus
    Call cmdDelRs_Click
End If

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
Dim strDate As String

With dgMain

'�P�@����
'If LastRow = Empty Then

    'dtpDeliveryTime.Visible = False
    If .DataSource Is Nothing Then Exit Sub
    If rsMain.RecordCount = 0 Then Exit Sub
    If rsMain.EOF Then Exit Sub
    If .Col = -1 Then Exit Sub
    
        If LastCol <> -1 Then '�O������^��
        
        '�~���ˬd
        If rsMain.Fields(LastCol).Name = "�f��" And Len(RTrim(rsMain("�f��"))) > 0 Then
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open "select descr,casecnt,innerpack,busr3=isnull(busr3,''),busr2 = isnull(busr2,''),busr1=isnull(busr1,'') from gv_skuxpack where sku = '" & rsMain("�f��") & "' and storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' ", cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF = True Then
                tmp_Rs.Close: .Col = LastCol: MsgBox "�L���f���I", 64, Me.Caption: .SetFocus: Exit Sub
            Else
                rsMain("�~�W") = RTrim(tmp_Rs("descr"))
                rsMain("�j���W��") = RTrim(tmp_Rs("busr3"))
                rsMain("�����W��") = RTrim(tmp_Rs("busr2"))
                rsMain("�p���W��") = RTrim(tmp_Rs("busr1"))
                rsMain("�j���J��") = RTrim(tmp_Rs("casecnt"))
                rsMain("�����J��") = RTrim(tmp_Rs("innerpack"))
                tmp_Rs.Close
                .Col = .Col + 1
            End If
        End If
    End If
    If rsMain("�j���ƶq") > 0 And rsMain("�j���J��") = 0 Then MsgBox "�j���J�Ƭ�0�A�L�k��J�j���ƶq�I", 16, "�`�N": rsMain("�j���ƶq") = 0: .SetFocus ': Exit Sub
    If rsMain("�����ƶq") > 0 And rsMain("�����J��") = 0 Then MsgBox "�����J�Ƭ�0�A�L�k��J�����ƶq�I", 16, "�`�N": rsMain("�����ƶq") = 0: .SetFocus ': Exit Sub
    If RTrim(GetWord(Trim(rsMain("�z�f�[�u")), 1, 10)) <> Trim(rsMain("�z�f�[�u")) Then MsgBox "�[�u���O���o�W�L10�Ӧr���I", 16, "�`�N": .SetFocus: rsMain("�z�f�[�u") = GetWord(Trim(rsMain("�z�f�[�u")), 1, 10) ': Exit Sub
    
    '�����\����
    If rsMain.Fields(.Col).Name = "�~�W" Or rsMain.Fields(.Col).Name = "�j���W��" Or rsMain.Fields(.Col).Name = "�����W��" Or rsMain.Fields(.Col).Name = "�����J��" Then .Col = .Col + 1
    If rsMain.Fields(.Col).Name = "�p���W��" Then .Col = .Col + 3
    If rsMain.Fields(.Col).Name = "�j���J��" Then .Col = .Col + 2
'End If

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_OP_ManualOrders = Nothing
End Sub

Private Sub cbo_Priority_Click()
Label3(3).Caption = "��f��"
frm_OP_ManualShipToOrders.ZOrder 1
Label3(21).Visible = False: txtShipToKey.Visible = False: cmdShipToList.Visible = False

If mySplit(Trim(cbo_Priority), " ", 0) = "R" Or mySplit(Trim(cbo_Priority), " ", 0) = "RC" Then

    Label3(3).Caption = "���f��"
ElseIf mySplit(Trim(cbo_Priority), " ", 0) = "A2B" Then

    Label3(3).Caption = "���f��"
    Label3(21).Visible = True: txtShipToKey.Visible = True: cmdShipToList.Visible = True
    
End If

End Sub

Private Sub cboSku_Click()

Dim rsTmp As New ADODB.Recordset, strLot6 As String

strLot6 = cboLot6.Text '�����ܧO

rsTmp.Open "select casecnt,descr = rtrim(isnull(descr,'')) from gv_skuxpack where storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' and sku = '" & cboSku & "' ", cn
If Not rsTmp.EOF Then
    txtCasecnt = rsTmp("casecnt"): txt_SkuDescr = rsTmp("descr")
Else
    txtCasecnt = 0:: txt_SkuDescr = ""
End If
rsTmp.Close

'���ܧO
str_SQL = "select distinct isnull(lotattribute.lottable06,'') as lottable06 from " & strWMSDB & "..lotxlocxid lotxlocxid join " & strWMSDB & "..lotattribute lotattribute on lotattribute.lot = lotattribute.lot where lotxlocxid.storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' and lotattribute.sku = '" & cboSku & "' order by lotattribute.lottable06 "
rsTmp.Open str_SQL, cn

If Not rsTmp.EOF Then rsTmp.MoveFirst

cboLot6.Clear
Do While Not rsTmp.EOF
    cboLot6.AddItem Trim(rsTmp("lottable06"))
    rsTmp.MoveNext
Loop

cboLot6 = strLot6 '�g�^�ܧO

End Sub

Private Sub cboSku_LostFocus()
Call cboSku_Click
'dg_OrderDetail.Col = 3: cboLot6.Text = dg_OrderDetail.Text  '�ܧO
End Sub

Private Sub cmbStorerkey_Click()

''���f��
'Dim rsTmp As New ADODB.Recordset
'str_SQL = "select sku = rtrim(sku) from sku where storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' order by sku"
'rsTmp.Open str_SQL, cn
'If rsTmp.EOF Then MsgBox "�䤣��ӳf�D�ӫ~�D�ɸ��", vbOKOnly, Me.Caption: Exit Sub
'rsTmp.MoveFirst
'
''txt_ConsigneeKey = ""
''txtShipToKey = ""
'
'cboSku.Clear
'Do While Not rsTmp.EOF
'    cboSku.AddItem rsTmp("sku")
'    rsTmp.MoveNext
'Loop
'
'rsTmp.Close: Set rsTmp = Nothing

End Sub

Private Sub cmd_AddNew_Click()
'�s�W

'�M���Ҧ����ȡA�]�t OrderDetail ����
Call Clear_AllField
fam_Orders.Enabled = True
fam_Orders.BackColor = &HFF80FF
fam_OrderDetail.BackColor = &HFF80FF

txt_Extern.Enabled = True
txt_OrderKey.Enabled = True
txt_QueryExternOrderKey.Text = ""
cmd_Modify.Enabled = False
cmd_AddNew.Enabled = False
cmd_Delete.Enabled = False
cmd_Save.Enabled = True
cmd_Cancel.Enabled = True

cmd_DetailAddNew.Enabled = True
cmd_DetailModify.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
cbo_Priority = ""
cmbStorerkey.Enabled = True
dgMain.Enabled = True

'���өw�q
str_SQL = "exec gs_ManualOrder_OrderDetail '" & txt_QueryExternOrderKey.Text & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Set rsMain = New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsMain)
tmp_Rs.Close

Set dgMain.DataSource = rsMain

SetDataGridColWidth Me.Caption, dgMain

txt_Extern.SetFocus
End Sub

Private Sub cmd_Cancel_Click()
'����

fam_Orders.Enabled = False
fam_Orders.BackColor = &H8000000A
fam_OrderDetail.BackColor = &H8000000A
txt_Extern.Enabled = False
txt_OrderKey.Enabled = False
cmd_Modify.Enabled = True
cmd_AddNew.Enabled = True
cmd_Save.Enabled = False
cmd_Cancel.Enabled = False

fam_DetailData.Enabled = False
dgMain.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False

Dim strTmp As String
strTmp = Trim(txt_QueryExternOrderKey)
'�M���Ҧ����ȡA�]�t OrderDetail ����
Call Clear_AllField

'�Y�O�ק�Ҧ��A���s���^��q����
If strTmp <> "" Then
   txt_QueryExternOrderKey.Text = strTmp
   Call cmd_OrdersQuery_Click
End If
End Sub

Private Sub cmd_Consigneelist_Click()
'��ܫȤ�ݿ�M��
'Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Consigneelist.Name)

If cmbStorerkey = "" Then MsgBox "�Х�����f�D�I", 64, Me.Caption: Exit Sub

'�I�s����
strDataList_Caller = Me.Name & " " & "txt_ConsigneeKey" & " " & mySplit(cmbStorerkey, " ", 0)
frm_ConsigneekeyQuery.Show vbModal

End Sub

Private Sub cmd_Delete_Click()
'�q�� >> �R��

Call cmd_OrdersQuery_Click

''�Q�פ��_���i�H��ʧ��
'If RTrim(cmbStorerkey) = "LLFA01" Then MsgBox "�Q�׳f�D�q��A���i��ʧ��!", vbOKOnly + vbCritical, "�q����@": Exit Sub
'If RTrim(cmbStorerkey) = "LMBO01" Then MsgBox "���_�f�D�q��A���i��ʧ��!", vbOKOnly + vbCritical, "�q����@": Exit Sub

If txtType = "�R��" Then Exit Sub

Dim rsTmp As New ADODB.Recordset, strRoute As String, Int_i As Long
Int_i = 0
rsTmp.Open "select orderkey from orders where orderkey = '" & txt_QueryExternOrderKey & "' ", cn
If rsTmp.EOF Then MsgBox "�䤣�즹�q��!", vbOKOnly, "�q��R��": rsTmp.Close: Exit Sub
rsTmp.Close

''�ˬd�q��O�_�w�^��WMS
'str_SQL = "select t2.receipt_no from trp02t t2 (nolock) join " & strWMSDB & "..orders o (nolock) on o.updatesource = t2.c_receipt_no where t2.c_receipt_no = '" & txt_QueryExternOrderKey & "' "
'rsTmp.Open str_SQL, cn
'If Not rsTmp.EOF Then MsgBox "���q��w�^��WMS�A�q��L�k�R��!", vbOKOnly, "�q��R��": rsTmp.Close: Exit Sub
'rsTmp.Close

'�ˬd�O�_�w�ƨ�
str_SQL = "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' "
rsTmp.Open str_SQL, cn
If Not rsTmp.EOF Then strRoute = rsTmp("route_no") & ""
If Not rsTmp.EOF Then MsgBox "�w�Ƹ��u�s��" & rsTmp("route_no") & "�A�q��L�k�R��!", vbOKOnly, "�q��R��": rsTmp.Close: Exit Sub
rsTmp.Close

If MsgBox("�T�w�R�����q��" & txt_QueryExternOrderKey & " (�t���έq��)? ", vbQuestion + vbYesNo, "�q��R��") <> vbYes Then Exit Sub
If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then If MsgBox("�R���x�W�Q�諸�q��ɡA�t�αN�۰ʵo�eMAIL�q���f�D�A�O�_�~��? ", vbQuestion + vbYesNo, "�q��R���q��") <> vbYes Then Exit Sub

Call DB_CheckConnectStatus

Tran_Level = cn.BeginTrans
     
    cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete TRP03T where receipt_no in (select receipt_no from trp02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete ORT03T where receipt_no in (select receipt_no from ort02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
     
    cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='�R��' ,editdate = getdate(),editwho = '" & User_id & "' where orderkey='" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords
    
    '��s���s���n���q
'    cn.Execute "exec gs_UpdateRoute '" & strRoute & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

txtType = "�R��"
cmbStorerkey.Enabled = True

'LTKK01�R��۰� Mail �q��
If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then Call SendMail(txt_QueryExternOrderKey)

Exit Sub

err_Handle:
If Tran_Level <> 0 Then Tran_Level = 0: cn.RollbackTrans

Dim tmpString As String
msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
CreateErrorLog Me.Name & "-�q��R��", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

End Sub

'Private Sub DeleteOrder()
'
'Call DB_CheckConnectStatus
'
'Tran_Level = cn.BeginTrans
'
'     cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='�R��' ,editdate = getdate()  where orderkey='" & txt_QueryExternOrderKey & "' and priority = '" & mySplit(Trim(cbo_Priority), " ", 0) & "' ", RowsAffect, adExecuteNoRecords
'
'cn.CommitTrans: Tran_Level = 0
'
'Exit Sub
'
'err_Handle:
'If Tran_Level <> 0 Then Tran_Level = 0: cn.RollbackTrans
'
'Dim tmpString As String
'msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'CreateErrorLog Me.Name & "-�q��R��", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
'MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'
'End Sub

Private Sub cmd_DetailAddNew_Click()
'�q��Ӷ� >> �s�W
intsrcSKUNowRow = 0
fam_DetailData.Enabled = True
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = True
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = True
txt_SkuDescr = ""

txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""
cboSku.SetFocus

End Sub

Private Sub cmd_DetailCancel_Click()
'�q��Ӷ� >> ����
If intsrcSKUNowRow = 0 Then
   cmd_DetailModify.Enabled = False
   cboSku.Text = "": txt_SkuDescr.Text = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""
Else
   '�ק�Ҧ��������G���^��Ӷ����
   cmd_DetailModify.Enabled = True
   With dg_OrderDetail
        .Row = intsrcSKUNowRow
        .Col = 1: .Text = cboSku.Text '�f��
        .Col = 2: .Text = Trim(txt_SkuDescr.Text)    '�~�W
        .Col = 3: .Text = cboLot6.Text   '�ܧO
        .Col = 4: .Text = txtOrderCS '�c��
        .Col = 5: .Text = txtOrderEA '�Ӽ�
        .Col = 6: .Text = txtCasecnt '�C�c
        .Col = 7: .Text = txtNotes '�Ƶ�
   End With
End If
cmd_DetailAddNew.Enabled = True
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
fam_DetailData.Enabled = False
End Sub

Private Sub cmd_DetailModify_Click()
'�q��Ӷ� >> �ק�
If intsrcSKUNowRow = 0 Then Exit Sub
fam_DetailData.Enabled = True
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = True
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = True

cboSku.SetFocus
End Sub

Private Sub cmd_DetailSave_Click()

txtNotes = myExCharFilter(txtNotes)

'�q��Ӷ� >> �s�W
If Len(Trim(cboSku.Text)) = 0 Then Exit Sub
If Len(RTrim(txtLot4)) > 0 And IsDate(Left(Trim(txtLot4), 4) & "/" & Mid(Trim(txtLot4), 5, 2) & "/" & Right(Trim(txtLot4), 2)) = False Then MsgBox "��ƿ��~�G�s�y����榡���~!!": Exit Sub
If Len(RTrim(txtLot5)) > 0 And IsDate(Left(Trim(txtLot5), 4) & "/" & Mid(Trim(txtLot5), 5, 2) & "/" & Right(Trim(txtLot5), 2)) = False Then MsgBox "��ƿ��~�G�������榡���~!!": Exit Sub
If Len(Trim(cboLot6)) = 0 Then: MsgBox "��ƿ��~�G�п�J�ܧO!!": cboLot6.SetFocus: Exit Sub
'And (Left(cbo_Priority, 1) = "A" Or Left(cbo_Priority, 1) = "I")
If Len(RTrim(txtOrderCS.Text)) = 0 Then txtOrderCS.Text = "0"
If Len(RTrim(txtOrderEA.Text)) = 0 Then txtOrderEA.Text = "0"
If CheckSKU(cboSku.Text) = 1 Then Exit Sub

'�t�e�ܧO�ˬd
If Len(Trim(cmdFacility)) = 0 Then
    cmdFacility = "�ըƹF�_��"
    If UCase(Right(Trim(cboLot6), 2)) = "-C" Then cmdFacility = "�ըƹF����"
    If UCase(Right(Trim(cboLot6), 2)) = "-S" Then cmdFacility = "�ըƹF�n��"
Else
    If UCase(Right(Trim(cboLot6), 2)) <> "-C" And UCase(Right(Trim(cboLot6), 2)) <> "-S" And cmdFacility <> "�ըƹF�_��" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": cboLot6.SetFocus: Exit Sub
    If UCase(Right(Trim(cboLot6), 2)) = "-C" And cmdFacility <> "�ըƹF����" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": cboLot6.SetFocus: Exit Sub
    If UCase(Right(Trim(cboLot6), 2)) = "-S" And cmdFacility <> "�ըƹF�n��" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": cboLot6.SetFocus: Exit Sub

End If

Dim lngOrderqty As Long
lngOrderqty = txtOrderCS.Text * txtCasecnt + txtOrderEA
If lngOrderqty = 0 Then
   msg_text = "��ƿ��~�G�q��ƶq���o�� 0"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txtOrderCS.SetFocus
   Exit Sub
End If

'1. �����Ʀs�J�� ROW
With dg_OrderDetail
     If intsrcSKUNowRow <> 0 Then '�ק�Ҧ��G�л\��Ӹ��
'        intsrcSKUNowRow = .Text
'        .Row = intsrcSKUNowRow
        If CheckSKU(cboSku.Text) = 1 Then Exit Sub
     Else
        Dim dbMaxNo As Double
        .Row = .Rows - 2: .Col = 0: dbMaxNo = Val(.Text)
        .Rows = .Rows + 1
        .Row = .Rows - 2
     End If
End With

'2. ��Ʀs�J dg_srcSKU ���w�� Row
With dg_OrderDetail
     .Col = 0    '�Ǹ�
     If intsrcSKUNowRow <> 0 Then
'        .Text = intsrcSKUNowRow    '�f����ƭק�Ҧ��A�u�έ�Ǹ�
     Else
        .Text = Format(CInt(dbMaxNo + 1), "0000")         '�f����Ʒs�W�Ҧ��A���ͷs�Ǹ�
     End If
     
     .Col = 1: .Text = UCase(RTrim(cboSku)) '�f��
     .Col = 2: .Text = txt_SkuDescr  '�~�W
     .Col = 3: .Text = cboLot6.Text   '�ܧO
     .Col = 4: .Text = txtOrderCS '�c��
     .Col = 5: .Text = txtOrderEA '�Ӽ�
     .Col = 6: .Text = txtCasecnt '�C�c
     .Col = 7: .Text = txtNotes '�Ƶ�
     .Col = 8: .Text = txtLot3 '�Ͳ��帹
     .Col = 9: .Text = txtLot4 '�s�y��
     .Col = 10: .Text = txtLot5 '�����
     
End With

'���]�ק�Ҧ��ѧO�X�ЭȡG�s�W�Ҧ�
intsrcSKUNowRow = 0

'�q����Ӹ�Ʒs�W�����A�M������
cboSku = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = "": txt_SkuDescr = "": txtLot3 = "": txtLot4 = "": txtLot5 = ""
cboSku.SetFocus

intsrcSKUNowRow = 0
fam_DetailData.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = True
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False

End Sub

Private Sub cmd_DetailDel_Click()
'�q��Ӷ� >> �R��
If intsrcSKUNowRow = 0 Then Exit Sub

Dim j As Integer
'1. �N�R���C��ƥѤU�@�C��ƨ��N
'   �ӫ᪺��ƦC���W���@�C
dg_OrderDetail.Visible = False
With dg_OrderDetail
     For iLoop = intsrcSKUNowRow To .Rows - 2   '�|���h�@��ťզC
         .Row = iLoop
         For j = 0 To .Cols - 1
             .Col = j
             .Text = .TextArray((.Row + 1) * .Cols + .Col)
         Next j
         '����̫�Ĥ@�C���W�����̫�ĤG�C�ɡA�|�O�˥ո�ƦC�A[�Ǹ�] ��줣�঳��
         '����ƪ��C�A[�Ǹ�] �������s�s��
         .Col = 0
         'If Val(.Text) = 0 Then .Text = "" Else .Text = .Row
     Next iLoop
'2. Grid �`�C�� - 1
     .Rows = .Rows - 1
     .Row = 1
     For iLoop = 0 To .Cols - 1
         .ColSel = iLoop
     Next iLoop
End With
dg_OrderDetail.Visible = True
'3. Reset �ܼ�
intsrcSKUNowRow = 0

'4. �M�������
cboSku.Text = "": txt_SkuDescr.Text = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""

fam_DetailData.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = True
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
fam_DetailData.Enabled = False

End Sub

'Private Sub cmd_DetailVerify_Click()
''�Ӷ�����
'If dg_OrderDetail.Rows = 2 Then Exit Sub
'If Trim(cmbStorerkey.Text) = "" Then
'   msg_text = "����J [�f�D] ���"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'End If
'
'Dim strSku As String, strErrorSKU As String
'Dim dbShipQty As Double
'
'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'strErrorSKU = ""
'With dg_OrderDetail
'     If .Rows = 2 Then Exit Sub
'     For iLoop = 1 To .Rows - 2
'         .Row = iLoop
'         .Col = 1: strSku = .Text
'         str_SQL = "Select *" & _
'                   "From BaseData_SKUPacking Where StorerKey = '" & cmbStorerkey.Text & "' and SKU = '" & strSku & "'"
'         tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'         If tmp_Rs.EOF Then
'            If strErrorSKU = "" Then
'               strErrorSKU = strSku
'            Else
'               strErrorSKU = strErrorSKU & "," & strSku
'            End If
'         Else
'            .Col = 3: .Text = tmp_Rs.Fields("�~�W").Value
''            .Col = 4: dbShipQty = .Text
''            .Col = 5: .Text = NumRound((dbShipQty / tmp_rs.Fields("�O�ഫ").Value), 2)
''            .Col = 6: .Text = NumRound((dbShipQty * tmp_rs.Fields("�����n").Value), 2)
''            .Col = 7: .Text = NumRound((dbShipQty * tmp_rs.Fields("��쭫�q").Value), 2)
'         End If
'         tmp_Rs.Close
'     Next iLoop
'End With
'End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_Modify_Click()
'�ק�
If Trim(txt_Extern.Text) = "" Then Exit Sub

'�Q�פ��i�H��ʧ��
If RTrim(cmbStorerkey) = "LLFA01" Then MsgBox "�Q�׳f�D�q��A�u���\�ק��f��P�Ƶ����A��L���Фŭק�!", vbOKOnly + vbExclamation, "�q����@"
'If RTrim(cmbStorerkey) = "LMBO01" Then MsgBox "���_�f�D�q��A���i��ʧ��!", vbOKOnly + vbCritical, "�q����@": Exit Sub

Call cmd_OrdersQuery_Click

fam_Orders.Enabled = True
fam_Orders.BackColor = &HFF8080
fam_OrderDetail.BackColor = &HFF8080

txt_Extern.Enabled = True
txt_OrderKey.Enabled = True
cmd_Modify.Enabled = False
cmd_Delete.Enabled = False
cmd_AddNew.Enabled = False
If txtType <> "�R��" Then cmd_Save.Enabled = True
cmd_Cancel.Enabled = True
txt_QueryExternOrderKey.Enabled = False
cmbStorerkey.Enabled = False
txt_OtQty.Enabled = True

'�q��Ӷ��s��\��]�w
cmd_DetailAddNew.Enabled = True
If intsrcSKUNowRow <> 0 Then
   cmd_DetailModify.Enabled = True
Else
   cmd_DetailModify.Enabled = False
End If
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
Call cmbStorerkey_Click
dgMain.Enabled = True

'���_�u��ק�header�A���ӭn���
If RTrim(cmbStorerkey) = "LMBO01" Then dgMain.Enabled = False
End Sub

Private Sub cmd_OrdersQuery_Click()

'�q��d��
fam_Orders.BackColor = &H8000000A
fam_OrderDetail.BackColor = &H8000000A
fam_Orders.Enabled = False
cmd_Modify.Enabled = False
cmd_AddNew.Enabled = True
cmd_Delete.Enabled = False
cmd_Save.Enabled = False
cmd_Cancel.Enabled = False
txt_QueryExternOrderKey.Enabled = True

fam_DetailData.Enabled = False
dgMain.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
cmbStorerkey.Enabled = True

txt_QueryExternOrderKey.Text = Trim(txt_QueryExternOrderKey.Text)
If Len(Trim(txt_QueryExternOrderKey.Text)) = 0 Then
   msg_text = "�Х���J�d�߱���G[TMS�渹]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_QueryExternOrderKey.SetFocus
   Exit Sub
End If

Dim strTmp As String
strTmp = Format(txt_QueryExternOrderKey.Text, "0000000000")

'�M���Ҧ����ȡA�]�t OrderDetail ����
Call Clear_AllField

txt_QueryExternOrderKey.Text = strTmp

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass

'���o�q�� Header
'str_SQL = "Select �f�D�渹,�f�D,�q���,�e�f��,�Ȥ�s��,����,�Ȥ�W��,�Ȥ�²��,�l���ϸ�,�B�e�ϰ�,�S��ݨD1,�S��ݨD2,�B�e�a�},�p���H,�q��,TMS�渹,b_phone2,�Ȥ�渹,�q�����O,��B��f�Ȥ�s��,�q�檬�A,�t�e�ܧO " & _
'          "From ManualOrder_Orders Where TMS�渹 = '" & txt_QueryExternOrderKey.Text & "'"

str_SQL = "Select Rtrim(a1.ExternOrderKey) As �f�D�渹 " & _
            ",Rtrim(a1.StorerKey) as �f�D " & _
            ",Convert(varchar(8),a1.OrderDate,112) as �q��� " & _
            ",Convert(varchar(8),a1.DeliveryDate,112) as �e�f�� " & _
            ",Rtrim(a1.ConsigneeKey) as �Ȥ�s�� " & _
            ",Isnull(Rtrim(Cast(a1.Notes as varchar(300))),'') as ���� " & _
            ",Case When b1.ConsigneeKey is null Then Rtrim(Isnull(a1.C_Company,'')) else Rtrim(Isnull(b1.Full_Name,'')) End as �Ȥ�W�� " & _
            ",Case When b1.ConsigneeKey is null Then '' else Rtrim(Isnull(b1.Short_Name,'')) End as �Ȥ�²�� " & _
            ",Case When b1.ZIP is not null Then Rtrim(b1.Zip) else Rtrim(Isnull(a1.C_Zip,'')) End as �l���ϸ� " & _
            ",Rtrim(Isnull(b1.Area_Code,'')) as �B�e�ϰ� " & _
            ",Rtrim(Isnull(b1.Extra_Demand_Code,'')) as �S��ݨD1 " & _
            ",Rtrim(Isnull(b1.Extra_Demand_Code2,'')) as �S��ݨD2 " & _
            ",Case When b1.Address is not null then Rtrim(b1.Address) else Rtrim(Isnull(a1.C_Address1,'')) +Rtrim(Isnull(a1.C_Address2,'')) End as �B�e�a�} " & _
            ",Case When b1.Contact is not null Then Rtrim(b1.Contact) else Rtrim(Isnull(a1.C_Contact1,'')) End as �p���H " & _
            ",Case When b1.phone is not null Then Rtrim(b1.phone) else Rtrim(Isnull(a1.C_Phone1,'')) End as �q�� " & _
            ",a1.OrderKey as TMS�渹 " & _
            ",B_phone2 = isnull(B_phone2,'') " & _
            ",�Ȥ�渹 = rtrim(a1.customerorderkey) " & _
            ",�q�����O = rtrim(a1.priority) " & _
            ",��B��f�Ȥ�s�� = rtrim(isnull(a1.b_company,'')) " & _
            ",�q�檬�A = rtrim(isnull(a1.type,'')) " & _
            ",�t�e�ܧO = rtrim(isnull(a1.facility,'')) " & _
            ",�ճ����O = rtrim(isnull(a1.b_city,''))" & _
            ",��� = case when isnull(cast(a1.otqty as char),'') = '' then '' else rtrim(cast(a1.otqty as char)) end ,���� = isnull(a1.GoodsBack,0) " & _
            "From Orders a1 Left outer join TRP01M b1 on b1.ConsigneeKey = a1.ConsigneeKey and b1.storerkey = a1.storerkey where a1.OrderKey = '" & txt_QueryExternOrderKey.Text & "'"

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_Modify.Enabled = False
   Screen.MousePointer = vbDefault
   Exit Sub
End If
'if tmp_rs("b_phone2") = 00 then label1.Caption ="�w��J�ƨ��t�ΡA�L�k�ܧ�!!"
txt_Extern.Text = tmp_Rs.Fields("�f�D�渹").Value
txt_OrderKey.Text = tmp_Rs.Fields("�Ȥ�渹").Value
cmbStorerkey.Text = tmp_Rs.Fields("�f�D").Value
txt_OrderDate.Text = tmp_Rs.Fields("�q���").Value
txt_DeliveryDate.Text = tmp_Rs.Fields("�e�f��").Value
txt_ConsigneeKey.Text = tmp_Rs.Fields("�Ȥ�s��").Value
txt_Description.Text = tmp_Rs.Fields("����").Value
txt_FullName.Text = tmp_Rs.Fields("�Ȥ�W��").Value
txt_Contact.Text = tmp_Rs.Fields("�p���H").Value
txt_Phone.Text = tmp_Rs.Fields("�q��").Value
txt_Address.Text = tmp_Rs.Fields("�B�e�a�}").Value
cbo_Priority.Text = tmp_Rs.Fields("�q�����O").Value
txtShipToKey.Text = tmp_Rs("��B��f�Ȥ�s��") & ""
txtType.Text = tmp_Rs("�q�檬�A")
cmdFacility.Text = tmp_Rs("�t�e�ܧO")
txt_B_city.Text = tmp_Rs("�ճ����O")
txt_OtQty.Text = tmp_Rs("���")

If RTrim(tmp_Rs("����")) = "1" Then Chk_receive.Value = 1 Else Chk_receive.Value = 0

If RTrim(cbo_Priority) = "A2B" Then Label3(21).Visible = True: txtShipToKey.Visible = True: cmdShipToList.Visible = True

If Len(RTrim(tmp_Rs("�S��ݨD1"))) > 0 Then
    For iLoop = 0 To cmb_ExtraDemand1.ListCount - 1
        If arExtraDemand(iLoop) = tmp_Rs.Fields("�S��ݨD1").Value Then
           cmb_ExtraDemand1.ListIndex = iLoop
           Exit For
        End If
    Next iLoop
End If

If Len(RTrim(tmp_Rs("�S��ݨD2"))) > 0 Then
    For iLoop = 0 To cmb_ExtraDemand2.ListCount - 1
        If arExtraDemand(iLoop) = tmp_Rs.Fields("�S��ݨD2").Value Then
           cmb_ExtraDemand2.ListIndex = iLoop
           Exit For
        End If
    Next iLoop
End If

txt_ShortName.Text = tmp_Rs.Fields("�Ȥ�²��").Value

For iLoop = 0 To cmb_ZIP.ListCount - 1
    If arZip(iLoop) = tmp_Rs.Fields("�l���ϸ�").Value Then
       cmb_ZIP.ListIndex = iLoop
       Exit For
    End If
Next iLoop
DoEvents: DoEvents

For iLoop = 0 To cmb_AreaCode.ListCount - 1
    If arAreaCode(iLoop) = tmp_Rs.Fields("�B�e�ϰ�").Value Then
       cmb_AreaCode.ListIndex = iLoop
       Exit For
    End If
Next iLoop
tmp_Rs.Close

'���o�q�� Detail >> �H OrderDetail ���D
'str_SQL = "exec gs_ManualOrder_OrderDetail '" & txt_QueryExternOrderKey.Text & "' "

str_SQL = "Select ���� = isnull(od.OrderLineNumber,'') " & _
            ",�f�� = Rtrim(od.SKU),�~�W = Rtrim(Isnull(s.Descr,'')),�ܧO = rtrim(isnull(od.lottable06,'')) " & _
            ",�j���ƶq = case when s.casecnt=0 then 0  else floor(od.originalQty/convert(int,s.casecnt)) end " & _
            ",�j���W�� = rtrim(isnull(s.busr3,'')) " & _
            ",�����ƶq = case when s.casecnt=0 then case when s.InnerPack=0 then 0 else floor(convert(int,od.originalQty)/convert(int,s.InnerPack)) end else case when s.InnerPack=0  then 0 else floor((convert(int,od.originalQty)%convert(int,s.casecnt))/convert(int,s.InnerPack)) end end " & _
            ",�����W�� = rtrim(isnull(s.busr2,'')) " & _
            ",�p���ƶq = case when s.casecnt=0 then  " & _
                            "case when s.InnerPack=0 then od.originalQty  " & _
                            "else convert(int,od.originalQty)%convert(int,s.InnerPack) end  " & _
                            "else case when s.InnerPack=0  then convert(int,od.originalQty)%convert(int,s.casecnt)  " & _
                            "else (convert(int,od.originalQty)%convert(int,s.casecnt))%convert(int,s.InnerPack) end End  " & _
            ",�p���W�� = rtrim(isnull(s.busr1,'')),�j���J�� = s.casecnt " & _
            ",�����J�� = s.innerpack,�Ͳ��帹 = rtrim(isnull(lottable03,'')) " & _
            ",�s�y�� = rtrim(isnull(convert(char(8),lottable04,112),'')) " & _
            ",����� = rtrim(isnull(convert(char(8),lottable05,112),'')) " & _
            ",�z�f�[�u = rtrim(isnull(updatesource,'')) " & _
            ",�Ƶ� = rtrim(isnull(od.notes,'')) " & _
            "From OrderDetail od (nolock) join gv_skuxpack s (nolock) on s.StorerKey = od.StorerKey and s.SKU = od.SKU where od.orderkey = '" & txt_QueryExternOrderKey.Text & "' " & _
            "Order by od.OrderLineNumber"

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q����Ӹ��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_Modify.Enabled = False
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Set rsMain = New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsMain)
tmp_Rs.Close

rsMain.MoveFirst
Set dgMain.DataSource = rsMain

SetDataGridColWidth Me.Caption, dgMain

'Do While Not tmp_Rs.EOF
'   With dg_OrderDetail
'       .Rows = .Rows + 1
'       .Row = .Rows - 2
'       .Col = 0: .Text = Rtrim(tmp_Rs("����"))
'       .Col = 1: .Text = tmp_Rs("�f��")
'       .Col = 2: .Text = tmp_Rs("�~�W")
'       .Col = 3: .Text = tmp_Rs("�ܧO") & ""
'       .Col = 4: .Text = tmp_Rs("�c��")
'       .Col = 5: .Text = tmp_Rs("�Ӽ�")
'       .Col = 6: .Text = tmp_Rs("�C�c")
'       .Col = 7: .Text = tmp_Rs("�Ƶ�") & ""
'       .Col = 8: .Text = tmp_Rs("�Ͳ��帹") & ""
'       .Col = 9: .Text = tmp_Rs("�s�y��") & ""
'       .Col = 10: .Text = tmp_Rs("�����") & ""
'
'  End With
'  tmp_Rs.MoveNext
'Loop
'tmp_Rs.Close

If txtType <> "�R��" Then cmd_Modify.Enabled = True
If txtType <> "�R��" Then cmd_Delete.Enabled = True
cmd_AddNew.Enabled = True
cmd_Cancel.Enabled = False

'�q��Ӷ��\����w
'fam_DetailData.Enabled = False
'cmd_DetailModify.Enabled = False
'cmd_DetailAddNew.Enabled = False
'cmd_DetailSave.Enabled = False
'cmd_DetailDel.Enabled = False
'fam_DetailData.Enabled = False
'intsrcSKUNowRow = 0

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q����@-�q��d��", Me.Caption, "cmd_OrdersQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub
Sub SendMail(strOrderkey As String)

''LTKK01�R��۰� Mail �q��
'If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then
'
'    Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String
'
'    'Ū��ini�Ѽ�
'    Dim objIni As New vbIniFile
'    objIni.FileName = App.Path & "/" & App.title & ".ini"
'
'    strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
'    strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
'    strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
'    strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
'    strSubject = "�q��R������"
'    strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
'    strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
'    strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
'    strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")
'
'    '�������w
'    strFrom = "Tkedi@bestlog.com.tw"
'    strTo = "jack@mail.kirin.com.tw,irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;simon@mail.kirin.com.tw"
'    strCC = "Tkedi@bestlog.com.tw"
'    strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
''    strSubject = "�q��R������"
'    strTextbody = "�����t�εo�e�H��!!"
'    strEmailID = "tkedi"
'    strEmailPW = "tkedibl01"
'    strAlways = "NO"
'
'    Set objIni = Nothing
'
'    Dim rsTmp As New ADODB.Recordset
'
'    If Len(RTrim(strFrom)) > 0 Then '���H���
'
'        str_SQL = "select �ܧO = 'BL01' " & _
'                ",�f�D�N�X = rtrim(o.storerkey) " & _
'                ",�q�渹�X��f�渹 = rtrim(od.externorderkey) + rtrim(od.externlineno) " & _
'                ",�a�}�O = substring(o.consigneekey,5,20) " & _
'                ",�Ƹ� = isnull(ss.storersku,od.sku) " & _
'                ",�ܧO_�x��O = 'BL01_'+ od.lottable06 " & _
'                ",�̤p���ƶq = isnull(od.originalqty,0) ,�q��� = convert(varchar,o.orderdate,111) " & _
'                ",�w�p��f�� =  convert(varchar,o.deliverydate,111) " & _
'                ",�R��� = convert(varchar,o.editdate,111) " & _
'                ",�Ȥ�q�渹�X = rtrim(o.customerorderkey) " & _
'                "From orders o join orderdetail od on o.orderkey = od.orderkey " & _
'                "left join storersku ss on ss.sku = od.sku and ss.storerkey = od.storerkey " & _
'                "Where o.type = '�R��' and o.orderkey = '" & strOrderkey & "' "
'
'        rsTmp.Open str_SQL, cn
'
'        '�p�G�L��Ƥ]�nmail
'        If Not rsTmp.EOF Or UCase(RTrim(strAlways)) = "YES" Then
'
'            strAddAttachment = "C:\LTKK01\�q��R������\�q��R������_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'
'            Call Recordset2Excel("�q��R������", rsTmp)
'            If Dir("C:\LTKK01\�q��R������", vbDirectory) = "" Then MkDirs "C:\LTKK01\�q��R������"
'            MyXlsApp.ActiveWorkbook.SaveAs strAddAttachment
'            MyXlsApp.Quit: Set MyXlsApp = Nothing
'
'            '�ǰe�l��
'            Dim objEmail As Object
'            Set objEmail = CreateObject("CDO.Message")
'
'            objEmail.From = strFrom
'            objEmail.To = strTo
'            objEmail.CC = strCC   ' �ƥ�
'            objEmail.BCC = strBCC ' �K��ƥ�
'            objEmail.Subject = strSubject
'            objEmail.TextBody = strTextbody
'            objEmail.AddAttachment strAddAttachment
'
'            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
'            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'            'SMTP ���A���ݭn���Ү�
'            If Len(RTrim(strEmailID)) > 0 Then
'                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
'                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
'            End If
'            objEmail.Configuration.Fields.Update
'            objEmail.Send
'
'            MsgBox "LTKK01�R����Ӹ�ơA�t�Τw�oMail�q��!", , "�R����Ӹ��"
'
'            Set objEmail = Nothing
'        End If
'    End If
'End If

End Sub
Private Sub cmd_Save_Click()
On Error GoTo err_Handle

'�Q�פ��i�H��ʧ��
'If RTrim(cmbStorerkey) = "LLFA01" Then MsgBox "�Q�׳f�D�q��A���i��ʧ��!", vbOKOnly + vbCritical, "�q����@": Exit Sub
'If RTrim(cmbStorerkey) = "LMBO01" Then MsgBox "���_�f�D�q��A���i��ʧ��!", vbOKOnly + vbCritical, "�q����@": Exit Sub

'Terry 20181220 user�n�D�������b
''�ˬd�ճ����O�A�ثe�u���Ȱ��f�D�A�ϥ�
'If Len(RTrim(txt_B_city)) > 0 And RTrim(Left(cmbStorerkey.Text, 6)) <> "LABT01" Then
'    MsgBox "�ճ����O�ثe�u���Ȱ��f�D�ϥΡA�L�k�s�ɡC�нT�{�ɮ�", vbOKOnly + vbCritical, "�q����@": Exit Sub
'End If

'�ˬd��ƬO�_���Ʀr
If Len(RTrim(txt_OtQty.Text)) > 0 And Not IsNumeric(txt_OtQty.Text) Then MsgBox "������п�J�Ʀr", vbOKOnly + vbCritical, "�`�N": txt_OtQty.SelStart = 0: txt_OtQty.SelLength = Len(txt_OtQty.Text): txt_OtQty.SetFocus: Exit Sub


rsMain.MoveFirst
Do While Not rsMain.EOF

If Trim(cbo_Priority) = "A2B" Then Exit Do
'�t�e�ܧO�ˬd
If Len(Trim(cmdFacility)) = 0 Then
    cmdFacility = "�ըƹF�_��"
    If UCase(Right(Trim(rsMain("�ܧO")), 2)) = "-C" Then cmdFacility = "�ըƹF����"
    If UCase(Right(Trim(rsMain("�ܧO")), 2)) = "-S" Then cmdFacility = "�ըƹF�n��"
Else
    If UCase(Right(Trim(rsMain("�ܧO")), 2)) <> "-C" And UCase(Right(Trim(rsMain("�ܧO")), 2)) <> "-S" And cmdFacility <> "�ըƹF�_��" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": Exit Sub
    If UCase(Right(Trim(rsMain("�ܧO")), 2)) = "-C" And cmdFacility <> "�ըƹF����" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": Exit Sub
    If UCase(Right(Trim(rsMain("�ܧO")), 2)) = "-S" And cmdFacility <> "�ըƹF�n��" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": Exit Sub

End If

rsMain.MoveNext
Loop

cmd_Save.Enabled = False
'�M���S��r��
Call myFormExCharFilter(Me)

'�ˮָ�ƥ��T
If CheckOrdersData() = False Then cmd_Save.Enabled = True: Exit Sub

Tran_Level = cn.BeginTrans

'�аO�­q�欰�R��
If txt_QueryExternOrderKey.Enabled = False And txtType <> "�R��" Then
     
'    Dim rsTmp As New ADODB.Recordset
'    rsTmp.Open "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' ", cn
'    If Not rsTmp.EOF Then MsgBox "���q��w�Ƹ��u�s�� " & rsTmp("route_no") & " �A�L�k�ק�!", vbOKOnly, "�s��": rsTmp.Close: Exit Sub
        
    Call DB_CheckConnectStatus
    
         cn.Execute "delete TRP03T where receipt_no in (select receipt_no from trp02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords

         cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete ORT03T where receipt_no in (select receipt_no from ort02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords

         cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='�R��' ,editdate = getdate() , editwho= '" & User_id & "' where orderkey='" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords

If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then Call SendMail(txt_QueryExternOrderKey)

End If

'�q���Ʀs��
If SaveOrdersData() = False Then
    cn.RollbackTrans: Tran_Level = 0
    MsgBox "�s�ɥ��ѡI", 16, "���~"
    Exit Sub
Else
    cn.CommitTrans: Tran_Level = 0
    txtType = "�R��"
    MsgBox "�q��ק�s�W�����C", vbOKOnly, Me.Caption
End If

Call cmd_OrdersQuery_Click

'�N���q��аO�borders..urgent_mark���
'�ˬd�O�_�����q��?
str_SQL = "select orderkey " & _
          "From orders " & _
          "where storerkey = 'LAPP01' and priority = 'I' and orderkey = '" & txt_QueryExternOrderKey & "' and type <> '�R��' and " & _
          "((convert(varchar(8),adddate,114) > '17:00:00' and convert(varchar(8),deliverydate,112) < = convert(varchar(8),getdate()+1,112) ) or " & _
          "(convert(varchar(8),adddate,114) > '17:30:00' and convert(varchar(8),deliverydate,112) < = convert(varchar(8),getdate()+2,112) ) or " & _
          "(convert(varchar(8),adddate,112) = convert(varchar(8),deliverydate,112)))"

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '���^��
    If MsgBox("�o�{���q��A�O�_�۰ʱN�q���s�����q��?", vbQuestion + vbYesNo, "�q����@") = vbYes Then
           '��surgent_mark���V:���q��
           cn.Execute "update orders set urgent_mark = 'V' where orderkey = '" & txt_QueryExternOrderKey & " ' ", RowsAffect, adExecuteNoRecords
    End If
End If

tmp_Rs.Close

fam_Orders.Enabled = False
fam_Orders.BackColor = &H8000000A
fam_OrderDetail.BackColor = &H8000000A
txt_Extern.Enabled = False
txt_OrderKey.Enabled = False
cmd_Modify.Enabled = True
cmd_AddNew.Enabled = True
cmd_Save.Enabled = False
cmd_Cancel.Enabled = False

fam_DetailData.Enabled = False
dgMain.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
cmbStorerkey.Enabled = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdShipToList_Click()
'��ܫȤ�ݿ�M��
'Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Consigneelist.Name)

If cmbStorerkey = "" Then MsgBox "�Х�����f�D�I", 64, Me.Caption: Exit Sub

'�I�s����
strDataList_Caller = Me.Name & " " & "txtShipToKey" & " " & mySplit(cmbStorerkey, " ", 0)
frm_ConsigneekeyQuery.Show vbModal
End Sub

Private Sub Command1_Click()

'�M���S��r��
Call myFormExCharFilter(Me)

If Len(RTrim(cmbStorerkey)) = 0 Then MsgBox "�п�ܳf�D�I", 64, Me.Caption: Exit Sub
If Len(RTrim(txt_ConsigneeKey)) = 0 Then MsgBox "�п�J�Ȥ�s���I", 64, Me.Caption: Exit Sub
If Len(RTrim(cmb_ZIP)) = 0 Then MsgBox "�п�J�l���ϸ��I", 64, Me.Caption: Exit Sub
If Len(RTrim(txt_FullName)) = 0 Then MsgBox "�п�J�Ȥ�W�١I", 64, Me.Caption: Exit Sub

Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select * from trp01m where storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' and consigneekey = '" & RTrim(txt_ConsigneeKey) & "' ", cn

If rsTmp.EOF Then

    str_SQL = "insert into trp01m(storerkey,consigneekey,area_code,zip,full_name,short_name,contact,phone,address,extra_demand_code,extra_demand_code2,addwho,editwho,updatesource) values('" & _
               mySplit(cmbStorerkey, " ", 0) & "','" & txt_ConsigneeKey & "','" & mySplit(cmb_AreaCode, " ", 0) & "','" & mySplit(cmb_ZIP, " ", 0) & "','" & txt_FullName & "','" & txt_ShortName & "','" & txt_Contact & "','" & txt_Phone & "','" & txt_Address & "','" & mySplit(cmb_ExtraDemand1, " ", 0) & "','" & mySplit(cmb_ExtraDemand1, " ", 0) & "','" & User_id & "','" & User_id & "','Manual') "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

Else
    MsgBox "�Ȥ�s�����ơA�нT�{�I", 64, Me.Caption

End If

rsTmp.Close

End Sub

Private Sub dg_OrderDetail_Click()
cboSku = "": txt_SkuDescr = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": txtLot3.Text = "": txtLot4.Text = "": txtLot5.Text = "": cboLot6.Text = ""

If fam_Orders.Enabled Then
   cmd_DetailModify.Enabled = False
   cmd_DetailAddNew.Enabled = True
   cmd_DetailSave.Enabled = False
   cmd_DetailDel.Enabled = False
   cmd_DetailCancel.Enabled = False
   fam_DetailData.Enabled = False
End If
With dg_OrderDetail
     intsrcSKUNowRow = 0
     .Col = 0    '����
     If Len(.Text) = 0 Then Exit Sub
     If .Row = 0 Then Exit Sub
     .Col = 1: cboSku = .Text
     .Col = 2: txt_SkuDescr = .Text
     .Col = 3: cboLot6.Text = .Text
     .Col = 4: txtOrderCS.Text = .Text
     .Col = 5: txtOrderEA.Text = .Text
     .Col = 6: txtCasecnt.Text = .Text
     .Col = 7: txtNotes.Text = .Text
     .Col = 8: txtLot3.Text = .Text
     .Col = 9: txtLot4.Text = .Text
     .Col = 10: txtLot5.Text = .Text
     
     intsrcSKUNowRow = .Row
     .Col = 0
     For iLoop = 0 To .Cols - 1
         .ColSel = iLoop
     Next iLoop
     If fam_Orders.Enabled Then
        cmd_DetailModify.Enabled = True
        cmd_DetailAddNew.Enabled = True
        cmd_DetailSave.Enabled = False
        cmd_DetailDel.Enabled = True
        cmd_DetailCancel.Enabled = False
        fam_DetailData.Enabled = False
    End If
End With

End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�q����@�@�~"
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

'�]�w�q����Ӯ榡
Call SetGDFormat_OrderDetail

Dim tmp_cnt As Double
'���X�Ҧ��l���ϸ� TRP02M
cmb_ZIP.Clear
str_SQL = "Select Rtrim(ZIP) as 'ZIP',Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP02M Order by ZIP"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arZip(1) As String
ReDim arZIPArea(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arZip(tmp_cnt) = tmp_Rs.Fields("ZIP").Value
      arZIPArea(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_ZIP.AddItem tmp_Rs.Fields("ZIP").Value & Space(5 - Len(Trim(tmp_Rs.Fields("ZIP").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arZip) Then
         ReDim Preserve arZip(UBound(arZip) + 10) As String
         ReDim Preserve arZIPArea(UBound(arZIPArea) + 10) As String
      End If
   Loop
End If

'���X�Ҧ��B�e�ϰ�N�X TRP03M
cmb_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'���X�Ҧ��S��ݨD--TRP04M
cmb_ExtraDemand1.Clear: cmb_ExtraDemand2.Clear
str_SQL = "Select Rtrim(Extra_Demand_Code) as 'ECode',Isnull(Rtrim(Description),'') as 'ECodeDescr' From TRP04M Order by Extra_Demand_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arExtraDemand(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arExtraDemand(tmp_cnt) = tmp_Rs.Fields("ECode").Value
      cmb_ExtraDemand1.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      cmb_ExtraDemand2.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arExtraDemand) Then
         ReDim Preserve arExtraDemand(UBound(arExtraDemand) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'���X�Ҧ��f�D
str_SQL = "Select Rtrim(Storerkey) + ' ' + rtrim(short_name) as 'Storer' From TRP16M where storerkey <> 'LTKK01' Order by Storerkey "
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn
If Not tmp_Rs.EOF Then

   Do While Not tmp_Rs.EOF
      cmbStorerkey.AddItem tmp_Rs("Storer")
      tmp_Rs.MoveNext
   Loop
End If
cmbStorerkey.ListIndex = -1
tmp_Rs.Close

cbo_Priority.AddItem "I �X�f"
cbo_Priority.AddItem "R �h�f"
cbo_Priority.AddItem "A ���"
cbo_Priority.AddItem "A2B ���f�t�e"
cbo_Priority.AddItem "RC ���f�J�w"
cbo_Priority.AddItem "C �V�w"

'cbo_Priority.AddItem "RS �h�f�~�X�w"

txt_B_city.AddItem ""
txt_B_city.AddItem "���@"
txt_B_city.AddItem "���d"

cbo_Priority = ""

End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub

'fam_OrderDetail.Width = Me.ScaleWidth - 60
'dg_OrderDetail.Width = fam_OrderDetail.Width - 180
'
'fam_OrderDetail.Height = Me.ScaleHeight - fam_Orders.Height - fam_Header.Height ' - 360
'dg_OrderDetail.Height = fam_OrderDetail.Height - fam_DetailData.Height - 240

dgMain.Width = Me.ScaleWidth - 60
dgMain.Height = Me.ScaleHeight - fam_Orders.Height - fam_Orders.Top - fam_Header.Height - fam_Header.Top

End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub SetGDFormat_OrderDetail()
'�W�١GSetGDFormat_OrdereDtail
'���O�G�Ƶ{��
'�\��G�M���ó]�w [�q��W�Ӹ��] ��ܮ榡
'�ѼơG�ǤJ�ȡG�L
Dim sub_var1 As Integer, sub_var2 As Integer
dg_OrderDetail.Visible = False
With dg_OrderDetail
     .FixedRows = 1: .Cols = 11
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

     '�]�w�C���榡
     .Row = 0
     .Col = 0: .Text = "����": .ColWidth(0) = 1000: .ColAlignment(0) = flexAlignLeftCenter
     .Col = 1: .Text = "�f��": .ColWidth(1) = 2400: .ColAlignment(1) = flexAlignLeftCenter
     .Col = 2: .Text = "�~�W": .ColWidth(2) = 3000: .ColAlignment(2) = flexAlignLeftCenter
     .Col = 3: .Text = "�ܧO": .ColWidth(3) = 600: .ColAlignment(3) = flexAlignCenterCenter
     .Col = 4: .Text = "�c��": .ColWidth(4) = 600: .ColAlignment(4) = flexAlignRightCenter
     .Col = 5: .Text = "�Ӽ�": .ColWidth(5) = 600: .ColAlignment(5) = flexAlignRightCenter
     .Col = 6: .Text = "�C�c": .ColWidth(6) = 600: .ColAlignment(6) = flexAlignRightCenter
     .Col = 7: .Text = "�Ƶ�": .ColWidth(7) = 3000: .ColAlignment(7) = flexAlignLeftCenter
     .Col = 8: .Text = "�Ͳ��帹": .ColWidth(8) = 1200: .ColAlignment(8) = flexAlignLeftCenter
     .Col = 9: .Text = "�s�y��": .ColWidth(9) = 800: .ColAlignment(9) = flexAlignLeftCenter
     .Col = 10: .Text = "�����": .ColWidth(10) = 800: .ColAlignment(10) = flexAlignLeftCenter

     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_OrderDetail.Visible = True
End Sub

Private Sub Clear_AllField()
'�M���Ҧ�����
'txt_Extern.Text = ""
cmbStorerkey = ""
cboSku = ""
'cbo_Priority.ListIndex = 0
'txt_OrderDate.Text = ""
'txt_DeliveryDate.Text = ""
'txt_Description.Text = ""
'txt_OrderKey.Text = ""
'txt_ConsigneeKey.Text = ""
'txt_FullName.Text = "": txt_ShortName.Text = ""
'txt_Contact.Text = "": txt_Phone.Text = ""
'txt_Address.Text = "": cmb_AreaCode.ListIndex = -1: cmb_ZIP.ListIndex = -1
'cmb_ExtraDemand1.ListIndex = -1: cmb_ExtraDemand2.ListIndex = -1
cmdFacility = ""
txt_B_city = ""
txt_OtQty = ""
Call ClearForm_AllField(frm_OP_ManualOrders)
fam_ConsigneeData.ZOrder 0

'�]�w�q����Ӯ榡
Call SetGDFormat_OrderDetail
cboSku.ListIndex = -1: txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
       Case "�q���"
            txt_OrderDate.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�e�f��"
            txt_DeliveryDate.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub txt_ConsigneeKey_GotFocus()
    fam_ConsigneeData.ZOrder 0
    fam_ConsigneeData.Enabled = True
End Sub

Public Sub txt_ConsigneeKey_LostFocus()

'�Ȥ�s��
fam_ConsigneeData.Enabled = False
If Trim(txt_ConsigneeKey.Text) = "" Then Exit Sub
txt_ConsigneeKey = myExCharFilter(txt_ConsigneeKey.Text)

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
str_SQL = "Select * From TRP01M Where storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' and ConsigneeKey = '" & txt_ConsigneeKey.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "��ƿ��~�G�Ȥ�s�� [" & txt_ConsigneeKey.Text & "] ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If
txt_FullName.Text = IIf(IsNull(tmp_Rs.Fields("Full_Name").Value), "", Trim(tmp_Rs.Fields("Full_Name").Value))
txt_ShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))
txt_Contact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_Phone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_Address.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
cmb_ZIP.ListIndex = -1
For iLoop = 0 To cmb_ZIP.ListCount - 1
   If arZip(iLoop) = Trim(tmp_Rs.Fields("ZIP").Value) Then
      cmb_ZIP.ListIndex = iLoop
      Exit For
   End If
Next iLoop
cmb_AreaCode.ListIndex = -1
For iLoop = 0 To cmb_AreaCode.ListCount - 1
    If arAreaCode(iLoop) = Trim(tmp_Rs.Fields("Area_Code").Value) Then
       cmb_AreaCode.ListIndex = iLoop
       Exit For
    End If
Next iLoop
cmb_ExtraDemand1.ListIndex = -1
For iLoop = 0 To cmb_ExtraDemand1.ListCount - 1
    If arExtraDemand(iLoop) = Trim(tmp_Rs.Fields("Extra_Demand_Code").Value) Then
       cmb_ExtraDemand1.ListIndex = iLoop
       Exit For
    End If
Next iLoop
cmb_ExtraDemand2.ListIndex = -1
For iLoop = 0 To cmb_ExtraDemand2.ListCount - 1
    If arExtraDemand(iLoop) = Trim(tmp_Rs.Fields("Extra_Demand_Code2").Value) Then
       cmb_ExtraDemand2.ListIndex = iLoop
       Exit For
    End If
Next iLoop
tmp_Rs.Close
End Sub

Private Sub txt_DeliveryDate_Click()
'�e�f��
If Trim(txt_DeliveryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2))
   End If
End If
mvDate.Top = fam_Orders.Top + txt_DeliveryDate.Top + txt_DeliveryDate.Height
mvDate.Left = fam_Orders.Left + txt_DeliveryDate.Left + txt_DeliveryDate.Width
mvDate.Tag = "�e�f��"
mvDate.Value = Now
mvDate.Visible = True
End Sub

Private Sub txt_OrderDate_Click()
'�q���
If Trim(txt_OrderDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_OrderDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_OrderDate.Text, 4) & "/" & Mid(txt_OrderDate.Text, 5, 2) & "/" & Right(txt_OrderDate.Text, 2))
   End If
End If
mvDate.Top = fam_Orders.Top + txt_OrderDate.Top + txt_OrderDate.Height
mvDate.Left = fam_Orders.Left + txt_OrderDate.Left + txt_OrderDate.Width
mvDate.Tag = "�q���"
mvDate.Value = Now
mvDate.Visible = True
End Sub

Private Sub txtLot_Change()

End Sub

Private Sub txtOrderCS_GotFocus()
txtOrderCS.SelStart = 0: txtOrderCS.SelLength = Len(txtOrderCS.Text)
End Sub

Private Sub txtOrdercs_KeyPress(KeyAscii As Integer)
'�q��Ӷ� >> �q��q
Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          txtOrderEA.SelStart = 0: txtOrderEA.SelLength = Len(txtOrderEA.Text): txtOrderEA.SetFocus
End Select

End Sub

Private Sub txt_QueryExternOrderKey_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then Call cmd_OrdersQuery_Click

End Sub

Private Sub txtOrderEA_GotFocus()
txtOrderEA.SelStart = 0: txtOrderEA.SelLength = Len(txtOrderEA.Text)
End Sub

Private Sub txtOrderea_KeyPress(KeyAscii As Integer)
'�q��Ӷ� >> �q��q
Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          cboLot6.SetFocus
End Select
End Sub

Private Sub txt_SKU_KeyPress(KeyAscii As Integer)
'�q��Ӷ���J >> �f��
If KeyAscii = vbKeyReturn Then
   txtOrderCS.SelStart = 0: txtOrderCS.SelLength = Len(txtOrderCS.Text): txtOrderCS.SetFocus
ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then '�p�g�r���אּ�j�g�r��
       KeyAscii = KeyAscii - 32
End If
End Sub

Private Function CheckSKU(ByVal strSku As String) As Integer
'�ˮֳf���O�_���T
CheckSKU = 1
If cmbStorerkey.Text = "" Then
   msg_text = "�f���ˮֲ��`�G�|����J [�f�D]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmbStorerkey.SetFocus
   Exit Function
End If

'���ҳf���O�_���T
str_SQL = "Select isnull(Rtrim(Descr),'') as 'Descr' From gv_SKUxpack Where StorerKey = '" & mySplit(cmbStorerkey.Text, " ", 0) & "' and SKU = '" & strSku & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�f�����ҿ��~�GStorer = [" & cmbStorerkey.Text & "] �L���f��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cboSku.SetFocus
   Exit Function
End If
txt_SkuDescr.Text = tmp_Rs.Fields("Descr").Value
CheckSKU = 0
tmp_Rs.Close

End Function

Private Sub txt_StorerKey_KeyPress(KeyAscii As Integer)
'�f�D
If KeyAscii >= 97 And KeyAscii <= 122 Then '�p�g�r���אּ�j�g�r��
   KeyAscii = KeyAscii - 32
End If
End Sub

Private Function CheckOrdersData() As Boolean

'��f����ˬd
rsMain.MoveFirst
Do While Not rsMain.EOF

    If Len(rsMain("�s�y��")) > 0 And Fun_ChkDateFormat(rsMain("�s�y��")) = 1 Then: MsgBox "�s�y��榡���~(YYYYMMDD)!", 16, Me.Caption: Exit Function
    If Len(rsMain("�����")) > 0 And Fun_ChkDateFormat(rsMain("�����")) = 1 Then: MsgBox "�����榡���~(YYYYMMDD)!", 16, Me.Caption: Exit Function

    '�������--�P�_SKU�O�_�s�b
    str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMain("�f��")) & "' and Storerkey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "�f�����s�b (" & Trim(rsMain("�f��")) & " )", 16, Me.Caption
        Exit Function
    End If

rsMain.MoveNext
Loop

'�q����@�@�~����s�ɮɡA�ˮ֭q���ƥ��T��
CheckOrdersData = False
If Trim(cmbStorerkey.Text) = "" Then
   msg_text = "��ƿ��~�G����J [�f�D] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   cmbStorerkey.Text = Trim(cmbStorerkey.Text)
End If
If Trim(cbo_Priority.Text) = "" Then
   msg_text = "��ƿ��~�G����J [�q�����O] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   cbo_Priority.Text = Trim(cbo_Priority.Text)
End If

If Fun_ChkDateFormat(Trim(txt_OrderDate.Text)) = 1 Then: MsgBox "�q���榡���~(YYYYMMDD)!", 16, Me.Caption: Exit Function
If Fun_ChkDateFormat(Trim(txt_DeliveryDate.Text)) = 1 Then: MsgBox "��f��榡���~(YYYYMMDD)!", 16, Me.Caption: Exit Function
If Trim(txt_OrderDate.Text) = "" Then
   msg_text = "��ƿ��~�G����J [�q���] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txt_OrderDate.Text = Trim(txt_OrderDate.Text)
End If
If Trim(txt_DeliveryDate.Text) = "" Then
   msg_text = "��ƿ��~�G����J [��f��] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txt_DeliveryDate.Text = Trim(txt_DeliveryDate.Text)
End If
If Trim(txt_ConsigneeKey.Text) = "" Then
   msg_text = "��ƿ��~�G����J [�Ȥ�s��] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txt_ConsigneeKey.Text = Trim(txt_ConsigneeKey.Text)
End If

If Trim(txtShipToKey.Text) = "" And mySplit(Trim(cbo_Priority), " ", 0) = "A2B" Then
   msg_text = "��ƿ��~�G����J [��B��f�Ȥ�s��] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txtShipToKey.Text = Trim(txtShipToKey.Text)
End If

If txt_Extern.Enabled = True Then
   If Trim(txt_Extern.Text) = "" Then
      msg_text = "��ƿ��~�G�s�W�q��ɥ�����J [�f�D�渹] ���"
      MsgBox msg_text, vbOKOnly + vbCritical, msg_title
      Exit Function
   Else
      txt_Extern.Text = Trim(txt_Extern.Text)
   End If
End If
If cmb_ZIP.ListIndex = -1 Then
   msg_text = "��ƿ��~�G�Ȥ�򥻸�ƥ������T�]�w [�l���ϸ�] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

'If dg_OrderDetail.Rows <= 2 Then
'   msg_text = "��ƿ��~�G�����s�W�q�� [���Ӹ��] ���"
'   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'   Exit Function
'End If

If rsMain Is Nothing Then
   msg_text = "��ƿ��~�G�����s�W�q�� [���Ӹ��] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

If rsMain.RecordCount = 0 Then
   msg_text = "��ƿ��~�G�����s�W�q�� [���Ӹ��] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

rsMain.MoveFirst
Do While Not rsMain.EOF
    If rsMain("�j���ƶq") < 0 Or rsMain("�����ƶq") < 0 Or rsMain("�p���ƶq") < 0 Then MsgBox "�ƶq�t�ƵL��!", 16, "�s��": Exit Function
    If rsMain("�j���ƶq") > 0 And rsMain("�j���J��") = 0 Then MsgBox "�j���J�Ƭ�0�A��J�j���ƶq�L��!", 16, "�s��": Exit Function
    If rsMain("�����ƶq") > 0 And rsMain("�����J��") = 0 Then MsgBox "�����J�Ƭ�0�A��J�����ƶq�L��!", 16, "�s��": Exit Function
    If rsMain("�j���ƶq") <> Int(rsMain("�j���J��")) And rsMain("�����ƶq") <> Int(rsMain("�����ƶq")) And rsMain("�p���ƶq") <> Int(rsMain("�p���ƶq")) Then MsgBox "�ƶq���঳�p���I!", 16, "�s��": Exit Function
    If rsMain("�j���ƶq") * rsMain("�j���J��") + rsMain("�����ƶq") * rsMain("�����J��") + rsMain("�p���ƶq") = 0 Then MsgBox "�q��q���ର0!", 16, "�s��": Exit Function
rsMain.MoveNext
Loop

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

'1.�ˮ֫Ȥ�s��
str_SQL = "Select Count(*) AS RecCount From TRP01M Where storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' and ConsigneeKey = '" & txt_ConsigneeKey.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCount").Value = 0 Then
   tmp_Rs.Close
   msg_text = "��ƿ��~�G�Ȥ�s�� [" & txt_ConsigneeKey.Text & "] ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
tmp_Rs.Close

'�O�_�w�Ƹ��s
tmp_Rs.Open "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' ", cn
If Not tmp_Rs.EOF Then MsgBox "���q��w�Ƹ��u�s�� " & tmp_Rs("route_no") & " �A�L�k�ק�!", vbOKOnly, "�s��": tmp_Rs.Close: Exit Function
tmp_Rs.Close

If mySplit(Trim(cmbStorerkey), " ", 0) = "LTKK01" Then

    '�x�W�Q��q�歫���ˬd
    If txt_QueryExternOrderKey.Enabled = True Then
       '�s�W�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and consigneekey = '" & txt_ConsigneeKey.Text & "' and convert(varchar(8),deliverydate,112) = '" & txt_DeliveryDate.Text & "' and rtrim(isnull(type,'')) <> '�R��' and priority ='" & mySplit(Trim(cbo_Priority.Text), " ", 0) & "' and cast(notes as varchar(300)) = '" & Trim(txt_Description) & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�x�W�Q��q�歫��!(�ۦP�Ȥ�s���B��f��B�q�����O�P�q��Ƶ�)", 64, "��ƿ��~": tmp_Rs.Close: Exit Function
    Else
        '�ק�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and consigneekey = '" & txt_ConsigneeKey.Text & "' and convert(varchar(8),deliverydate,112) = '" & txt_DeliveryDate.Text & "' and rtrim(isnull(type,'')) <> '�R��' and priority ='" & mySplit(Trim(cbo_Priority.Text), " ", 0) & "' and cast(notes as varchar(300)) = '" & txt_Description & "' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�x�W�Q��q�歫��!(�ۦP�Ȥ�s���B��f��B�q�����O�P�q��Ƶ�)", 64, "��ƿ��~": tmp_Rs.Close: Exit Function
    End If

ElseIf mySplit(Trim(cmbStorerkey), " ", 0) = "LNIP01" Then
    If txt_QueryExternOrderKey.Enabled = True Then
       '�s�W�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�x�W�ߨ��q�渹�X�����б���(���\����)�A�нT�{�q���ƵL�~�I", 64, "�`�N"
    Else
        '�ק�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�x�W�ߨ��q�渹�X�����б���(���\����)�A�нT�{�q���ƵL�~�I", 64, "�`�N"
    End If

ElseIf mySplit(Trim(cmbStorerkey), " ", 0) = "LABT01" Then
    If txt_QueryExternOrderKey.Enabled = True Then
       '�s�W�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�Ȱ��q�渹�X����!", 64, "��ƿ��~": tmp_Rs.Close: Exit Function
    Else
        '�ק�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�Ȱ��q�渹�X�����б���(���\����)�A�нT�{�q���ƵL�~�I�åB�t�@�f�D�渹��ưO�o�ץ��I", 64, "�`�N"
    End If
    
Else
     '��L�f�D�q�歫���ˬd
    If txt_QueryExternOrderKey.Enabled = True Then
       '�s�W�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�f�D�q�渹�X����!", 64, "��ƿ��~": tmp_Rs.Close: Exit Function
    Else
        '�ק�q���
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "�f�D�q�渹�X����!", 64, "��ƿ��~": tmp_Rs.Close: Exit Function
    End If
    
End If
tmp_Rs.Close

' mark by gemini
''3.�ק�G�ˮ֪��A�A�f�D�渹���D�A�Ҧ��q�楲����� [�ݱƨ����A] �~�i�i��ק�
'If txt_Extern.Enabled = False Then
'   str_SQL = "Select Count(*) as RecCount From TRP02T Where StorerKey = '" & txt_StorerKey.Text & "' and Extern = '" & txt_Extern.Text & "'"
'   tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'   If tmp_rs.Fields("RecCount").Value > 0 Then
'      tmp_rs.Close
'      msg_text = "��ƿ��~�G�f�D�q�� [" & txt_Extern.Text & "] ���A�㦡 [�w�ƨ�]�A�����\�ק�"
'      MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'      Exit Function
'   End If
'End If

''4.�Ӷ�������ҡA�p�� [�c��]�B[�O��]�B[���n]�B[���q]
'Dim strSKU As String, strErrorSKU As String
'Dim dbShipQty As Double
'strErrorSKU = ""
'With dg_OrderDetail
'     If .Rows = 2 Then Exit Function
'     For iLoop = 1 To .Rows - 2
'         .Row = iLoop
'         .Col = 1: strSKU = .Text
'         If CheckSKU(.Text) = 1 Then Exit Function
'         .Col = 4: dbShipQty = Val(.Text)
'         str_SQL = "Select * " & _
'                   "From BaseData_SKUPacking Where StorerKey = '" & cmbStorerkey.Text & "' and SKU = '" & strSKU & "'"
'         tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'         If tmp_rs.EOF Then
'            If strErrorSKU = "" Then
'               strErrorSKU = strSKU
'            Else
'               strErrorSKU = strErrorSKU & "," & strSKU
'            End If
'         Else
'
'            .Col = 5
'            If tmp_rs.Fields("�O�ഫ").Value = 0 Then
'            .Text = 0
'            Else
'            .Text = NumRound((dbShipQty / tmp_rs.Fields("�O�ഫ").Value), 2)
'            End If
'
'            .Col = 6: .Text = NumRound((dbShipQty * tmp_rs.Fields("�����n").Value), 2)
'            .Col = 7: .Text = NumRound((dbShipQty * tmp_rs.Fields("��쭫�q").Value), 2)
'         End If
'         tmp_rs.Close
'     Next iLoop
'End With

'Terry 20181217 user�n�D�}��C��
''�q�����O�ˬd�AC��u������B�Q��M�Ȱ��ϥ�
'If mySplit(Trim(cbo_Priority), " ", 0) = "C" Then
'    If Left(RTrim(cmbStorerkey.Text), 6) <> "LKAO01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LABT01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LTKK01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LCHF01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LPSI01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LNVA01" Then
'        msg_text = "��ƿ��~�G [�q�����O]=C�A�ثe���w���Ȱ��B�Q��B�����B�ʨơB�g�����Ϊ���f�D�ϥ�"
'        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'        Exit Function
'    End If
'End If

CheckOrdersData = True
Exit Function

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q����@-�s��-����ˮ�", Me.Caption, "Form SubProgram CheckOrdersData", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Function

Private Function SaveOrdersData() As Boolean
'�q���Ʀs��
On Error GoTo err_Handle
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
SaveOrdersData = False
Dim rsTmp As New ADODB.Recordset, intTmp As Long, strConsigneeKey As String, str_receive As String

If Chk_receive.Value = 1 Then str_receive = 1 Else str_receive = 0

'��Ʈw���ʥ��--�_�I
'Tran_Level = cn.BeginTrans

Dim strOrderkey As String

    '1.���s���q��s��
    str_SQL = "select isnull(max(orderkey),0) from orders"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strOrderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
    tmp_Rs.Close
    
    Dim intPointer As Integer
    intPointer = 1
    
    '3.�q����Y��Ʀs��
    If Left(RTrim(cmbStorerkey.Text), 6) = "LMBO01" And mySplit(Trim(cbo_Priority), " ", 0) = "A2B" Then
    '���_A2B�q��A�q���T��B�I���C��L�f�D���OA�I��
    Call txtShipToKey_LostFocus
        str_SQL = "Insert into Orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_Contact1,C_Company,C_Address1,C_Address2, " & _
                 " C_ZIP,C_Phone1,AddWho,EditWho,DoRoute,Notes,customerorderkey,type,b_company,facility,externconsigneekey,updatesource,b_city,otqty,GoodsBack) Values ('" & _
                 strOrderkey & "','" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & txt_Extern & "','" & _
                 Left(txt_OrderDate.Text, 4) & "/" & Mid(txt_OrderDate.Text, 5, 2) & "/" & Right(txt_OrderDate.Text, 2) & "','" & _
                 Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "'," & _
                 "'" & mySplit(Trim(cbo_Priority), " ", 0) & "','" & txt_ConsigneeKey.Text & "','" & txt_ShipToContact.Text & "','" & txt_ShipToShortName.Text & "','" & myExCharFilter(Trim(GetWord(txt_ShipToAddress.Text, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(txt_ShipToAddress.Text, intPointer, 45))) & "','" & Trim(txt_ShipToZIP.Text) & "','" & _
                 txt_ShipToPhone.Text & "','" & User_id & "','" & User_id & "','Y','" & txt_Description.Text & "','" & UCase(txt_OrderKey.Text) & "','','" & txtShipToKey.Text & "','" & cmdFacility & "','" & txt_ConsigneeKey.Text & "','ManualOrder','" & RTrim(txt_B_city) & "','" & txt_OtQty & "','" & str_receive & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    Else
        str_SQL = "Insert into Orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_Contact1,C_Company,C_Address1,C_address2, " & _
                 " C_ZIP,C_Phone1,AddWho,EditWho,DoRoute,Notes,customerorderkey,type,b_company,facility,externconsigneekey,updatesource,b_city,otqty,GoodsBack) Values ('" & _
                 strOrderkey & "','" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & txt_Extern & "','" & _
                 Left(txt_OrderDate.Text, 4) & "/" & Mid(txt_OrderDate.Text, 5, 2) & "/" & Right(txt_OrderDate.Text, 2) & "','" & _
                 Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "'," & _
                 "'" & mySplit(Trim(cbo_Priority), " ", 0) & "','" & txt_ConsigneeKey.Text & "','" & txt_Contact.Text & "','" & txt_ShortName.Text & "','" & myExCharFilter(Trim(GetWord(txt_Address.Text, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(txt_Address.Text, intPointer, 45))) & "','" & arZip(cmb_ZIP.ListIndex) & "','" & _
                 txt_Phone.Text & "','" & User_id & "','" & User_id & "','Y','" & txt_Description.Text & "','" & UCase(txt_OrderKey.Text) & "','','" & txtShipToKey.Text & "','" & cmdFacility & "','" & txt_ConsigneeKey.Text & "','ManualOrder','" & RTrim(txt_B_city) & "','" & txt_OtQty & "','" & str_receive & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    End If
    
'�O�_�O�ק�
If txt_QueryExternOrderKey.Enabled = False And txt_QueryExternOrderKey <> strOrderkey Then

'���_�q��ק�A�h�NTMS�渹��s�^Custorders��
 If RTrim(cmbStorerkey) = "LMBO01" Then
    cn.Execute "update custorders set orderkey = '" & strOrderkey & "' where orderkey = '" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords
    cn.Execute "update custorderdetail set orderkey = '" & strOrderkey & "' where orderkey = '" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords
 End If
 
'����l�q����
    str_SQL = "select externordertype = isnull(o.externordertype,''), amount = isnull(o.amount,0) , billtokey=isnull(o.billtokey,'') , b_contact1 = isnull(o.b_contact1,'') ,  door = isnull(o.door,'') , stop = isnull(o.stop,'') , facility = isnull(o.facility,'') , o.invoiceno,o.externconsigneekey,o.b_vat,ordergroup = isnull(o.ordergroup,''),BuyerPo = isnull(o.BuyerPo,''),b_city = isnull(o.b_city,''),o.otqty " & _
    ", Cash = o.Cash " & _
    ", Bill = o.Bill " & _
    ", ReceiveCash = o.ReceiveCash " & _
    ", ReceiveBill = o.ReceiveBill " & _
    ", PayStatus = o.PayStatus " & _
    ", B_Contact2 = isnull(o.B_Contact2,'') " & _
    ",o.urgent_mark,o.reserve_mark , od.* " & _
    "from orders o join orderdetail od on o.orderkey = od.orderkey where o.orderkey = '" & txt_QueryExternOrderKey & "' "
    
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

'��s��L���OTQTY,B_City, GoodsBack �{�b��ܵe���W�A�ҥH�εe���W�����D�C�o�䤣�έ쥻����ư���s

    str_SQL = "update orders set externordertype = '" & RTrim(tmp_Rs("externordertype")) & "'" & _
              ", amount = " & tmp_Rs("amount") & _
              ", billtokey = '" & tmp_Rs("billtokey") & "' " & _
              ", b_contact1 = '" & tmp_Rs("b_contact1") & "' " & _
              ", door = '" & tmp_Rs("door") & "' " & _
              ", stop = '" & tmp_Rs("stop") & "' " & _
              ", b_vat = '" & tmp_Rs("b_vat") & "' " & _
              ", Updatesource = '" & txt_QueryExternOrderKey & "' " & _
              ", invoiceno = '" & tmp_Rs("invoiceno") & "' " & _
              ", externconsigneekey = '" & tmp_Rs("externconsigneekey") & "' " & _
              ", ordergroup = '" & tmp_Rs("ordergroup") & "' " & _
              ", BuyerPo = '" & tmp_Rs("BuyerPo") & "' " & _
              ", Cash = " & tmp_Rs("Cash") & _
              ", Bill = " & tmp_Rs("Bill") & _
              ", ReceiveCash = " & tmp_Rs("ReceiveCash") & _
              ", ReceiveBill = " & tmp_Rs("ReceiveBill") & _
              ", PayStatus = '" & tmp_Rs("PayStatus") & "' " & _
              ", B_Contact2= '" & tmp_Rs("B_Contact2") & "' " & _
              ",urgent_mark= '" & tmp_Rs("urgent_mark") & "' " & _
              ",reserve_mark= '" & tmp_Rs("reserve_mark") & "' " & _
              "where orderkey = '" & strOrderkey & "' "
'    tmp_Rs.Close
    
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
End If

'4.�B�z�q��Ӷ��s��
Dim strLineNo As String, strSku As String, dbOrderCS As Double, dbOrderIP As Double, dbOrderEA As Double, dbCasecnt As Double, dbInnerpack As Double, strTKLOC As String, strNotes As String, strLot3 As String, strLot4 As String, strLot5 As String

rsMain.MoveFirst
Do While Not rsMain.EOF
    strLineNo = Trim(myExCharFilter(rsMain("����")))
    strSku = Trim(myExCharFilter(rsMain("�f��")))        '�f��
    strTKLOC = Trim(myExCharFilter(rsMain("�ܧO")))    'TKLOC
    dbOrderCS = Trim(myExCharFilter(rsMain("�j���ƶq")))      '�c��
    dbOrderIP = Trim(myExCharFilter(rsMain("�����ƶq")))      '�Ӽ�
    dbOrderEA = Trim(myExCharFilter(rsMain("�p���ƶq")))      '�Ӽ�
    dbCasecnt = Trim(myExCharFilter(rsMain("�j���J��")))      '�C�c
    dbInnerpack = Trim(myExCharFilter(rsMain("�����J��")))      '�C��
    strNotes = Trim(myExCharFilter(rsMain("�Ƶ�")))      '�Ƶ�
    strLot3 = Trim(myExCharFilter(rsMain("�Ͳ��帹")))      '�Ͳ��帹
    strLot4 = Trim(myExCharFilter(rsMain("�s�y��")))      '�s�y��
    strLot5 = Trim(myExCharFilter(rsMain("�����")))      '�����
    
    '�s�W�Ginsert into OrderDetail
    str_SQL = "insert into OrderDetail(StorerKey,OrderKey,OrderLineNumber,ExternOrderKey,SKU,OriginalQty,OpenQty,ShippedQty,UOM,AddWho,EditWho,lottable06,notes,updatesource,lottable03,lottable04,lottable05) Values ('" & _
              mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & strOrderkey & "','" & strLineNo & "','" & txt_Extern & "','" & strSku & "'," & (dbOrderCS * dbCasecnt) + (dbOrderIP * dbInnerpack) + dbOrderEA & "," & (dbOrderCS * dbCasecnt) + (dbOrderIP * dbInnerpack) + dbOrderEA & ",0" & _
              ",'EA','" & User_id & "','" & User_id & "','" & UCase(strTKLOC) & "','" & strNotes & "','" & Trim(myExCharFilter(rsMain("�z�f�[�u"))) & "','" & strLot3 & "','" & strLot4 & "','" & strLot5 & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�L�s�y��h��s��null
    If Len(Trim(strLot4)) = 0 Then cn.Execute "update orderdetail set lottable04 = null where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' ", RowsAffect, adExecuteNoRecords
    
    '�L�����h��s��null
    If Len(Trim(strLot5)) = 0 Then cn.Execute "update orderdetail set lottable05 = null where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' ", RowsAffect, adExecuteNoRecords

'�q��ק�ɸɩ��Ӹ��
If txt_QueryExternOrderKey.Enabled = False And txt_QueryExternOrderKey <> strOrderkey Then

    '�ɭq����Ӹ��
    tmp_Rs.MoveFirst
    Do While Not tmp_Rs.EOF
        If strLineNo = RTrim(tmp_Rs("orderlinenumber")) Then
            cn.Execute "update orderdetail set ExternLineNo = '" & tmp_Rs("ExternLineNo") & "' " & _
            ",RetailSKU = '" & tmp_Rs("RetailSKU") & "' " & _
            ",PickCode = '" & tmp_Rs("PickCode") & "' " & _
            ",Facility = '" & tmp_Rs("Facility") & "' " & _
            ",UnitPrice = '" & tmp_Rs("UnitPrice") & "' " & _
            ",OtherUOM = '" & tmp_Rs("OtherUOM") & "' " & _
            ",UOM = '" & tmp_Rs("UOM") & "' " & _
            "where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' ", RowsAffect, adExecuteNoRecords
        End If

    tmp_Rs.MoveNext
    Loop
        
End If

rsMain.MoveNext
Loop

'�q��ק��Close
If txt_QueryExternOrderKey.Enabled = False And txt_QueryExternOrderKey <> strOrderkey Then tmp_Rs.Close

'��spackkey
str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & strOrderkey & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�ɫȤ��� LABT����ƥN�ɻ�
cn.Execute "exec gs_OrdersUpdate '" & mySplit(cmbStorerkey.Text, " ", 0) & "' ", RowsAffect, adExecuteNoRecords

'7. DB Transaction Commit
'cn.CommitTrans: Tran_Level = 0

'8.�M���ù�
Call Clear_AllField
txt_QueryExternOrderKey.Text = strOrderkey
SaveOrdersData = True
txt_QueryExternOrderKey.Enabled = True
Exit Function

err_Handle:
   If Tran_Level <> 0 Then cn.RollbackTrans: Tran_Level = 0

   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q����@-�s��-��Ʀs��", Me.Caption, "Form SubProgram SaveOrdersData", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Function

Private Sub txtShipToKey_GotFocus()
    frm_OP_ManualShipToOrders.ZOrder 0
    fam_ConsigneeData.Enabled = False
    Call txtShipToKey_LostFocus
End Sub

Public Sub txtShipToKey_LostFocus()

'��B��f�Ȥ�s��
If Trim(txtShipToKey.Text) = "" Then Exit Sub

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
str_SQL = "Select * From TRP01M Where storerkey = '" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "' and ConsigneeKey = '" & txtShipToKey.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "��ƿ��~�G�Ȥ�s�� [" & txtShipToKey.Text & "] ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If

txt_ShipToFullName.Text = IIf(IsNull(tmp_Rs.Fields("Full_Name").Value), "", Trim(tmp_Rs.Fields("Full_Name").Value))
txt_ShipToShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))
txt_ShipToContact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_ShipToPhone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_ShipToAddress.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
txt_ShipToZIP = tmp_Rs("zip") & ""
txt_ShipToAreaCode = tmp_Rs("Area_Code") & ""
txt_ShipToExtraDemand1 = tmp_Rs("Extra_Demand_code2") & ""
txt_ShipToExtraDemand2 = tmp_Rs("Extra_Demand_code2") & ""

tmp_Rs.Close

End Sub

Private Sub cboLot6_KeyPress(KeyAscii As Integer)
'�p�g�r���אּ�j�g�r��
If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

