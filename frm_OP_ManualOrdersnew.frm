VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_OP_ManualOrders 
   Caption         =   "�q����@�@�~ "
   ClientHeight    =   8310
   ClientLeft      =   240
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11550
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   4080
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   4410
      _ExtentX        =   7779
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
      StartOfWeek     =   72810497
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame fam_Orders 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   3195
      Left            =   0
      TabIndex        =   30
      Top             =   840
      Width           =   11520
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
         TabIndex        =   101
         Top             =   960
         Width           =   9180
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
         ItemData        =   "frm_OP_ManualOrders.frx":0000
         Left            =   960
         List            =   "frm_OP_ManualOrders.frx":0002
         TabIndex        =   80
         Top             =   600
         Width           =   1575
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
         Left            =   5400
         TabIndex        =   60
         Top             =   1260
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
         Left            =   7065
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   59
         Top             =   1245
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
         ItemData        =   "frm_OP_ManualOrders.frx":0004
         Left            =   960
         List            =   "frm_OP_ManualOrders.frx":0006
         TabIndex        =   58
         Top             =   285
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
         TabIndex        =   57
         Top             =   600
         Width           =   975
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
         TabIndex        =   56
         Top             =   285
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
         TabIndex        =   55
         Top             =   1245
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
         TabIndex        =   54
         Top             =   285
         Width           =   1575
      End
      Begin VB.Frame fam_ConsigneeData 
         BackColor       =   &H00004040&
         BorderStyle     =   0  '�S���ؽu
         Height          =   1470
         Left            =   120
         TabIndex        =   35
         Top             =   1605
         Width           =   10000
         Begin VB.ComboBox cmb_ExtraDemand2 
            BackColor       =   &H8000000B&
            Height          =   300
            Left            =   5235
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   45
            Top             =   1125
            Width           =   4230
         End
         Begin VB.ComboBox cmb_ExtraDemand1 
            BackColor       =   &H8000000A&
            Height          =   300
            Left            =   1005
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   60
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
         TabIndex        =   34
         Top             =   1260
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
         TabIndex        =   33
         Top             =   285
         Width           =   975
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
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtFacility 
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
         Left            =   6015
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   31
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame frm_OP_ManualShipToOrders 
         BackColor       =   &H00004040&
         BorderStyle     =   0  '�S���ؽu
         Enabled         =   0   'False
         Height          =   1470
         Left            =   120
         TabIndex        =   61
         Top             =   1605
         Width           =   10000
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
            Top             =   255
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
            TabIndex        =   64
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   79
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
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
            Top             =   585
            Width           =   390
         End
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
         TabIndex        =   102
         Top             =   1020
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
         TabIndex        =   90
         Top             =   660
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
         Left            =   3720
         TabIndex        =   89
         Top             =   1305
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
         TabIndex        =   88
         Top             =   345
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q�渹�X"
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
         TabIndex        =   87
         Top             =   345
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
         TabIndex        =   86
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��f��"
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
         TabIndex        =   85
         Top             =   660
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
         TabIndex        =   84
         Top             =   345
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
         TabIndex        =   83
         Top             =   345
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
         TabIndex        =   82
         Top             =   660
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
         TabIndex        =   81
         Top             =   660
         Width           =   780
      End
   End
   Begin VB.Frame fam_Header 
      Height          =   870
      Left            =   0
      TabIndex        =   17
      Top             =   -75
      Width           =   11520
      Begin VB.ComboBox cboKey 
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
         ItemData        =   "frm_OP_ManualOrders.frx":0008
         Left            =   120
         List            =   "frm_OP_ManualOrders.frx":000A
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "�渹���O"
         Top             =   300
         Width           =   1335
      End
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
         Left            =   5040
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   240
         Width           =   960
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
         Left            =   6120
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
         Top             =   240
         Width           =   960
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
         Left            =   8280
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   240
         Width           =   960
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
         Left            =   7200
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
         Top             =   240
         Width           =   960
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
         Left            =   10440
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   16
         Top             =   240
         Width           =   960
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
         Left            =   9360
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         Top             =   240
         Width           =   960
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
         Left            =   4200
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txt_QueryExternOrderKey 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   1
         Top             =   300
         Width           =   2550
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000080&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Height          =   585
         Left            =   4995
         Top             =   195
         Width           =   6480
      End
   End
   Begin VB.Frame fam_OrderDetail 
      BackColor       =   &H8000000B&
      Height          =   4365
      Left            =   0
      TabIndex        =   18
      Top             =   3960
      Width           =   11520
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
         Picture         =   "frm_OP_ManualOrders.frx":000C
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         Left            =   120
         TabIndex        =   19
         Top             =   120
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
            TabIndex        =   99
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
            TabIndex        =   97
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
            TabIndex        =   95
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
            TabIndex        =   93
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
            TabIndex        =   91
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
            ItemData        =   "frm_OP_ManualOrders.frx":08D6
            Left            =   3480
            List            =   "frm_OP_ManualOrders.frx":08D8
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
            TabIndex        =   9
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
            ItemData        =   "frm_OP_ManualOrders.frx":08DA
            Left            =   600
            List            =   "frm_OP_ManualOrders.frx":08DC
            TabIndex        =   8
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
            TabIndex        =   100
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
            TabIndex        =   98
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
            TabIndex        =   96
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
            TabIndex        =   94
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
            TabIndex        =   92
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
         TabIndex        =   15
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
str_SQL = "select distinct isnull(lotattribute.lottable06,'') as lottable06 from wms..lotxlocxid lotxlocxid join wms..lotattribute lotattribute on lotattribute.lot = lotattribute.lot where lotxlocxid.storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' and lotattribute.sku = '" & cboSku & "' order by lotattribute.lottable06 "
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

'���f��
Dim rsTmp As New ADODB.Recordset
str_SQL = "select sku = rtrim(sku) from sku where storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' order by sku"
rsTmp.Open str_SQL, cn
If rsTmp.EOF Then MsgBox "�䤣��ӳf�D�ӫ~�D�ɸ��", vbOKOnly, Me.Caption: Exit Sub
rsTmp.MoveFirst

cboSku.Clear
Do While Not rsTmp.EOF
    cboSku.AddItem rsTmp("sku")
    rsTmp.MoveNext
Loop

rsTmp.Close: Set rsTmp = Nothing

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
Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Consigneelist.Name)
End Sub

Private Sub cmd_Delete_Click()
'�q�� >> �R��
If Len(RTrim(txt_Extern)) = 0 Or txtType = "�R��" Then Exit Sub

Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select * from orders where orderkey = '" & txt_QueryExternOrderKey & "' ", cn
If rsTmp.EOF Then MsgBox "�䤣�즹�q��!", vbOKOnly, "�q��R��": rsTmp.Close: Exit Sub
rsTmp.Close

rsTmp.Open "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' ", cn
If Not rsTmp.EOF Then MsgBox "�w�Ƹ��u�s��" & rsTmp("route_no") & "�A�q��L�k�R��!", vbOKOnly, "�q��R��": rsTmp.Close: Exit Sub
rsTmp.Close

If MsgBox("�T�w�R�����q��(�t���έq��)? ", vbQuestion + vbYesNo, "�q��R��") <> vbYes Then Exit Sub

Call DB_CheckConnectStatus

Tran_Level = cn.BeginTrans
     
    cn.Execute "delete TRP02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP03T where receipt_no in (select receipt_no from trp02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='�R��' ,editdate = getdate(),editwho = '" & User_id & "' where orderkey='" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0
txtType = "�R��"
cmbStorerkey.Enabled = True

'LTKK01�R��۰� Mail �q��
If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then Call SendMail(txt_QueryExternOrderKey)

Exit Sub

err_Handle:
If Tran_Level <> 0 Then Tran_Level = 0: cn.RollbackTrans

Dim tmpString As String
msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
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

'�t�e�ܧO
If Len(Trim(txtFacility)) = 0 Then
    txtFacility = "�ըƹF�_��"
    If UCase(Right(Trim(cboLot6), 2)) = "-C" Then txtFacility = "�ըƹF����"
    If UCase(Right(Trim(cboLot6), 2)) = "-S" Then txtFacility = "�ըƹF�n��"
Else
    If UCase(Right(Trim(cboLot6), 2)) <> "-C" And UCase(Right(Trim(cboLot6), 2)) <> "-S" And txtFacility <> "�ըƹF�_��" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": cboLot6.SetFocus: Exit Sub
    If UCase(Right(Trim(cboLot6), 2)) = "-C" And txtFacility <> "�ըƹF����" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": cboLot6.SetFocus: Exit Sub
    If UCase(Right(Trim(cboLot6), 2)) = "-S" And txtFacility <> "�ըƹF�n��" Then: MsgBox "�t�e�ܧO�P�ӫ~�ܧO����!!", 64, "�ܧO���~!!": cboLot6.SetFocus: Exit Sub

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

Private Sub cmd_DetailVerify_Click()
'�Ӷ�����
If dg_OrderDetail.Rows = 2 Then Exit Sub
If Trim(cmbStorerkey.Text) = "" Then
   msg_text = "����J [�f�D] ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Dim strSKU As String, strErrorSKU As String
Dim dbShipQty As Double

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
strErrorSKU = ""
With dg_OrderDetail
     If .Rows = 2 Then Exit Sub
     For iLoop = 1 To .Rows - 2
         .Row = iLoop
         .Col = 1: strSKU = .Text
         str_SQL = "Select *" & _
                   "From BaseData_SKUPacking Where StorerKey = '" & cmbStorerkey.Text & "' and SKU = '" & strSKU & "'"
         tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
         If tmp_rs.EOF Then
            If strErrorSKU = "" Then
               strErrorSKU = strSKU
            Else
               strErrorSKU = strErrorSKU & "," & strSKU
            End If
         Else
            .Col = 3: .Text = tmp_rs.Fields("�~�W").Value
'            .Col = 4: dbShipQty = .Text
'            .Col = 5: .Text = NumRound((dbShipQty / tmp_rs.Fields("�O�ഫ").Value), 2)
'            .Col = 6: .Text = NumRound((dbShipQty * tmp_rs.Fields("�����n").Value), 2)
'            .Col = 7: .Text = NumRound((dbShipQty * tmp_rs.Fields("��쭫�q").Value), 2)
         End If
         tmp_rs.Close
     Next iLoop
End With
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_Modify_Click()
'�ק�
If Trim(txt_Extern.Text) = "" Then Exit Sub

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

Dim strTmp As String, strTmp1 As String
'strTmp = Format(txt_QueryExternOrderKey.Text, "0000000000")
strTmp = txt_QueryExternOrderKey.Text
strTmp1 = cboKey.Text

'�M���Ҧ����ȡA�]�t OrderDetail ����
Call Clear_AllField
fam_Orders.Caption = ""
txt_QueryExternOrderKey.Text = strTmp
cboKey.Text = strTmp1

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass

'���o�q�� Header
If cboKey.Text = "" Then
    str_SQL = "Select �f�D�渹,�f�D,�q���,�e�f��,�Ȥ�s��,����,�Ȥ�W��,�Ȥ�²��,�l���ϸ�,�B�e�ϰ�,�S��ݨD1,�S��ݨD2,�B�e�a�},�p���H,�q��,TMS�渹,b_phone2,�Ȥ�渹,�q�����O,��B��f�Ȥ�s��,�q�檬�A,�t�e�ܧO " & _
    "From ManualOrder_Orders Where (TMS�渹 = '" & txt_QueryExternOrderKey.Text & "' or �f�D�渹 = '" & txt_QueryExternOrderKey.Text & "') "
ElseIf cboKey.Text = "TMS�渹" Then
    str_SQL = "Select �f�D�渹,�f�D,�q���,�e�f��,�Ȥ�s��,����,�Ȥ�W��,�Ȥ�²��,�l���ϸ�,�B�e�ϰ�,�S��ݨD1,�S��ݨD2,�B�e�a�},�p���H,�q��,TMS�渹,b_phone2,�Ȥ�渹,�q�����O,��B��f�Ȥ�s��,�q�檬�A,�t�e�ܧO " & _
    "From ManualOrder_Orders Where TMS�渹 = '" & Format(txt_QueryExternOrderKey.Text, "0000000000") & "' "
    txt_QueryExternOrderKey = Format(txt_QueryExternOrderKey.Text, "0000000000")
Else
    str_SQL = "Select �f�D�渹,�f�D,�q���,�e�f��,�Ȥ�s��,����,�Ȥ�W��,�Ȥ�²��,�l���ϸ�,�B�e�ϰ�,�S��ݨD1,�S��ݨD2,�B�e�a�},�p���H,�q��,TMS�渹,b_phone2,�Ȥ�渹,�q�����O,��B��f�Ȥ�s��,�q�檬�A,�t�e�ܧO " & _
    "From ManualOrder_Orders Where �f�D�渹 = '" & txt_QueryExternOrderKey.Text & "' "

End If

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
tmp_rs.CursorLocation = adUseClient
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_Modify.Enabled = False:   Screen.MousePointer = vbDefault
   Exit Sub
End If

If tmp_rs.RecordCount > 1 Then MsgBox "�ۦP�渹�ⵧ��ơA�п�ܳ渹���O!!", 64, "�q��d��": cmd_Modify.Enabled = False:   Screen.MousePointer = vbDefault: Exit Sub

'if tmp_rs("b_phone2") = 00 then label1.Caption ="�w��J�ƨ��t�ΡA�L�k�ܧ�!!"
txt_Extern.Text = tmp_rs.Fields("�f�D�渹").Value
txt_OrderKey.Text = tmp_rs.Fields("�Ȥ�渹").Value
cmbStorerkey.Text = tmp_rs.Fields("�f�D").Value
txt_OrderDate.Text = tmp_rs.Fields("�q���").Value
txt_DeliveryDate.Text = tmp_rs.Fields("�e�f��").Value
txt_ConsigneeKey.Text = tmp_rs.Fields("�Ȥ�s��").Value
txt_Description.Text = tmp_rs.Fields("����").Value
txt_FullName.Text = tmp_rs.Fields("�Ȥ�W��").Value
txt_Contact.Text = tmp_rs.Fields("�p���H").Value
txt_Phone.Text = tmp_rs.Fields("�q��").Value
txt_Address.Text = tmp_rs.Fields("�B�e�a�}").Value
cbo_Priority.Text = tmp_rs.Fields("�q�����O").Value
txtShipToKey.Text = tmp_rs("��B��f�Ȥ�s��") & ""
txtType.Text = tmp_rs("�q�檬�A")
txtFacility.Text = tmp_rs("�t�e�ܧO")
fam_Orders.Caption = tmp_rs("TMS�渹")

If RTrim(cbo_Priority) = "A2B" Then Label3(21).Visible = True: txtShipToKey.Visible = True: cmdShipToList.Visible = True

If Len(RTrim(tmp_rs("�S��ݨD1"))) > 0 Then
    For iLoop = 0 To cmb_ExtraDemand1.ListCount - 1
        If arExtraDemand(iLoop) = tmp_rs.Fields("�S��ݨD1").Value Then
           cmb_ExtraDemand1.ListIndex = iLoop
           Exit For
        End If
    Next iLoop
End If

If Len(RTrim(tmp_rs("�S��ݨD2"))) > 0 Then
    For iLoop = 0 To cmb_ExtraDemand2.ListCount - 1
        If arExtraDemand(iLoop) = tmp_rs.Fields("�S��ݨD2").Value Then
           cmb_ExtraDemand2.ListIndex = iLoop
           Exit For
        End If
    Next iLoop
End If

txt_ShortName.Text = tmp_rs.Fields("�Ȥ�²��").Value

For iLoop = 0 To cmb_ZIP.ListCount - 1
    If arZip(iLoop) = tmp_rs.Fields("�l���ϸ�").Value Then
       cmb_ZIP.ListIndex = iLoop
       Exit For
    End If
Next iLoop
DoEvents: DoEvents

For iLoop = 0 To cmb_AreaCode.ListCount - 1
    If arAreaCode(iLoop) = tmp_rs.Fields("�B�e�ϰ�").Value Then
       cmb_AreaCode.ListIndex = iLoop
       Exit For
    End If
Next iLoop
tmp_rs.Close

'���o�q�� Detail >> �H OrderDetail ���D
str_SQL = "Select �f�D�渹,����,�f��,�~�W,�q��q,�z�f�q,�C�c,�z�f�c��,�z�f�O��,�z�f���n,�z�f���q,�ܧO,�c��,�Ӽ�,�Ƶ�,�Ͳ��帹,�s�y��,����� " & _
          "From ManualOrder_OrderDetail Where TMS�渹 = '" & fam_Orders.Caption & "' Order by ����"
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q����Ӹ��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_Modify.Enabled = False
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Do While Not tmp_rs.EOF
   With dg_OrderDetail
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0: .Text = RTrim(tmp_rs("����"))
       .Col = 1: .Text = tmp_rs("�f��")
       .Col = 2: .Text = tmp_rs("�~�W")
       .Col = 3: .Text = tmp_rs("�ܧO")
       .Col = 4: .Text = tmp_rs("�c��")
       .Col = 5: .Text = tmp_rs("�Ӽ�")
       .Col = 6: .Text = tmp_rs("�C�c")
       .Col = 7: .Text = tmp_rs("�Ƶ�")
       .Col = 8: .Text = tmp_rs("�Ͳ��帹")
       .Col = 9: .Text = tmp_rs("�s�y��")
       .Col = 10: .Text = tmp_rs("�����")

  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close

If txtType <> "�R��" Then cmd_Modify.Enabled = True
If txtType <> "�R��" Then cmd_Delete.Enabled = True
cmd_AddNew.Enabled = True
cmd_Cancel.Enabled = False

'�q��Ӷ��\����w
fam_DetailData.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
fam_DetailData.Enabled = False
intsrcSKUNowRow = 0

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "-�q����@-�q��d��", Me.Caption, "cmd_OrdersQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub
Sub SendMail(strOrderkey As String)

'LTKK01�R��۰� Mail �q��
If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then
    
    Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String
    
    'Ū��ini�Ѽ�
    Dim objIni As New vbIniFile
    objIni.FileName = App.Path & "/" & App.title & ".ini"
    
    strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
    strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
    strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
    strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
    strSubject = "�q��R������"
    strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
    strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
    strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
    strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")
    
    Set objIni = Nothing
    
    Dim rsTmp As New ADODB.Recordset
    
    If Len(RTrim(strFrom)) > 0 Then '���H���
    
        str_SQL = "select �ܧO = 'BL01' " & _
                ",�f�D�N�X = rtrim(o.storerkey) " & _
                ",�q�渹�X��f�渹 = rtrim(od.externorderkey) + rtrim(od.externlineno) " & _
                ",�a�}�O = substring(o.consigneekey,5,20) " & _
                ",�Ƹ� = od.sku " & _
                ",�ܧO_�x��O = 'BL01_'+ od.lottable06 " & _
                ",�̤p���ƶq = isnull(od.originalqty,0) ,�q��� = convert(varchar,o.orderdate,111) " & _
                ",�w�p��f�� =  convert(varchar,o.deliverydate,111) " & _
                ",�R��� = convert(varchar,o.editdate,111) " & _
                ",�Ȥ�q�渹�X = rtrim(o.customerorderkey) " & _
                "From orders o join orderdetail od on o.orderkey = od.orderkey " & _
                "Where o.type = '�R��' and o.orderkey = '" & strOrderkey & "' "

        rsTmp.Open str_SQL, cn
        
        '�p�G�L��Ƥ]�nmail
        If Not rsTmp.EOF Or UCase(RTrim(strAlways)) = "YES" Then
            
            strAddAttachment = "C:\LTKK01\�q��R������\�q��R������_" & Format(Now, "yyyymmddhhMMss") & ".xls"
            
            Call Recordset2Excel("�q��R������", rsTmp)
            If Dir("C:\LTKK01\�q��R������", vbDirectory) = "" Then MkDirs "C:\LTKK01\�q��R������"
            MyXlsApp.ActiveWorkbook.SaveAs strAddAttachment
            MyXlsApp.Quit: Set MyXlsApp = Nothing
    
            '�ǰe�l��
            Dim objEmail As Object
            Set objEmail = CreateObject("CDO.Message")
        
            objEmail.From = strFrom
            objEmail.To = strTo
            objEmail.CC = strCC   ' �ƥ�
            objEmail.BCC = strBCC ' �K��ƥ�
            objEmail.Subject = strSubject
            objEmail.TextBody = strTextbody
            objEmail.AddAttachment strAddAttachment
        
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.bestlog.com.tw"
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            'SMTP ���A���ݭn���Ү�
            If Len(RTrim(strEmailID)) > 0 Then
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
            End If
            objEmail.Configuration.Fields.Update
            objEmail.Send
        
            MsgBox "LTKK01�R����Ӹ�ơA�t�Τw�oMail�q��!", , "�R����Ӹ��"
        
            Set objEmail = Nothing
        End If
    End If
End If

End Sub
Private Sub cmd_Save_Click()

'�M���S��r��
Call myFormExCharFilter(Me)

'�ˮָ�ƥ��T
If CheckOrdersData() = False Then Exit Sub

'�R���­q��
If txt_QueryExternOrderKey.Enabled = False And txtType <> "�R��" Then
     
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Open "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' ", cn
    If Not rsTmp.EOF Then MsgBox "���q��w�Ƹ��u�s�� " & rsTmp("route_no") & " �A�L�k�ק�!", vbOKOnly, "�s��": rsTmp.Close: Exit Sub
        
    Call DB_CheckConnectStatus
    
    Tran_Level = cn.BeginTrans
    
         cn.Execute "delete TRP02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP03T where receipt_no in (select receipt_no from trp02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='�R��' ,editdate = getdate() , editwho= '" & User_id & "' where orderkey='" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans: Tran_Level = 0
    txtType = "�R��"

If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then Call SendMail(txt_QueryExternOrderKey)

End If

'�q���Ʀs��
If SaveOrdersData() = False Then Exit Sub

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
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
cmbStorerkey.Enabled = True

End Sub

Private Sub cmdShipToList_Click()
'��ܫȤ�ݿ�M��
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmdShipToList.Name)
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
    cn.Execute str_SQL

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
Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arZip(1) As String
ReDim arZIPArea(1) As String
If Not tmp_rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_rs.EOF
      arZip(tmp_cnt) = tmp_rs.Fields("ZIP").Value
      arZIPArea(tmp_cnt) = tmp_rs.Fields("AreaCode").Value
      cmb_ZIP.AddItem tmp_rs.Fields("ZIP").Value & Space(5 - Len(Trim(tmp_rs.Fields("ZIP").Value))) & tmp_rs.Fields("Descr").Value
      tmp_rs.MoveNext
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
Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_rs.EOF
      arAreaCode(tmp_cnt) = tmp_rs.Fields("AreaCode").Value
      cmb_AreaCode.AddItem tmp_rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_rs.Fields("AreaCode").Value))) & tmp_rs.Fields("Descr").Value
      tmp_rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If
tmp_rs.Close
'���X�Ҧ��S��ݨD--TRP04M
cmb_ExtraDemand1.Clear: cmb_ExtraDemand2.Clear
str_SQL = "Select Rtrim(Extra_Demand_Code) as 'ECode',Isnull(Rtrim(Description),'') as 'ECodeDescr' From TRP04M Order by Extra_Demand_Code"
Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arExtraDemand(1) As String
If Not tmp_rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_rs.EOF
      arExtraDemand(tmp_cnt) = tmp_rs.Fields("ECode").Value
      cmb_ExtraDemand1.AddItem tmp_rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_rs.Fields("ECode").Value))) & tmp_rs.Fields("ECodeDescr").Value
      cmb_ExtraDemand2.AddItem tmp_rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_rs.Fields("ECode").Value))) & tmp_rs.Fields("ECodeDescr").Value
      tmp_rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arExtraDemand) Then
         ReDim Preserve arExtraDemand(UBound(arExtraDemand) + 10) As String
      End If
   Loop
End If
tmp_rs.Close

'���X�Ҧ��f�D
str_SQL = "Select Rtrim(Storerkey) + ' ' + rtrim(short_name) as 'Storer' From TRP16M Order by Storerkey "
Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn
If Not tmp_rs.EOF Then

   Do While Not tmp_rs.EOF
      cmbStorerkey.AddItem tmp_rs("Storer")
      tmp_rs.MoveNext
   Loop
End If
cmbStorerkey.ListIndex = -1
tmp_rs.Close

cbo_Priority.AddItem "I �X�f"
cbo_Priority.AddItem "R �h�f"
cbo_Priority.AddItem "A �Q�����"
cbo_Priority.AddItem "A2B ���f�t�e"
cbo_Priority.AddItem "RC ���f�J�w"
cbo_Priority.AddItem "RS �h�f�~�X�w"
cbo_Priority = ""

cboKey.AddItem "TMS�渹"
'cboKey.AddItem "�q�渹�X"
'cboKey = ""

End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub

fam_OrderDetail.Width = Me.ScaleWidth - 60
dg_OrderDetail.Width = fam_OrderDetail.Width - 180

fam_OrderDetail.Height = Me.ScaleHeight - fam_Orders.Height - fam_Header.Height ' - 360
dg_OrderDetail.Height = fam_OrderDetail.Height - fam_DetailData.Height - 240

End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_OP_ManualOrders = Nothing
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

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
If Trim(txt_ConsigneeKey.Text) = "" Then Exit Sub
txt_ConsigneeKey = myExCharFilter(txt_ConsigneeKey.Text)

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
str_SQL = "Select * From TRP01M Where storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' and ConsigneeKey = '" & txt_ConsigneeKey.Text & "'"
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "��ƿ��~�G�Ȥ�s�� [" & txt_ConsigneeKey.Text & "] ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If
txt_FullName.Text = IIf(IsNull(tmp_rs.Fields("Full_Name").Value), "", Trim(tmp_rs.Fields("Full_Name").Value))
txt_ShortName.Text = IIf(IsNull(tmp_rs.Fields("Short_Name").Value), "", Trim(tmp_rs.Fields("Short_Name").Value))
txt_Contact.Text = IIf(IsNull(tmp_rs.Fields("Contact").Value), "", Trim(tmp_rs.Fields("Contact").Value))
txt_Phone.Text = IIf(IsNull(tmp_rs.Fields("Phone").Value), "", Trim(tmp_rs.Fields("Phone").Value))
txt_Address.Text = IIf(IsNull(tmp_rs.Fields("Address").Value), "", Trim(tmp_rs.Fields("Address").Value))
cmb_ZIP.ListIndex = -1
For iLoop = 0 To cmb_ZIP.ListCount - 1
   If arZip(iLoop) = tmp_rs.Fields("ZIP").Value Then
      cmb_ZIP.ListIndex = iLoop
      Exit For
   End If
Next iLoop
cmb_AreaCode.ListIndex = -1
For iLoop = 0 To cmb_AreaCode.ListCount - 1
    If arAreaCode(iLoop) = tmp_rs.Fields("Area_Code").Value Then
       cmb_AreaCode.ListIndex = iLoop
       Exit For
    End If
Next iLoop
cmb_ExtraDemand1.ListIndex = -1
For iLoop = 0 To cmb_ExtraDemand1.ListCount - 1
    If arExtraDemand(iLoop) = tmp_rs.Fields("Extra_Demand_Code").Value Then
       cmb_ExtraDemand1.ListIndex = iLoop
       Exit For
    End If
Next iLoop
cmb_ExtraDemand2.ListIndex = -1
For iLoop = 0 To cmb_ExtraDemand2.ListCount - 1
    If arExtraDemand(iLoop) = tmp_rs.Fields("Extra_Demand_Code2").Value Then
       cmb_ExtraDemand2.ListIndex = iLoop
       Exit For
    End If
Next iLoop
tmp_rs.Close
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
'�d�߱���G�f�D�渹
If KeyAscii = vbKeyReturn Then
   Call cmd_OrdersQuery_Click
End If
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

Private Function CheckSKU(ByVal strSKU As String) As Integer
'�ˮֳf���O�_���T
CheckSKU = 1
If cmbStorerkey.Text = "" Then
   msg_text = "�f���ˮֲ��`�G�|����J [�f�D]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmbStorerkey.SetFocus
   Exit Function
End If

'���ҳf���O�_���T
str_SQL = "Select isnull(Rtrim(Descr),'') as 'Descr' From SKU Where StorerKey = '" & mySplit(cmbStorerkey.Text, " ", 0) & "' and SKU = '" & strSKU & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�f�����ҿ��~�GStorer = [" & cmbStorerkey.Text & "] �L���f��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cboSku.SetFocus
   Exit Function
End If
txt_SkuDescr.Text = tmp_rs.Fields("Descr").Value
CheckSKU = 0
tmp_rs.Close

End Function

Private Sub txt_StorerKey_KeyPress(KeyAscii As Integer)
'�f�D
If KeyAscii >= 97 And KeyAscii <= 122 Then '�p�g�r���אּ�j�g�r��
   KeyAscii = KeyAscii - 32
End If
End Sub

Private Function CheckOrdersData() As Boolean

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

If dg_OrderDetail.Rows <= 2 Then
   msg_text = "��ƿ��~�G�����s�W�q�� [���Ӹ��] ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)

'1.�ˮ֫Ȥ�s��
str_SQL = "Select Count(*) AS RecCount From TRP01M Where storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' and ConsigneeKey = '" & txt_ConsigneeKey.Text & "'"
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_rs.Fields("RecCount").Value = 0 Then
   tmp_rs.Close
   msg_text = "��ƿ��~�G�Ȥ�s�� [" & txt_ConsigneeKey.Text & "] ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
tmp_rs.Close

'�O�_�w�Ƹ��s
tmp_rs.Open "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' ", cn
If Not tmp_rs.EOF Then MsgBox "���q��w�Ƹ��u�s�� " & tmp_rs("route_no") & " �A�L�k�ק�!", vbOKOnly, "�s��": tmp_rs.Close: Exit Function
tmp_rs.Close

If mySplit(Trim(cmbStorerkey), " ", 0) = "LTKK01" Then

    '�x�W�Q��q�歫���ˬd
    If txt_QueryExternOrderKey.Enabled = True Then
       '�s�W�q���
       tmp_rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and consigneekey = '" & txt_ConsigneeKey.Text & "' and convert(varchar(8),deliverydate,112) = '" & txt_DeliveryDate.Text & "' and rtrim(isnull(type,'')) <> '�R��' and priority ='" & mySplit(Trim(cbo_Priority.Text), " ", 0) & "' and cast(notes as varchar(300)) = '" & Trim(txt_Description) & "' ", cn
       If tmp_rs.Fields("RecCount").Value <> 0 Then MsgBox "�x�W�Q��q�歫��!(�ۦP�Ȥ�s���B��f��B�q�����O�P�q��Ƶ�)", 64, "��ƿ��~": tmp_rs.Close: Exit Function
    Else
        '�ק�q���
       tmp_rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and consigneekey = '" & txt_ConsigneeKey.Text & "' and convert(varchar(8),deliverydate,112) = '" & txt_DeliveryDate.Text & "' and rtrim(isnull(type,'')) <> '�R��' and priority ='" & mySplit(Trim(cbo_Priority.Text), " ", 0) & "' and cast(notes as varchar(300)) = '" & txt_Description & "' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_rs.Fields("RecCount").Value <> 0 Then MsgBox "�x�W�Q��q�歫��!(�ۦP�Ȥ�s���B��f��B�q�����O�P�q��Ƶ�)", 64, "��ƿ��~": tmp_rs.Close: Exit Function
    End If
Else
     '��L�f�D�q�歫���ˬd
    If txt_QueryExternOrderKey.Enabled = True Then
       '�s�W�q���
       tmp_rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' ", cn
       If tmp_rs.Fields("RecCount").Value <> 0 Then MsgBox "�f�D�q�渹�X����!", 64, "��ƿ��~": tmp_rs.Close: Exit Function
    Else
        '�ק�q���
       tmp_rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '�R��' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_rs.Fields("RecCount").Value <> 0 Then MsgBox "�f�D�q�渹�X����!", 64, "��ƿ��~": tmp_rs.Close: Exit Function
    End If

End If
tmp_rs.Close

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
CheckOrdersData = True
Exit Function

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "-�q����@-�s��-����ˮ�", Me.Caption, "Form SubProgram CheckOrdersData", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Function

Private Function SaveOrdersData() As Boolean
'�q���Ʀs��
On Error GoTo err_Handle
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
SaveOrdersData = False

'��Ʈw���ʥ��--�_�I
Tran_Level = cn.BeginTrans

Dim strOrderkey As String

'1.���s���q��s��

    str_SQL = "select isnull(max(orderkey),0) from orders"
    Call Confirm_Recordset_Closed(tmp_rs)
    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strOrderkey = StrPadLeft(Val(Trim(tmp_rs.Fields(0))) + 1, 10, 0)
    tmp_rs.Close
    
'2.�q����Y��Ʀs��
   str_SQL = "Insert into Orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_Contact1,C_Company,C_Address1," & _
             " C_ZIP,C_Phone1,AddWho,EditWho,DoRoute,Notes,customerorderkey,type,b_company,facility) Values ('" & _
             strOrderkey & "','" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & txt_Extern & "','" & _
             Left(txt_OrderDate.Text, 4) & "/" & Mid(txt_OrderDate.Text, 5, 2) & "/" & Right(txt_OrderDate.Text, 2) & "','" & _
             Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "'," & _
             "'" & mySplit(Trim(cbo_Priority), " ", 0) & "','" & txt_ConsigneeKey.Text & "','" & txt_Contact.Text & "','" & txt_ShortName.Text & "','" & txt_Address.Text & "','" & arZip(cmb_ZIP.ListIndex) & "','" & _
             txt_Phone.Text & "','" & User_id & "','" & User_id & "','Y','" & txt_Description.Text & "','" & UCase(txt_OrderKey.Text) & "','','" & txtShipToKey.Text & "','" & txtFacility & "' )"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�O�_�O�ק�
If txt_QueryExternOrderKey.Enabled = False And txt_QueryExternOrderKey <> strOrderkey Then

'����l�q����
    tmp_rs.Open "select amount = isnull(amount,0) , billtokey=isnull(billtokey,'') , b_contact1 = isnull(b_contact1,'') ,  door = isnull(door,'') , stop = isnull(stop,'') , facility = isnull(facility,'') from orders where orderkey = '" & txt_QueryExternOrderKey & "'", cn, adOpenForwardOnly, adLockReadOnly

'��s��L���
    str_SQL = "update orders set amount = " & tmp_rs("amount") & _
              ", billtokey = '" & tmp_rs("billtokey") & "' " & _
              ", b_contact1 = '" & tmp_rs("b_contact1") & "' " & _
              ", door = '" & tmp_rs("door") & "' " & _
              ", stop = '" & tmp_rs("stop") & "' " & _
              ", facility = '" & tmp_rs("facility") & "' " & _
              ", Updatesource = '" & txt_QueryExternOrderKey & "' " & _
              "where orderkey = '" & strOrderkey & "' "
    tmp_rs.Close
    
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
End If

'4.�B�z�q��Ӷ��s��
Dim strLineNo As String, strSKU As String, dbOrderCS As Double, dbOrderEA As Double, dbCasecnt As Double, strTKLOC As String, strNotes As String, strLot3 As String, strLot4 As String, strLot5 As String
With dg_OrderDetail
     For iLoop = 1 To .Rows - 2
         .Row = iLoop
         .Col = 0: strLineNo = .Text
         .Col = 1: strSKU = Trim(.Text)        '�f��
         .Col = 3: strTKLOC = myExCharFilter(.Text)    'TKLOC
         .Col = 4: dbOrderCS = myExCharFilter(.Text)     '�c��
         .Col = 5: dbOrderEA = myExCharFilter(.Text)     '�Ӽ�
         .Col = 6: dbCasecnt = myExCharFilter(.Text)     '�C�c
         .Col = 7: strNotes = myExCharFilter(.Text)     '�Ƶ�
         .Col = 8: strLot3 = myExCharFilter(.Text)     '�Ͳ��帹
         .Col = 9: strLot4 = myExCharFilter(.Text)     '�s�y��
         .Col = 10: strLot5 = myExCharFilter(.Text)     '�����
            '�s�W�Ginsert into OrderDetail
            str_SQL = "insert into OrderDetail (StorerKey,OrderKey,OrderLineNumber,ExternOrderKey,SKU,OriginalQty,OpenQty,ShippedQty,UOM,AddWho,EditWho,lottable06,notes,lottable03,lottable04,lottable05) Values ('" & _
                      mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & strOrderkey & "','" & strLineNo & "','" & txt_Extern & "','" & strSKU & "'," & dbOrderCS * dbCasecnt + dbOrderEA & "," & dbOrderCS * dbCasecnt + dbOrderEA & ",0" & _
                      ",'EA','" & User_id & "','" & User_id & "','" & UCase(strTKLOC) & "','" & strNotes & "','" & strLot3 & "','" & strLot4 & "','" & strLot5 & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�L�����h��s��null
            If Len(Trim(strLot4)) = 0 Then cn.Execute "update orderdetail set lottable04 = null where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' "
            '�L�����h��s��null
            If Len(Trim(strLot5)) = 0 Then cn.Execute "update orderdetail set lottable05 = null where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' "
            
     Next iLoop
End With

'��spackkey
str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & strOrderkey & "' "

cn.Execute str_SQL

'�ɫȤ���
cn.Execute "exec gs_OrdersUpdate '" & mySplit(cmbStorerkey.Text, " ", 0) & "' "

'7. DB Transaction Commit
cn.CommitTrans: Tran_Level = 0

MsgBox "�q��s�W����(" & strOrderkey & ")�C", vbOKOnly, Me.Caption

'8.�M���ù�
Call Clear_AllField
txt_QueryExternOrderKey.Text = strOrderkey
SaveOrdersData = True
txt_QueryExternOrderKey.Enabled = True
Exit Function

err_Handle:
   If Tran_Level <> 0 Then cn.RollbackTrans: Tran_Level = 0

   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "-�q����@-�s��-��Ʀs��", Me.Caption, "Form SubProgram SaveOrdersData", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Function

Private Sub txtShipToKey_GotFocus()
    frm_OP_ManualShipToOrders.ZOrder 0
    fam_ConsigneeData.Enabled = False
End Sub

Public Sub txtShipToKey_LostFocus()

'��B��f�Ȥ�s��
If Trim(txtShipToKey.Text) = "" Then Exit Sub

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
str_SQL = "Select * From TRP01M Where storerkey = '" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "' and ConsigneeKey = '" & txtShipToKey.Text & "'"
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "��ƿ��~�G�Ȥ�s�� [" & txtShipToKey.Text & "] ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If

txt_ShipToFullName.Text = IIf(IsNull(tmp_rs.Fields("Full_Name").Value), "", Trim(tmp_rs.Fields("Full_Name").Value))
txt_ShipToShortName.Text = IIf(IsNull(tmp_rs.Fields("Short_Name").Value), "", Trim(tmp_rs.Fields("Short_Name").Value))
txt_ShipToContact.Text = IIf(IsNull(tmp_rs.Fields("Contact").Value), "", Trim(tmp_rs.Fields("Contact").Value))
txt_ShipToPhone.Text = IIf(IsNull(tmp_rs.Fields("Phone").Value), "", Trim(tmp_rs.Fields("Phone").Value))
txt_ShipToAddress.Text = IIf(IsNull(tmp_rs.Fields("Address").Value), "", Trim(tmp_rs.Fields("Address").Value))
txt_ShipToZIP = tmp_rs("zip") & ""
txt_ShipToAreaCode = tmp_rs("Area_Code") & ""
txt_ShipToExtraDemand1 = tmp_rs("Extra_Demand_code2") & ""
txt_ShipToExtraDemand2 = tmp_rs("Extra_Demand_code2") & ""

tmp_rs.Close

End Sub

Private Sub cboLot6_KeyPress(KeyAscii As Integer)
'�p�g�r���אּ�j�g�r��
If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

