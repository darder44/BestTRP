VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_OP_CutOrders 
   Caption         =   "����h���q�����"
   ClientHeight    =   7140
   ClientLeft      =   210
   ClientTop       =   855
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11475
   Begin TabDlg.SSTab SSTab1 
      Height          =   7080
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12488
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "�q��C��"
      TabPicture(0)   =   "frm_OP_CutOrders.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape2(0)"
      Tab(0).Control(1)=   "Shape1"
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(6)=   "Label1(19)"
      Tab(0).Control(7)=   "Shape2(1)"
      Tab(0).Control(8)=   "dg_TRP02W"
      Tab(0).Control(9)=   "cmd_Tab1_ResetRS"
      Tab(0).Control(10)=   "cmd_FilterAndSort"
      Tab(0).Control(11)=   "txt_Tab0_TotalCase"
      Tab(0).Control(12)=   "txt_Tab0_TotalPallet"
      Tab(0).Control(13)=   "txt_Tab0_TotalVolumn"
      Tab(0).Control(14)=   "txt_Tab0_TotalWeight"
      Tab(0).Control(15)=   "txt_Tab0_OrderCount"
      Tab(0).Control(16)=   "cmd_Tab0_DisplaySelectedOrder"
      Tab(0).Control(17)=   "cmd_Tab0_DisplayOrders"
      Tab(0).Control(18)=   "cmd_Exit(0)"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "�ݤ��έq��"
      TabPicture(1)   =   "frm_OP_CutOrders.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fam_Tab1_Orders"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fam_Tab1_OrderDetail"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "�q����Ω��� + �d��"
      TabPicture(2)   =   "frm_OP_CutOrders.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dg_CutOrderDetail"
      Tab(2).Control(1)=   "dg_CutOrders"
      Tab(2).Control(2)=   "fam_Tab2_Delete"
      Tab(2).Control(3)=   "fam_Tab2_Qoery"
      Tab(2).ControlCount=   4
      Begin VB.Frame fam_Tab1_OrderDetail 
         BackColor       =   &H00808000&
         Caption         =   "�q�����"
         ForeColor       =   &H00400040&
         Height          =   4470
         Left            =   270
         TabIndex        =   54
         Top             =   2400
         Width           =   10875
         Begin VB.TextBox txt_Tab1_SelectedPalletQty 
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
            Left            =   4365
            TabIndex        =   63
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_SelectedVolumn 
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
            Left            =   3420
            TabIndex        =   62
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_SelectedWeight 
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
            Left            =   2475
            TabIndex        =   61
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_SelectedCaseQty 
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
            Left            =   1515
            TabIndex        =   60
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_CutCaseQty 
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
            Left            =   6480
            TabIndex        =   59
            Top             =   450
            Width           =   700
         End
         Begin VB.CommandButton cmd_Tab1_CutQty 
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
            Height          =   560
            Left            =   7200
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   58
            Top             =   180
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab1_CutPalletQty 
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
            Left            =   6480
            TabIndex        =   57
            Top             =   150
            Width           =   700
         End
         Begin VB.CommandButton cmd_Tab1_CutOrders 
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
            Height          =   560
            Left            =   9600
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   56
            Top             =   180
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_ClearQty 
            BackColor       =   &H00FF80FF&
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
            Height          =   560
            Left            =   8385
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   55
            Top             =   180
            Width           =   1200
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_SelectedOrderDetail 
            Height          =   3645
            Left            =   45
            TabIndex        =   64
            Top             =   765
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   6429
            _Version        =   393216
            Cols            =   9
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
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
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   18
            Left            =   1860
            TabIndex        =   71
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�O��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   20
            Left            =   4650
            TabIndex        =   70
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
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   21
            Left            =   3735
            TabIndex        =   69
            Top             =   240
            Width           =   420
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
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   22
            Left            =   2805
            TabIndex        =   68
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "2.�c�Ƥ���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   195
            Index           =   23
            Left            =   5475
            TabIndex        =   67
            Top             =   510
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "1.�O�Ƥ���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   195
            Index           =   24
            Left            =   5475
            TabIndex        =   66
            Top             =   225
            Width           =   1005
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
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   25
            Left            =   195
            TabIndex        =   65
            Top             =   510
            Width           =   1260
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
         Height          =   630
         Index           =   0
         Left            =   -64890
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   53
         Top             =   495
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab0_DisplayOrders 
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
         Height          =   600
         Left            =   -74670
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   52
         Top             =   495
         Width           =   2250
      End
      Begin VB.CommandButton cmd_Tab0_DisplaySelectedOrder 
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
         Height          =   600
         Left            =   -72135
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   50
         Top             =   495
         Width           =   2250
      End
      Begin VB.TextBox txt_Tab0_OrderCount 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -73845
         TabIndex        =   49
         Top             =   6615
         Width           =   915
      End
      Begin VB.TextBox txt_Tab0_TotalWeight 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -69630
         TabIndex        =   48
         Top             =   6615
         Width           =   1290
      End
      Begin VB.TextBox txt_Tab0_TotalVolumn 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -67365
         TabIndex        =   47
         Top             =   6615
         Width           =   1290
      End
      Begin VB.TextBox txt_Tab0_TotalPallet 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -65130
         TabIndex        =   46
         Top             =   6615
         Width           =   1290
      End
      Begin VB.TextBox txt_Tab0_TotalCase 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -71955
         TabIndex        =   45
         Top             =   6615
         Width           =   1290
      End
      Begin VB.Frame fam_Tab1_Orders 
         BackColor       =   &H8000000C&
         Caption         =   "�q����"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   1875
         Left            =   255
         TabIndex        =   11
         Top             =   480
         Width           =   10905
         Begin VB.TextBox txt_Tab1_Storer 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   29
            Top             =   285
            Width           =   825
         End
         Begin VB.TextBox txt_Tab1_OrderKey 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   2745
            TabIndex        =   28
            Top             =   285
            Width           =   1050
         End
         Begin VB.TextBox txt_Tab1_Extern 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   4710
            TabIndex        =   27
            Top             =   285
            Width           =   1410
         End
         Begin VB.TextBox txt_Tab1_FullName 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   26
            Top             =   585
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_Address 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   25
            Top             =   870
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_ExtraDemand1 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   24
            Top             =   1155
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_ExtraDemand2 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   23
            Top             =   1440
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_ZIP 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   7170
            TabIndex        =   22
            Top             =   585
            Width           =   1680
         End
         Begin VB.TextBox txt_Tab1_AreaCode 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   9795
            TabIndex        =   21
            Top             =   585
            Width           =   825
         End
         Begin VB.TextBox txt_Tab1_VehicleType 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   7170
            TabIndex        =   20
            Top             =   870
            Width           =   1680
         End
         Begin VB.CheckBox chk_Tab1_MultiCustomer 
            BackColor       =   &H8000000C&
            Caption         =   "���e�Ȥ�"
            Height          =   180
            Left            =   8910
            TabIndex        =   19
            Top             =   1200
            Width           =   1260
         End
         Begin VB.TextBox txt_Tab1_ChannelType 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   7170
            TabIndex        =   18
            Top             =   1155
            Width           =   1680
         End
         Begin VB.TextBox txt_Tab1_OrderDate 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   7185
            TabIndex        =   17
            Top             =   285
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   9405
            TabIndex        =   16
            Top             =   270
            Width           =   1230
         End
         Begin VB.TextBox txt_Tab1_Weight 
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
            Left            =   7650
            TabIndex        =   15
            Top             =   1440
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_Volumn 
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
            Left            =   8595
            TabIndex        =   14
            Top             =   1440
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_PalletQty 
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
            Left            =   9540
            TabIndex        =   13
            Top             =   1440
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_EXEConfirm 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   9795
            TabIndex        =   12
            Top             =   885
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f        �D"
            Height          =   180
            Index           =   4
            Left            =   225
            TabIndex        =   44
            Top             =   345
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��s��"
            Height          =   180
            Index           =   5
            Left            =   1965
            TabIndex        =   43
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D�渹"
            Height          =   180
            Index           =   6
            Left            =   3930
            TabIndex        =   42
            Top             =   345
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�W��"
            Height          =   180
            Index           =   7
            Left            =   225
            TabIndex        =   41
            Top             =   645
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�e�f�a�}"
            Height          =   180
            Index           =   8
            Left            =   225
            TabIndex        =   40
            Top             =   915
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�S��ݨD 1"
            Height          =   180
            Index           =   9
            Left            =   90
            TabIndex        =   39
            Top             =   1215
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�S��ݨD 2"
            Height          =   180
            Index           =   10
            Left            =   90
            TabIndex        =   38
            Top             =   1485
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�l���ϸ�"
            Height          =   180
            Index           =   11
            Left            =   6390
            TabIndex        =   37
            Top             =   645
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϽX"
            Height          =   180
            Index           =   12
            Left            =   9030
            TabIndex        =   36
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���إN�X"
            Height          =   180
            Index           =   13
            Left            =   6390
            TabIndex        =   35
            Top             =   915
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q�����A"
            Height          =   180
            Index           =   14
            Left            =   6390
            TabIndex        =   34
            Top             =   1215
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q����"
            Height          =   180
            Index           =   15
            Left            =   6390
            TabIndex        =   33
            Top             =   345
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�f���"
            Height          =   180
            Index           =   16
            Left            =   8625
            TabIndex        =   32
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���q/���n/�O��"
            Height          =   180
            Index           =   17
            Left            =   6405
            TabIndex        =   31
            Top             =   1515
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "EXE�^��"
            Height          =   180
            Index           =   26
            Left            =   9030
            TabIndex        =   30
            Top             =   960
            Width           =   690
         End
      End
      Begin VB.CommandButton cmd_FilterAndSort 
         BackColor       =   &H00FF80FF&
         Caption         =   "�z �� �� ��"
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
         Left            =   -69615
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
         Top             =   510
         Width           =   2160
      End
      Begin VB.CommandButton cmd_Tab1_ResetRS 
         Appearance      =   0  '����
         BackColor       =   &H00C0C0FF&
         Caption         =   "�����z��Ƨ�"
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
         Left            =   -67410
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   8
         Top             =   510
         Width           =   2160
      End
      Begin VB.Frame fam_Tab2_Qoery 
         BackColor       =   &H00404000&
         Height          =   2160
         Left            =   -65790
         TabIndex        =   4
         Top             =   3270
         Width           =   1995
         Begin VB.CommandButton cmd_Tab2_ExternQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�f�D�渹�d��"
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
            Picture         =   "frm_OP_CutOrders.frx":0054
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   6
            Top             =   1215
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab2_Extern 
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
            TabIndex        =   5
            Top             =   630
            Width           =   1605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D�渹"
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
            TabIndex        =   7
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.Frame fam_Tab2_Delete 
         Appearance      =   0  '����
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -65805
         TabIndex        =   2
         Top             =   5520
         Visible         =   0   'False
         Width           =   1995
         Begin VB.CommandButton cmd_Tab2_CutOrderDelete 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�R�����έq��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   90
            Picture         =   "frm_OP_CutOrders.frx":035E
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   3
            ToolTipText     =   "�R��"
            Top             =   180
            Width           =   1800
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_CutOrders 
         Height          =   2775
         Left            =   -74790
         TabIndex        =   1
         Top             =   450
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   9
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
      Begin MSDataGridLib.DataGrid dg_CutOrderDetail 
         Height          =   3600
         Left            =   -74790
         TabIndex        =   10
         Top             =   3285
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6350
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
      Begin MSDataGridLib.DataGrid dg_TRP02W 
         Height          =   5205
         Left            =   -74745
         TabIndex        =   51
         Top             =   1305
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   9181
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
      Begin VB.Shape Shape2 
         BackColor       =   &H00004000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H008080FF&
         BorderWidth     =   2
         Height          =   720
         Index           =   1
         Left            =   -74730
         Top             =   435
         Width           =   2385
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�q�浧��"
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
         Index           =   19
         Left            =   -74745
         TabIndex        =   76
         Top             =   6690
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�`���q"
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
         Index           =   0
         Left            =   -70320
         TabIndex        =   75
         Top             =   6675
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�`���n"
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
         Index           =   1
         Left            =   -68055
         TabIndex        =   74
         Top             =   6675
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�`�O��"
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
         Index           =   2
         Left            =   -65820
         TabIndex        =   73
         Top             =   6690
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�`�c��"
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
         Index           =   3
         Left            =   -72645
         TabIndex        =   72
         Top             =   6675
         Width           =   630
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Height          =   735
         Left            =   -69690
         Top             =   450
         Width           =   4530
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   720
         Index           =   0
         Left            =   -72210
         Top             =   435
         Width           =   2385
      End
   End
End
Attribute VB_Name = "frm_OP_CutOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private blTRP02WEventEnable As Boolean
Private rs_TRP02W As ADODB.Recordset
Private rs_CutOrderDetail As ADODB.Recordset   '�w�����q����Τ��q�����

Private dbCut_TotalCaseQty As Double
Private dbCut_TotalWeight As Double
Private dbCut_TotalVolumn As Double
Private dbCut_TotalPalletQty As Double

Private Sub cmd_FilterAndSort_Click()
'�q��C�� >> �z��Ƨ�
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_TRP02W"

If ShowForm_RS_FilterAndSort(rs_TRP02W, "�ݱƨ��q��", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = 2

End Sub

Private Sub cmd_Tab0_DisplayOrders_Click()
'�q��C�� >> ��ܫݱƨ��q��
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_TRP02W.DataSource = Nothing
Set rs_TRP02W = Nothing

str_SQL = "Select �q��s��,�e�f��,�Ȥ�s��,�f�D�渹,�c��,���q,���n,�O��,ZIP,�ϽX,�Ȥ�W��,�q���,�f�D,�ѧO,�q��Ƶ�,EXE�^�� " & _
          "From CutOrders_SourceOrder Order by �O�� DESC"
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
Call Replication_Recordset(tmp_Rs, rs_TRP02W)
tmp_Rs.Close

With dg_TRP02W
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_TRP02W.MoveFirst
blTRP02WEventEnable = False
Set dg_TRP02W.DataSource = rs_TRP02W
With dg_TRP02W
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1100       '�q��s��
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900        '�e�f��
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1200       '�Ȥ�s��
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 900        '�f�D�渹
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 800        '�c��
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 800        '���q
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800        '���n
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 800        '�O��
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 500        'ZIP
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 500       '�ϽX
    .Columns(10).Alignment = dbgCenter
    .Columns(11).Width = 3500      '�Ȥ�W��
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1000      '�q���
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 700       '�f�D
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1100      '�ѧO
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1500      '�q��Ƶ�
    .Columns(15).Alignment = dbgLeft
    .Columns(14).Width = 900       'EXE�^��
    .Columns(14).Alignment = dbgLeft
End With
blTRP02WEventEnable = True
'�����q��Ҧ��Ӷ��`�p��ƭ�
str_SQL = "Select count(�q��s��) as �q�浧��,sum(�c��) as �`�c��,sum(���q) as �`���q,sum(���n) as �`���n,sum(�O��) as �`�O�� " & _
          "From CutOrders_SourceOrder  "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   txt_Tab0_OrderCount.Text = tmp_Rs.Fields("�q�浧��").Value
   txt_Tab0_TotalCase.Text = tmp_Rs.Fields("�`�c��").Value
   txt_Tab0_TotalWeight.Text = tmp_Rs.Fields("�`���q").Value
   txt_Tab0_TotalVolumn.Text = tmp_Rs.Fields("�`���n").Value
   txt_Tab0_TotalPallet.Text = tmp_Rs.Fields("�`�O��").Value
End If
tmp_Rs.Close

'�M����
Call Clear_SelectedOrderData
'�]�w�����έq�椧�q��W��
Call SetGrid_Format_SelectedOrderDetail
'�]�w���έq��C��
Call SetGrid_Format_CutOrderList
'�]�w�w�������έq����Ӫ�
Call CreateRS_CutOrderDetail

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

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_Tab0_DisplaySelectedOrder_Click()
'�q��C�� >> ��ܭq�����
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub
If dg_TRP02W.SelBookmarks.Count = 0 Then
   msg_text = "�L���w������q��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Screen.MousePointer = vbHourglass
'�]�w���έq��C��
Call SetGrid_Format_CutOrderList
'�]�w�w�������έq����Ӫ�
Call CreateRS_CutOrderDetail

'�M����
Call Clear_SelectedOrderData
SSTab1.Tab = 1
DoEvents: DoEvents

Dim strOrderkey As String
strOrderkey = rs_TRP02W.Fields("�q��s��").Value

str_SQL = "Select �f�D,�q��s��,�f�D�渹,�q���,�X�f��,�ϽX,�Ȥ�W��,�e�f�a�},�S��ݨD1,�S��ݨD2,�l���ϸ�,���إN�X,�q��,���e,���q,���n,�O��,EXE�^�� " & _
          "From CutOrders_SelectedOrders Where �q��s�� = '" & strOrderkey & "'"
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
txt_Tab1_Storer.Text = tmp_Rs.Fields("�f�D").Value
txt_Tab1_OrderKey.Text = tmp_Rs.Fields("�q��s��").Value
txt_Tab1_Extern.Text = tmp_Rs.Fields("�f�D�渹").Value
txt_Tab1_OrderDate.Text = tmp_Rs.Fields("�q���").Value
txt_Tab1_DeliveryDate.Text = tmp_Rs.Fields("�X�f��").Value
txt_Tab1_FullName.Text = tmp_Rs.Fields("�Ȥ�W��").Value
txt_Tab1_Address.Text = tmp_Rs.Fields("�e�f�a�}").Value
txt_Tab1_ExtraDemand1.Text = tmp_Rs.Fields("�S��ݨD1").Value
txt_Tab1_ExtraDemand2.Text = tmp_Rs.Fields("�S��ݨD2").Value
txt_Tab1_ZIP.Text = tmp_Rs.Fields("�l���ϸ�").Value & ""
txt_Tab1_AreaCode.Text = tmp_Rs.Fields("�ϽX").Value
txt_Tab1_VehicleType.Text = tmp_Rs.Fields("���إN�X").Value
txt_Tab1_ChannelType.Text = tmp_Rs.Fields("�q��").Value
If tmp_Rs.Fields("���e").Value = "Y" Then
   chk_Tab1_MultiCustomer.Value = vbChecked
Else
   chk_Tab1_MultiCustomer.Value = vbUnchecked
End If
txt_Tab1_Weight.Text = tmp_Rs.Fields("���q").Value
txt_Tab1_Volumn.Text = tmp_Rs.Fields("���n").Value
txt_Tab1_PalletQty.Text = tmp_Rs.Fields("�O��").Value
txt_Tab1_EXEConfirm.Text = tmp_Rs.Fields("EXE�^��").Value
tmp_Rs.Close

'�]�w�����έq�椧�q��W��
Call SetGrid_Format_SelectedOrderDetail
str_SQL = "Select ����,�f��,�~�W,�q��q,�c��,���q,���n,�O��,�C�O�c��,�C�c�Ӽ� " & _
          "From CutOrders_SelectedOrderDetail Where �q��s�� = '" & strOrderkey & "' order by ����"
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
   With dg_SelectedOrderDetail
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
        .Col = 7    '�O��
        .Text = tmp_Rs.Fields("�O��").Value
        .Col = 13   '�C�O�c��
        .Text = tmp_Rs.Fields("�C�O�c��").Value
        '�C�c�Ӽ�
        .Col = 14: .Text = tmp_Rs("�C�c�Ӽ�")
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
Set tmp_Rs = Nothing
Screen.MousePointer = vbDefault
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

Private Sub cmd_Tab1_ClearQty_Click()
'�ݤ��έq�� >> �M�� [�O�Ƥ���][�c�Ƥ���] ����
txt_Tab1_CutPalletQty.Text = ""
txt_Tab1_CutCaseQty.Text = ""
'�M�����έӼ�
dg_SelectedOrderDetail.Col = 12: dg_SelectedOrderDetail.Text = ""
txt_Tab1_CutCaseQty.SetFocus
'RUN Button [�ƶq����] Click
Call cmd_Tab1_CutQty_Click
End Sub

Private Sub cmd_Tab1_CutOrders_Click()
'�ݤ��έq�� >> ���έq��
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub

Dim intTRP02WBookMark As String     '���b�i�� [�q����Χ@�~] ���q���ƦC
Dim strCutOrder_SrcKey As String    '���b�i�� [�q����Χ@�~] ���q��s��
Dim dbMaxKey As Double              '�s�q��s���G���X key
Dim strCutOrder_NewKey As String    '�s���ΥX�Ӥ��q��� [�q��s��]
Dim i As Double
Dim int_CS As Integer, int_CutCS As Integer

On Error GoTo err_Handle
If Len(Trim(txt_Tab1_OrderKey.Text)) = 0 Then Exit Sub

If dg_TRP02W.SelBookmarks.Count = 0 Then
   msg_text = "�{�ǿ��~�G������q��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'�ˬd�O���I������Τ��q��Ӷ�
dg_SelectedOrderDetail.Visible = False
Dim dbCount As Double
dbCount = 0
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         .Row = i: .Col = 1
         If Len(Trim(.Text)) <> 0 Then
            dbCount = dbCount + 1
         End If
     Next i
End With
dg_SelectedOrderDetail.Visible = True
If dbCount = 0 Then
   msg_text = "��ƿ��~�G����������Τ��q���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If

'�ˬd�`�ƶq�O�_������μƶq�A���������渹���� edit by eric
For i = 1 To dg_SelectedOrderDetail.Rows - 1
    dg_SelectedOrderDetail.Row = i:
    dg_SelectedOrderDetail.Col = 4: int_CS = int_CS + Val(dg_SelectedOrderDetail.Text)
    dg_SelectedOrderDetail.Col = 8: int_CutCS = int_CutCS + Val(dg_SelectedOrderDetail.Text)
Next

If int_CutCS = int_CS Then
   msg_text = "�`���νc�� ���� �q���`�c�ơA�нT�{!"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If

Screen.MousePointer = vbHourglass
'��Ʈw���ʥ��--�_�I
Tran_Level = 0
Tran_Level = cn.BeginTrans

'���s���ΥX�Ӥ��q��M�w�� [�q��s��]
strCutOrder_SrcKey = txt_Tab1_OrderKey.Text
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
blTRP02WEventEnable = False
rs_TRP02W.Filter = adFilterNone
rs_TRP02W.Filter = "�q��s�� = '" & strCutOrder_SrcKey & "'"
If rs_TRP02W.RecordCount = 0 Then
   msg_text = "��p���A�䤣��ŦX���󪺭�q���Ƴ�"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   rs_TRP02W.Filter = adFilterNone
   rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   Exit Sub
Else
   '���ͤ��έq��-Header
   intTRP02WBookMark = rs_TRP02W.Bookmark
   With dg_CutOrders
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0
        .Text = .Rows - 2
        .Col = 1: .Text = strCutOrder_NewKey '�q��s��
        .Col = 2: .Text = rs_TRP02W.Fields("�e�f��").Value '�e�f��
        .Col = 3: .Text = rs_TRP02W.Fields("�Ȥ�s��").Value '�Ȥ�s��
        .Col = 4: .Text = rs_TRP02W.Fields("�f�D�渹").Value '�f�D�渹
        .Col = 5: .Text = txt_Tab1_SelectedCaseQty.Text '�c��
        .Col = 6: .Text = txt_Tab1_SelectedWeight.Text '���q
        .Col = 7: .Text = txt_Tab1_SelectedVolumn.Text '���n
        .Col = 8: .Text = txt_Tab1_SelectedPalletQty.Text '�O��
        .Col = 9: .Text = rs_TRP02W.Fields("ZIP").Value '�l���ϸ�
        .Col = 10: .Text = rs_TRP02W.Fields("�ϽX").Value '�ϽX
        .Col = 11: .Text = rs_TRP02W.Fields("�Ȥ�W��").Value '�Ȥ�W��
        .Col = 12: .Text = rs_TRP02W.Fields("�q���").Value '�q���
        .Col = 13: .Text = rs_TRP02W.Fields("�f�D").Value '�f�D
        .Col = 14: .Text = "���έq��" '�ѧO
        .Col = 15: .Text = rs_TRP02W.Fields("EXE�^��").Value 'EXE�^��
        .Col = 0
        For i = 0 To .Cols - 1
            .ColSel = i
        Next i
        
        '���ͷs���q����--TRP02W
        str_SQL = "Insert into TRP02W (StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Description,Case_cnt,Weight,Volumn_Weight,Pallet_Qty,EXTERN,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,exe_confirm) " & _
                  "Select StorerKey,'" & strCutOrder_NewKey & "',C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Description," & _
                  txt_Tab1_SelectedCaseQty.Text & "," & txt_Tab1_SelectedWeight.Text & "," & txt_Tab1_SelectedVolumn.Text & "," & txt_Tab1_SelectedPalletQty.Text & ",EXTERN,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,exe_confirm " & _
                  "From TRP02W Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '��s��q�椧�έp�Ʀr--TRP02W
        str_SQL = "Update TRP02W Set Case_cnt=Case_cnt-" & txt_Tab1_SelectedCaseQty.Text & "," & _
                  "Weight=Weight-" & txt_Tab1_SelectedWeight.Text & ",Volumn_Weight=Volumn_Weight-" & txt_Tab1_SelectedVolumn.Text & ",Pallet_Qty=Pallet_Qty-" & txt_Tab1_SelectedPalletQty.Text & " " & _
                  "Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End With
End If
rs_TRP02W.Filter = adFilterNone
rs_TRP02W.Sort = "�q��s�� ASC"
Do While Not rs_TRP02W.EOF
   If rs_TRP02W.Bookmark = intTRP02WBookMark Then
        Do While dg_TRP02W.SelBookmarks.Count <> 0
           dg_TRP02W.SelBookmarks.Remove 0
        Loop
        '�ϥ���ܥ��b�i�� [�q����Χ@�~] ���q���ƦC
        dg_TRP02W.SelBookmarks.Add rs_TRP02W.Bookmark
      Exit Do
   End If
   rs_TRP02W.MoveNext
Loop
blTRP02WEventEnable = True

'���έq�椧 OrderDetail
Dim dbsrcQty As Double, dbCutQty As Double, dbSeqNo As Double, dbCutEAQty As Long, dbCasecntQty As Integer
dbSeqNo = 0
dg_SelectedOrderDetail.Visible = False
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         .Row = i: .Col = 1
         If .Text <> "" Then   '�Ӷ��Q����i�����
            .Col = 0: dbSeqNo = .Text          '�O�d��q�涵���s���H������
            .Col = 4: dbsrcQty = Val(.Text)    '��q��c��
            .Col = 8: dbCutQty = Val(.Text)    '���νc��
            .Col = 12: dbCutEAQty = Val(.Text) '���έӼ�
            .Col = 14: dbCasecntQty = Val(.Text) '�~���C�c�Ӽ�
            If dbsrcQty = dbCutQty Then        '�Y�������c�ƶi����ΡA���O�ǳƫ���R�����Ӷ�
               .Col = 1: .Text = "X"
               Call InsertInto_CutOrderDetail(strCutOrder_NewKey, dbSeqNo)
               str_SQL = "Update TRP03W Set Receipt_No = '" & strCutOrder_NewKey & "' " & _
                         "Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Else
               '��s �ݤ��έq�����
               .Col = 1: .Text = ""        '�M�����O�A��s����
               .Col = 4: dbsrcQty = Val(.Text)    '��q��c��
               .Col = 8: dbCutQty = Val(.Text)    '���νc��
               .Col = 4: .Text = dbsrcQty - dbCutQty
               .Col = 5: dbsrcQty = Val(.Text)    '��q�歫�q
               .Col = 9: dbCutQty = Val(.Text)    '���έ��q
               .Col = 5: .Text = dbsrcQty - dbCutQty
               .Col = 6: dbsrcQty = Val(.Text)    '��q����n
               .Col = 10: dbCutQty = Val(.Text)   '���Χ��n
               .Col = 6: .Text = dbsrcQty - dbCutQty
               .Col = 7: dbsrcQty = Val(.Text)    '��q��O��
               .Col = 11: dbCutQty = Val(.Text)   '���ΪO��
               .Col = 7: .Text = dbsrcQty - dbCutQty
               Call InsertInto_CutOrderDetail(strCutOrder_NewKey, dbSeqNo)
               
               '��s TRP03W ��ƶq
               str_SQL = "Update TRP03W Set Order_Qty =  "
'               .Col = 4: str_SQL = str_SQL & .Text & ",Weight = "'mark by gemini
               .Col = 4: str_SQL = str_SQL & (.Text * dbCasecntQty) & ",Weight = " 'add by gemini
               .Col = 5: str_SQL = str_SQL & .Text & ",Volumn_Weight = "
               .Col = 6: str_SQL = str_SQL & .Text & ",Pallet_Qty = "
               .Col = 7: str_SQL = str_SQL & .Text & " "
               str_SQL = str_SQL & "Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
               
'               '�N�c�ƴ���^�Ӽ� by gemini 20071212
'               str_SQL = "Update TRP03W Set TRP03W.Order_Qty = TRP03W.Order_Qty * s1.casecnt " & _
'                        "from trp03w trp03w join sku s on trp03w.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = trp03w.storerkey " & _
'                        "Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
'                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
               '�s�W�s�q�椧�q��Ӷ�
               str_SQL = "Insert into TRP03W (StorerKey,Extern,Receipt_No,Seq_No,Product_No,Ship_Unit,Order_Qty,Weight,Volumn_Weight,Pallet_Qty,Description) " & _
                         "Select StorerKey,Extern,'" & strCutOrder_NewKey & "',Seq_No,Product_No,Ship_Unit,"
'               .Col = 8: str_SQL = str_SQL & .Text & ","'mark by gemini
               str_SQL = str_SQL & (dbCutEAQty) & "," 'add by gemini
               .Col = 9: str_SQL = str_SQL & .Text & ","
               .Col = 10: str_SQL = str_SQL & .Text & ","
               .Col = 11: str_SQL = str_SQL & .Text & ","
               str_SQL = str_SQL & "Description From TRP03W Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
               
'               '�N�c�ƴ���^�Ӽ� by gemini 20071212
'               str_SQL = "Update TRP03W Set TRP03W.Order_Qty = TRP03W.Order_Qty * s1.casecnt " & _
'                        "from trp03w trp03w join sku s on trp03w.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = trp03w.storerkey " & _
'                        "Where Receipt_No = '" & strCutOrder_NewKey & "' and SEQ_NO = " & dbSeqNo & ""
'               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
               .Col = 8: .Text = ""   '���νc��
               .Col = 9: .Text = ""   '���έ��q
               .Col = 10: .Text = ""  '���Χ��n
               .Col = 11: .Text = ""  '���ΪO��
               .Col = 12: .Text = ""  '���έӼ�
            End If
         End If
     Next i
End With

'�R���H���ƶq���Τ��q��Ӷ�
Dim j As Double
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         For j = 1 To .Rows - 2
             .Row = j: .Col = 1
             If .Text = "X" Then
                Call Delete_GridRow(dg_SelectedOrderDetail, j)
                Exit For
             End If
         Next j
     Next i
     '���s���ͭq��[�`�έp���
     txt_Tab1_Weight.Text = 0
     txt_Tab1_Volumn.Text = 0
     txt_Tab1_PalletQty.Text = 0
     For i = 1 To .Rows - 2
         .Row = i
         .Col = 5: txt_Tab1_Weight.Text = Val(txt_Tab1_Weight.Text) + Val(.Text)
         .Col = 6: txt_Tab1_Volumn.Text = Val(txt_Tab1_Volumn.Text) + Val(.Text)
         .Col = 7: txt_Tab1_PalletQty.Text = Val(txt_Tab1_PalletQty.Text) + Val(.Text)
     Next i
End With
dg_SelectedOrderDetail.Visible = True

'�M������-��������έp
txt_Tab1_SelectedCaseQty.Text = ""
dbCut_TotalCaseQty = 0
txt_Tab1_SelectedWeight.Text = ""
dbCut_TotalWeight = 0
txt_Tab1_SelectedVolumn.Text = ""
dbCut_TotalVolumn = 0
txt_Tab1_SelectedPalletQty.Text = ""
dbCut_TotalPalletQty = 0

'�Ӷ����μƶq���G�O�ơA�c��
txt_Tab1_CutCaseQty.Text = ""
txt_Tab1_CutPalletQty.Text = ""
If dg_SelectedOrderDetail.Rows = 2 And txt_Tab1_Weight.Text = 0 And txt_Tab1_Volumn.Text = 0 And txt_Tab1_PalletQty.Text = 0 Then
   
   '�w�������Τ��q��G�R�� TRP02W & TRP03W
   str_SQL = "Delete From TRP02W Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Delete From TRP03W Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�M�� [�ݤ��έq��] �������q������ƭ�
   Call Clear_SelectedOrderData
   DoEvents
   SSTab1.Tab = 0
   DoEvents
End If

cn.CommitTrans: Tran_Level = 0
Screen.MousePointer = vbDefault

'��椧�q��Ӽ��ˬd
Dim rsTmp As New ADODB.Recordset
'str_SQL = " select * from gv_CheckOrderQty g join trp02w t2 on g.tms�渹 = t2.c_receipt_no where TMS�渹 = '" & strCutOrder_SrcKey & "' "

'����쥻�渹��c_receipt_no edit by Eric 20141230
str_SQL = "select " & _
            "TMS�渹 = od.orderkey " & _
            ",�q��q = sum(od.originalqty) " & _
            ",�ݱƨ��q = isnull((select sum(isnull(t3.order_qty,0)) from trp03w t3(nolock) join trp02w t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0) " & _
            ",�w�ƨ��q = isnull((select sum(isnull(t3.order_qty,0)) from trp03t t3(nolock) join trp02t t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0) " & _
            "from orderdetail od(nolock) join orders o(nolock) on o.orderkey = od.orderkey and isnull(o.type,'')<>'�R��' and priority not in ( 'R','A2B','RC') " & _
            "where od.orderkey in (select c_receipt_no from trp02w(nolock) where receipt_no = '" & strCutOrder_SrcKey & "') " & _
            "group by od.orderkey " & _
            "having sum(od.originalqty)<>isnull((select sum(isnull(t3.order_qty,0)) from trp03w t3(nolock) join trp02w t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0)+isnull((select sum(isnull(t3.order_qty,0)) from trp03t t3(nolock) join trp02t t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0) "
rsTmp.Open str_SQL, cn

If Not rsTmp.EOF Then MsgBox "�Ȥ��l�q��q�P���q��q���šA�нT�{!", vbOKOnly, Me.Caption
rsTmp.Close: Set rsTmp = Nothing

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   dg_SelectedOrderDetail.Visible = True
   blTRP02WEventEnable = True
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ݤ��έq��-�q�����", Me.Caption, "cmd_Tab1_CutOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_CutQty_Click()
'�ݤ��έq�� >> �ƶq����
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub
If txt_Tab1_SelectedPalletQty = 0 Then MsgBox "�C�O�c�Ƭ�0�A�L�k�i��q����ΡC", 64, Me.Caption: Exit Sub

Dim tmpQty As Double

cmd_Tab1_CutOrders.Enabled = False
cmd_Tab1_ClearQty.Enabled = False
If Val(txt_Tab1_CutCaseQty.Text) = 0 And Val(txt_Tab1_CutPalletQty.Text) = 0 Then
   '��ƶq�M����ܡG���������
   dg_SelectedOrderDetail.Col = 1: dg_SelectedOrderDetail.Text = ""   '�������
End If

If Val(txt_Tab1_CutPalletQty.Text) > 0 Then
   dg_SelectedOrderDetail.Col = 7: tmpQty = Val(dg_SelectedOrderDetail.Text)
   If Val(txt_Tab1_CutPalletQty.Text) > tmpQty Then
      msg_text = "��ƿ��~�G���ΪO�� �j�� �~���`�O��"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      cmd_Tab1_ClearQty.Enabled = True
      cmd_Tab1_CutOrders.Enabled = True
      Exit Sub
   End If
   '����J���ΪO�ơG�H�O�Ƭ��ǡA�M�� [���νc��] ����
   dg_SelectedOrderDetail.Col = 11
   dg_SelectedOrderDetail.Text = txt_Tab1_CutPalletQty.Text
   dg_SelectedOrderDetail.Col = 8
   dg_SelectedOrderDetail.Text = ""
   
   '�p����έӼ�
   dg_SelectedOrderDetail.Col = 13: tmpQty = Val(dg_SelectedOrderDetail.Text) * Val(txt_Tab1_CutPalletQty)
   dg_SelectedOrderDetail.Col = 14: tmpQty = Val(dg_SelectedOrderDetail.Text) * tmpQty
   dg_SelectedOrderDetail.Col = 12: dg_SelectedOrderDetail.Text = tmpQty: If dg_SelectedOrderDetail.Text = 0 Then dg_SelectedOrderDetail.Text = ""
Else
   dg_SelectedOrderDetail.Col = 4: tmpQty = Val(dg_SelectedOrderDetail.Text)
   If Val(txt_Tab1_CutCaseQty.Text) > tmpQty Then
      msg_text = "��ƿ��~�G���νc�� �j�� �~���`�c��"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      cmd_Tab1_ClearQty.Enabled = True
      cmd_Tab1_CutOrders.Enabled = True
      Exit Sub
   End If

   '��J���νc�ơG�c��
   dg_SelectedOrderDetail.Col = 11
   dg_SelectedOrderDetail.Text = ""
   dg_SelectedOrderDetail.Col = 8
   dg_SelectedOrderDetail.Text = txt_Tab1_CutCaseQty.Text
   
    '�p����έӼ�
   dg_SelectedOrderDetail.Col = 14: tmpQty = Val(dg_SelectedOrderDetail.Text) * Val(txt_Tab1_CutCaseQty)
   dg_SelectedOrderDetail.Col = 12: dg_SelectedOrderDetail.Text = tmpQty: If dg_SelectedOrderDetail.Text = 0 Then dg_SelectedOrderDetail.Text = ""
End If

'�ˬd���έӼƬO�_�����
dg_SelectedOrderDetail.Col = 12
If Val(dg_SelectedOrderDetail.Text) <> Int(Val(dg_SelectedOrderDetail.Text)) Then MsgBox "���έӼƤ��঳�p���I!", vbOKOnly, Me.Caption: Call cmd_Tab1_ClearQty_Click

'�p�������q��Ӷ����[�` [�c��] [���q] [�~�n] [�O��]
Call Calculate_SelectedPrderDetail

'�M�����ζq����
txt_Tab1_CutCaseQty.Text = ""
txt_Tab1_CutPalletQty.Text = ""
cmd_Tab1_ClearQty.Enabled = True
cmd_Tab1_CutOrders.Enabled = True
End Sub

Private Sub cmd_Tab1_ResetRS_Click()
'�����z��Ƨ�
If rs_TRP02W Is Nothing Then Exit Sub

'�����z�����A���]�ƧǨ̾�
 blTRP02WEventEnable = False
 rs_TRP02W.Filter = adFilterNone
 rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
 blTRP02WEventEnable = True
End Sub

Private Sub cmd_Tab2_CutOrderDelete_Click()
'�q����Ω��� >> �R��

Dim dbDeleteRow As Double, strOrderkey As String, strStorerkey As String, strExtern As String, strC_Receipt_no As String
With dg_CutOrders
     dbDeleteRow = .Row
     .Col = 1: strOrderkey = .Text      '�q��s�� Receipt_No
     .Col = 4: strExtern = .Text        '�f�D�渹 Extern
     .Col = 13: strStorerkey = .Text    '�f�D  StorerKey
     .Col = 16: strC_Receipt_no = .Text '��lTMS�渹 C_Receipt_no
     
     If .Text = "" Then Exit Sub
     If Left(strOrderkey, 2) <> "CT" Then MsgBox "�D���έq��L�k�R��!", vbOKOnly, Me.Caption: Exit Sub
     
     msg_text = "�R���@�~�G�T�{�R����������έq��G" & strOrderkey
     If MsgBox(msg_text, vbOKCancel + vbInformation, msg_title) = vbCancel Then Exit Sub
End With

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

'�ˮֱ��R�����q��G�H��lTMS�渹���d�߱���
str_SQL = "Select Count(*) as RecCount From TRP02W Where c_receipt_no = '" & strC_Receipt_no & "' and StorerKey = '" & strStorerkey & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCount").Value = 1 Then
   tmp_Rs.Close
   msg_text = "�q��s���G" & strOrderkey & " �����\�R���A�]���lTMS�渹�u���������q����!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
ElseIf tmp_Rs.Fields("RecCount").Value = 0 Then
   tmp_Rs.Close
   msg_text = "�q��s���G" & strOrderkey & " �w���s�b�A�Э��s����d��!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
tmp_Rs.Close

'���̤p�q��s���G�����Q�R���q��Ҧ������ءB�ƶq
Dim strToOrderKey As String
str_SQL = "Select Min(Receipt_No) as �����q��s�� From TRP02W Where C_Receipt_no = '" & strC_Receipt_no & "' and StorerKey = '" & strStorerkey & "' and Receipt_No <> '" & strOrderkey & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   strToOrderKey = tmp_Rs.Fields("�����q��s��").Value
Else
   tmp_Rs.Close
   msg_text = "�䤣��i�H�������R�����q�涵�����ؼЭq��!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
tmp_Rs.Close

'�H�U������ Gemini @20080324
Tran_Level = 0
Tran_Level = cn.BeginTrans
'��s�����q�椧������� TRP02W
With dg_CutOrders
     .Row = dbDeleteRow
     str_SQL = "Update TRP02W Set Case_cnt=Case_cnt+"
     .Col = 5: str_SQL = str_SQL & .Text & ",Weight=Weight+"
     .Col = 6: str_SQL = str_SQL & .Text & ",Volumn_Weight=Volumn_Weight+"
     .Col = 7: str_SQL = str_SQL & .Text & ",Pallet_Qty=Pallet_Qty+"
     .Col = 5: str_SQL = str_SQL & .Text & " "
     str_SQL = str_SQL & "Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strToOrderKey & "'"
     cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End With

'��s�����q�椧������� TRP03W
rs_CutOrderDetail.Filter = adFilterNone
rs_CutOrderDetail.Filter = "�q��s�� = '" & strOrderkey & "'"
If rs_CutOrderDetail.EOF Then
   msg_text = "��p���A�䤣��ŦX���󪺤l�q����Ӹ�Ƴ�"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   rs_CutOrderDetail.Filter = adFilterNone
   rs_CutOrderDetail.Sort = "�q��s��,���� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   Exit Sub
Else
   Do While Not rs_CutOrderDetail.EOF
      '���ݱ����q��s�����L�ۦP�����B�f�����q��Ӷ� TRP03W
      str_SQL = "Select Count(*) AS RecCount From TRP03W " & _
                "Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strToOrderKey & "' and " & _
                "      Seq_No = " & rs_CutOrderDetail.Fields("����").Value & " and Product_No = '" & rs_CutOrderDetail.Fields("�f��").Value & "'"
      tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
      If tmp_Rs.Fields("RecCount").Value = 0 Then
         '�s�W�Ӷ� TRP03W
         str_SQL = "Insert into TRP03W (StorerKey,EXTERN,Receipt_No,Seq_No,Product_No,Ship_Unit,Order_Qty,Weight,Volumn_Weight,Pallet_Qty,Description) " & _
                   "Select StorerKey,EXTERN,'" & strToOrderKey & "',Seq_No,Product_No,Ship_Unit,Order_Qty,Weight,Volumn_Weight,Pallet_Qty,Description " & _
                   "From TRP03W Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strOrderkey & "' and " & _
                   "      Seq_No = " & rs_CutOrderDetail.Fields("����").Value & " and Product_No = '" & rs_CutOrderDetail.Fields("�f��").Value & "'"
      Else
         '��s�Ӷ� TRP03W
         str_SQL = "Update TRP03W Set Order_Qty = Order_Qty + " & rs_CutOrderDetail.Fields("�c��").Value & "," & _
                   "Weight = Weight + " & rs_CutOrderDetail.Fields("���q").Value & "," & _
                   "Volumn_Weight = Volumn_Weight + " & rs_CutOrderDetail.Fields("���n").Value & "," & _
                   "Pallet_Qty = Pallet_Qty + " & rs_CutOrderDetail.Fields("�O��").Value & " " & _
                   "Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strToOrderKey & "' and " & _
                   "      Seq_No = " & rs_CutOrderDetail.Fields("����").Value & " and Product_No = '" & rs_CutOrderDetail.Fields("�f��").Value & "'"
      End If
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
      tmp_Rs.Close
      rs_CutOrderDetail.MoveNext
   Loop
   '�R���Ӷ�
   rs_CutOrderDetail.MoveFirst
   Do While Not rs_CutOrderDetail.EOF
      rs_CutOrderDetail.Delete
      rs_CutOrderDetail.MoveFirst
   Loop
   str_SQL = "Delete From TRP03W Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strOrderkey & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End If
rs_CutOrderDetail.Filter = adFilterNone
rs_CutOrderDetail.Sort = "�q��s��,���� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj

'�R���l�q����Y
Call Delete_GridRow(dg_CutOrders, dbDeleteRow)
'�R���q��D�� TRP02W
str_SQL = "Delete From TRP02W Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strOrderkey & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans
Tran_Level = 0
Screen.MousePointer = vbDefault

'��椧�q��Ӽ��ˬd
Dim rsTmp As New ADODB.Recordset
str_SQL = " select * from gv_CheckOrderQty g join trp02w t2 on g.tms�渹 = t2.c_receipt_no where TMS�渹 = '" & strToOrderKey & "' "
rsTmp.Open str_SQL, cn
If Not rsTmp.EOF Then MsgBox "�Ȥ��l�q��q�P���q��q���šA�нT�{!", vbOKOnly, Me.Caption
rsTmp.Close: Set rsTmp = Nothing

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q����Ω���-�R��", Me.Caption, "cmd_Tab2_CutOrderDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_ExternQuery_Click()
'�q����Ω���+�d�� >> �d��
If Len(Trim(txt_Tab2_Extern.Text)) = 0 Then Exit Sub

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle

'�]�w���έq��C��
Call SetGrid_Format_CutOrderList
'�]�w�w�������έq����Ӫ�
Call CreateRS_CutOrderDetail

str_SQL = "Select �q��s��,�e�f��,�Ȥ�s��,�f�D�渹,�c��,���q,���n,�O��,ZIP,�ϽX,�Ȥ�W��,�q���,�f�D,�ѧO,EXE�^��,��lTMS�渹 " & _
          "From CutOrders_SourceOrder Where �f�D�渹 like '" & Trim(txt_Tab2_Extern.Text) & "%' Order by �q��s�� "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����(TRP02W)"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Do While Not tmp_Rs.EOF
   With dg_CutOrders
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0
        .Text = .Rows - 2
        .Col = 1    '�q��s��
        .Text = tmp_Rs.Fields("�q��s��").Value
        .Col = 2    '�e�f��
        .Text = tmp_Rs.Fields("�e�f��").Value
        .Col = 3    '�Ȥ�s��
        .Text = tmp_Rs.Fields("�Ȥ�s��").Value
        .Col = 4    '�f�D�渹
        .Text = tmp_Rs.Fields("�f�D�渹").Value
        .Col = 5    '�c��
        .Text = tmp_Rs.Fields("�c��").Value
        .Col = 6    '���q
        .Text = tmp_Rs.Fields("���q").Value
        .Col = 7    '���n
        .Text = tmp_Rs.Fields("���n").Value
        .Col = 8    '�O��
        .Text = tmp_Rs.Fields("�O��").Value
        .Col = 9    '�l���ϸ�
        .Text = tmp_Rs.Fields("ZIP").Value
        .Col = 10   '�ϽX
        .Text = tmp_Rs.Fields("�ϽX").Value
        .Col = 11   '�Ȥ�W��
        .Text = tmp_Rs.Fields("�Ȥ�W��").Value
        .Col = 12   '�q���
        .Text = tmp_Rs.Fields("�q���").Value
        .Col = 13   '�f�D
        .Text = tmp_Rs.Fields("�f�D").Value
        .Col = 14   '�ѧO
        .Text = tmp_Rs.Fields("�ѧO").Value
        .Col = 15   'EXE�^��
        .Text = tmp_Rs.Fields("EXE�^��").Value
        .Col = 16: .Text = tmp_Rs.Fields("��lTMS�渹").Value '��lTMS�渹
        .Col = 0
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close

'TRP03W
str_SQL = "Select �q��s��,����,�f��,�~�W,�c��,���q,���n,�O�� " & _
          "From CutOrders_SelectedOrderDetail Where �f�D�渹 like '" & Trim(txt_Tab2_Extern.Text) & "%' order by �q��s��,����"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����Ӹ��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Do While Not tmp_Rs.EOF
   rs_CutOrderDetail.AddNew
   rs_CutOrderDetail.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
   rs_CutOrderDetail.Fields("����").Value = tmp_Rs.Fields("����").Value
   rs_CutOrderDetail.Fields("�f��").Value = tmp_Rs.Fields("�f��").Value
   rs_CutOrderDetail.Fields("�~�W").Value = tmp_Rs.Fields("�~�W").Value
   rs_CutOrderDetail.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
   rs_CutOrderDetail.Fields("���q").Value = tmp_Rs.Fields("���q").Value
   rs_CutOrderDetail.Fields("���n").Value = tmp_Rs.Fields("���n").Value
   rs_CutOrderDetail.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
   rs_CutOrderDetail.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q����ΦW��+�d��-�f�D�渹�d��", Me.Caption, "cmd_Tab2_ExternQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub dg_CutOrders_Click()
'�w�������Τ��l�q��C��
'�I�@���G���
Dim i As Double
With dg_CutOrders
     .Col = 1   '�q��s��
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub dg_SelectedOrderDetail_Click()
'�ݤ��Τ��q��G�q����Ӷ���
'�I�@���G����A���D�M�� [���μƶq] �_�h�����O�� [���] ���A

txt_Tab1_CutPalletQty.Text = ""
txt_Tab1_CutCaseQty.Text = ""

Dim i As Integer
Dim tmpQty As Double
With dg_SelectedOrderDetail
     .Col = 2   '�f��
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 1
     If Len(.Text) = 0 Then
        .Text = "V"
        .Col = 4   '��ܩҿ�����c��
        tmpQty = .Text
        dbCut_TotalCaseQty = dbCut_TotalCaseQty + .Text
        txt_Tab1_SelectedCaseQty.Text = dbCut_TotalCaseQty
        .Col = 8: .Text = tmpQty
        txt_Tab1_CutCaseQty.Text = tmpQty
        
        '�p��ҿ�����Ӽ�
        .Col = 14: tmpQty = .Text * tmpQty
        .Col = 12: .Text = tmpQty
        
        .Col = 5   '��ܩҿ�������q
        tmpQty = .Text
        dbCut_TotalWeight = dbCut_TotalWeight + .Text
        txt_Tab1_SelectedWeight.Text = dbCut_TotalWeight
        .Col = 9: .Text = tmpQty
        
        .Col = 6   '��ܩҿ�������n
        tmpQty = .Text
        dbCut_TotalVolumn = dbCut_TotalVolumn + .Text
        txt_Tab1_SelectedVolumn.Text = dbCut_TotalVolumn
        .Col = 10: .Text = tmpQty
        
        .Col = 7   '��ܩҿ�����O��
        tmpQty = .Text
        dbCut_TotalPalletQty = dbCut_TotalPalletQty + .Text
        txt_Tab1_SelectedPalletQty.Text = dbCut_TotalPalletQty
        .Col = 11: .Text = tmpQty
        txt_Tab1_CutPalletQty.Text = tmpQty
     Else
        .Col = 11   '���Τ��O��
        If Val(.Text) <> 0 Then
           txt_Tab1_CutPalletQty.Text = .Text
        End If
        .Col = 8   '���Τ��c��
        If Val(.Text) <> 0 Then
           txt_Tab1_CutCaseQty.Text = .Text
        End If
     End If
     '�ϥտ������Ʀ�
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub dg_TRP02W_HeadClick(ByVal ColIndex As Integer)
'�H�ƹ��I�� dg_TRP02W �����D��
Dim OrderFieldName As String
If TypeName(rs_TRP02W) <> "Nothing" Then
   OrderFieldName = "[" & dg_TRP02W.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_TRP02W.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_TRP02W.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

Private Sub dg_TRP02W_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If blTRP02WEventEnable Then
   With dg_TRP02W
        'Do While .SelBookmarks.Count <> 0
        '   dg_TRP02W.SelBookmarks.Remove 0
        'Loop
        '�ϥ���ܿ������ƦC
        dg_TRP02W.SelBookmarks.Add rs_TRP02W.Bookmark
   End With
End If
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "����h���q�����"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'�]�w�����έq�椧�q��W��
Call SetGrid_Format_SelectedOrderDetail

'�]�w�w�������έq��C��
Call SetGrid_Format_CutOrderList
'�]�w�w�������έq����Ӫ�
Call CreateRS_CutOrderDetail

End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   dg_TRP02W.Width = dg_TRP02W.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_TRP02W.Height = dg_TRP02W.Height - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(19).Top = Label1(19).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_OrderCount.Top = txt_Tab0_OrderCount.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(3).Top = Label1(3).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalCase.Top = txt_Tab0_TotalCase.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(0).Top = Label1(0).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalWeight.Top = txt_Tab0_TotalWeight.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(1).Top = Label1(1).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalVolumn.Top = txt_Tab0_TotalVolumn.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(2).Top = Label1(2).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalPallet.Top = txt_Tab0_TotalPallet.Top - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab1_Orders.Left = fam_Tab1_Orders.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_OrderDetail.Height = fam_Tab1_OrderDetail.Height - (dbsrcFormHeight - Me.ScaleHeight)
   fam_Tab1_OrderDetail.Width = fam_Tab1_OrderDetail.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_SelectedOrderDetail.Height = dg_SelectedOrderDetail.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_SelectedOrderDetail.Width = dg_SelectedOrderDetail.Width - (dbsrcFormWidth - Me.ScaleWidth)
   
   dg_CutOrders.Width = dg_CutOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_CutOrderDetail.Width = dg_CutOrderDetail.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_CutOrderDetail.Height = dg_CutOrderDetail.Height - (dbsrcFormHeight - Me.ScaleHeight)
   fam_Tab2_Qoery.Left = fam_Tab2_Qoery.Left - (dbsrcFormWidth - Me.ScaleWidth)
   fam_Tab2_Delete.Left = fam_Tab2_Delete.Left - (dbsrcFormWidth - Me.ScaleWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   dg_TRP02W.Width = dg_TRP02W.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_TRP02W.Height = dg_TRP02W.Height + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(19).Top = Label1(19).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_OrderCount.Top = txt_Tab0_OrderCount.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(3).Top = Label1(3).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalCase.Top = txt_Tab0_TotalCase.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(0).Top = Label1(0).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalWeight.Top = txt_Tab0_TotalWeight.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(1).Top = Label1(1).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalVolumn.Top = txt_Tab0_TotalVolumn.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(2).Top = Label1(2).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalPallet.Top = txt_Tab0_TotalPallet.Top + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab1_Orders.Left = fam_Tab1_Orders.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_OrderDetail.Height = fam_Tab1_OrderDetail.Height + (Me.ScaleHeight - dbsrcFormHeight)
   fam_Tab1_OrderDetail.Width = fam_Tab1_OrderDetail.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_SelectedOrderDetail.Height = dg_SelectedOrderDetail.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_SelectedOrderDetail.Width = dg_SelectedOrderDetail.Width + (Me.ScaleWidth - dbsrcFormWidth)
   
   dg_CutOrders.Width = dg_CutOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_CutOrderDetail.Width = dg_CutOrderDetail.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_CutOrderDetail.Height = dg_CutOrderDetail.Height + (Me.ScaleHeight - dbsrcFormHeight)
   fam_Tab2_Qoery.Left = fam_Tab2_Qoery.Left + (Me.ScaleWidth - dbsrcFormWidth)
   fam_Tab2_Delete.Left = fam_Tab2_Delete.Left + (Me.ScaleWidth - dbsrcFormWidth)
   
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
Set frm_OP_CutOrders = Nothing
End Sub

Private Sub Clear_SelectedOrderData()
'�M�� [�ݤ��έq��] Orders ������
dbCut_TotalCaseQty = 0
txt_Tab1_SelectedCaseQty.Text = ""
dbCut_TotalWeight = 0
txt_Tab1_SelectedWeight.Text = ""
dbCut_TotalVolumn = 0
txt_Tab1_SelectedVolumn.Text = ""
dbCut_TotalPalletQty = 0
txt_Tab1_SelectedPalletQty.Text = ""

txt_Tab1_CutCaseQty.Text = ""
txt_Tab1_CutPalletQty.Text = ""

txt_Tab1_Storer.Text = ""
txt_Tab1_OrderKey.Text = ""
txt_Tab1_Extern.Text = ""
txt_Tab1_OrderDate.Text = ""
txt_Tab1_DeliveryDate.Text = ""
txt_Tab1_FullName.Text = ""
txt_Tab1_Address.Text = ""
txt_Tab1_ExtraDemand1.Text = ""
txt_Tab1_ExtraDemand2.Text = ""
txt_Tab1_ZIP.Text = ""
txt_Tab1_AreaCode.Text = ""
txt_Tab1_VehicleType.Text = ""
txt_Tab1_ChannelType.Text = ""
chk_Tab1_MultiCustomer.Value = vbUnchecked
txt_Tab1_Weight.Text = ""
txt_Tab1_Volumn.Text = ""
txt_Tab1_PalletQty.Text = ""
 txt_Tab1_EXEConfirm.Text = ""
End Sub

Private Sub SetGrid_Format_SelectedOrderDetail()
'����@���ݤ��έq�椧���ة���
Dim sub_var1 As Integer, sub_var2 As Integer
dg_SelectedOrderDetail.Visible = False
With dg_SelectedOrderDetail
     .Rows = 2: .Cols = 15
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
     .ColWidth(2) = 800
     .ColWidth(3) = 2200
     .ColWidth(4) = 600
     .ColWidth(5) = 700
     .ColWidth(6) = 700
     .ColWidth(7) = 700
     .ColWidth(8) = 850
     .ColWidth(9) = 850
     .ColWidth(10) = 850
     .ColWidth(11) = 850
     .ColWidth(12) = 850
     .ColWidth(13) = 850
     .ColWidth(14) = 850
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "����"
     .Col = 1: .Text = "��"
     .Col = 2: .Text = "�f��"
     .Col = 3: .Text = "�~�W"
     .Col = 4: .Text = "�c��"
     .Col = 5: .Text = "���q"
     .Col = 6: .Text = "���n"
     .Col = 7: .Text = "�O��"
     .Col = 8: .Text = "���νc��"
     .Col = 9: .Text = "���έ��q"
     .Col = 10: .Text = "���Χ��n"
     .Col = 11: .Text = "���ΪO��"
     .Col = 12: .Text = "���έӼ�"
     .Col = 13: .Text = "�C�O�c��"
     .Col = 14: .Text = "�C�c�Ӽ�"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignLeftCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignRightCenter
     .ColAlignment(10) = flexAlignRightCenter
     .ColAlignment(11) = flexAlignRightCenter
     .ColAlignment(12) = flexAlignRightCenter
     .ColAlignment(13) = flexAlignRightCenter
     .ColAlignment(14) = flexAlignRightCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignLeftCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_SelectedOrderDetail.Visible = True
End Sub

Private Sub Delete_GridRow(ByRef dgDataGrid As MSHFlexGrid, ByVal intRow As Double)
'�ݤ��έq�涵��(Detail) ��ƧR��
If intRow = 0 Then Exit Sub

Dim i As Double, j As Integer

'1. �N�R���C��ƥѤU�@�C��ƨ��N
'   �ӫ᪺��ƦC���W���@�C
With dgDataGrid
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

Private Sub Calculate_SelectedPrderDetail()
'�p�������q��Ӷ��G�c�ơA���q�A�~�n�A�O��

dbCut_TotalCaseQty = 0
txt_Tab1_SelectedCaseQty.Text = ""
dbCut_TotalWeight = 0
txt_Tab1_SelectedWeight.Text = ""
dbCut_TotalVolumn = 0
txt_Tab1_SelectedVolumn.Text = ""
dbCut_TotalPalletQty = 0
txt_Tab1_SelectedPalletQty.Text = ""

Dim dbCaseQty As Double, dbWeight As Double, dbVolumn As Double, dbPalletQty As Double, tmpQty As Long, dbCutEAQty As Long, dbTiHiQty As Long, dbCasecntQty As Integer
Dim dbCutPLQty As Double, dbCutCSQty As Double
Dim i As Double
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         .Row = i
         .Col = 1
         If .Text <> "" Then   '�Q���
            .Col = 4: dbCaseQty = Val(.Text)     '�c��
            .Col = 5: dbWeight = Val(.Text)      '���q
            .Col = 6: dbVolumn = Val(.Text)      '���n
            .Col = 7: dbPalletQty = Val(.Text)   '�O��
            .Col = 12: dbCutEAQty = Val(.Text) '���έӼ�
            .Col = 13: dbTiHiQty = Val(.Text) '�C�O�c��
            .Col = 14: dbCasecntQty = Val(.Text) '�C�c�Ӽ�
            .Col = 11   '���ΪO��
            If Val(.Text) <> 0 Then '�����ΪO��
               dbCutPLQty = Val(.Text)
               '���ΪO�ƴ��⤧�c��
               .Col = 8: .Text = dbCutEAQty / dbCasecntQty
               dbCut_TotalCaseQty = dbCut_TotalCaseQty + .Text
               
              '���ΪO�ƴ��⤧���q
              .Col = 9: .Text = ((dbCutPLQty / dbPalletQty) * dbWeight)
               dbCut_TotalWeight = dbCut_TotalWeight + .Text
               
               '���νc�ƴ��⤧���n
               .Col = 10: .Text = ((dbCutPLQty / dbPalletQty) * dbVolumn)
               dbCut_TotalVolumn = dbCut_TotalVolumn + .Text
               
               dbCut_TotalPalletQty = dbCut_TotalPalletQty + dbCutPLQty
            Else
               .Col = 8   '���νc��
               If Val(.Text) <> 0 Then
                  dbCutCSQty = Val(.Text)
                  dbCut_TotalCaseQty = dbCut_TotalCaseQty + dbCutCSQty
                 .Col = 9   '���νc�ƴ��⤧���q
                 .Text = ((dbCutCSQty / dbCaseQty) * dbWeight)
                  dbCut_TotalWeight = dbCut_TotalWeight + ((dbCutCSQty / dbCaseQty) * dbWeight)
                 .Col = 10   '���νc�ƴ��⤧���n
                 .Text = ((dbCutCSQty / dbCaseQty) * dbVolumn)
                  dbCut_TotalVolumn = dbCut_TotalVolumn + ((dbCutCSQty / dbCaseQty) * dbVolumn)
                 
                 '���νc�ƴ��⤧�O��
                 .Col = 11: .Text = (dbCutEAQty / dbTiHiQty / dbCasecntQty)
                  dbCut_TotalPalletQty = dbCut_TotalPalletQty + .Text
               End If
            End If
         Else
            .Col = 9: .Text = ""
            .Col = 10: .Text = ""
         End If
     Next i
End With
'��ܿ�����Ӷ��U��줧�[�`��
txt_Tab1_SelectedCaseQty.Text = dbCut_TotalCaseQty
txt_Tab1_SelectedWeight.Text = dbCut_TotalWeight
txt_Tab1_SelectedVolumn.Text = dbCut_TotalVolumn
txt_Tab1_SelectedPalletQty.Text = dbCut_TotalPalletQty

End Sub
Private Sub SetGrid_Format_CutOrderList()
'�w������Τ��q��C��
Dim sub_var1 As Integer, sub_var2 As Integer
dg_CutOrders.Visible = False
With dg_CutOrders
     .Rows = 2: .Cols = 17
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
     .ColWidth(0) = 500
     .ColWidth(1) = 1100
     .ColWidth(2) = 900
     .ColWidth(3) = 1200
     .ColWidth(4) = 800
     .ColWidth(5) = 800
     .ColWidth(6) = 800
     .ColWidth(7) = 800
     .ColWidth(8) = 800
     .ColWidth(9) = 400
     .ColWidth(10) = 500
     .ColWidth(11) = 3500
     .ColWidth(12) = 1000
     .ColWidth(13) = 700
     .ColWidth(14) = 800
     .ColWidth(15) = 800
     .ColWidth(16) = 1200
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "����"
     .Col = 1: .Text = "�q��s��"
     .Col = 2: .Text = "�e�f��"
     .Col = 3: .Text = "�Ȥ�s��"
     .Col = 4: .Text = "�f�D�渹"
     .Col = 5: .Text = "�c��"
     .Col = 6: .Text = "���q"
     .Col = 7: .Text = "���n"
     .Col = 8: .Text = "�O��"
     .Col = 9: .Text = "ZIP"
     .Col = 10: .Text = "�ϽX"
     .Col = 11: .Text = "�Ȥ�W��"
     .Col = 12: .Text = "�q���"
     .Col = 13: .Text = "�f�D"
     .Col = 14: .Text = "�ѧO"
     .Col = 15: .Text = "EXE�^��"
     .Col = 16: .Text = "��lTMS�渹"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignCenterCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignLeftCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .ColAlignment(10) = flexAlignCenterCenter
     .ColAlignment(11) = flexAlignLeftCenter
     .ColAlignment(12) = flexAlignLeftCenter
     .ColAlignment(13) = flexAlignLeftCenter
     .ColAlignment(14) = flexAlignLeftCenter
     .ColAlignment(15) = flexAlignLeftCenter
     .ColAlignment(16) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_CutOrders.Visible = True
End Sub

Private Sub CreateRS_CutOrderDetail()
'�w������Τ��q�����
Call ReDim_Recordset(rs_CutOrderDetail)
With rs_CutOrderDetail
     .Fields.Append "�q��s��", adVarChar, 10
     .Fields.Append "����", adDouble
     .Fields.Append "�f��", adVarChar, 20
     .Fields.Append "�~�W", adVarChar, 60
     .Fields.Append "�c��", adDouble
     .Fields.Append "���q", adDouble
     .Fields.Append "���n", adDouble
     .Fields.Append "�O��", adDouble
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
Set dg_CutOrderDetail.DataSource = rs_CutOrderDetail
'�]�w������
With dg_CutOrderDetail
    .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
    .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
    .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
    .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
    .Columns(0).Width = 1000        '�q��s��
    .Columns(0).Alignment = dbgLeft
    .Columns(1).Width = 800         '����
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2000         '�f��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 2400        '�~�W
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 800         '�c��
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 800         '���q
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 800         '���n
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800         '�O��
    .Columns(7).Alignment = dbgRight
End With
End Sub

Private Sub InsertInto_CutOrderDetail(strOrderkey As String, SeqNo As Double)
'�N [�ݤ��έq��]-�q����� �ثe��ƦC
'  ��J [���έq�����] �����Ӷ��� Recordset
rs_CutOrderDetail.AddNew
rs_CutOrderDetail.Fields("�q��s��").Value = strOrderkey
rs_CutOrderDetail.Fields("����").Value = SeqNo
With dg_SelectedOrderDetail
     .Col = 2
     rs_CutOrderDetail.Fields("�f��").Value = .Text
     .Col = 3
     rs_CutOrderDetail.Fields("�~�W").Value = .Text
     .Col = 8
     rs_CutOrderDetail.Fields("�c��").Value = .Text
     .Col = 9
     rs_CutOrderDetail.Fields("���q").Value = .Text
     .Col = 10
     rs_CutOrderDetail.Fields("���n").Value = .Text
     .Col = 11
     rs_CutOrderDetail.Fields("�O��").Value = .Text
End With
rs_CutOrderDetail.Update
End Sub

Public Sub frm_OP_CutOrders_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
'��椽�ΰƵ{���A�� frm_RS_FilterAndSort ���I�s
'�ǤJ�ȡGstrCode      �ʧ@�ѧO�X
'                     [FILTER] �ۭq�z��    [SORT] �Ƨ�
'        strReturn    �z�� or �Ƨ� ���]�w�r��

Select Case strCode
       Case "FILTER"  '�ۭq�z��
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP02W"   '�w�s�d�ߩ��Ӹ��
                        blTRP02WEventEnable = False
                        rs_TRP02W.Filter = adFilterNone
                        rs_TRP02W.Filter = strReturn
                        If rs_TRP02W.RecordCount = 0 Then
                           msg_text = "��p���A�䤣��ŦX���󪺸�Ƴ�"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_TRP02W.Filter = adFilterNone
                           rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
                           blTRP02WEventEnable = True
                           Exit Sub
                        End If
                        blTRP02WEventEnable = True
            End Select
       Case "SORT"    '�Ƨ�
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP02W"   '�ܯ��p����Ӹ��
                        blTRP02WEventEnable = False
                        rs_TRP02W.Sort = strReturn
                        blTRP02WEventEnable = True
            End Select
End Select
End Sub

Private Sub txt_Tab2_Extern_KeyPress(KeyAscii As Integer)
'�q����Ω��� + �d�� >> �f�D�渹
If KeyAscii = vbKeyReturn Then
   cmd_Tab2_ExternQuery.SetFocus
End If
End Sub
