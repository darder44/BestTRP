VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Query_PalletRent 
   Caption         =   "�����p��"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11325
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "�Ȥ᯲���p��"
      TabPicture(0)   =   "frm_Query_PalletRent.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmnDialog"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_DateS"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Query"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmd_Exit(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_HeadExcel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_DateE"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lst_Cust"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtUnitPrice"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "�q�������p��"
      TabPicture(1)   =   "frm_Query_PalletRent.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtUnitPrite1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtCarNo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_DateS_Tab1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmd_Query_Tab1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd_Exit(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt_DateE_Tab1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label17"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label15"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label14"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label13"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Shape1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label11"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.TextBox txtUnitPrite1 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -69240
         TabIndex        =   11
         Top             =   1020
         Width           =   1395
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   1080
         Width           =   1395
      End
      Begin VB.TextBox txtCarNo 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -71040
         TabIndex        =   10
         Top             =   1020
         Width           =   1395
      End
      Begin VB.TextBox txt_DateS_Tab1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74640
         TabIndex        =   8
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '����
         BackColor       =   &H00C0C000&
         ForeColor       =   &H80000008&
         Height          =   4800
         Left            =   -74880
         TabIndex        =   31
         Top             =   1860
         Width           =   11085
         Begin VB.Frame Frame5 
            Caption         =   "�W�����l"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   2055
            Begin VB.TextBox txt_SumA_Tab1 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   960
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.Label Label6 
               BackStyle       =   0  '�z��
               Caption         =   "���l"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   44
               Top             =   420
               Width           =   615
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "�������l"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2520
            TabIndex        =   32
            Top             =   240
            Width           =   4815
            Begin VB.TextBox txt_Sum_Tab1 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   3720
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.TextBox txt_Out_Tab1 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   2280
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.TextBox txt_In_Tab1 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   720
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.Label Label9 
               BackStyle       =   0  '�z��
               Caption         =   "���l"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3120
               TabIndex        =   38
               Top             =   420
               Width           =   615
            End
            Begin VB.Label Label8 
               BackStyle       =   0  '�z��
               Caption         =   "�٤J"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1680
               TabIndex        =   37
               Top             =   420
               Width           =   615
            End
            Begin VB.Label Label7 
               BackStyle       =   0  '�z��
               Caption         =   "�ɥX"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   420
               Width           =   495
            End
         End
         Begin MSDataGridLib.DataGrid dg_PalletCDS 
            Height          =   3360
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   5927
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   16
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
               Size            =   9.75
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
      Begin VB.CommandButton cmd_Query_Tab1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�d  ��"
         DownPicture     =   "frm_Query_PalletRent.frx":0038
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
         Left            =   -67440
         Picture         =   "frm_Query_PalletRent.frx":17BA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   12
         Top             =   600
         Width           =   1050
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
         Height          =   990
         Index           =   1
         Left            =   -65040
         Picture         =   "frm_Query_PalletRent.frx":1BFC
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   600
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
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
         Height          =   990
         Left            =   -66240
         Picture         =   "frm_Query_PalletRent.frx":203E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   13
         Top             =   600
         Width           =   1050
      End
      Begin VB.TextBox txt_DateE_Tab1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -72840
         TabIndex        =   9
         Top             =   1020
         Width           =   1395
      End
      Begin VB.ComboBox lst_Cust 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txt_DateE 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   1395
      End
      Begin VB.CommandButton cmd_HeadExcel 
         BackColor       =   &H00C0FFFF&
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
         Height          =   990
         Left            =   8760
         Picture         =   "frm_Query_PalletRent.frx":2908
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   660
         Width           =   1050
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
         Height          =   990
         Index           =   0
         Left            =   9960
         Picture         =   "frm_Query_PalletRent.frx":31D2
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         Top             =   660
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Query 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�d  ��"
         DownPicture     =   "frm_Query_PalletRent.frx":3614
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
         Picture         =   "frm_Query_PalletRent.frx":4D96
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   660
         Width           =   1050
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '����
         BackColor       =   &H00C0C000&
         ForeColor       =   &H80000008&
         Height          =   4800
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   11085
         Begin VB.Frame Frame4 
            Caption         =   "�W�����l"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   2055
            Begin VB.TextBox txt_SumA 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   960
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.Label Label12 
               BackStyle       =   0  '�z��
               Caption         =   "���l"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   30
               Top             =   420
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "�������l"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2520
            TabIndex        =   21
            Top             =   240
            Width           =   4815
            Begin VB.TextBox txt_In 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   720
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.TextBox txt_Out 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   2160
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.TextBox txt_Sum 
               Alignment       =   1  '�a�k���
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   3600
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   360
               Width           =   795
            End
            Begin VB.Label Label2 
               BackStyle       =   0  '�z��
               Caption         =   "�ɥX"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   420
               Width           =   495
            End
            Begin VB.Label Label3 
               BackStyle       =   0  '�z��
               Caption         =   "�٤J"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   26
               Top             =   420
               Width           =   615
            End
            Begin VB.Label Label4 
               BackStyle       =   0  '�z��
               Caption         =   "���l"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   12
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3000
               TabIndex        =   25
               Top             =   420
               Width           =   615
            End
         End
         Begin MSDataGridLib.DataGrid dg_PalletDetail 
            Height          =   3360
            Left            =   240
            TabIndex        =   7
            Top             =   1320
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   5927
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   16
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
               Size            =   9.75
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
      Begin VB.TextBox txt_DateS 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   1395
      End
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   8640
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label17 
         BackStyle       =   0  '�z��
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69120
         TabIndex        =   47
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  '�z��
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   46
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71040
         TabIndex        =   45
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '�z��
         Caption         =   "~~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         TabIndex        =   41
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '�z��
         Caption         =   "���w����d��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   40
         Top             =   720
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00008080&
         BorderWidth     =   2
         Height          =   1260
         Index           =   0
         Left            =   -74760
         Top             =   480
         Width           =   10965
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '�z��
         Caption         =   "~~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72600
         TabIndex        =   39
         Top             =   1260
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '�z��
         Caption         =   "~~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�W��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '�z��
         Caption         =   "���w����d��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00008080&
         BorderWidth     =   2
         Height          =   1260
         Index           =   4
         Left            =   240
         Top             =   480
         Width           =   10965
      End
   End
End
Attribute VB_Name = "frm_Query_PalletRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private disp_rsd As ADODB.Recordset
Private disp_rsd_Tab1 As ADODB.Recordset
Private i As Integer

Private Sub cmd_Exit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmd_HeadExcel_Click()
    If IsNumeric(txtUnitPrice.Text) = False Then MsgBox "����������~!!", vbOKOnly, Me.Caption: txtUnitPrice.SetFocus: Exit Sub
    If disp_rsd Is Nothing Then Exit Sub
    If disp_rsd.RecordCount = 0 Then Exit Sub
    '�N��Ƽg�Jexcel��
    Dim MyXlsApp As Excel.Application   '�}��excel��
    Dim objFld As Field
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '�s�WWookbooks
    MyXlsApp.Workbooks.Add
    '�s�WSheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "�Ȥ᯲���p��"
    MyXlsApp.ActiveSheet.Name = "�Ȥ᯲���p��"
    i = 3
    MyXlsApp.Cells(i, 1).Value = "��b��_"
    MyXlsApp.Cells(i, 2).Value = Trim(Me.txt_DateS.Text)
    MyXlsApp.Cells(i, 3).Value = "�W�����l"
    MyXlsApp.Cells(i, 4).Value = Trim(Me.txt_SumA.Text)
    MyXlsApp.Cells(i, 5).Value = "�Ȥ�W��"
    MyXlsApp.Cells(i, 6).Value = Trim(Me.lst_Cust.Text)
    i = i + 1
    MyXlsApp.Cells(i, 1).Value = "��b�騴"
    MyXlsApp.Cells(i, 2).Value = Trim(Me.txt_DateE.Text)
    MyXlsApp.Cells(i, 3).Value = "�������l"
    MyXlsApp.Cells(i, 4).Value = Trim(Me.txt_Sum.Text)
    MyXlsApp.Cells(i, 5).Value = "��b���"
    MyXlsApp.Cells(i, 6).Value = Format(Now, "yyyymmdd")
    i = i + 2
    MyXlsApp.Cells(i, 1).Value = "���"
    MyXlsApp.Cells(i, 2).Value = "�ɤJ"
    MyXlsApp.Cells(i, 3).Value = "�٦^"
    MyXlsApp.Cells(i, 4).Value = "��鵲�l"
    MyXlsApp.Cells(i, 5).Value = Trim(Me.txt_SumA.Text)
    MyXlsApp.Cells(i, 6).Value = "����"
    i = i + 1
    disp_rsd.MoveFirst
    '���,�ɥX,�٤J,��鵲�l,�֭p���l,����
    Dim intSum2 As Integer, intSum3 As Integer, intSum4 As Integer, intSum6 As Double
    Do While Not disp_rsd.EOF
    MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
    MyXlsApp.Cells(i, 1).Value = Trim(disp_rsd.Fields(1))
    MyXlsApp.Cells(i, 2).Value = disp_rsd.Fields(2): intSum2 = intSum2 + disp_rsd.Fields(2)
    MyXlsApp.Cells(i, 3).Value = disp_rsd.Fields(3): intSum3 = intSum3 + disp_rsd.Fields(3)
    MyXlsApp.Cells(i, 4).Value = disp_rsd.Fields(4): intSum4 = intSum4 + disp_rsd.Fields(4)
    MyXlsApp.Cells(i, 5).Value = MyXlsApp.Cells(i - 1, 5).Value + MyXlsApp.Cells(i, 4).Value
    MyXlsApp.Cells(i, 6).Value = MyXlsApp.Cells(i, 5).Value * txtUnitPrice.Text: intSum6 = intSum6 + MyXlsApp.Cells(i, 5).Value * txtUnitPrice.Text

    disp_rsd.MoveNext
    i = i + 1
    Loop
    MyXlsApp.Cells(6, 5).Value = "�֭p���l"
    MyXlsApp.Cells(i, 1).Value = "�`�p"
    MyXlsApp.Cells(i, 2).Value = intSum2
    MyXlsApp.Cells(i, 3).Value = intSum3
    MyXlsApp.Cells(i, 4).Value = intSum4
    MyXlsApp.Cells(i, 6).Value = intSum6
    '�X���x�s��
    MyXlsApp.Range("A1:F1").Select
    With MyXlsApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    MyXlsApp.Selection.Merge
    MyXlsApp.Selection.Font.Bold = True
    With MyXlsApp.Selection.Font
        .Name = "�s�ө���"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Cells(1, 1).Value = "�ըƹF���y-�Ȥ᯲���p��"
'    '�X���x�s��
'    MyXlsApp.Range("F3:G3").Select
'    MyXlsApp.Selection.Merge
'    MyXlsApp.Range("F4:G4").Select
'    MyXlsApp.Selection.Merge
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A6:F" & i & ",A3:F4").Select
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
End Sub

Private Sub cmd_Query_Click()

    If Len(Trim(Me.txt_DateS)) = 0 Or Len(Trim(Me.txt_DateE)) = 0 Then
        msg_text = "�п�J���w����d��!!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_DateS.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If

    If Len(Trim(Me.lst_Cust.Text)) = 0 Then
        msg_text = "�п�J�Ȥ�W��!!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.lst_Cust.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    '��where
    Dim strWhere As String
    Dim strTmp As String
    If Len(Trim(Me.txt_DateS.Text)) > 0 And Len(Trim(Me.txt_DateE.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),adddate,112)  Between '" & Trim(Me.txt_DateS.Text) & "' and '" & Trim(Me.txt_DateE.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS.Text)) > 0 And Len(Trim(Me.txt_DateE.Text)) = 0 Then
       strWhere = " Convert(Varchar(8),adddate,112) = '" & Trim(Me.txt_DateS.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS.Text)) = 0 And Len(Trim(Me.txt_DateE.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),adddate,112) = '" & Trim(Me.txt_DateE.Text) & "' "
    End If
    
    If Len(Trim(Me.lst_Cust.Text)) > 0 Then
        If Len(Trim(strWhere)) > 0 Then
            strWhere = strWhere & " and Customer like '" & Me.lst_Cust.Text & "%'"
        Else
            strWhere = strWhere & " Customer like '" & Me.lst_Cust.Text & "%'"
        End If
    End If

    If Len(Trim(strWhere)) > 0 Then
        strWhere = "where " & strWhere
    End If
    '�������l
    str_SQL = "select isnull(sum(QtyIn),0) as �ɥX,isnull(sum(QtyOut),0) as �٤J,isnull(sum(QtyIn),0)-isnull(sum(QtyOut),0) as ���l from Pallet_utlCst  " & strWhere & ""
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    Me.txt_In.Text = Trim(tmp_Rs.Fields(0))
    Me.txt_Out.Text = Trim(tmp_Rs.Fields(1))
    Dim int_sum As Integer
    int_sum = Trim(tmp_Rs.Fields(2))
    tmp_Rs.Close
    Screen.MousePointer = vbDefault
    '��������
    str_SQL = "select c.YMD as ��� " & _
            ",sum(isnull(p.QtyIn,0)) as �ɥX " & _
            ",sum(isnull(p.QtyOut,0)) as �٤J " & _
            ", sum(isnull(p.QtyIn,0)) - sum(isnull(p.QtyOut,0)) as ���p�p " & _
            "from Pallet_utlCst p right join Calender c on c.ymd = Convert(char(8),p.adddate,112) and rtrim(p.customer) like '" & lst_Cust.Text & "%' " & _
            "where c.YMD Between '" & txt_DateS.Text & "' and '" & txt_DateE.Text & "' " & _
            "group by  c.YMD order by c.YMD "
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       Set dg_PalletDetail.DataSource = Nothing
       Exit Sub
    End If
    Call ReDim_Recordset(disp_rsd)
    Call Replication_Recordset(tmp_Rs, disp_rsd)
    tmp_Rs.Close
    disp_rsd.MoveFirst
    Set dg_PalletDetail.DataSource = disp_rsd
    With dg_PalletDetail
         .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
         .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
         .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
         .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
    End With
    With dg_PalletDetail
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000       'sku
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 1000
        .Columns(3).Alignment = dbgRight
        .Columns(4).Width = 1000
        .Columns(4).Alignment = dbgRight
'        .Columns(5).Width = 800
'        .Columns(5).Alignment = dbgLeft
'        .Columns(6).Width = 600
'        .Columns(6).Alignment = dbgRight
'        .Columns(7).Width = 600
'        .Columns(7).Alignment = dbgRight
'        .Columns(8).Width = 2000
'        .Columns(8).Alignment = dbgLeft
    End With
    Screen.MousePointer = vbDefault
    '�W�����l
    strWhere = ""
    strTmp = ""
    If Len(Trim(Me.txt_DateS.Text)) > 0 And Len(Trim(Me.txt_DateE.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),adddate,112)< '" & Trim(Me.txt_DateS.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS.Text)) > 0 And Len(Trim(Me.txt_DateE.Text)) = 0 Then
       strWhere = " Convert(Varchar(8),adddate,112)< '" & Trim(Me.txt_DateS.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS.Text)) = 0 And Len(Trim(Me.txt_DateE.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),adddate,112)< '" & Trim(Me.txt_DateE.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS.Text)) = 0 And Len(Trim(Me.txt_DateE.Text)) = 0 Then
       Me.txt_SumA.Text = "0"
       Me.txt_Sum.Text = int_sum
       Exit Sub
    End If
    If Len(Trim(Me.lst_Cust.Text)) > 0 Then
        If Len(Trim(strWhere)) > 0 Then
            strWhere = strWhere & " and Customer like '" & Me.lst_Cust.Text & "%'"
        Else
            strWhere = strWhere & " Customer like '" & Me.lst_Cust.Text & "%'"
        End If
    End If
    
    If Len(Trim(strWhere)) > 0 Then
        strWhere = "where " & strWhere
    End If
    str_SQL = "select sum(QtyIn) as �ɥX,sum(QtyOut) as �٤J,isnull(sum(QtyIn),0)-isnull(sum(QtyOut),0) as ���l from Pallet_utlCst  " & strWhere & ""
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    Me.txt_SumA.Text = Trim(tmp_Rs.Fields(2))
    Me.txt_Sum.Text = Trim(tmp_Rs.Fields(2)) + int_sum
    tmp_Rs.Close
    Screen.MousePointer = vbDefault
    SSTab1.SetFocus
End Sub

Private Sub cmd_Query_Tab1_Click()
    Dim strWhere As String
    Dim strTmp As String
    Screen.MousePointer = 11
    
        If Len(Trim(Me.txt_DateS_Tab1)) = 0 Or Len(Trim(Me.txt_DateE_Tab1)) = 0 Then
        msg_text = "�п�J���w����d��!!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_DateS.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Len(Trim(Me.txtCarNo.Text)) = 0 Then
        msg_text = "�п�J����!!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.lst_Cust.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    '��where
    If Len(Trim(Me.txt_DateS_Tab1.Text)) > 0 And Len(Trim(Me.txt_DateE_Tab1.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),checkDate,112)  Between '" & Trim(Me.txt_DateS_Tab1.Text) & "' and '" & Trim(Me.txt_DateE_Tab1.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS_Tab1.Text)) > 0 And Len(Trim(Me.txt_DateE_Tab1.Text)) = 0 Then
       strWhere = " Convert(Varchar(8),checkDate,112) = '" & Trim(Me.txt_DateS_Tab1.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS_Tab1.Text)) = 0 And Len(Trim(Me.txt_DateE_Tab1.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),checkDate,112) = '" & Trim(Me.txt_DateE_Tab1.Text) & "' "
    End If
    
    If Len(Trim(Me.txtCarNo.Text)) > 0 Then
        If Len(Trim(strWhere)) > 0 Then
            strWhere = strWhere & " and CarNo = '" & Me.txtCarNo.Text & "'"
        Else
            strWhere = strWhere & " CarNo = '" & Me.txtCarNo.Text & "'"
        End If
    End If

    If Len(Trim(strWhere)) > 0 Then
        strWhere = "where " & strWhere
    End If
    
    '�������l
    str_SQL = "select isnull(sum(QtyIn),0) as �ɥX,isnull(sum(QtyOut),0) as �٤J,isnull(sum(QtyIn),0)-isnull(sum(QtyOut),0) as ���l from Pallet_cst  " & strWhere & ""
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    Me.txt_In_Tab1.Text = Trim(tmp_Rs.Fields(0))
    Me.txt_Out_Tab1.Text = Trim(tmp_Rs.Fields(1))
    Dim int_sum As Integer
    int_sum = Trim(tmp_Rs.Fields(2))
    tmp_Rs.Close
    Screen.MousePointer = vbDefault
    '��������
'    str_SQL = "select Convert(Varchar,CheckDate,112) as ���,CarNo as ����,CheckNo as �渹,isnull(QtyIn,0) as �ɥX,isnull(QtyOut,0) as �٤J from dbo.Pallet_cst  " & strwhere & " order by Convert(Varchar,CheckDate,111),CheckNo"
    
    str_SQL = "select c.YMD as ��� " & _
        ",sum(isnull(p.QtyIn,0)) as �ɥX " & _
        ",sum(isnull(p.QtyOut,0)) as �٤J " & _
        ",sum(isnull(p.QtyIn,0)) - sum(isnull(p.QtyOut,0)) as ���p�p " & _
        "from Pallet_cst p right join Calender c on c.ymd = Convert(char(8),p.checkDate,112) and rtrim(p.CarNo) = '" & txtCarNo.Text & "' " & _
        "where c.YMD Between '" & txt_DateS_Tab1.Text & "' and '" & txt_DateE_Tab1.Text & "' " & _
        "group by  c.YMD order by c.YMD "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       Set dg_PalletDetail.DataSource = Nothing
       MsgBox "�d�L��ơI", 64, Me.Caption
       Exit Sub
    End If
    Call ReDim_Recordset(disp_rsd_Tab1)
    Call Replication_Recordset(tmp_Rs, disp_rsd_Tab1)
    tmp_Rs.Close
    disp_rsd_Tab1.MoveFirst
    Set dg_PalletCDS.DataSource = disp_rsd_Tab1
    With dg_PalletCDS
         .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
         .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
         .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
         .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
    End With
    With dg_PalletCDS
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000       'sku
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1000
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 1000
        .Columns(3).Alignment = dbgRight
        .Columns(4).Width = 1000
        .Columns(4).Alignment = dbgRight
'        .Columns(5).Width = 800
'        .Columns(5).Alignment = dbgLeft
        
    End With
    Screen.MousePointer = vbDefault
    '�W�����l
    strWhere = ""
    strTmp = ""
    If Len(Trim(Me.txt_DateS_Tab1.Text)) > 0 And Len(Trim(Me.txt_DateE_Tab1.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),checkDate,112) < '" & Trim(Me.txt_DateS_Tab1.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS_Tab1.Text)) > 0 And Len(Trim(Me.txt_DateE_Tab1.Text)) = 0 Then
       strWhere = " Convert(Varchar(8),checkDate,112)< '" & Trim(Me.txt_DateS_Tab1.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS_Tab1.Text)) = 0 And Len(Trim(Me.txt_DateE_Tab1.Text)) > 0 Then
       strWhere = " Convert(Varchar(8),checkDate,112)< '" & Trim(Me.txt_DateE_Tab1.Text) & "' "
    ElseIf Len(Trim(Me.txt_DateS_Tab1.Text)) = 0 And Len(Trim(Me.txt_DateE_Tab1.Text)) = 0 Then
       Me.txt_SumA_Tab1.Text = "0"
       Me.txt_Sum_Tab1.Text = int_sum
       Exit Sub
    End If
    
    If Len(Trim(strWhere)) > 0 Then
        strWhere = "where CarNo = '" & Me.txtCarNo.Text & "' and " & strWhere
    End If
    str_SQL = "select sum(QtyIn) as �ɥX,sum(QtyOut) as �٤J,isnull(sum(QtyIn),0)-isnull(sum(QtyOut),0) as ���l from Pallet_cst  " & strWhere & ""
    Screen.MousePointer = vbHourglass
    Call Confirm_Recordset_Closed(tmp_Rs)
    cn.CommandTimeout = 120
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    Me.txt_SumA_Tab1.Text = Trim(tmp_Rs.Fields(2))
    Me.txt_Sum_Tab1.Text = Trim(tmp_Rs.Fields(2)) + int_sum
    tmp_Rs.Close
    Screen.MousePointer = vbDefault
    SSTab1.SetFocus
End Sub

Private Sub Command1_Click()
    If IsNumeric(txtUnitPrice1.Text) = False Then MsgBox "����������~!!", vbOKOnly, Me.Caption: txtUnitPrice1.SetFocus: Exit Sub
    If disp_rsd_Tab1 Is Nothing Then Exit Sub
    If disp_rsd_Tab1.RecordCount = 0 Then Exit Sub
    '�N��Ƽg�Jexcel��
    Dim MyXlsApp As Excel.Application   '�}��excel��
    Dim objFld As Field
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '�s�WWookbooks
    MyXlsApp.Workbooks.Add
    '�s�WSheets
    'MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "�̪O�����p��"
    MyXlsApp.ActiveSheet.Name = "�̪O�����p��"
    i = 3
    MyXlsApp.Cells(i, 1).Value = "��b��_"
    MyXlsApp.Cells(i, 2).Value = Trim(Me.txt_DateS_Tab1.Text)
    MyXlsApp.Cells(i, 3).Value = "�W�����l"
    MyXlsApp.Cells(i, 4).Value = Trim(Me.txt_SumA_Tab1.Text)
    MyXlsApp.Cells(i, 5).Value = "�Ȥ�W��"
    MyXlsApp.Cells(i, 6).Value = txtCarNo.Text
    i = i + 1
    MyXlsApp.Cells(i, 1).Value = "��b�騴"
    MyXlsApp.Cells(i, 2).Value = Trim(Me.txt_DateE_Tab1.Text)
    MyXlsApp.Cells(i, 3).Value = "�������l"
    MyXlsApp.Cells(i, 4).Value = Trim(Me.txt_Sum_Tab1.Text)
    MyXlsApp.Cells(i, 5).Value = "��b���"
    MyXlsApp.Cells(i, 6).Value = Format(Now, "yyyymmdd")
    i = i + 2
    ' Convert(VarChar, checkdate, 112) As ���,CarNo As ����, CheckNo As �渹, isnull(QtyIn, 0) As �ɥX, isnull(QtyOut, 0) As �٤J
    MyXlsApp.Cells(i, 1).Value = "���"
    MyXlsApp.Cells(i, 2).Value = "�ɥX"
    MyXlsApp.Cells(i, 3).Value = "�٤J"
    MyXlsApp.Cells(i, 4).Value = "��鵲�l"
    MyXlsApp.Cells(i, 5).Value = Trim(Me.txt_SumA_Tab1.Text)
    MyXlsApp.Cells(i, 6).Value = "����"
   
    i = i + 1
    disp_rsd_Tab1.MoveFirst
    
    '���,�ɥX,�٤J,��鵲�l,�֭p���l,����
    Dim intSum2 As Integer, intSum3 As Integer, intSum4 As Integer, intSum6 As Double
    Do While Not disp_rsd_Tab1.EOF
    MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
    MyXlsApp.Cells(i, 1).Value = Trim(disp_rsd_Tab1.Fields(1))
    MyXlsApp.Cells(i, 2).Value = disp_rsd_Tab1.Fields(2): intSum2 = intSum2 + disp_rsd_Tab1.Fields(2)
    MyXlsApp.Cells(i, 3).Value = disp_rsd_Tab1.Fields(3): intSum3 = intSum3 + disp_rsd_Tab1.Fields(3)
    MyXlsApp.Cells(i, 4).Value = disp_rsd_Tab1.Fields(4): intSum4 = intSum4 + disp_rsd_Tab1.Fields(4)
    MyXlsApp.Cells(i, 5).Value = MyXlsApp.Cells(i - 1, 5).Value + MyXlsApp.Cells(i, 4).Value
    MyXlsApp.Cells(i, 6).Value = MyXlsApp.Cells(i, 5).Value * txtUnitPrice1.Text: intSum6 = intSum6 + MyXlsApp.Cells(i, 5).Value * txtUnitPrice1.Text

    disp_rsd_Tab1.MoveNext
    i = i + 1
    Loop
    MyXlsApp.Cells(6, 5).Value = "�֭p���l"
    MyXlsApp.Cells(i, 1).Value = "�`�p"
    MyXlsApp.Cells(i, 2).Value = intSum2
    MyXlsApp.Cells(i, 3).Value = intSum3
    MyXlsApp.Cells(i, 4).Value = intSum4
    MyXlsApp.Cells(i, 6).Value = intSum6
    '�X���x�s��
    MyXlsApp.Range("A1:F1").Select
    With MyXlsApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    MyXlsApp.Selection.Merge
    MyXlsApp.Selection.Font.Bold = True
    With Selection.Font
        .Name = "�s�ө���"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Cells(1, 1).Value = "�ըƹF���y-�q�������p��"
'    '�X���x�s��
'    MyXlsApp.Range("F3:G3").Select
'    MyXlsApp.Selection.Merge
'    MyXlsApp.Range("F4:G4").Select
'    MyXlsApp.Selection.Merge
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A6:F" & i & ",A3:F4").Select
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
    
    
    '���,����,�渹,�Z�O,�ɥX,�٤J
'    Do While Not disp_rsd_Tab1.EOF
'        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
'        MyXlsApp.Cells(i, 1).Value = Trim(disp_rsd_Tab1.Fields(1))
'        MyXlsApp.Cells(i, 2).NumberFormatLocal = "@"
'        MyXlsApp.Cells(i, 2).Value = Trim(disp_rsd_Tab1.Fields(2))
'        MyXlsApp.Cells(i, 3).NumberFormatLocal = "@"
'        MyXlsApp.Cells(i, 3).Value = disp_rsd_Tab1.Fields(3)
'        MyXlsApp.Cells(i, 4).Value = disp_rsd_Tab1.Fields(4)
'        MyXlsApp.Cells(i, 5).Value = disp_rsd_Tab1.Fields(5)
'        disp_rsd_Tab1.MoveNext
'        i = i + 1
'    Loop
'    '�X���x�s��
'    MyXlsApp.Range("A1:G1").Select
'    With MyXlsApp.Selection
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .ShrinkToFit = False
'        .MergeCells = False
'    End With
'    MyXlsApp.Selection.Merge
'    MyXlsApp.Selection.Font.Bold = True
'    With Selection.Font
'        .Name = "�s�ө���"
'        .Size = 18
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = xlAutomatic
'    End With
'    MyXlsApp.Cells(1, 1).Value = "�ըƹF���y-IDS�̪O��b��"
'    '�X���x�s��
'    MyXlsApp.Range("F3:G3").Select
'    MyXlsApp.Selection.Merge
'    MyXlsApp.Range("F4:G4").Select
'    MyXlsApp.Selection.Merge
'
'    '�����ϥ�
'    MyXlsApp.Cells.Select
'    With MyXlsApp.Selection.Interior
'        .ColorIndex = 2
'        .Pattern = xlSolid
'    End With
'    '�e�uh
'    MyXlsApp.Range("A6:E" & i - 1 & ",A3:F4").Select
'    MyXlsApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    MyXlsApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With MyXlsApp.Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With MyXlsApp.Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With MyXlsApp.Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With MyXlsApp.Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With MyXlsApp.Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With MyXlsApp.Selection.Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    MyXlsApp.Visible = True
End Sub

Private Sub Form_Activate()
    '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
    Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
    msg_title = "�����p��"
End Sub

Private Sub Form_Load()
    Me.Height = 7600: Me.Width = 11500
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select code from CodeLkup where listname='Cust_CDS'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
       Do While Not tmp_Rs.EOF
          lst_Cust.AddItem Trim(tmp_Rs.Fields("code"))
          tmp_Rs.MoveNext
       Loop
    End If
    tmp_Rs.Close
End Sub


