VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_BaseData_OPCode 
   Caption         =   "   �@   �~   �N  �X   ��   ��   ��   �@"
   ClientHeight    =   6405
   ClientLeft      =   810
   ClientTop       =   1530
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   10665
   Begin TabDlg.SSTab SSTab1 
      Height          =   6360
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   11218
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "�򥻥N�X1"
      TabPicture(0)   =   "frm_BaseData_OPCode.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd_Exit(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "��������"
      TabPicture(1)   =   "frm_BaseData_OPCode.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(1)"
      Tab(1).Control(1)=   "dg_Tab1CarType"
      Tab(1).Control(2)=   "cmd_Tab1CarType_Show"
      Tab(1).Control(3)=   "cmd_Exit(2)"
      Tab(1).Control(4)=   "cmd_Tab1CarType_Save"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "�S��ݨD"
      TabPicture(2)   =   "frm_BaseData_OPCode.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Label1(0)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "�B�e�ϰ�"
      TabPicture(3)   =   "frm_BaseData_OPCode.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(2)"
      Tab(3).Control(1)=   "dg_Tab3Area"
      Tab(3).Control(2)=   "cmd_Tab3Area_Show"
      Tab(3).Control(3)=   "cmd_Exit(3)"
      Tab(3).Control(4)=   "cmd_Tab3Area_Save"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "�l���ϸ�"
      TabPicture(4)   =   "frm_BaseData_OPCode.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmd_Tab4Zip_Show"
      Tab(4).Control(1)=   "cmd_Exit(4)"
      Tab(4).Control(2)=   "cmd_Tab4Zip_Save"
      Tab(4).Control(3)=   "dg_Tab4Zip"
      Tab(4).Control(4)=   "Label1(3)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "�x�}�ϽX"
      TabPicture(5)   =   "frm_BaseData_OPCode.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(4)"
      Tab(5).Control(1)=   "dg_Tab5GridCode"
      Tab(5).Control(2)=   "cmd_Tab5GridCode_Show"
      Tab(5).Control(3)=   "cmd_Exit(5)"
      Tab(5).Control(4)=   "cmd_Tab5GridCodeSave"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "�򥻥N�X2"
      TabPicture(6)   =   "frm_BaseData_OPCode.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1(4)"
      Tab(6).Control(1)=   "Frame1(5)"
      Tab(6).Control(2)=   "cmd_Exit(6)"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   " "
      TabPicture(7)   =   "frm_BaseData_OPCode.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(5)"
      Tab(7).Control(1)=   "dg_Tab7_TRP17M"
      Tab(7).Control(2)=   "cmd_Tab7_DisPlay"
      Tab(7).Control(3)=   "cmd_Tab7_Delete"
      Tab(7).Control(4)=   "cmd_Tab7_Save"
      Tab(7).Control(5)=   "cmd_Exit(7)"
      Tab(7).ControlCount=   6
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "���}"
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
         Index           =   7
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":00E0
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   62
         Top             =   4800
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab7_Save 
         BackColor       =   &H00FFC0C0&
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
         Height          =   945
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":0522
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   61
         Top             =   2640
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab7_Delete 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�R��"
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
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":082C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   60
         Top             =   3720
         Width           =   1035
      End
      Begin VB.Frame Frame3 
         Caption         =   "�S��ݨD�Ӷ�"
         Height          =   5535
         Left            =   -69600
         TabIndex        =   56
         Top             =   720
         Width           =   5055
         Begin VB.CommandButton cmd_Tab2_Delete 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�R  ��"
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
            Left            =   1440
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   63
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab2SpecDemandDetail_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   2895
            Picture         =   "frm_BaseData_OPCode.frx":0B36
            TabIndex        =   58
            Top             =   240
            Width           =   1830
         End
         Begin VB.CommandButton cmd_Tab2SpecDemandDetail_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   255
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   57
            Top             =   240
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab2SpecDemandDetail 
            Height          =   4650
            Left            =   120
            TabIndex        =   59
            Top             =   780
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   8202
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.CommandButton cmd_Tab7_DisPlay 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��ܩҦ����"
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
         Left            =   -68310
         Picture         =   "frm_BaseData_OPCode.frx":0E40
         TabIndex        =   49
         Top             =   480
         Width           =   1830
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
         Height          =   480
         Index           =   6
         Left            =   -65625
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   48
         Top             =   435
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         Caption         =   "���`��]"
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
         Height          =   2565
         Index           =   5
         Left            =   -74805
         TabIndex        =   44
         Top             =   930
         Width           =   5100
         Begin VB.CommandButton cmd_Tab6RSC_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":114A
            TabIndex        =   46
            Top             =   285
            Width           =   1500
         End
         Begin VB.CommandButton cmd_Tab6RSC_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":1454
            TabIndex        =   45
            Top             =   285
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab6RSC 
            Height          =   1785
            Left            =   90
            TabIndex        =   47
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "���`�d��"
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
         Height          =   2565
         Index           =   4
         Left            =   -69660
         TabIndex        =   40
         Top             =   930
         Width           =   5100
         Begin VB.CommandButton cmd_Tab6RBC_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":175E
            TabIndex        =   42
            Top             =   285
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab6RBC_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":1A68
            TabIndex        =   41
            Top             =   285
            Width           =   1500
         End
         Begin MSDataGridLib.DataGrid dg_Tab6RBC 
            Height          =   1785
            Left            =   90
            TabIndex        =   43
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.CommandButton cmd_Tab5GridCodeSave 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�s  ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -65985
         Picture         =   "frm_BaseData_OPCode.frx":1D72
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   39
         Top             =   3840
         Width           =   1125
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
         Height          =   1050
         Index           =   5
         Left            =   -66000
         Picture         =   "frm_BaseData_OPCode.frx":207C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   38
         Top             =   5055
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Tab5GridCode_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��ܩҦ����"
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
         Left            =   -68355
         Picture         =   "frm_BaseData_OPCode.frx":24BE
         TabIndex        =   35
         Top             =   600
         Width           =   1830
      End
      Begin VB.CommandButton cmd_Tab4Zip_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��ܩҦ����"
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
         Left            =   -68040
         Picture         =   "frm_BaseData_OPCode.frx":27C8
         TabIndex        =   32
         Top             =   510
         Width           =   1830
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
         Height          =   1050
         Index           =   4
         Left            =   -65865
         Picture         =   "frm_BaseData_OPCode.frx":2AD2
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   31
         Top             =   4710
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Tab4Zip_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�s  ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -65865
         Picture         =   "frm_BaseData_OPCode.frx":2F14
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   30
         Top             =   3495
         Width           =   1125
      End
      Begin VB.CommandButton cmd_Tab3Area_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�s  ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -65880
         Picture         =   "frm_BaseData_OPCode.frx":321E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   27
         Top             =   3525
         Width           =   1125
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
         Height          =   1050
         Index           =   3
         Left            =   -65895
         Picture         =   "frm_BaseData_OPCode.frx":3528
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   26
         Top             =   4740
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Tab3Area_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��ܩҦ����"
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
         Left            =   -68070
         Picture         =   "frm_BaseData_OPCode.frx":396A
         TabIndex        =   25
         Top             =   540
         Width           =   1830
      End
      Begin VB.CommandButton cmd_Tab1CarType_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�s  ��"
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
         Left            =   -66105
         Picture         =   "frm_BaseData_OPCode.frx":3C74
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   21
         Top             =   3630
         Width           =   1050
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
         Index           =   2
         Left            =   -66120
         Picture         =   "frm_BaseData_OPCode.frx":3F7E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   20
         Top             =   4830
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab1CarType_Show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��ܩҦ����"
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
         Left            =   -68430
         Picture         =   "frm_BaseData_OPCode.frx":43C0
         TabIndex        =   19
         Top             =   540
         Width           =   1830
      End
      Begin VB.Frame Frame1 
         Caption         =   "�h�B�u��"
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
         Height          =   2565
         Index           =   3
         Left            =   5325
         TabIndex        =   15
         Top             =   3615
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Move_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":46CA
            TabIndex        =   17
            Top             =   285
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab0Move_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":49D4
            TabIndex        =   16
            Top             =   285
            Width           =   1500
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Move 
            Height          =   1770
            Left            =   90
            TabIndex        =   18
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3122
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "���Τ覡"
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
         Height          =   2565
         Index           =   2
         Left            =   5340
         TabIndex        =   11
         Top             =   855
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Employ_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":4CDE
            TabIndex        =   13
            Top             =   285
            Width           =   1500
         End
         Begin VB.CommandButton cmd_Tab0Employ_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":4FE8
            TabIndex        =   12
            Top             =   285
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Employ 
            Height          =   1785
            Left            =   90
            TabIndex        =   14
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "�˨��覡"
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
         Height          =   2565
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   3615
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Load_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":52F2
            TabIndex        =   9
            Top             =   285
            Width           =   1500
         End
         Begin VB.CommandButton cmd_Tab0Load_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":55FC
            TabIndex        =   8
            Top             =   285
            Width           =   1050
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Load 
            Height          =   1785
            Left            =   90
            TabIndex        =   10
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "���[�Φ�"
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
         Height          =   2565
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   855
         Width           =   5100
         Begin VB.CommandButton cmd_Tab0Box_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   3960
            Picture         =   "frm_BaseData_OPCode.frx":5906
            TabIndex        =   5
            Top             =   285
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab0Box_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   105
            Picture         =   "frm_BaseData_OPCode.frx":5C10
            TabIndex        =   3
            Top             =   285
            Width           =   1500
         End
         Begin MSDataGridLib.DataGrid dg_Tab0Box 
            Height          =   1785
            Left            =   90
            TabIndex        =   4
            Top             =   675
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
         Height          =   480
         Index           =   0
         Left            =   9375
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   1
         Top             =   360
         Width           =   1050
      End
      Begin MSDataGridLib.DataGrid dg_Tab1CarType 
         Height          =   5130
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab3Area 
         Height          =   5130
         Left            =   -74745
         TabIndex        =   28
         Top             =   960
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab4Zip 
         Height          =   5130
         Left            =   -74715
         TabIndex        =   33
         Top             =   930
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab5GridCode 
         Height          =   5130
         Left            =   -74790
         TabIndex        =   36
         Top             =   1020
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_Tab7_TRP17M 
         Height          =   5130
         Left            =   -74640
         TabIndex        =   50
         Top             =   900
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
      Begin VB.Frame Frame2 
         Caption         =   "�Ȥ�S��ݨD"
         Height          =   5535
         Left            =   -74880
         TabIndex        =   52
         Top             =   720
         Width           =   5055
         Begin VB.CommandButton cmd_Tab2SpecDemand_Save 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�s  ��"
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
            Left            =   255
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   54
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmd_Tab2SpecDemand_Show 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��ܩҦ����"
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
            Left            =   2895
            Picture         =   "frm_BaseData_OPCode.frx":5F1A
            TabIndex        =   53
            Top             =   240
            Width           =   1830
         End
         Begin MSDataGridLib.DataGrid dg_Tab2SpecDemand 
            Height          =   4650
            Left            =   120
            TabIndex        =   55
            Top             =   780
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   8202
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483624
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~�G�Ȥ��\�s�W�B�ק�A�����\�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   0
         Left            =   -74595
         TabIndex        =   24
         Top             =   375
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~�G�Ȥ��\�s�W�B�ק�A�����\�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   5
         Left            =   -74580
         TabIndex        =   51
         Top             =   495
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~�G�Ȥ��\�s�W�B�ק�A�����\�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   4
         Left            =   -74745
         TabIndex        =   37
         Top             =   570
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~�G�Ȥ��\�s�W�B�ק�A�����\�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   3
         Left            =   -74670
         TabIndex        =   34
         Top             =   480
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~�G�Ȥ��\�s�W�B�ק�A�����\�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   2
         Left            =   -74700
         TabIndex        =   29
         Top             =   510
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~�G�Ȥ��\�s�W�B�ק�A�����\�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   -74700
         TabIndex        =   23
         Top             =   555
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~�G�Ȥ��\�s�W�B�ק�A�����\�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   13
         Left            =   1815
         TabIndex        =   6
         Top             =   465
         Width           =   4845
      End
   End
End
Attribute VB_Name = "frm_BaseData_OPCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private intloop As Integer
'�򥻥N�X���@ 1
Private rs_Tab0CarBox As ADODB.Recordset       '���[�Φ�
Private rs_Tab0Load As ADODB.Recordset         '�˨��覡
Private rs_Tab0Employ As ADODB.Recordset       '���Τ覡
Private rs_Tab0Move As ADODB.Recordset         '�h�B�u��

Private rs_Tab1CarType As ADODB.Recordset      '���إN�X
Private rs_Tab2SpecDemand As ADODB.Recordset   '�S��ݨD
Private rs_Tab3Area As ADODB.Recordset         '�B�e�ϰ�
Private rs_Tab4Zip As ADODB.Recordset          '�l���ϸ�
Private rs_Tab5GridCode As ADODB.Recordset     '�x�}�ϽX
Private rs_Tab7_TRP17M  As ADODB.Recordset     '�p�O�N�X
Private rs_Tab2SpecDemandDetail As ADODB.Recordset   '�S��ݨD�Ӷ�
'�򥻥N�X���@ 2
Private rs_Tab6RSC As ADODB.Recordset          '���`��]
Private rs_Tab6RBC As ADODB.Recordset          '���`�d��


Private Sub cmd_Tab0Box_Save_Click()
'�򥻥N�X1 >> ���[�Φ� >> �s��
If rs_Tab0CarBox Is Nothing Then
   msg_text = "�Х��d�ߩҦ� ���[�Φ� ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0CarBox.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� ���[�Φ� ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0CarBox.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0CarBox.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0CarBox.Fields("���[�Φ�").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'CARBOXTYPE' And Code = '" & rs_Tab0CarBox.Fields("�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'CARBOXTYPE','" & rs_Tab0CarBox.Fields("�N�X").Value & "','" & rs_Tab0CarBox.Fields("���[�Φ�").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0CarBox.MoveNext
Loop
rs_Tab0CarBox.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-���[�Φ�-�s��", Me.Caption, "cmd_Tab0Box_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Box_Show_Click()
'�򥻥N�X1 >> ���[�Φ� ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Box.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0CarBox)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS ���[�Φ� " & _
          "From CodeLKUP Where ListName = 'CARBOXTYPE'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0CarBox)
tmp_Rs.Close

With dg_Tab0Box
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab0CarBox.MoveFirst
Set dg_Tab0Box.DataSource = rs_Tab0CarBox
With dg_Tab0Box
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '���[�Φ�
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���[�Φ�-��ܩҦ����", Me.Caption, "cmd_Tab0Box_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Employ_Save_Click()
'�򥻥N�X1 >> ���Τ覡 >> �s��
If rs_Tab0Employ Is Nothing Then
   msg_text = "�Х��d�ߩҦ� ���Τ覡 ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0Employ.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� ���Τ覡 ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0Employ.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0Employ.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0Employ.Fields("���Τ覡").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'EMPLOYTYPE' And Code = '" & rs_Tab0Employ.Fields("�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'EMPLOYTYPE','" & rs_Tab0Employ.Fields("�N�X").Value & "','" & rs_Tab0Employ.Fields("���Τ覡").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0Employ.MoveNext
Loop
rs_Tab0Employ.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-���Τ覡-�s��", Me.Caption, "cmd_Tab0Employ_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Employ_Show_Click()
'�򥻥N�X1 >> ���Τ覡 ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Employ.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0Employ)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS ���Τ覡 " & _
          "From CodeLKUP Where ListName = 'EMPLOYTYPE'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0Employ)
tmp_Rs.Close

With dg_Tab0Employ
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab0Employ.MoveFirst
Set dg_Tab0Employ.DataSource = rs_Tab0Employ
With dg_Tab0Employ
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '���Τ覡
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���Τ覡-��ܩҦ����", Me.Caption, "cmd_Tab0Box_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Load_Save_Click()
'�򥻥N�X1 >> �˨��覡 >> �s��
If rs_Tab0Load Is Nothing Then
   msg_text = "�Х��d�ߩҦ� �˨��覡 ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0Load.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� �˨��覡 ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0Load.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0Load.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0Load.Fields("�˨��覡").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'LOADUNLOADTYPE' And Code = '" & rs_Tab0Load.Fields("�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'LOADUNLOADTYPE','" & rs_Tab0Load.Fields("�N�X").Value & "','" & rs_Tab0Load.Fields("�˨��覡").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0Load.MoveNext
Loop
rs_Tab0Load.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-�˨��覡-�s��", Me.Caption, "cmd_Tab0Load_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Load_Show_Click()
'�򥻥N�X1 >> �˨��覡 ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Load.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0Load)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS �˨��覡 " & _
          "From CodeLKUP Where ListName = 'LOADUNLOADTYPE'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0Load)
tmp_Rs.Close

With dg_Tab0Load
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab0Load.MoveFirst
Set dg_Tab0Load.DataSource = rs_Tab0Load
With dg_Tab0Load
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '�˨��覡
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�˨��覡-��ܩҦ����", Me.Caption, "cmd_Tab0Load_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Move_Save_Click()
'�򥻥N�X1 >> �h�B�u�� >> �s��
If rs_Tab0Move Is Nothing Then
   msg_text = "�Х��d�ߩҦ� �h�B�u�� ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab0Move.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� �h�B�u�� ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab0Move.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab0Move.EOF
   str_SQL = "Update CodeLKUP " & _
             "Set Description = '" & rs_Tab0Move.Fields("�h�B�u��").Value & "',EditWho='" & User_id & "',EditDate=Getdate() " & _
             "Where ListName = 'MOVETOOL' And Code = '" & rs_Tab0Move.Fields("�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP(ListName,Code,Description,AddWho,EditWho,EditDate) Values (" & _
                "'MOVETOOL','" & rs_Tab0Move.Fields("�N�X").Value & "','" & rs_Tab0Move.Fields("�h�B�u��").Value & "','" & _
                User_id & "','" & User_id & "',Getdate())"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab0Move.MoveNext
Loop
rs_Tab0Move.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-�˨��覡-�s��", Me.Caption, "cmd_Tab0Move_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0Move_Show_Click()
'�򥻥N�X1 >> �h�B�覡 ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0Move.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0Move)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS �h�B�u�� " & _
          "From CodeLKUP Where ListName = 'MOVETOOL'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0Move)
tmp_Rs.Close

With dg_Tab0Move
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab0Move.MoveFirst
Set dg_Tab0Move.DataSource = rs_Tab0Move
With dg_Tab0Move
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '�h�B�u��
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�h�B�u��-��ܩҦ����", Me.Caption, "cmd_Tab0Move_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1CarType_Save_Click()
'�������� >> �s��
If rs_Tab1CarType Is Nothing Then
   msg_text = "�Х��d�ߩҦ�����������ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab1CarType.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ�����������ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab1CarType.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab1CarType.EOF
   str_SQL = "Update TRP15M " & _
             "Set Description = '" & rs_Tab1CarType.Fields("��������").Value & "' ,Car_Type = '" & rs_Tab1CarType("�p�O���O") & "' " & _
             "Where Vehicle_Type = '" & rs_Tab1CarType.Fields("�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP15M(Vehicle_Type,Description,Car_Type) Values (" & _
                "'" & rs_Tab1CarType.Fields("�N�X").Value & "','" & rs_Tab1CarType("��������") & "','" & rs_Tab1CarType("�p�O���O") & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab1CarType.MoveNext
Loop
rs_Tab1CarType.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-��������-�s��", Me.Caption, "cmd_Tab1CarType_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1CarType_Show_Click()
'���إN�X >> ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab1CarType.DataSource = Nothing
Call ReDim_Recordset(rs_Tab1CarType)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Vehicle_Type) AS �N�X, RTRIM(Description) AS �������� ,�p�O���O = car_type " & _
          "From TRP15M Order by Vehicle_Type"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab1CarType)
tmp_Rs.Close

With dg_Tab1CarType
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_Tab1CarType.MoveFirst
Set dg_Tab1CarType.DataSource = rs_Tab1CarType
With dg_Tab1CarType
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 500       '�N�X
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 5000      '��������
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000      '�p�O���O
    .Columns(3).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-��������-��ܩҦ����", Me.Caption, "cmd_Tab1CarType_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Delete_Click()

If rs_Tab2SpecDemandDetail Is Nothing Then Exit Sub

Tran_Level = cn.BeginTrans

If MsgBox("�T�{�R���f�D�G" & rs_Tab2SpecDemandDetail.Fields("�f�D") & "�F�Ȥ�s�աG" & rs_Tab2SpecDemandDetail("�Ȥ�s��") & "�F�~���G" & rs_Tab2SpecDemandDetail("�~��") & " !!", vbOKCancel, "�R��") <> vbOK Then Exit Sub

cn.Execute "delete trp18m where storerkey = '" & rs_Tab2SpecDemandDetail.Fields("�f�D") & "' and consigneekey = '" & rs_Tab2SpecDemandDetail.Fields("�Ȥ�s��") & "' and code = '" & rs_Tab2SpecDemandDetail.Fields("�~��") & "' ", RowsAffect, adExecuteNoRecords

If RowsAffect = 1 Then
    cn.CommitTrans: Tran_Level = 0
    rs_Tab2SpecDemandDetail.Delete
Else
    cn.RollbackTrans
End If

End Sub

Private Sub cmd_Tab2SpecDemand_Save_Click()
'�������� >> �s��
If rs_Tab2SpecDemand Is Nothing Then
   msg_text = "�Х��d�ߩҦ��S��ݨD��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab2SpecDemand.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ��S��ݨD��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab2SpecDemand.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab2SpecDemand.EOF
   str_SQL = "Update TRP04M " & _
             "Set Description = '" & rs_Tab2SpecDemand.Fields("�S��ݨD").Value & "' " & _
             "Where Extra_Demand_Code = '" & rs_Tab2SpecDemand.Fields("�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP04M(Extra_Demand_Code,Description) Values (" & _
                "'" & rs_Tab2SpecDemand.Fields("�N�X").Value & "','" & rs_Tab2SpecDemand.Fields("�S��ݨD").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab2SpecDemand.MoveNext
Loop
rs_Tab2SpecDemand.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-�S��ݨD-�s��", Me.Caption, "cmd_Tab2SpecDemand_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2SpecDemand_Show_Click()
'�S��ݨD >> ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab2SpecDemand.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2SpecDemand)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Extra_Demand_Code) AS �N�X, RTRIM(Description) AS �S��ݨD " & _
          "From TRP04M Order by Extra_Demand_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab2SpecDemand)
tmp_Rs.Close

With dg_Tab2SpecDemand
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab2SpecDemand.MoveFirst
Set dg_Tab2SpecDemand.DataSource = rs_Tab2SpecDemand
With dg_Tab2SpecDemand
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 6000       '�S��ݨD����
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�S��ݨD-��ܩҦ����", Me.Caption, "cmd_Tab2SpecDemand_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2SpecDemandDetail_Save_Click()
'�S��ݨD-->�S��ݨD�Ӷ�>> �s��
If rs_Tab2SpecDemandDetail Is Nothing Then
   msg_text = "�Х��d�ߩҦ��S��ݨD��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab2SpecDemandDetail.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ��S��ݨD��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab2SpecDemandDetail.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab2SpecDemandDetail.EOF
   str_SQL = "Update TRP18M " & _
             "Set Description = '" & rs_Tab2SpecDemandDetail.Fields("�ݨD����").Value & "'," & _
             " Consigneekey = '" & rs_Tab2SpecDemandDetail.Fields("�Ȥ�s��").Value & "'," & _
             " Code = '" & rs_Tab2SpecDemandDetail.Fields("�~��").Value & "'," & _
             " Storerkey = '" & rs_Tab2SpecDemandDetail.Fields("�f�D").Value & "'" & _
             " Where Code = '" & rs_Tab2SpecDemandDetail("�~��") & "' and Storerkey = '" & rs_Tab2SpecDemandDetail("�f�D") & "' and consigneekey = '" & rs_Tab2SpecDemandDetail("�Ȥ�s��") & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP18M(Storerkey,Consigneekey,Code,Description) Values (" & _
                "'" & UCase(rs_Tab2SpecDemandDetail("�f�D")) & "','" & UCase(rs_Tab2SpecDemandDetail.Fields("�Ȥ�s��").Value) & "'," & _
                "'" & UCase(rs_Tab2SpecDemandDetail("�~��")) & "','" & rs_Tab2SpecDemandDetail.Fields("�ݨD����").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab2SpecDemandDetail.MoveNext
Loop

rs_Tab2SpecDemandDetail.MoveFirst
cn.CommitTrans: Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-�S��ݨD�Ӷ�-�s��", Me.Caption, "cmd_Tab2SpecDemandDetail_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2SpecDemandDetail_Show_Click()
'�S��ݨD�Ӷ� >> ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Set dg_Tab2SpecDemandDetail.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2SpecDemandDetail)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT " & _
        "RTRIM(Storerkey) AS �f�D, " & _
        "RTRIM(consigneekey) AS �Ȥ�s�� , " & _
        "RTRIM(code) as �~�� , " & _
        " isnull(RTRIM(Description),'') AS �ݨD����  " & _
        "From TRP18M "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If tmp_rs.EOF Then
'   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
'   Exit Sub
'End If

Call Replication_Recordset(tmp_Rs, rs_Tab2SpecDemandDetail)
tmp_Rs.Close

With dg_Tab2SpecDemandDetail
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

If Not rs_Tab2SpecDemandDetail.EOF Then rs_Tab2SpecDemandDetail.MoveFirst

Set dg_Tab2SpecDemandDetail.DataSource = rs_Tab2SpecDemandDetail
With dg_Tab2SpecDemandDetail
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 800       '�f�D
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '�Ȥ�s��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '�~��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000       '�Ƶ�
    .Columns(4).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�S��ݨD-��ܩҦ����", Me.Caption, "cmd_Tab2SpecDemand_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3Area_Save_Click()
'�B�e�ϰ� >> �s��
If rs_Tab3Area Is Nothing Then
   msg_text = "�Х��d�ߩҦ�[�B�e�ϰ�]��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab3Area.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ�[�B�e�ϰ�]��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab3Area.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab3Area.EOF
   str_SQL = "Update TRP03M " & _
             "Set Description = '" & rs_Tab3Area.Fields("�B�e�ϰ�").Value & "', "
   If Trim(rs_Tab3Area.Fields("�̤j���ح���").Value) = "" Then
      str_SQL = str_SQL & "Max_Size_Limit = null,"
   Else
      str_SQL = str_SQL & "Max_Size_Limit = " & Val(rs_Tab3Area.Fields("�̤j���ح���").Value) & ","
   End If
   If Trim(rs_Tab3Area.Fields("�̤p���ح���").Value) = "" Then
      str_SQL = str_SQL & "Min_Size_Limit = null "
   Else
      str_SQL = str_SQL & "Min_Size_Limit = " & Val(rs_Tab3Area.Fields("�̤p���ح���").Value) & " "
   End If
   str_SQL = str_SQL & _
             "Where Area_Code = '" & rs_Tab3Area.Fields("�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP03M(Area_Code,Max_Size_Limit,Min_Size_Limit,Description) Values (" & _
                "'" & rs_Tab3Area.Fields("�N�X").Value & "',"
      If Trim(rs_Tab3Area.Fields("�̤j���ح���").Value) = "" Then
         str_SQL = str_SQL & "null,"
      Else
         str_SQL = str_SQL & Val(rs_Tab3Area.Fields("�̤j���ح���").Value) & ", "
      End If
      If Trim(rs_Tab3Area.Fields("�̤p���ح���").Value) = "," Then
         str_SQL = str_SQL & "null,'"
      Else
         str_SQL = str_SQL & Val(rs_Tab3Area.Fields("�̤p���ح���").Value) & ",'"
      End If
      str_SQL = str_SQL & rs_Tab3Area.Fields("�B�e�ϰ�").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab3Area.MoveNext
Loop
rs_Tab3Area.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-�B�e�ϰ�-�s��", Me.Caption, "cmd_Tab3Area_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3Area_Show_Click()
'�S��ݨD >> ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab3Area.DataSource = Nothing
Call ReDim_Recordset(rs_Tab3Area)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(Area_Code) AS �N�X, RTRIM(Isnull(Cast(MAX_SIZE_LIMIT as varchar(300)),'')) AS �̤j���ح���,RTRIM(Isnull(Cast(MIN_SIZE_LIMIT as varchar(300)),'')) AS �̤p���ح���, " & _
          "RTRIM(Isnull(Description,'')) AS �B�e�ϰ� " & _
          "From TRP03M Order by Area_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab3Area)
tmp_Rs.Close

With dg_Tab3Area
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab3Area.MoveFirst
Set dg_Tab3Area.DataSource = rs_Tab3Area
With dg_Tab3Area
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1300       '�̤j���ح���
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1300       '�̤p���ح���
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 4000       '�B�e�ϰ�
    .Columns(4).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�B�e�ϰ�-��ܩҦ����", Me.Caption, "cmd_Tab3Area_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4Zip_Save_Click()
'�l���ϸ� >> �s��
If rs_Tab4Zip Is Nothing Then
   msg_text = "�Х��d�ߩҦ� [�l���ϸ�] ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab4Zip.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� [�l���ϸ�] ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab4Zip.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab4Zip.EOF
   str_SQL = "Update TRP02M " & _
             "Set Description = '" & rs_Tab4Zip.Fields("����").Value & "',city = '" & rs_Tab4Zip.Fields("����").Value & "',dcode = '" & rs_Tab4Zip("�H�t��۽X") & "',E_Abb = '" & rs_Tab4Zip("�Y�g") & "', "
   If Trim(rs_Tab4Zip.Fields("�B�e�ϰ�N�X").Value) = "" Then
      str_SQL = str_SQL & "Area_Code = null "
   Else
      str_SQL = str_SQL & "Area_Code = '" & Trim(rs_Tab4Zip.Fields("�B�e�ϰ�N�X").Value) & "' "
   End If
   str_SQL = str_SQL & _
             "Where ZIP = '" & rs_Tab4Zip.Fields("�l���ϸ�").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP02M(ZIP,city,dcode,Area_Code,Description,E_Abb) Values (" & _
                "'" & rs_Tab4Zip.Fields("�l���ϸ�").Value & "', '" & rs_Tab4Zip.Fields("����").Value & "','" & rs_Tab4Zip.Fields("�H�t��۽X").Value & "',"
      If Trim(rs_Tab4Zip.Fields("�B�e�ϰ�N�X").Value) = "" Then
         str_SQL = str_SQL & "null,'"
      Else
         str_SQL = str_SQL & "'" & Trim(rs_Tab4Zip.Fields("�B�e�ϰ�N�X").Value) & "', '"
      End If
      str_SQL = str_SQL & rs_Tab4Zip.Fields("����").Value & "','" & rs_Tab4Zip.Fields("�Y�g") & "') "
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab4Zip.MoveNext
Loop
rs_Tab4Zip.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-�B�e�ϰ�-�s��", Me.Caption, "cmd_Tab3Area_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4Zip_Show_Click()
'�l���ϸ� >> ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab4Zip.DataSource = Nothing
Call ReDim_Recordset(rs_Tab4Zip)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(ZIP) AS �l���ϸ�,RTRIM(Area_Code) AS �B�e�ϰ�N�X,RTRIM(city) AS ����,RTRIM(Isnull(Description,'')) AS ����,�H�t��۽X=rtrim(isnull(dcode,'')),�Y�g = isnull(E_Abb,'') " & _
          "From TRP02M Order by ZIP "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab4Zip)
tmp_Rs.Close

With dg_Tab4Zip
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab4Zip.MoveFirst
Set dg_Tab4Zip.DataSource = rs_Tab4Zip
With dg_Tab4Zip
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�l���ϸ�
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '�B�e�ϰ�N�X
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '����
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1500       '����
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1000       '�H�t��۽X
    .Columns(5).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�l���ϸ�-��ܩҦ����", Me.Caption, "cmd_Tab4ZIP_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5GridCode_Show_Click()
'�x�}�ϽX >> ��ܩҦ����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab5GridCode.DataSource = Nothing
Call ReDim_Recordset(rs_Tab5GridCode)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select Rtrim(Grid_Code) as �x�}�ϽX , Rtrim(Isnull(Grid_Type ,'')) as ���O,Rtrim(Isnull(X_Coordinate,'')) as X�y��," & _
          "   Rtrim(Isnull(Y_Coordinate,'')) as Y�y�� , Rtrim(Isnull(Description,'')) as ���� " & _
          "From TRP14M order by GRID_CODE"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab5GridCode)
tmp_Rs.Close

With dg_Tab5GridCode
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab5GridCode.MoveFirst
Set dg_Tab5GridCode.DataSource = rs_Tab5GridCode
With dg_Tab5GridCode
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�x�}�ϽX
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '�x�}�����O
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       'X�y��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000       'Y�y��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 3000       '����
    .Columns(5).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�x�}�ϽX-��ܩҦ����", Me.Caption, "cmd_Tab5GridCode_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab5GridCodeSave_Click()
'�x�}�ϽX >> �s��
If rs_Tab5GridCode Is Nothing Then
   msg_text = "�Х��d�ߩҦ� [�x�}�ϽX] ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab5GridCode.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� [�x�}�ϽX] ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab5GridCode.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab5GridCode.EOF
   str_SQL = "Update TRP14M " & _
             "Set Grid_Type = '" & rs_Tab5GridCode.Fields("���O").Value & "',X_Coordinate = '" & rs_Tab5GridCode.Fields("X�y��").Value & "'," & _
             " Y_Coordinate = '" & rs_Tab5GridCode.Fields("Y�y��").Value & "',Description = '" & rs_Tab5GridCode.Fields("����").Value & "' " & _
             "Where Grid_Code = '" & rs_Tab5GridCode.Fields("�x�}�ϽX").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP14M(Grid_Code,Grid_Type,X_Coordinate,Y_Coordinate,Description) Values (" & _
                "'" & rs_Tab4Zip.Fields("�x�}�ϽX").Value & "','" & rs_Tab5GridCode.Fields("���O").Value & "','" & rs_Tab5GridCode.Fields("X�y��").Value & "','" & _
                rs_Tab5GridCode.Fields("Y�y��").Value & "','" & rs_Tab5GridCode.Fields("����").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab5GridCode.MoveNext
Loop
rs_Tab5GridCode.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-�x�}�ϽX-�s��", Me.Caption, "cmd_Tab5GridCode_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6RBC_Save_Click()
'�򥻥N�X2 >> ���`�d�� >> �s��
If rs_Tab6RBC Is Nothing Then
   msg_text = "�Х��d�ߩҦ� ���`�d�� ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab6RBC.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� ���[�d�� ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab6RBC.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab6RBC.EOF
   str_SQL = "Update TRP06M " & _
             "Set Description = '" & rs_Tab6RBC.Fields("���`�d��").Value & "' " & _
             "Where RBC_Code = '" & rs_Tab6RBC.Fields("�d�ݥN�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP06M(RBC_Code,Description) Values (" & _
                "'" & rs_Tab6RBC.Fields("�d�ݥN�X").Value & "','" & rs_Tab6RBC.Fields("���`�d��").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab6RBC.MoveNext
Loop
rs_Tab6RBC.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-���`�d��-�s��", Me.Caption, "cmd_Tab6RBC_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6RBC_Show_Click()
'�򥻥N�X2 >> ���`�d��
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab6RBC.DataSource = Nothing
Call ReDim_Recordset(rs_Tab6RBC)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(RBC_Code) AS �d�ݥN�X, RTRIM(Description) AS ���`�d�� " & _
          "From TRP06M Order by RBC_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab6RBC)
tmp_Rs.Close

With dg_Tab6RBC
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab6RBC.MoveFirst
Set dg_Tab6RBC.DataSource = rs_Tab6RBC
With dg_Tab6RBC
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�d�ݥN�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2500       '���`�d��
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���`�d��-��ܩҦ����", Me.Caption, "cmd_Tab6RBC_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmd_Tab6RSC_Save_Click()
'�򥻥N�X2 >> ���`��] >> �s��
If rs_Tab6RSC Is Nothing Then
   msg_text = "�Х��d�ߩҦ� ���`��] ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If rs_Tab6RSC.RecordCount = 0 Then
   msg_text = "�Х��d�ߩҦ� ���[�Φ� ��ơA�T�{��A���� [�s��] �@�~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents

Tran_Level = 0
Tran_Level = cn.BeginTrans

rs_Tab6RSC.MoveFirst
Call DB_CheckConnectStatus
Do While Not rs_Tab6RSC.EOF
   str_SQL = "Update TRP05M " & _
             "Set Description = '" & rs_Tab6RSC.Fields("���`��]").Value & "' " & _
             "Where RSC_Code = '" & rs_Tab6RSC.Fields("���`�N�X").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '�䤣��i��s����ƦC >> �s�W�������
   If RowsAffect = 0 Then
      str_SQL = "Insert into TRP05M(RSC_Code,Description) Values (" & _
                "'" & rs_Tab6RSC.Fields("���`�N�X").Value & "','" & rs_Tab6RSC.Fields("���`��]").Value & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_Tab6RSC.MoveNext
Loop
rs_Tab6RSC.MoveFirst
cn.CommitTrans
Tran_Level = 0

msg_text = "�s�ɧ@�~����"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

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
   CreateErrorLog Me.Name & "-���`��]-�s��", Me.Caption, "cmd_Tab6RSC_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab6RSC_Show_Click()
'�򥻥N�X2 >> ���`��]
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab6RSC.DataSource = Nothing
Call ReDim_Recordset(rs_Tab6RSC)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT RTRIM(RSC_Code) AS ���`�N�X, RTRIM(Description) AS ���`��] " & _
          "From TRP05M Order by RSC_Code "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab6RSC)
tmp_Rs.Close

With dg_Tab6RSC
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab6RSC.MoveFirst
Set dg_Tab6RSC.DataSource = rs_Tab6RSC
With dg_Tab6RSC
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '���`�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2500       '���`��]
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���`��]-��ܩҦ����", Me.Caption, "cmd_Tab6RSC_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab7_Delete_Click()

If rs_Tab7_TRP17M Is Nothing Then Exit Sub

Call ReDim_Recordset(tmp_Rs)

str_SQL = "select * from sdn05t s5 join sdn02t s2 on s5.sdn_no = s2.receipt_no and s2.storerkey = '" & rs_Tab7_TRP17M.Fields("�f�D") & "' and s5.costcode = '" & rs_Tab7_TRP17M.Fields("�N�X") & "' "
tmp_Rs.Open str_SQL, cn
If Not tmp_Rs.EOF Then MsgBox "�ϥΤ��p�O�N�X�L�k�R��!!", 64, "�R��": Exit Sub

If MsgBox("�T�{�R���p�O�N�X " & rs_Tab7_TRP17M.Fields("�f�D") & "-" & rs_Tab7_TRP17M("�N�X") & " !!", vbOKCancel, "�R��") <> vbOK Then Exit Sub

cn.Execute "delete trp17m where storerkey = '" & rs_Tab7_TRP17M.Fields("�f�D") & "' and costcode = '" & rs_Tab7_TRP17M.Fields("�N�X") & "' ", RowsAffect, adExecuteNoRecords

rs_Tab7_TRP17M.Delete

tmp_Rs.Close

End Sub

Private Sub cmd_Tab7_DisPlay_Click()
    '���إN�X >> ��ܩҦ����
    On Error GoTo err_Handle
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_Tab1CarType.DataSource = Nothing
    Call ReDim_Recordset(rs_Tab7_TRP17M)
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "SELECT �f�D = rtrim(storerkey) ,RTRIM(CostCode) AS �N�X,rtrim(CostKind) as �д����O,��� = rtrim(UOM) ,Receivable as �������,Payable as ���I���,rtrim(AreaStart) as �_�I,rtrim(AreaEnd) as ���I,rtrim(CostName) as �p�O�W��,rtrim(CostNote) as ���� " & _
              "From TRP17M Order by storerkey,CostCode"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C���"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
'    Call Replication_Recordset(tmp_rs, rs_Tab7_TRP17M)
    Call OffLineRecordset(tmp_Rs, rs_Tab7_TRP17M)
    tmp_Rs.Close
    
    With dg_Tab1CarType
         .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
         .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
         .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
         .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
    End With
    rs_Tab7_TRP17M.MoveFirst
    Set dg_Tab7_TRP17M.DataSource = rs_Tab7_TRP17M
    With dg_Tab7_TRP17M
        .RowHeight = 250
        .Columns(0).Width = 800        '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 800      '����
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000       '�N�X
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 800       '�N�X
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1000      '�Ȥ�W��
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 800      '�������
        .Columns(5).Alignment = dbgRight
        .Columns(6).Width = 800      '���I���
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 1000      '�_�I
        .Columns(7).Alignment = dbgLeft
        .Columns(8).Width = 1000      '���I
        .Columns(8).Alignment = dbgLeft
        .Columns(9).Width = 3000      '����
        .Columns(9).Alignment = dbgLeft
    End With
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-��������-��ܩҦ����", Me.Caption, "cmd_Tab1CarType_Show_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab7_Save_Click()
    '�������� >> �s��
    If rs_Tab7_TRP17M Is Nothing Then
        msg_text = "�Х��d�ߩҦ���ơA�T�{��A���� [�s��] �@�~"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
        If rs_Tab7_TRP17M.RecordCount = 0 Then
        msg_text = "�Х��d�ߩҦ���ơA�T�{��A���� [�s��] �@�~"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    On Error GoTo err_Handle
    Screen.MousePointer = vbHourglass
    dg_Tab7_TRP17M.Enabled = False
    DoEvents: DoEvents
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    rs_Tab7_TRP17M.MoveFirst
    Call DB_CheckConnectStatus
    Do While Not rs_Tab7_TRP17M.EOF
    
    If Len(Trim(rs_Tab7_TRP17M.Fields("�f�D"))) = 0 Then MsgBox "�п�J�f�D���", 64, "�s��": Tran_Level = 0: cn.RollbackTrans: Screen.MousePointer = 0: dg_Tab7_TRP17M.Enabled = True: Exit Sub
    If Len(Trim(rs_Tab7_TRP17M.Fields("�N�X"))) = 0 Then MsgBox "�п�J�N�X���", 64, "�s��": Tran_Level = 0: cn.RollbackTrans: Screen.MousePointer = 0: dg_Tab7_TRP17M.Enabled = True: Exit Sub
    
       str_SQL = "Update TRP17M " & _
                  "Set Storerkey = '" & Trim(rs_Tab7_TRP17M.Fields("�f�D")) & "' ,CostName = '" & rs_Tab7_TRP17M.Fields("�p�O�W��").Value & "',Receivable = '" & rs_Tab7_TRP17M.Fields("�������").Value & "', " & _
                  "Payable = '" & rs_Tab7_TRP17M.Fields("���I���").Value & "',AreaStart = '" & rs_Tab7_TRP17M.Fields("�_�I").Value & "'," & _
                  "AreaEnd = '" & rs_Tab7_TRP17M.Fields("���I").Value & "',CostNote = '" & rs_Tab7_TRP17M.Fields("����").Value & "'," & _
                  "CostKind = '" & rs_Tab7_TRP17M.Fields("�д����O").Value & "' ,UOM = '" & rs_Tab7_TRP17M("���") & "' " & _
                  "Where Storerkey = '" & rs_Tab7_TRP17M.Fields("�f�D") & "' and CostCode = '" & Trim(rs_Tab7_TRP17M.Fields("�N�X").Value) & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '�䤣��i��s����ƦC >> �s�W�������
        If RowsAffect = 0 Then
            str_SQL = "Insert into TRP17M (Storerkey,CostCode,CostName,Receivable,Payable,AreaStart,AreaEnd,CostNote,CostKind,adduser,UOM) Values (" & _
                      "'" & Trim(rs_Tab7_TRP17M.Fields("�f�D").Value) & "','" & Trim(rs_Tab7_TRP17M.Fields("�N�X").Value) & "','" & rs_Tab7_TRP17M.Fields("�p�O�W��").Value & "', '" & rs_Tab7_TRP17M.Fields("�������").Value & "', " & _
                      "'" & rs_Tab7_TRP17M.Fields("���I���").Value & "','" & rs_Tab7_TRP17M.Fields("�_�I").Value & "'," & _
                      "'" & rs_Tab7_TRP17M.Fields("���I").Value & "','" & rs_Tab7_TRP17M.Fields("����").Value & "','" & rs_Tab7_TRP17M.Fields("�д����O").Value & "','" & User_id & "' , '" & rs_Tab7_TRP17M("���") & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
        rs_Tab7_TRP17M.MoveNext
    Loop
    rs_Tab7_TRP17M.MoveFirst
    cn.CommitTrans: Tran_Level = 0
    
    dg_Tab7_TRP17M.Enabled = True
    msg_text = "�s�ɧ@�~����"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    
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
    CreateErrorLog Me.Name & "-�p�O�N�X-�s��", Me.Caption, "cmd_Tab1CarType_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�@�~�N�X��ƺ��@"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 6405
dbsrcFormWidth = 10665
Me.Height = 6915: Me.Width = 10800: Me.Left = 0
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

SSTab1.Tab = 0

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
Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_BaseData_OPCode = Nothing
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub
