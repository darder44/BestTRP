VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_BaseData_Code 
   Caption         =   "�t �� �N �X �� �@"
   ClientHeight    =   5490
   ClientLeft      =   1770
   ClientTop       =   1815
   ClientWidth     =   8175
   Icon            =   "frm_BaseData_Code.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   8175
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   9631
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "�N�X���O"
      TabPicture(0)   =   "frm_BaseData_Code.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmd_List_ListName"
      Tab(0).Control(1)=   "dg_ListName"
      Tab(0).Control(2)=   "cmd_Exit(0)"
      Tab(0).Control(3)=   "cmd_Save_ListName"
      Tab(0).Control(4)=   "Label1(13)"
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "���q���"
      TabPicture(1)   =   "frm_BaseData_Code.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_Save_UserCompany"
      Tab(1).Control(1)=   "cmd_Exit(1)"
      Tab(1).Control(2)=   "cmd_List_UserCompany"
      Tab(1).Control(3)=   "dg_UserCompany"
      Tab(1).Control(4)=   "Label1(14)"
      Tab(1).Control(5)=   "Label1(4)"
      Tab(1).Control(6)=   "Label1(3)"
      Tab(1).Control(7)=   "Label1(2)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "�s�ո��"
      TabPicture(2)   =   "frm_BaseData_Code.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(5)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(6)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(7)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(15)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "dg_UserGroup"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmd_Save_UserGroup"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmd_Exit(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmd_List_UserGroup"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Menu"
      TabPicture(3)   =   "frm_BaseData_Code.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmd_List_MenuForm"
      Tab(3).Control(1)=   "cmd_Exit(4)"
      Tab(3).Control(2)=   "cmd_Save_MenuForm"
      Tab(3).Control(3)=   "dg_MenuForm"
      Tab(3).Control(4)=   "Label1(10)"
      Tab(3).Control(5)=   "Label1(11)"
      Tab(3).Control(6)=   "Label1(12)"
      Tab(3).Control(7)=   "Label1(17)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frm_BaseData_Code.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.CommandButton cmd_List_MenuForm 
         BackColor       =   &H00C0FFFF&
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
         Height          =   945
         Left            =   -68025
         Picture         =   "frm_BaseData_Code.frx":0956
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   26
         Top             =   3900
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
         Index           =   4
         Left            =   -68040
         Picture         =   "frm_BaseData_Code.frx":0C60
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   25
         Top             =   2430
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Save_MenuForm 
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
         Left            =   -68040
         Picture         =   "frm_BaseData_Code.frx":10A2
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   24
         Top             =   1245
         Width           =   1050
      End
      Begin VB.CommandButton cmd_List_UserGroup 
         BackColor       =   &H00C0FFFF&
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
         Height          =   945
         Left            =   6810
         Picture         =   "frm_BaseData_Code.frx":13AC
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   16
         Top             =   3825
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
         Left            =   6810
         Picture         =   "frm_BaseData_Code.frx":16B6
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   15
         Top             =   2355
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Save_UserGroup 
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
         Left            =   6810
         Picture         =   "frm_BaseData_Code.frx":1AF8
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   1185
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Save_UserCompany 
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
         Left            =   -68190
         Picture         =   "frm_BaseData_Code.frx":1E02
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   1185
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
         Index           =   1
         Left            =   -68190
         Picture         =   "frm_BaseData_Code.frx":210C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
         Top             =   2355
         Width           =   1050
      End
      Begin VB.CommandButton cmd_List_UserCompany 
         BackColor       =   &H00C0FFFF&
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
         Height          =   945
         Left            =   -68190
         Picture         =   "frm_BaseData_Code.frx":254E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
         Top             =   3825
         Width           =   1050
      End
      Begin VB.CommandButton cmd_List_ListName 
         BackColor       =   &H00C0FFFF&
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
         Height          =   945
         Left            =   -68190
         Picture         =   "frm_BaseData_Code.frx":2858
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   3825
         Width           =   1050
      End
      Begin MSDataGridLib.DataGrid dg_ListName 
         Height          =   3585
         Left            =   -74715
         TabIndex        =   3
         Top             =   1215
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   6324
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         Index           =   0
         Left            =   -68190
         Picture         =   "frm_BaseData_Code.frx":2B62
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   2355
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Save_ListName 
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
         Left            =   -68190
         Picture         =   "frm_BaseData_Code.frx":2FA4
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   1
         Top             =   1185
         Width           =   1050
      End
      Begin MSDataGridLib.DataGrid dg_UserCompany 
         Height          =   3585
         Left            =   -74715
         TabIndex        =   8
         Top             =   1215
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   6324
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_UserGroup 
         Height          =   3585
         Left            =   285
         TabIndex        =   17
         Top             =   1215
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   6324
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin MSDataGridLib.DataGrid dg_MenuForm 
         Height          =   3585
         Left            =   -74880
         TabIndex        =   27
         Top             =   1290
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6324
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�`�N�GMenu / Form �N�X�H�s�W����h"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   10
         Left            =   -74220
         TabIndex        =   31
         Top             =   4995
         Width           =   4260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "TableName�GLogictown.dbo.CodeLKUP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Index           =   11
         Left            =   -74220
         TabIndex        =   30
         Top             =   495
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "ListName = ""APMENU"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   12
         Left            =   -74220
         TabIndex        =   29
         Top             =   825
         Width           =   3075
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   17
         Left            =   -68130
         TabIndex        =   28
         Top             =   765
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   15
         Left            =   6795
         TabIndex        =   23
         Top             =   690
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   14
         Left            =   -68205
         TabIndex        =   22
         Top             =   690
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ި�@�~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   13
         Left            =   -68205
         TabIndex        =   21
         Top             =   690
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�`�N�G�s�եN�X�H�s�W����h"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   7
         Left            =   705
         TabIndex        =   20
         Top             =   4920
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "TableName�GLogictown.dbo.CodeLKUP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Index           =   6
         Left            =   705
         TabIndex        =   19
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "ListName = ""USERGROUP"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   705
         TabIndex        =   18
         Top             =   810
         Width           =   3630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "ListName = ""USERCOMPANY"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   4
         Left            =   -74295
         TabIndex        =   13
         Top             =   810
         Width           =   4080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "TableName�GLogictown.dbo.CodeLKUP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Index           =   3
         Left            =   -74295
         TabIndex        =   12
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�`�N�G���q�N�X�H�s�W����h"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   2
         Left            =   -74295
         TabIndex        =   11
         Top             =   4920
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�`�N�G���O�N�X�H�s�W����h"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   -74295
         TabIndex        =   6
         Top             =   4920
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "TableName�GLogictown.dbo.CodeList"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Index           =   0
         Left            =   -74295
         TabIndex        =   4
         Top             =   660
         Width           =   4740
      End
   End
End
Attribute VB_Name = "frm_BaseData_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs_ListName As ADODB.Recordset      '�N�X���O�]�w���
Private rs_UserCompany As ADODB.Recordset   '�ϥΪ̸��--���q
Private rs_UserGroup As ADODB.Recordset     '�ϥΪ̸��--�s��
Private rs_StorerMap As ADODB.Recordset     '���q�BExceed Storer ����
Private rs_MenuForm As ADODB.Recordset      '�l�t�� Menu & Form �N�X�]�w���

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_List_ListName_Click()
'�N�X���O >> ��ܩҦ����
Set dg_ListName.DataSource = Nothing
Screen.MousePointer = vbHourglass
str_SQL = "Select Rtrim(ListName) as 'ListName' , Rtrim(Description) as 'Descr' From CodeList Order by ListName"
On Error GoTo err_Handle
Call Confirm_Recordset_Closed(tmp_Rs)
Screen.MousePointer = vbHourglass
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�N�X���O�]�w���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_ListName)
With dg_ListName
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_ListName.MoveFirst
Set dg_ListName.DataSource = rs_ListName
With dg_ListName
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1800       '���O�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 3300       '���O�N�X����
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[�N�X���O]-[��ܩҦ����]", Me.Caption, "cmd_List_ListName_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_List_MenuForm_Click()
'Menu/Form >> ��ܩҦ����
Set dg_MenuForm.DataSource = Nothing
Screen.MousePointer = vbHourglass
str_SQL = "Select Rtrim(Code) as 'APCode' , Rtrim(Description) as 'Descr' From CodeLKUP Where ListName = 'APMENU' Order by Description"
On Error GoTo err_Handle
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L Menu & Form �N�X�]�w���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_MenuForm)
With dg_MenuForm
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_MenuForm.MoveFirst
Set dg_MenuForm.DataSource = rs_MenuForm
With dg_MenuForm
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 3000       '�s�եN�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 3500       '�s�եN�X����
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[Menu & Form]-[��ܩҦ����]", Me.Caption, "cmd_List_MenuForm_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_List_UserCompany_Click()
'���q��� >> ��ܩҦ����
Set dg_UserCompany.DataSource = Nothing
Screen.MousePointer = vbHourglass
str_SQL = "Select Rtrim(Code) as 'Code' , Rtrim(Description) as 'Descr' From CodeLKUP Where ListName = 'USERCOMPANY' Order by Code"
On Error GoTo err_Handle
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L���q�N�X�]�w���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_UserCompany)
With dg_UserCompany
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_UserCompany.MoveFirst
Set dg_UserCompany.DataSource = rs_UserCompany
With dg_UserCompany
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1800       '���q�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 3300       '���q�N�X����
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[���q���]-[��ܩҦ����]", Me.Caption, "cmd_List_UserCompany_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_List_UserGroup_Click()
'�s�ո�� >> ��ܩҦ����
Set dg_UserGroup.DataSource = Nothing
Screen.MousePointer = vbHourglass
str_SQL = "Select Rtrim(Code) as 'Code' , Rtrim(Description) as 'Descr' From CodeLKUP Where ListName = 'USERGROUP' Order by Code"
On Error GoTo err_Handle
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�s�եN�X�]�w���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_UserGroup)
With dg_UserGroup
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_UserGroup.MoveFirst
Set dg_UserGroup.DataSource = rs_UserGroup
With dg_UserGroup
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1800       '�s�եN�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 3300       '�s�եN�X����
    .Columns(2).Alignment = dbgLeft
End With
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[�s�ո��]-[��ܩҦ����]", Me.Caption, "cmd_List_UserGroup_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Save_ListName_Click()
'�N�X���O >> �s��
If rs_ListName Is Nothing Then Exit Sub
If rs_ListName.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle
If Tran_Level = 0 Then
   Tran_Level = cn.BeginTrans
End If
rs_ListName.MoveFirst
Do While Not rs_ListName.EOF
   str_SQL = "Update CodeList Set Description='" & rs_ListName.Fields("Descr").Value & "',EditDate=Getdate(),EditWho='" & User_id & "' " & _
             "Where ListName = '" & rs_ListName.Fields("ListName").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeList (ListName,Description,AddDate,AddWho,EditDate,EditWho) Values ('" & _
                rs_ListName.Fields("ListName").Value & "','" & rs_ListName.Fields("Descr").Value & "',Getdate(),'" & User_id & "',Getdate(),'" & User_id & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
      
   End If
   rs_ListName.MoveNext
Loop
cn.CommitTrans
Tran_Level = 0

Set dg_ListName.DataSource = Nothing
Set rs_ListName = Nothing

Exit Sub

err_Handle:
     If Tran_Level <> 0 Then
        cn.RollbackTrans
        Tran_Level = 0
     End If
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[�N�X���O]-[�s��]", Me.Caption, "cmd_Save_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Save_MenuForm_Click()
'Menu/Form >> �s��
If rs_MenuForm Is Nothing Then Exit Sub
If rs_MenuForm.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle
If Tran_Level = 0 Then
   Tran_Level = cn.BeginTrans
End If
rs_MenuForm.MoveFirst
Do While Not rs_MenuForm.EOF
   str_SQL = "Update CodeLKUP Set Description='" & rs_MenuForm.Fields("Descr").Value & "',EditDate=Getdate(),EditWho='" & User_id & "' " & _
             "Where ListName = 'APMENU' and Code = '" & rs_MenuForm.Fields("APCode").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP (ListName,Code,Description,AddDate,AddWho,EditDate,EditWho) Values ('APMENU','" & _
                rs_MenuForm.Fields("APCode").Value & "','" & rs_MenuForm.Fields("Descr").Value & "',Getdate(),'" & User_id & "',Getdate(),'" & User_id & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_MenuForm.MoveNext
Loop
cn.CommitTrans
Tran_Level = 0

Set dg_MenuForm.DataSource = Nothing
Set rs_MenuForm = Nothing

Exit Sub

err_Handle:
     If Tran_Level <> 0 Then
         cn.RollbackTrans
         Tran_Level = 0
     End If
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[Menu/Form]-[�s��]", Me.Caption, "cmd_Save_MenuForm_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub



Private Sub cmd_Save_UserCompany_Click()
'���q��� >> �s��
If rs_UserCompany Is Nothing Then Exit Sub
If rs_UserCompany.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle
If Tran_Level = 0 Then
   Tran_Level = cn.BeginTrans
End If
rs_UserCompany.MoveFirst
Do While Not rs_UserCompany.EOF
   str_SQL = "Update CodeLKUP Set Description='" & rs_UserCompany.Fields("Descr").Value & "',EditDate=Getdate(),EditWho='" & User_id & "' " & _
             "Where ListName = 'USERCOMPANY' and Code = '" & rs_UserCompany.Fields("Code").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP (ListName,Code,Description,AddDate,AddWho,EditDate,EditWho) Values ('USERCOMPANY','" & _
                rs_UserCompany.Fields("Code").Value & "','" & rs_UserCompany.Fields("Descr").Value & "',Getdate(),'" & User_id & "',Getdate(),'" & User_id & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_UserCompany.MoveNext
Loop
cn.CommitTrans
Tran_Level = 0

Set dg_UserCompany.DataSource = Nothing
Set rs_UserCompany = Nothing

Exit Sub

err_Handle:
     If Tran_Level <> 0 Then
        cn.RollbackTrans
        Tran_Level = 0
     End If
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[���q���]-[�s��]", Me.Caption, "cmd_Save_UserCompany_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Save_UserGroup_Click()
'�s�ո�� >> �s��
If rs_UserGroup Is Nothing Then Exit Sub
If rs_UserGroup.RecordCount = 0 Then Exit Sub
On Error GoTo err_Handle
If Tran_Level = 0 Then
   Tran_Level = cn.BeginTrans
End If
rs_UserGroup.MoveFirst
Do While Not rs_UserGroup.EOF
   str_SQL = "Update CodeLKUP Set Description='" & rs_UserGroup.Fields("Descr").Value & "',EditDate=Getdate(),EditWho='" & User_id & "' " & _
             "Where ListName = 'USERGROUP' and Code = '" & rs_UserGroup.Fields("Code").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   If RowsAffect = 0 Then
      str_SQL = "Insert into CodeLKUP (ListName,Code,Description,AddDate,AddWho,EditDate,EditWho) Values ('USERGROUP','" & _
                rs_UserGroup.Fields("Code").Value & "','" & rs_UserGroup.Fields("Descr").Value & "',Getdate(),'" & User_id & "',Getdate(),'" & User_id & "')"
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End If
   rs_UserGroup.MoveNext
Loop
cn.CommitTrans
Tran_Level = 0

Set dg_UserGroup.DataSource = Nothing
Set rs_UserGroup = Nothing

Exit Sub

err_Handle:
     If Tran_Level <> 0 Then
        cn.RollbackTrans
        Tran_Level = 0
     End If
     Dim tmpString As String
     Screen.MousePointer = vbDefault
     tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
     msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Desc:" & err.Description
     CreateErrorLog Me.Name & "-[�s�ո��]-[�s��]", Me.Caption, "cmd_Save_UserGroup_Click", tmpString
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub



Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�l�t�ΥN�X���@"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
Me.Height = 6000: Me.Width = 8300
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_BaseData_Code = Nothing
End Sub

