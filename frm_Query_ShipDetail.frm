VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Query_ShipDetail 
   Caption         =   "�t�m���Ӭd��"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   15.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Query_ShipDetail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   12600
   WindowState     =   2  '�̤j��
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   7440
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5520
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
      StartOfWeek     =   61472769
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   3
      Top             =   3120
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   20
         TabAction       =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmd2Excel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��Excel"
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
         Left            =   9840
         Picture         =   "frm_Query_ShipDetail.frx":08CA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   8
         Top             =   2040
         Width           =   1065
      End
      Begin VB.CommandButton cmdExit 
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
         Height          =   870
         Left            =   11040
         Picture         =   "frm_Query_ShipDetail.frx":1BC4
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
         Top             =   2040
         Width           =   1065
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0FFC0&
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
         Height          =   870
         Left            =   11040
         Picture         =   "frm_Query_ShipDetail.frx":2B7D6
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         Top             =   1080
         Width           =   1065
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00C0E0FF&
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
         Height          =   870
         Left            =   9840
         Picture         =   "frm_Query_ShipDetail.frx":2BAE8
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   1080
         Width           =   1065
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   4895
         _Version        =   393216
         TabHeight       =   520
         TabMaxWidth     =   4410
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "�����J"
         TabPicture(0)   =   "frm_Query_ShipDetail.frx":2BDF2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label3(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label3(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(21)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label1(24)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label1(25)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label1(7)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label1(10)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label3(4)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label3(5)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Text10"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Text9"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Text8"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Text7"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Text6"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Text5"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Text4"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Text3"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Frame4"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Text1"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Text2"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Text13"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Text14"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).ControlCount=   25
         TabCaption(1)   =   "������"
         TabPicture(1)   =   "frm_Query_ShipDetail.frx":2BE0E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "OrderType"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "WH"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Storerkey"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Level"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Lot6"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Loc"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label1(11)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label1(9)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label1(8)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label1(5)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label1(4)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label1(6)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "�����"
         TabPicture(2)   =   "frm_Query_ShipDetail.frx":2BE2A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "List1"
         Tab(2).ControlCount=   1
         Begin VB.ListBox OrderType 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   -73320
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   57
            Top             =   780
            Width           =   1455
         End
         Begin VB.TextBox Text14 
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
            MaxLength       =   20
            TabIndex        =   51
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text13 
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
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   50
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text2 
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
            MaxLength       =   8
            TabIndex        =   49
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
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
            Left            =   1080
            MaxLength       =   8
            TabIndex        =   48
            Top             =   480
            Width           =   1485
         End
         Begin VB.ListBox List1 
            Columns         =   6
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            Left            =   -74880
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   45
            Top             =   480
            Width           =   8895
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   4560
            TabIndex        =   35
            Top             =   360
            Width           =   4455
            Begin VB.CommandButton cmdSave 
               Caption         =   "�x�s"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   47
               Top             =   180
               Width           =   855
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "�R��"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3120
               TabIndex        =   43
               Top             =   180
               Width           =   855
            End
            Begin VB.Frame Frame3 
               Caption         =   "�s�W�`�ΰѼ�"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   120
               TabIndex        =   37
               Top             =   1080
               Width           =   4215
               Begin VB.TextBox Text12 
                  BeginProperty Font 
                     Name            =   "�s�ө���"
                     Size            =   11.25
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   720
                  MaxLength       =   60
                  TabIndex        =   40
                  Top             =   720
                  Width           =   3405
               End
               Begin VB.TextBox Text11 
                  BeginProperty Font 
                     Name            =   "�s�ө���"
                     Size            =   11.25
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   720
                  MaxLength       =   40
                  TabIndex        =   39
                  Top             =   240
                  Width           =   1965
               End
               Begin VB.CommandButton cmdAddnew 
                  Caption         =   "�s�W"
                  BeginProperty Font 
                     Name            =   "�s�ө���"
                     Size            =   9
                     Charset         =   136
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3000
                  TabIndex        =   38
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Alignment       =   2  '�m�����
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '�z��
                  Caption         =   "�W��"
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
                  Left            =   120
                  TabIndex        =   42
                  Top             =   300
                  Width           =   480
               End
               Begin VB.Label Label1 
                  Alignment       =   2  '�m�����
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '�z��
                  Caption         =   "����"
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
                  Index           =   3
                  Left            =   120
                  TabIndex        =   41
                  Top             =   780
                  Width           =   480
               End
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   11.25
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   36
               Text            =   "Combo1"
               Top             =   600
               Width           =   4245
            End
            Begin VB.Label Label1 
               Alignment       =   2  '�m�����
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�`�ΰѼ�"
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
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   960
            End
         End
         Begin VB.ListBox WH 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   -71760
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   32
            Top             =   780
            Width           =   1455
         End
         Begin VB.ListBox Storerkey 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   -74880
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   31
            Top             =   780
            Width           =   1455
         End
         Begin VB.ListBox Level 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   -68640
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   27
            Top             =   780
            Width           =   1455
         End
         Begin VB.ListBox Lot6 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   -70200
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   26
            Top             =   780
            Width           =   1455
         End
         Begin VB.ListBox Loc 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   -67080
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   25
            Top             =   780
            Width           =   1455
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   16
            Top             =   1560
            Width           =   1485
         End
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   15
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text5 
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   14
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text6 
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
            TabIndex        =   13
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text7 
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
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   12
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text8 
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
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text9 
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   10
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text10 
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
            TabIndex        =   9
            Top             =   2280
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���O"
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
            Left            =   -73320
            TabIndex        =   58
            Top             =   480
            Width           =   480
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
            Index           =   5
            Left            =   2655
            TabIndex        =   55
            Top             =   2340
            Width           =   240
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
            Index           =   4
            Left            =   2655
            TabIndex        =   54
            Top             =   1980
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q�渹�X"
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
            TabIndex        =   53
            Top             =   900
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��f���"
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
            TabIndex        =   52
            Top             =   540
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�w�O"
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
            Left            =   -71760
            TabIndex        =   34
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   8
            Left            =   -74880
            TabIndex        =   33
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�h��"
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
            Left            =   -68640
            TabIndex        =   30
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ܧO"
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
            Left            =   -70200
            TabIndex        =   29
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�x��"
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
            Left            =   -67080
            TabIndex        =   28
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "WMS�渹"
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
            Left            =   75
            TabIndex        =   24
            Top             =   1260
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�t�m�帹"
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
            Left            =   120
            TabIndex        =   23
            Top             =   1620
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���~�s��"
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
            TabIndex        =   22
            Top             =   1980
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�̪O�s��"
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
            Left            =   120
            TabIndex        =   21
            Top             =   2340
            Width           =   960
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
            Index           =   0
            Left            =   2655
            TabIndex        =   20
            Top             =   540
            Width           =   240
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
            Index           =   1
            Left            =   2655
            TabIndex        =   19
            Top             =   900
            Width           =   240
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
            Index           =   2
            Left            =   2655
            TabIndex        =   18
            Top             =   1260
            Width           =   240
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
            Index           =   3
            Left            =   2655
            TabIndex        =   17
            Top             =   1620
            Width           =   240
         End
      End
      Begin VB.Label Label2 
         Caption         =   "�t�m���Ӭd��"
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
         Height          =   735
         Left            =   9840
         TabIndex        =   46
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   7770
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "���A"
            TextSave        =   "���A"
            Object.ToolTipText     =   "���A"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   15611
            MinWidth        =   2646
            Object.ToolTipText     =   "��Ƶ���"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.ToolTipText     =   "�ϥΪ�"
         EndProperty
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
   End
End
Attribute VB_Name = "frm_Query_ShipDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object

Private Sub cmd2Excel_Click()

'��ƱƧ�
'�פJExcel
Recordset2Excel Left(Combo1.Text, InStr(Combo1.Text + " ", " ") - 1), rsMain
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmdAddnew_Click()
On Error GoTo err_Handle
If Len(RTrim(Text11.Text)) = 0 Then MsgBox "�п�J�ѼƦW��!!", 64, "�s�W�`�ΰѼ�": Exit Sub
Screen.MousePointer = 11

Text11.Text = Replace(Replace(Text11.Text, ",", ""), " ", "")

Dim rsTmp As New ADODB.Recordset
With rsTmp
    .CursorLocation = 3
    .Open "select * from codelkup where listname = 'QueryShipDetail' and code = '" & Text11.Text & "'", cn
    If .RecordCount > 0 Then
        MsgBox "�`�ΰѼƦW�٤w�s�b�A�п�ܨ�L�W��!!", 64, Me.Caption
    Else
    
        Call Save
        MsgBox """" & Text11.Text & """ �ѼƷs�W����!!", 64, "�`�ΰѼƷs�W"
        Combo1.AddItem (RTrim(Text11.Text) & " " & RTrim(Text12.Text))
        Combo1.ListIndex = Combo1.ListCount - 1
        
    End If

    .Close: Set rsTmp = Nothing
End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdSave_Click()
On Error GoTo err_Handle
If Len(RTrim(Combo1)) = 0 Then MsgBox "�п�ܰѼƦW��!!", 64, "�x�s�`�ΰѼ�": Exit Sub
Screen.MousePointer = 11

Dim rsTmp As New ADODB.Recordset
With rsTmp
    .CursorLocation = 3
    .Open "select * from codelkup where listname = 'QueryShipDetail' and code = '" & mySplit(Combo1, " ", 0) & "'", cn
    If .RecordCount = 0 Then
        MsgBox "�䤣��`�ΰѼ�!!", 64, Me.Caption
    Else
        cn.Execute "delete codelkup where listname = 'QueryShipDetail' and code = '" & mySplit(Combo1, " ", 0) & "'", RowsAffect, adExecuteNoRecords
        Call Save
        MsgBox """" & mySplit(Combo1, " ", 0) & """ �Ѽ��x�s����!!", 64, "�`�ΰѼ��x�s"
        
    End If

    .Close: Set rsTmp = Nothing
End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub Save()

Dim strSelected As String, i As Integer
strSelected = " " '���קK���ŭȪ��ﶵ����A�G�O�d�@�Ӧr��

'�f�D�����
For i = 0 To Storerkey.ListCount - 1
    If Storerkey.Selected(i) Then strSelected = strSelected & Trim(Storerkey.List(i)) & ","
Next
strSelected = Left(strSelected, Len(strSelected) - 1) & "| "

'�ܧO�����
For i = 0 To Lot6.ListCount - 1
    If Lot6.Selected(i) Then strSelected = strSelected & Lot6.List(i) & ","
Next
strSelected = Left(strSelected, Len(strSelected) - 1) & "| "

'�h�ƿ����
For i = 0 To Level.ListCount - 1
    If Level.Selected(i) Then strSelected = strSelected & Level.List(i) & ","
Next
strSelected = Left(strSelected, Len(strSelected) - 1) & "| "

'�x�Ͽ����
For i = 0 To Loc.ListCount - 1
    If Loc.Selected(i) Then strSelected = strSelected & Loc.List(i) & ","
Next
strSelected = Left(strSelected, Len(strSelected) - 1) & "| "

'�������
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then strSelected = strSelected & List1.List(i) & ","
Next
strSelected = Left(strSelected, Len(strSelected) - 1) & "| "

'�w�O�����
For i = 0 To WH.ListCount - 1
    If WH.Selected(i) Then strSelected = strSelected & WH.List(i) & ","
Next
strSelected = Left(strSelected, Len(strSelected) - 1) & "| "

str_SQL = "insert into codelkup (listname,code,[description],notes, adddate , addwho , editwho) " & _
            "values ('QueryShipDetail','" & RTrim(Text11) & "','" & RTrim(Text12) & "','" & strSelected & "' , getdate () , '" & User_id & "' ,'" & User_id & "')"

cn.BeginTrans: Tran_Level = 1
cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
cn.CommitTrans: Tran_Level = 0

Combo1 = Text11 & " " & Text12

End Sub
Private Sub cmdDelete_Click()

If Combo1.ListIndex = -1 Then Exit Sub
str_SQL = "delete codelkup where listname = 'QueryShipDetail' and code = '" & Left(Combo1.Text, InStr(Combo1.Text + " ", " ") - 1) & "'"

cn.BeginTrans
cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
cn.CommitTrans
Combo1.RemoveItem Combo1.ListIndex

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
If List1.SelCount = 0 Then MsgBox "�п�ܿ�X���I", vbOKOnly + vbInformation, Me.Caption: Exit Sub
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Screen.MousePointer = 11

'�������
Dim strSelected As String, i As Integer, strTmp As String
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
        If (List1.List(i) = "�t�m�c��" Or List1.List(i) = "�t�m�Ӽ�" Or List1.List(i) = "�t�m�`�Ӽ�" Or List1.List(i) = "�t�m�`���q" Or List1.List(i) = "�t�m�`���n") Then
            strSelected = strSelected & List1.List(i) & "= sum(" & List1.List(i) & "),"
        Else
            strSelected = strSelected & List1.List(i) & ","
        End If
    End If
Next
strSelected = Left(strSelected, Len(strSelected) - 1)

str_SQL = "select " & strSelected & " from gv_QueryShipDetail where 1 = 1 "

'�f�D�����
strTmp = ""
For i = 0 To Storerkey.ListCount - 1
    If Storerkey.Selected(i) Then
            strTmp = strTmp & "'" & Storerkey.List(i) & "',"
    End If
Next
If Len(strTmp) > 0 Then str_SQL = str_SQL & "and �f�D in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'���O�����
strTmp = ""
For i = 0 To OrderType.ListCount - 1
    If OrderType.Selected(i) Then
            strTmp = strTmp & "'" & OrderType.List(i) & "',"
    End If
Next
If Len(strTmp) > 0 Then str_SQL = str_SQL & "and �q�����O in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'�w�O�����
strTmp = ""
For i = 0 To WH.ListCount - 1
    If WH.Selected(i) Then
            strTmp = strTmp & "'" & WH.List(i) & "',"
    End If
Next
If Len(strTmp) > 0 Then str_SQL = str_SQL & "and �w�O in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'�ܧO�����
strTmp = ""
For i = 0 To Lot6.ListCount - 1
    If Lot6.Selected(i) Then
            strTmp = strTmp & "'" & Lot6.List(i) & "',"
    End If
Next
If Len(strTmp) > 0 Then str_SQL = str_SQL & "and �t�m�ܧO in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'�h�ƿ����
strTmp = ""
For i = 0 To Level.ListCount - 1
    If Level.Selected(i) Then
            strTmp = strTmp & "'" & Level.List(i) & "',"
    End If
Next
If Len(strTmp) > 0 Then str_SQL = str_SQL & "and �h�� in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'�x�Ͽ����
strTmp = ""
For i = 0 To Loc.ListCount - 1
    If Loc.Selected(i) Then
            strTmp = strTmp & "'" & Loc.List(i) & "',"
    End If
Next

If Len(strTmp) > 0 Then str_SQL = str_SQL & "and left(�x��,2) in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'�帹
If (Len(RTrim(Text3.Text)) > 0 And Len(RTrim(Text4.Text)) = 0) Or (Len(RTrim(Text3.Text)) = 0 And Len(RTrim(Text4.Text)) > 0) Then str_SQL = str_SQL & "and �t�m�帹 = '" & RTrim(Text3.Text) & RTrim(Text4.Text) & "' "
If (Len(RTrim(Text3.Text)) > 0 And Len(RTrim(Text4.Text)) > 0) Then str_SQL = str_SQL & "and �t�m�帹 between '" & RTrim(Text3.Text) & "'and'" & RTrim(Text4.Text) & "' "

'WMS�渹
If (Len(RTrim(Text5.Text)) > 0 And Len(RTrim(Text6.Text)) = 0) Or (Len(RTrim(Text5.Text)) = 0 And Len(RTrim(Text6.Text)) > 0) Then str_SQL = str_SQL & "and WMS�渹 = '" & RTrim(Text5.Text) & RTrim(Text6.Text) & "' "
If (Len(RTrim(Text5.Text)) > 0 And Len(RTrim(Text6.Text)) > 0) Then str_SQL = str_SQL & "and WMS�渹 between '" & RTrim(Text5.Text) & "'and'" & RTrim(Text6.Text) & "' "

'�~��
If (Len(RTrim(Text7.Text)) > 0 And Len(RTrim(Text8.Text)) = 0) Or (Len(RTrim(Text7.Text)) = 0 And Len(RTrim(Text8.Text)) > 0) Then str_SQL = str_SQL & "and �~�� = '" & RTrim(Text7.Text) & RTrim(Text8.Text) & "' "
If (Len(RTrim(Text7.Text)) > 0 And Len(RTrim(Text8.Text)) > 0) Then str_SQL = str_SQL & "and �~�� between '" & RTrim(Text7.Text) & "'and'" & RTrim(Text8.Text) & "' "

'�̪O�s��
If (Len(RTrim(Text9.Text)) > 0 And Len(RTrim(Text10.Text)) = 0) Or (Len(RTrim(Text9.Text)) = 0 And Len(RTrim(Text10.Text)) > 0) Then str_SQL = str_SQL & "and �̪O�s�� = '" & RTrim(Text9.Text) & RTrim(Text10.Text) & "' "
If (Len(RTrim(Text9.Text)) > 0 And Len(RTrim(Text10.Text)) > 0) Then str_SQL = str_SQL & "and �̪O�s�� between '" & RTrim(Text9.Text) & "'and'" & RTrim(Text10.Text) & "' "

'�J�w���
If (Len(RTrim(Text1.Text)) > 0 And Len(RTrim(Text2.Text)) = 0) Or (Len(RTrim(Text1.Text)) = 0 And Len(RTrim(Text2.Text)) > 0) Then str_SQL = str_SQL & "and ��f�� = '" & RTrim(Text1.Text) & RTrim(Text2.Text) & "' "
If (Len(RTrim(Text1.Text)) > 0 And Len(RTrim(Text2.Text)) > 0) Then str_SQL = str_SQL & "and ��f�� between '" & RTrim(Text1.Text) & "'and'" & RTrim(Text2.Text) & "' "

'�f�D�渹
If (Len(RTrim(Text13.Text)) > 0 And Len(RTrim(Text14.Text)) = 0) Or (Len(RTrim(Text13.Text)) = 0 And Len(RTrim(Text14.Text)) > 0) Then str_SQL = str_SQL & "and �f�D�渹 = '" & RTrim(Text13.Text) & RTrim(Text14.Text) & "' "
If (Len(RTrim(Text13.Text)) > 0 And Len(RTrim(Text14.Text)) > 0) Then str_SQL = str_SQL & "and �f�D�渹 between '" & RTrim(Text13.Text) & "'and'" & RTrim(Text14.Text) & "' "

'Group by ��
strSelected = ""
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
        If (List1.List(i) = "�t�m�c��" Or List1.List(i) = "�t�m�Ӽ�" Or List1.List(i) = "�t�m�`�Ӽ�" Or List1.List(i) = "�t�m�`���q" Or List1.List(i) = "�t�m�`���n") Then
        Else
            strSelected = strSelected & List1.List(i) & ","
        End If
    End If
Next
strSelected = Left(strSelected, Len(strSelected) - 1)

'Group by
str_SQL = str_SQL & "Group by " & strSelected

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "�~��"

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

    .ColumnHeaders = True        '���D�����
    .RowHeight = 300

End With

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub Combo1_Click()

Dim rsTmp As New ADODB.Recordset
With rsTmp
    .CursorLocation = 3
    .Open "select code , notes from codelkup where listname = 'QueryShipDetail' and code = '" & Left(Combo1.Text, InStr(Combo1.Text + " ", " ") - 1) & "'", cn
    If .EOF = True Then MsgBox "�䤣��`�Ϊ��Ѽ�( " & Left(Combo1.Text, InStr(Combo1.Text + " ", " ") - 1) & " )!!", vbOKOnly, Me.Caption: Exit Sub
    
    Text11 = mySplit(Combo1, " ", 0)
    Text12 = mySplit(Combo1, " ", 1)
    
    Dim arrTmp, arrTmp1, i As Integer, j As Integer, objTmp As Object
    arrTmp = Split(rsTmp("notes"), "|")
    
    arrTmp1 = Split(arrTmp(0), ",")
    '�f�D
    Set objTmp = Storerkey
    
    With objTmp
        For j = 0 To .ListCount - 1
            .ListIndex = j: .Selected(j) = False
            
            For i = 0 To UBound(arrTmp1)
                If Trim(.Text) = Trim(arrTmp1(i)) Then .Selected(j) = True: Exit For
            Next i
        Next j
    End With

    '�ܧO
    arrTmp1 = Split(arrTmp(1), ",")
    Set objTmp = Lot6
    
    With objTmp
        For j = 0 To .ListCount - 1
            .ListIndex = j: .Selected(j) = False
            
            For i = 0 To UBound(arrTmp1)
                If Trim(.Text) = Trim(arrTmp1(i)) Then .Selected(j) = True: Exit For
            Next i
        Next j
    End With

    '�h��
    arrTmp1 = Split(arrTmp(2), ",")
    Set objTmp = Level

    With objTmp
        For j = 0 To .ListCount - 1
            .ListIndex = j: .Selected(j) = False

            For i = 0 To UBound(arrTmp1)
                If Trim(.Text) = Trim(arrTmp1(i)) Then .Selected(j) = True: Exit For
            Next i
        Next j
    End With

    '�x��
    arrTmp1 = Split(arrTmp(3), ",")
    Set objTmp = Loc

    With objTmp
        For j = 0 To .ListCount - 1
            .ListIndex = j: .Selected(j) = False

            For i = 0 To UBound(arrTmp1)
                If Trim(.Text) = Trim(arrTmp1(i)) Then .Selected(j) = True: Exit For
            Next i
        Next j
    End With
    
    '���
    arrTmp1 = Split(arrTmp(4), ",")
    Set objTmp = List1
    
    With objTmp
        For j = 0 To .ListCount - 1
            .ListIndex = j: .Selected(j) = False
            
            For i = 0 To UBound(arrTmp1)
                If Trim(.Text) = Trim(arrTmp1(i)) Then .Selected(j) = True: Exit For
            Next i
        Next j
    End With
    
    '�w�O
    If UBound(arrTmp) < 4 Then
        arrTmp1 = Split(arrTmp(5), ",")
        Set objTmp = WH
        
        With objTmp
            For j = 0 To .ListCount - 1
                .ListIndex = j: .Selected(j) = False
                
                For i = 0 To UBound(arrTmp1)
                    If Trim(.Text) = Trim(arrTmp1(i)) Then .Selected(j) = True: Exit For
                Next i
            Next j
        End With
    End If
    .Close: Set rsTmp = Nothing
    
End With

End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdQuery_Click
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    dgMain.Height = Frame2.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth
    dgMain.Width = Frame2.Width - 240
End If

End Sub

Private Sub cmdReset_Click()
Dim i As Integer
'���]
Call ClearForm_AllField(Me)

'�f�D
For i = 0 To Storerkey.ListCount - 1
Storerkey.Selected(i) = False
Next

'�w�O
For i = 0 To WH.ListCount - 1
WH.Selected(i) = False
Next

'�ܧO
For i = 0 To Lot6.ListCount - 1
Lot6.Selected(i) = False
Next

'�h��
For i = 0 To Level.ListCount - 1
Level.Selected(i) = False
Next

'�x��
For i = 0 To Loc.ListCount - 1
Loc.Selected(i) = False
Next

'���
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next

End Sub

Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)

If dgMain.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMain.Sort = dgMain.Columns(ColIndex).Caption & " DESC"
    dgMain.ClearSelCols
    intColumnIndex = 255

Else
    rsMain.Sort = dgMain.Columns(ColIndex).Caption
    dgMain.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgmain_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub cmdExit_Click()
Unload Me '�������{��
'End �������ε{��
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

'���list
Dim rsTmp As New ADODB.Recordset
With rsTmp
    .CursorLocation = 3
    .Open "select * from gv_QueryShipDetail where 1 = 2", cn
    For i = 0 To .Fields.Count - 1
        List1.AddItem rsTmp(i).Name
    Next
    .Close
    
    .Open "select distinct storerkey from wms..orders order by storerkey ", cn
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        Storerkey.AddItem RTrim(rsTmp("storerkey"))
        rsTmp.MoveNext
    Loop
    .Close

    .Open "select distinct rtrim(o.type) as OrderType from wms..orders o ", cn
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        OrderType.AddItem RTrim(rsTmp("OrderType"))
        rsTmp.MoveNext
    Loop
    .Close

    WH.AddItem "�ըƹF�_��"
    WH.AddItem "�ըƹF����"
    WH.AddItem "�ըƹF�n��"
    
    .Open "select distinct lottable06 from wms..lotattribute l order by lottable06 ", cn
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        Lot6.AddItem RTrim(rsTmp("lottable06"))
        rsTmp.MoveNext
    Loop
    .Close
    
    .Open "select distinct loclevel from wms..loc order by loclevel ", cn
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        Level.AddItem RTrim(rsTmp("loclevel"))
        rsTmp.MoveNext
    Loop
    .Close
    
    .Open "select distinct left(loc,2) as loc  from wms..lotxloc order by loc", cn
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        Loc.AddItem rsTmp("loc")
        rsTmp.MoveNext
    Loop
    .Close
    
'�`�����
    .Open "select code , description from codelkup where listname = 'QueryShipDetail' order by code", cn
    If .EOF = False Then
    .MoveFirst
    For i = 0 To .RecordCount - 1
        Combo1.AddItem (RTrim(rsTmp("code")) & " " & RTrim(rsTmp("Description")))
        .MoveNext
    Next
    Combo1.ListIndex = 0
    End If
    .Close: Set rsTmp = Nothing
    
End With

SSTab1.Tab = 0
Text1 = Format(Now - 30, "YYYYMMDD")
Text2 = Format(Now, "YYYYMMDD")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub text1_Click()
Set objMvdateTarget = Text1
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub text2_Click()
Set objMvdateTarget = Text2
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
