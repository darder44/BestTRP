VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_SDNConfirm 
   Caption         =   "ñ��T�{"
   ClientHeight    =   8925
   ClientLeft      =   135
   ClientTop       =   975
   ClientWidth     =   14160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   14160
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   7200
      TabIndex        =   123
      Top             =   6480
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
      StartOfWeek     =   174456833
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "���`ñ��T�{"
      TabPicture(0)   =   "frm_OP_SDNConfirm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "�B�O���Ӭd��"
      TabPicture(1)   =   "frm_OP_SDNConfirm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame10"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frm_OP_SDNConfirm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape1(4)"
      Tab(2).Control(1)=   "Label3(23)"
      Tab(2).Control(2)=   "Label3(24)"
      Tab(2).Control(3)=   "Label3(25)"
      Tab(2).Control(4)=   "Label3(26)"
      Tab(2).Control(5)=   "Label3(35)"
      Tab(2).Control(6)=   "Frame5"
      Tab(2).Control(7)=   "Frame1"
      Tab(2).Control(8)=   "Frame6"
      Tab(2).Control(9)=   "txt_Tab02_C_Route_No"
      Tab(2).Control(10)=   "txt_Tab02_Receiver"
      Tab(2).Control(11)=   "txt_Tab02_Driver"
      Tab(2).Control(12)=   "txt_Tab02_Delivery_Date"
      Tab(2).Control(13)=   "txt_Tab02_C_VEHICLE_ID_NO"
      Tab(2).Control(14)=   "cmd_Tab2_SelectCar"
      Tab(2).Control(15)=   "Frame7"
      Tab(2).Control(16)=   "Frame8"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "ñ����ӽT�{"
      TabPicture(3)   =   "frm_OP_SDNConfirm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra_MultiOrder_Header"
      Tab(3).Control(1)=   "fra_MultiOrder_Detail"
      Tab(3).Control(2)=   "fra_OneOrder_Header"
      Tab(3).Control(3)=   "fra_Function"
      Tab(3).Control(4)=   "fra_OneOrder_Detail"
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame14 
         BackColor       =   &H80000004&
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   120
         TabIndex        =   234
         Top             =   360
         Width           =   10095
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
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
            Index           =   2
            Left            =   8880
            Picture         =   "frm_OP_SDNConfirm.frx":0070
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   251
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton cmdQueryT0 
            BackColor       =   &H00FFFFC0&
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
            Left            =   7680
            Picture         =   "frm_OP_SDNConfirm.frx":29C82
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   250
            Top             =   120
            Width           =   1065
         End
         Begin VB.ComboBox cboCarT0 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5280
            TabIndex        =   248
            Text            =   "cboCarT0"
            Top             =   600
            Width           =   2325
         End
         Begin VB.ComboBox cboStorerT0 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5280
            TabIndex        =   246
            Text            =   "cboStorerT0"
            Top             =   240
            Width           =   2325
         End
         Begin VB.TextBox txtDeliveryDateST0 
            Alignment       =   2  '�m�����
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
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   239
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtDeliveryDateET0 
            Alignment       =   2  '�m�����
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
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   238
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtRouteET0 
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
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   237
            Top             =   600
            Width           =   1365
         End
         Begin VB.TextBox txtRouteST0 
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   236
            Top             =   600
            Width           =   1365
         End
         Begin VB.CommandButton cmdPickSave 
            BackColor       =   &H00FFFF80&
            Caption         =   "�s��"
            Enabled         =   0   'False
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
            Left            =   4680
            Picture         =   "frm_OP_SDNConfirm.frx":29F8C
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   235
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1065
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
            Index           =   35
            Left            =   4560
            TabIndex        =   249
            Top             =   660
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
            Index           =   34
            Left            =   4560
            TabIndex        =   247
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   33
            Left            =   2535
            TabIndex        =   243
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   32
            Left            =   120
            TabIndex        =   242
            Top             =   285
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�G�����s"
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
            Left            =   120
            TabIndex        =   241
            Top             =   645
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Left            =   2535
            TabIndex        =   240
            Top             =   660
            Width           =   360
         End
      End
      Begin VB.Frame Frame13 
         Height          =   3255
         Left            =   120
         TabIndex        =   232
         Top             =   5040
         Width           =   8295
         Begin VB.TextBox txtCustomerOrderkey 
            BackColor       =   &H0000FFFF&
            Height          =   270
            Left            =   6240
            TabIndex        =   252
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdOpenOrderT0 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�˵�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1200
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   245
            Top             =   180
            Width           =   960
         End
         Begin VB.CommandButton cmdDeliveryokT0 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���`�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   244
            Top             =   180
            Width           =   960
         End
         Begin MSDataGridLib.DataGrid dgOrderT0 
            Height          =   2295
            Left            =   120
            TabIndex        =   233
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ֳt�T�{�ȭ��Q��B���_�B���L�S�B�Ȱ�...���f�D�I"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   36
            Left            =   2235
            TabIndex        =   264
            Top             =   300
            Visible         =   0   'False
            Width           =   5520
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Route"
         Height          =   3615
         Left            =   120
         TabIndex        =   230
         Top             =   1440
         Width           =   12495
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00FFFF80&
            Caption         =   "�R���p�O"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   298
            ToolTipText     =   "�R���p�O�\��"
            Top             =   2640
            Width           =   1065
         End
         Begin VB.CommandButton cmdRecalculate 
            BackColor       =   &H00FFFF80&
            Caption         =   "���s�p�O"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   297
            ToolTipText     =   "����p�O�\��"
            Top             =   2040
            Width           =   1065
         End
         Begin VB.CommandButton cmdPremiamAPnew 
            BackColor       =   &H0080FFFF&
            Caption         =   "ĳ�����I���u"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   281
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtPointT0 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1965
            TabIndex        =   277
            Top             =   3240
            Width           =   960
         End
         Begin VB.TextBox txtTransCubeT0 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   510
            TabIndex        =   276
            Top             =   3240
            Width           =   960
         End
         Begin VB.CommandButton cmdTransCube 
            BackColor       =   &H00FFFFC0&
            Caption         =   "���I�����(��)"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   275
            Top             =   1440
            Width           =   1035
         End
         Begin VB.TextBox txtTotalCubeT0 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   7350
            TabIndex        =   271
            Top             =   3240
            Width           =   1080
         End
         Begin VB.TextBox txtTotalPLT0 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4260
            TabIndex        =   267
            Top             =   3240
            Width           =   1080
         End
         Begin VB.TextBox txtTotalCST0 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   5790
            TabIndex        =   266
            Top             =   3240
            Width           =   1080
         End
         Begin VB.TextBox txtTotalWgtT0 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   8925
            TabIndex        =   265
            Top             =   3240
            Width           =   1080
         End
         Begin VB.CommandButton cmdTKPremiamAR 
            BackColor       =   &H00FFFF80&
            Caption         =   "TKĳ���������u"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   254
            ToolTipText     =   "�Ȱw��x�W�Q��"
            Top             =   840
            Width           =   1065
         End
         Begin VB.CommandButton cmdPremiamAP 
            BackColor       =   &H0080FFFF&
            Caption         =   "ĳ�����I���u"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   6000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   253
            Top             =   960
            Visible         =   0   'False
            Width           =   1065
         End
         Begin MSDataGridLib.DataGrid dgRouteT0 
            Height          =   2295
            Left            =   1260
            TabIndex        =   231
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�I��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   39
            Left            =   1590
            TabIndex        =   279
            Top             =   3285
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   38
            Left            =   120
            TabIndex        =   278
            Top             =   3285
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   37
            Left            =   6960
            TabIndex        =   272
            Top             =   3285
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���q"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   8
            Left            =   9150
            TabIndex        =   270
            Top             =   3285
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�c��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   9
            Left            =   5400
            TabIndex        =   269
            Top             =   3285
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�w����G�O��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   11
            Left            =   3120
            TabIndex        =   268
            Top             =   3285
            Width           =   1080
         End
      End
      Begin VB.Frame Frame10 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   162
         Top             =   3000
         Width           =   8295
         Begin MSDataGridLib.DataGrid dgMain1 
            Height          =   2295
            Left            =   120
            TabIndex        =   163
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
      Begin VB.Frame Frame9 
         BackColor       =   &H80000004&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   -74880
         TabIndex        =   148
         Top             =   360
         Width           =   13935
         Begin VB.ListBox Car_Num 
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
            ItemData        =   "frm_OP_SDNConfirm.frx":2A296
            Left            =   7920
            List            =   "frm_OP_SDNConfirm.frx":2A298
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   280
            ToolTipText     =   "���x��=�w�s-�w�t-�w�z = 0"
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox cboStorerkey 
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
            Left            =   5640
            TabIndex        =   192
            Top             =   240
            Width           =   1605
         End
         Begin VB.TextBox txt2RouteS 
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   189
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txt2RouteE 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   188
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtEarning 
            Alignment       =   1  '�a�k���
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   186
            Text            =   "0"
            Top             =   2160
            Width           =   1125
         End
         Begin VB.TextBox txtAR 
            Alignment       =   1  '�a�k���
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   184
            Text            =   "0"
            Top             =   1680
            Width           =   1125
         End
         Begin VB.TextBox txtAP 
            Alignment       =   1  '�a�k���
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   182
            Text            =   "0"
            Top             =   1920
            Width           =   1125
         End
         Begin VB.ComboBox cboCostkind 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            TabIndex        =   178
            Top             =   600
            Width           =   1605
         End
         Begin VB.TextBox txtSignDateE 
            Alignment       =   2  '�m�����
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
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   175
            Top             =   2040
            Width           =   1485
         End
         Begin VB.TextBox txtSignDateS 
            Alignment       =   2  '�m�����
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
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   174
            Top             =   2040
            Width           =   1485
         End
         Begin VB.ComboBox cboCar 
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
            Left            =   5640
            TabIndex        =   172
            Top             =   960
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox txtRouteE 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   169
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox txtRouteS 
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   168
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryE 
            Alignment       =   2  '�m�����
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
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   165
            Top             =   1680
            Width           =   1485
         End
         Begin VB.TextBox txtDeliveryS 
            Alignment       =   2  '�m�����
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
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   164
            Top             =   1680
            Width           =   1485
         End
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H00FFFFC0&
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
            Left            =   10320
            Picture         =   "frm_OP_SDNConfirm.frx":2A29A
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   157
            Top             =   240
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
            Left            =   12720
            Picture         =   "frm_OP_SDNConfirm.frx":2A5A4
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   156
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
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
            Index           =   0
            Left            =   12720
            Picture         =   "frm_OP_SDNConfirm.frx":2A8B6
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   155
            Top             =   1200
            Width           =   1065
         End
         Begin VB.CommandButton cmd2Excel 
            BackColor       =   &H00FFC0C0&
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
            Left            =   11520
            Picture         =   "frm_OP_SDNConfirm.frx":544C8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   154
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtExternS 
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
            Left            =   1200
            TabIndex        =   153
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox txtExternE 
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
            Left            =   3000
            TabIndex        =   152
            Top             =   960
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveToText 
            BackColor       =   &H00C0E0FF&
            Caption         =   "�|�p���"
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
            Left            =   11520
            Picture         =   "frm_OP_SDNConfirm.frx":557C2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   151
            Top             =   1200
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtOrderkeyE 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   150
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtOrderkeyS 
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   149
            Top             =   1320
            Width           =   1485
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
            Index           =   29
            Left            =   4800
            TabIndex        =   193
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   28
            Left            =   2640
            TabIndex        =   191
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�G�����s"
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
            Left            =   120
            TabIndex        =   190
            Top             =   285
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�禬"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   26
            Left            =   4530
            TabIndex        =   187
            Top             =   2220
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   25
            Left            =   4530
            TabIndex        =   185
            Top             =   1740
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���I"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   24
            Left            =   4530
            TabIndex        =   183
            Top             =   1980
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�д����O"
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
            Left            =   4560
            TabIndex        =   179
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   22
            Left            =   120
            TabIndex        =   177
            Top             =   1380
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   21
            Left            =   2640
            TabIndex        =   176
            Top             =   2100
            Width           =   360
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
            Index           =   6
            Left            =   7320
            TabIndex        =   173
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   20
            Left            =   120
            TabIndex        =   171
            Top             =   645
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   19
            Left            =   2640
            TabIndex        =   170
            Top             =   660
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ñ����"
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
            Left            =   120
            TabIndex        =   167
            Top             =   2085
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Left            =   2640
            TabIndex        =   166
            Top             =   1740
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   17
            Left            =   2655
            TabIndex        =   161
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Index           =   16
            Left            =   120
            TabIndex        =   160
            Top             =   1005
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
            Index           =   0
            Left            =   120
            TabIndex        =   159
            Top             =   1740
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
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
            Left            =   2640
            TabIndex        =   158
            Top             =   1380
            Width           =   360
         End
      End
      Begin VB.Frame fra_OneOrder_Detail 
         Appearance      =   0  '����
         BackColor       =   &H00404000&
         ForeColor       =   &H80000008&
         Height          =   3060
         Left            =   -74880
         TabIndex        =   26
         Top             =   5640
         Width           =   11160
         Begin VB.TextBox txt_OneOrder_SignQty 
            BackColor       =   &H0000FFFF&
            Height          =   270
            Left            =   1500
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1515
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cmb_OneOrder_RSCCode 
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNConfirm.frx":55ACC
            Left            =   1530
            List            =   "frm_OP_SDNConfirm.frx":55ACE
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   28
            Top             =   2355
            Visible         =   0   'False
            Width           =   2490
         End
         Begin VB.ComboBox cmb_OneOrder_RBCCode 
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNConfirm.frx":55AD0
            Left            =   1500
            List            =   "frm_OP_SDNConfirm.frx":55AD2
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   27
            Top             =   1995
            Visible         =   0   'False
            Width           =   2340
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gd_OneOrder_OrderDetail 
            Height          =   3270
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   5768
            _Version        =   393216
            ScrollBars      =   2
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "�p�O����"
         Height          =   2535
         Left            =   -74280
         TabIndex        =   120
         Top             =   6120
         Width           =   12975
         Begin VB.TextBox Text4 
            BackColor       =   &H0000FFFF&
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
            Left            =   960
            TabIndex        =   122
            Top             =   840
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab2_DelCost 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�R���p�OCTrl+D"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   114
            Top             =   960
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab2_AddCost 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�s�W�p�OCTrl+A"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   113
            Top             =   240
            Width           =   1035
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab2_SDN_Cost 
            Height          =   2145
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   3784
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Cols            =   14
            FixedCols       =   0
            BackColorSel    =   10354595
            ForeColorSel    =   8454016
            BackColorBkg    =   -2147483644
            AllowBigSelection=   0   'False
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   14
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "�q��˸����e"
         Height          =   3135
         Left            =   -74280
         TabIndex        =   119
         Top             =   2280
         Width           =   12975
         Begin VB.TextBox Text3 
            BackColor       =   &H0000FFFF&
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
            Left            =   840
            TabIndex        =   121
            Top             =   960
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab2_DelOrder 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�R���q�� Ctrl+D"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   111
            Top             =   960
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Tab2_AddOrder 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�s�W�q��CTrl+A"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11760
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   110
            Top             =   240
            Width           =   1035
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab2_SDN_Detail 
            Height          =   2760
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   4868
            _Version        =   393216
            Cols            =   14
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   14
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
      End
      Begin VB.CommandButton cmd_Tab2_SelectCar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�H"
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
         Height          =   255
         Left            =   -69720
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   105
         Top             =   1680
         Width           =   330
      End
      Begin VB.TextBox txt_Tab02_C_VEHICLE_ID_NO 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -70680
         TabIndex        =   104
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_Tab02_Delivery_Date 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -72960
         TabIndex        =   103
         Top             =   1680
         Width           =   1200
      End
      Begin VB.TextBox txt_Tab02_Driver 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -68160
         TabIndex        =   106
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_Tab02_Receiver 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -65880
         TabIndex        =   108
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_Tab02_C_Route_No 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   -63360
         TabIndex        =   102
         Top             =   1680
         Width           =   1320
      End
      Begin VB.Frame Frame6 
         Height          =   525
         Left            =   -74280
         TabIndex        =   92
         Top             =   5520
         Width           =   5355
         Begin VB.OptionButton Op_Tab2_WT 
            Caption         =   "Option1"
            Height          =   255
            Left            =   4920
            TabIndex        =   98
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_CBM 
            Caption         =   "Option1"
            Height          =   255
            Left            =   3360
            TabIndex        =   97
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_CS 
            Caption         =   "Option1"
            Height          =   255
            Left            =   1800
            TabIndex        =   96
            Top             =   210
            Width           =   255
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Case 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   975
            TabIndex        =   95
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Volumn 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   2505
            TabIndex        =   94
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Weight 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4080
            TabIndex        =   93
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���q"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   13
            Left            =   3705
            TabIndex        =   101
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   14
            Left            =   2115
            TabIndex        =   100
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�`�p�G�c��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   15
            Left            =   75
            TabIndex        =   99
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '����
         BackColor       =   &H00004000&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -74280
         TabIndex        =   85
         Top             =   480
         Width           =   12960
         Begin VB.CommandButton cmd_Tab2_AddNew 
            BackColor       =   &H00C0FFC0&
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
            Height          =   450
            Left            =   4560
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   91
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Modify 
            BackColor       =   &H00C0E0FF&
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
            Height          =   450
            Left            =   3360
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   90
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Save 
            BackColor       =   &H00C0C0FF&
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
            Height          =   450
            Left            =   5760
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   89
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Delete 
            BackColor       =   &H000080FF&
            Caption         =   "�R  ��"
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
            Height          =   450
            Left            =   8160
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   88
            Top             =   195
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
            Height          =   450
            Index           =   1
            Left            =   9360
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   87
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Cancel 
            BackColor       =   &H00C0FFFF&
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
            Height          =   450
            Left            =   6960
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   86
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         Height          =   525
         Left            =   -66660
         TabIndex        =   75
         Top             =   5520
         Width           =   5355
         Begin VB.OptionButton Op_Tab2_SumWT 
            Caption         =   "Option1"
            Height          =   255
            Left            =   5040
            TabIndex        =   81
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_SumCBM 
            Caption         =   "Option1"
            Height          =   255
            Left            =   3480
            TabIndex        =   80
            Top             =   210
            Width           =   255
         End
         Begin VB.OptionButton Op_Tab2_SumCS 
            Caption         =   "Option1"
            Height          =   255
            Left            =   1920
            TabIndex        =   79
            Top             =   210
            Width           =   255
         End
         Begin VB.TextBox txt_Tab2_sum_Case 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1095
            TabIndex        =   78
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_sum_CBM 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   2625
            TabIndex        =   77
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_sum_WT 
            Alignment       =   2  '�m�����
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4200
            TabIndex        =   76
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���q"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   12
            Left            =   3825
            TabIndex        =   84
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���n"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   10
            Left            =   2235
            TabIndex        =   83
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p�p�G�c��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   7
            Left            =   195
            TabIndex        =   82
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Frame fra_Function 
         Height          =   615
         Left            =   -74880
         TabIndex        =   50
         Top             =   360
         Width           =   11895
         Begin VB.CommandButton cmdUnRouteConfirm 
            BackColor       =   &H00FFFFC0&
            Caption         =   "     ����     �X���T�{"
            Enabled         =   0   'False
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
            Left            =   8640
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmdShipNotes 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�ɦL�X�f��"
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
            Left            =   9360
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   194
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdCost 
            BackColor       =   &H00FFFFC0&
            Caption         =   "�B�O���@"
            Enabled         =   0   'False
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
            Left            =   8040
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   6
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmdCarNOChange 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�����ܧ�"
            Enabled         =   0   'False
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
            Left            =   6720
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   4
            Top             =   120
            Width           =   1245
         End
         Begin VB.ComboBox cmbOrderkey 
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
            ItemData        =   "frm_OP_SDNConfirm.frx":55AD4
            Left            =   120
            List            =   "frm_OP_SDNConfirm.frx":55AD6
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   0
            Top             =   165
            Width           =   1455
         End
         Begin VB.CommandButton cmdNotYetOrder 
            BackColor       =   &H00C0FFFF&
            Caption         =   "�ݽT�{ñ��"
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
            Left            =   5400
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   3
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FF80FF&
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
            Height          =   450
            Index           =   0
            Left            =   10680
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   7
            Top             =   120
            Width           =   1110
         End
         Begin VB.CommandButton cmd_OrderQuery 
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
            Height          =   465
            Left            =   4320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   2
            Top             =   120
            Width           =   1005
         End
         Begin VB.TextBox txt_OrderKey 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   1
            Top             =   165
            Width           =   2745
         End
      End
      Begin VB.Frame fra_OneOrder_Header 
         Appearance      =   0  '����
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4665
         Left            =   -74880
         TabIndex        =   30
         Top             =   960
         Width           =   11895
         Begin VB.TextBox txt_Cartype 
            Enabled         =   0   'False
            Height          =   270
            Left            =   9240
            TabIndex        =   295
            Text            =   "txt_Cartype"
            Top             =   4320
            Width           =   975
         End
         Begin VB.TextBox txt_ReserveMark 
            Height          =   270
            Left            =   4920
            TabIndex        =   294
            Text            =   "txt_ReserveMark"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txt_UrgentMark 
            Height          =   270
            Left            =   3720
            TabIndex        =   293
            Text            =   "txt_UrgentMark"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txt_Stairs 
            Alignment       =   1  '�a�k���
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
            Left            =   10560
            TabIndex        =   292
            ToolTipText     =   "���_�W�U��"
            Top             =   4320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cbx_B_city 
            Height          =   300
            ItemData        =   "frm_OP_SDNConfirm.frx":55AD8
            Left            =   10680
            List            =   "frm_OP_SDNConfirm.frx":55AE5
            TabIndex        =   291
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txt_ReceiveCash 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   288
            ToolTipText     =   "��ڦ��{���B"
            Top             =   3420
            Width           =   1335
         End
         Begin VB.TextBox txt_Cash 
            Alignment       =   1  '�a�k���
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
            Left            =   7680
            TabIndex        =   286
            ToolTipText     =   "�U�f���{���B"
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txt_Externordertype 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   284
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txt_BranchId 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   282
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txt_C_Receipt_No 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   274
            Top             =   2040
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton cmdSDNBack 
            BackColor       =   &H0000FFFF&
            Caption         =   "ñ�檬�A"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   9480
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   263
            Top             =   2400
            Width           =   645
         End
         Begin VB.TextBox txt_OneOrder_ConsigneeKey1 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   262
            Top             =   1440
            Width           =   1080
         End
         Begin VB.TextBox txt_OneOrder_FullName1 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   261
            Top             =   1440
            Width           =   4680
         End
         Begin VB.TextBox txt_OneOrder_OrderKey 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   260
            Top             =   240
            Width           =   1080
         End
         Begin VB.TextBox txt_Storer 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   259
            Top             =   240
            Width           =   2040
         End
         Begin VB.TextBox txt_OneOrder_Address1 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   257
            Top             =   1740
            Width           =   4680
         End
         Begin VB.TextBox txt_Zip1 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   256
            Top             =   1740
            Width           =   1080
         End
         Begin VB.CommandButton cmdReceiptDetail 
            BackColor       =   &H00C0C0C0&
            Caption         =   "���f����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9240
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   255
            Top             =   3120
            Width           =   1005
         End
         Begin VB.ComboBox cboInvBack 
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
            ItemData        =   "frm_OP_SDNConfirm.frx":55B01
            Left            =   9240
            List            =   "frm_OP_SDNConfirm.frx":55B03
            TabIndex        =   20
            Top             =   3675
            Width           =   855
         End
         Begin VB.TextBox txt_SDNNote 
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
            Left            =   1680
            TabIndex        =   21
            ToolTipText     =   "�t�e���`�ݸԭz���`�o�ͭ�]"
            Top             =   4300
            Width           =   4815
         End
         Begin VB.TextBox txt_C_ROUTE_NO 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   146
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txt_ZIP 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   1140
            Width           =   1080
         End
         Begin VB.TextBox txt_OneOrder_DeliveryDate 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   143
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_RouteNo 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   141
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_StorerKey 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   139
            Top             =   240
            Width           =   1080
         End
         Begin VB.TextBox txt_Priority 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   540
            Width           =   1080
         End
         Begin VB.TextBox txt_TRPHandle 
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
            Left            =   1680
            TabIndex        =   14
            Top             =   3420
            Width           =   4815
         End
         Begin VB.TextBox txt_SortingCost 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   18
            ToolTipText     =   "�t�e���`�ҭl�ͥX���z�f�O"
            Top             =   4020
            Width           =   1335
         End
         Begin VB.TextBox txt_CustHandle 
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
            Left            =   1680
            TabIndex        =   13
            Top             =   3120
            Width           =   4815
         End
         Begin VB.TextBox txt_TRPCost 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   17
            ToolTipText     =   "�t�e���`�ҭl�ͥX���t�e�O"
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox txt_INVHandle 
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
            Left            =   1680
            TabIndex        =   16
            Top             =   4020
            Width           =   4815
         End
         Begin VB.TextBox txt_TotalCost 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7680
            TabIndex        =   19
            ToolTipText     =   "�t�e���`�ҭl�ͥX���O�ΦX�p"
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox txt_Advance 
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
            Left            =   1680
            TabIndex        =   15
            Top             =   3720
            Width           =   4815
         End
         Begin VB.TextBox txt_OneOrder_StorerOrderKey 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   128
            Top             =   240
            Width           =   1680
         End
         Begin VB.ComboBox cmbScan 
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
            Left            =   5400
            TabIndex        =   10
            Text            =   "cmbScan"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txt_OneOrder_CustomerOrderkey1 
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
            Left            =   7560
            TabIndex        =   11
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_Status 
            BackColor       =   &H80000000&
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
            Left            =   10680
            TabIndex        =   40
            ToolTipText     =   "���⬰���;��⬰�F��"
            Top             =   1750
            Width           =   1080
         End
         Begin VB.CommandButton cmd_OneOrder_Expect 
            BackColor       =   &H000080FF&
            Caption         =   "���`�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   22
            Top             =   3120
            Width           =   1485
         End
         Begin VB.CommandButton cmd_OneOrder_Deliveryok 
            BackColor       =   &H00FF8080&
            Caption         =   "���`�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   12
            Top             =   2400
            Width           =   1485
         End
         Begin VB.CommandButton cmd_OneOrder_NoDelivery 
            BackColor       =   &H008080FF&
            Caption         =   "���X�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   23
            ToolTipText     =   "�I��""���X�q��""�A�t�αN���_�p�O"
            Top             =   3720
            Width           =   1485
         End
         Begin VB.TextBox txt_OneOrder_Address 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1140
            Width           =   4680
         End
         Begin VB.TextBox txt_OneOrder_FullName 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   840
            Width           =   4680
         End
         Begin VB.TextBox txt_OneOrder_TRPCompany 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_Driver 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1740
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_VehicleID 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txt_OneOrder_OrderDate 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1150
            Width           =   1080
         End
         Begin VB.TextBox txt_OneOrder_ConsigneeKey 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   840
            Width           =   1080
         End
         Begin VB.TextBox txt_OneOrder_Description 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   540
            Width           =   5760
         End
         Begin VB.TextBox txt_OneOrder_ArriveDate 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1450
            Width           =   1080
         End
         Begin MSComCtl2.DTPicker dtp_OneOrder_SignDate 
            Height          =   375
            Left            =   960
            TabIndex        =   8
            Top             =   2520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   -2147483643
            CustomFormat    =   "yyyy/MM/dd HH:mm"
            Format          =   174129155
            UpDown          =   -1  'True
            CurrentDate     =   39438
         End
         Begin MSComCtl2.DTPicker dtpSDNSendDate 
            Height          =   375
            Left            =   3600
            TabIndex        =   9
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   174129155
            UpDown          =   -1  'True
            CurrentDate     =   39438
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p�O�N�X"
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
            Index           =   36
            Left            =   9240
            TabIndex        =   296
            Top             =   4080
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ճ�����"
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
            Index           =   34
            Left            =   9840
            TabIndex        =   290
            Top             =   2085
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ꦬ�f��"
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
            Index           =   33
            Left            =   6720
            TabIndex        =   289
            Top             =   3465
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�����f��"
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
            Index           =   32
            Left            =   6720
            TabIndex        =   287
            Top             =   3180
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�q�����O"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   9480
            TabIndex        =   285
            Top             =   885
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�����q"
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
            Left            =   240
            TabIndex        =   283
            Top             =   2100
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�_�I"
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
            Height          =   195
            Index           =   30
            Left            =   3240
            TabIndex        =   273
            Top             =   1140
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���I"
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
            Height          =   195
            Index           =   27
            Left            =   3240
            TabIndex        =   258
            Top             =   1680
            UseMnemonic     =   0   'False
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�o���^��"
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
            Left            =   9240
            TabIndex        =   181
            Top             =   3480
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ñ��Ƶ�"
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
            Left            =   120
            TabIndex        =   180
            Top             =   4380
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�G�����s"
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
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   147
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
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
            Index           =   9
            Left            =   120
            TabIndex        =   144
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���u�s��"
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
            Left            =   120
            TabIndex        =   142
            Top             =   600
            Width           =   780
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
            Left            =   2880
            TabIndex        =   140
            Top             =   300
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
            Index           =   54
            Left            =   9840
            TabIndex        =   138
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "TMS�渹"
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
            Index           =   53
            Left            =   9840
            TabIndex        =   136
            Top             =   300
            Width           =   765
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   240
            X2              =   11640
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����B�z�覡"
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
            Index           =   47
            Left            =   120
            TabIndex        =   135
            Top             =   3480
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���`�z�f�O"
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
            Index           =   51
            Left            =   6720
            TabIndex        =   134
            Top             =   4080
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�^�гB�z�覡"
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
            Index           =   46
            Left            =   120
            TabIndex        =   133
            Top             =   3180
            Width           =   1560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���`�t�e�O"
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
            Index           =   50
            Left            =   6720
            TabIndex        =   132
            Top             =   3780
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�w�s�վ�覡"
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
            Index           =   52
            Left            =   120
            TabIndex        =   131
            Top             =   4080
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�O�ΦX�p"
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
            Index           =   49
            Left            =   6720
            TabIndex        =   130
            Top             =   4380
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ﵽ�覡"
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
            Index           =   48
            Left            =   120
            TabIndex        =   129
            Top             =   3780
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q����"
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
            Left            =   9840
            TabIndex        =   44
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���y"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   45
            Left            =   4920
            TabIndex        =   127
            Top             =   2580
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ñ��^��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   10
            Left            =   2760
            TabIndex        =   126
            Top             =   2580
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�禬�渹"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   44
            Left            =   6720
            TabIndex        =   125
            Top             =   2580
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ñ�檬�A"
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
            Index           =   11
            Left            =   9840
            TabIndex        =   49
            Top             =   1800
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�ñ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   48
            Top             =   2580
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�餽�q"
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
            Index           =   8
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�r�p�m�W"
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
            Index           =   7
            Left            =   120
            TabIndex        =   46
            Top             =   1800
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e����"
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
            Index           =   6
            Left            =   120
            TabIndex        =   45
            Top             =   1500
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��f���"
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
            Left            =   9840
            TabIndex        =   43
            Top             =   1500
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
            Left            =   2880
            TabIndex        =   42
            Top             =   600
            Width           =   780
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
            Index           =   13
            Left            =   6960
            TabIndex        =   41
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.Frame fra_MultiOrder_Detail 
         Appearance      =   0  '����
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   3780
         Left            =   -74880
         TabIndex        =   51
         Top             =   3240
         Visible         =   0   'False
         Width           =   10920
         Begin VB.ComboBox cmb_MultiOrder_RBCCode 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNConfirm.frx":55B05
            Left            =   1545
            List            =   "frm_OP_SDNConfirm.frx":55B07
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   54
            Top             =   2010
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.ComboBox cmb_MultiOrder_RSCCode 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   300
            ItemData        =   "frm_OP_SDNConfirm.frx":55B09
            Left            =   1575
            List            =   "frm_OP_SDNConfirm.frx":55B0B
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   53
            Top             =   2370
            Visible         =   0   'False
            Width           =   2490
         End
         Begin VB.TextBox txt_MultiOrder_SignQty 
            Alignment       =   1  '�a�k���
            Height          =   270
            Left            =   1545
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1770
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gd_MultiOrder_OrderDetail 
            Height          =   3510
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   6191
            _Version        =   393216
            FixedCols       =   0
            ScrollBars      =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fra_MultiOrder_Header 
         Appearance      =   0  '����
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   -75000
         TabIndex        =   56
         Top             =   4080
         Visible         =   0   'False
         Width           =   10935
         Begin VB.CommandButton cmd_MultiOrder_NoDelivery 
            BackColor       =   &H008080FF&
            Caption         =   "���X�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   7635
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   67
            Top             =   1470
            Width           =   1485
         End
         Begin VB.CommandButton cmd_MultiOrder_Deliveryok 
            BackColor       =   &H00FF8080&
            Caption         =   "���`�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5940
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   66
            Top             =   1470
            Width           =   1485
         End
         Begin VB.CommandButton cmd_MultiOrder_Expect 
            BackColor       =   &H000080FF&
            Caption         =   "���`�T�{"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   9330
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   65
            Top             =   1470
            Width           =   1485
         End
         Begin VB.TextBox txt_MultiOrder_SignDate 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            Left            =   5400
            TabIndex        =   64
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txt_MultiOrder_Status 
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
            Left            =   6735
            TabIndex        =   63
            Top             =   450
            Width           =   1215
         End
         Begin VB.TextBox txt_MultiOrder_ArriveDate 
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
            Left            =   9375
            TabIndex        =   62
            Top             =   450
            Width           =   1215
         End
         Begin VB.TextBox txt_MultiOrder_ConsigneeKey 
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
            Left            =   885
            TabIndex        =   61
            Top             =   135
            Width           =   1575
         End
         Begin VB.TextBox txt_MultiOrder_OrderDate 
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
            Left            =   9375
            TabIndex        =   60
            Top             =   150
            Width           =   1215
         End
         Begin VB.TextBox txt_MultiOrder_StorerKey 
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
            Left            =   6735
            TabIndex        =   59
            Top             =   150
            Width           =   1020
         End
         Begin VB.TextBox txt_MultiOrder_FullName 
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
            Left            =   2490
            TabIndex        =   58
            Top             =   135
            Width           =   2775
         End
         Begin VB.TextBox txt_MultiOrder_Address 
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
            Left            =   885
            TabIndex        =   57
            Top             =   435
            Width           =   4395
         End
         Begin MSDataGridLib.DataGrid dg_MultiOrder 
            Height          =   1185
            Left            =   60
            TabIndex        =   68
            Top             =   840
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   2090
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
         Begin MSComCtl2.DTPicker dtp_MultiOrder_SignDate 
            Height          =   375
            Left            =   7560
            TabIndex        =   124
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
            Format          =   173932547
            UpDown          =   -1  'True
            CurrentDate     =   39438
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�ñ�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   18
            Left            =   5940
            TabIndex        =   74
            Top             =   990
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���A"
            Height          =   180
            Index           =   17
            Left            =   6345
            TabIndex        =   73
            Top             =   525
            Width           =   360
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
            Index           =   16
            Left            =   90
            TabIndex        =   72
            Top             =   195
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�e�f��"
            Height          =   180
            Index           =   15
            Left            =   8805
            TabIndex        =   71
            Top             =   510
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q���"
            Height          =   180
            Index           =   14
            Left            =   8805
            TabIndex        =   70
            Top             =   225
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D"
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
            Index           =   12
            Left            =   6330
            TabIndex        =   69
            Top             =   210
            Width           =   390
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   5895
         Left            =   3120
         TabIndex        =   195
         Top             =   1920
         Visible         =   0   'False
         Width           =   13695
         Begin VB.TextBox Text1 
            BackColor       =   &H0000FFFF&
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
            Left            =   1800
            TabIndex        =   229
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Frame Frame3 
            Caption         =   "�p�p"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   525
            Left            =   8880
            TabIndex        =   219
            Top             =   3360
            Width           =   4875
            Begin VB.TextBox txt_Tab0_sum_WT 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   3720
               TabIndex        =   225
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_sum_CBM 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   2145
               TabIndex        =   224
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_sum_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   615
               TabIndex        =   223
               Top             =   165
               Width           =   840
            End
            Begin VB.OptionButton Op_SumCS_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   1440
               TabIndex        =   222
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton Op_SumCBM_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   3000
               TabIndex        =   221
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton Op_SumWT_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   4560
               TabIndex        =   220
               Top             =   210
               Width           =   255
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   3
               Left            =   195
               TabIndex        =   228
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   2
               Left            =   1755
               TabIndex        =   227
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3345
               TabIndex        =   226
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame Frame2 
            Height          =   525
            Left            =   3480
            TabIndex        =   213
            Top             =   3720
            Width           =   7635
            Begin VB.OptionButton Op_CS_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   4200
               TabIndex        =   216
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton Op_CBM_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   5760
               TabIndex        =   215
               Top             =   210
               Width           =   255
            End
            Begin VB.OptionButton OpWT_del 
               Caption         =   "Option1"
               Height          =   255
               Left            =   7320
               TabIndex        =   214
               Top             =   210
               Width           =   255
            End
            Begin VB.Label Lb_Route 
               Caption         =   "Route"
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
               Height          =   165
               Left            =   1200
               TabIndex        =   217
               Top             =   210
               Width           =   1215
            End
         End
         Begin VB.TextBox Text2_del 
            BackColor       =   &H0000FFFF&
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
            Left            =   1920
            TabIndex        =   212
            Top             =   4800
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab0_SumQty 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   10080
            TabIndex        =   211
            Top             =   2880
            Width           =   1005
         End
         Begin VB.Frame Frame11 
            Height          =   1005
            Left            =   3960
            TabIndex        =   196
            Top             =   240
            Width           =   8085
            Begin VB.TextBox txt_DeliveryDate_End 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2445
               TabIndex        =   205
               Top             =   630
               Width           =   1125
            End
            Begin VB.TextBox txt_DeliveryDate_Start 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1050
               TabIndex        =   204
               Top             =   630
               Width           =   1125
            End
            Begin VB.TextBox txt_RouteNo_Start 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1050
               TabIndex        =   203
               Top             =   240
               Width           =   1125
            End
            Begin VB.TextBox txt_RouteNo_End 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2445
               TabIndex        =   202
               Top             =   240
               Width           =   1125
            End
            Begin VB.CheckBox ck_confirm 
               Caption         =   "���T�{ñ��"
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
               Left            =   3720
               TabIndex        =   201
               Top             =   285
               Width           =   1455
            End
            Begin VB.OptionButton Op_UnCheck 
               Caption         =   "����z"
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
               Height          =   345
               Left            =   7080
               TabIndex        =   200
               Top             =   210
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CheckBox ck_back 
               Caption         =   "���^��ñ��"
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
               Left            =   3720
               TabIndex        =   199
               Top             =   675
               Width           =   1455
            End
            Begin VB.OptionButton Op_OnCheck 
               Caption         =   "�w��z"
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
               Height          =   255
               Left            =   7080
               TabIndex        =   198
               Top             =   645
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txt_Tab0_C_VEHICLE_ID_NO 
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5880
               TabIndex        =   197
               Top             =   240
               Width           =   1125
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
               Index           =   19
               Left            =   2205
               TabIndex        =   210
               Top             =   555
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�X�����"
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
               Index           =   20
               Left            =   150
               TabIndex        =   209
               Top             =   675
               Width           =   840
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
               Index           =   21
               Left            =   2205
               TabIndex        =   208
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���u�s��"
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
               Index           =   22
               Left            =   150
               TabIndex        =   207
               Top             =   285
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "��   ��"
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
               Index           =   39
               Left            =   5190
               TabIndex        =   206
               Top             =   285
               Width           =   600
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_SDN_Head_del 
            Height          =   1920
            Left            =   120
            TabIndex        =   218
            Top             =   1320
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   3387
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Cols            =   14
            FixedCols       =   0
            BackColorSel    =   10354595
            ForeColorSel    =   8454016
            BackColorBkg    =   -2147483644
            AllowBigSelection=   0   'False
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   14
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�����G"
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
         Index           =   35
         Left            =   -71340
         TabIndex        =   118
         Top             =   1725
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�e�f��G"
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
         Index           =   26
         Left            =   -73800
         TabIndex        =   117
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�r�p�H�G"
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
         Index           =   25
         Left            =   -69000
         TabIndex        =   116
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��ڤH�G"
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
         Index           =   24
         Left            =   -66720
         TabIndex        =   115
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���u�s���G"
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
         Index           =   23
         Left            =   -64440
         TabIndex        =   107
         Top             =   1725
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '���z��
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Height          =   735
         Index           =   4
         Left            =   -74280
         Top             =   1440
         Width           =   12930
      End
   End
End
Attribute VB_Name = "frm_OP_SDNConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private iLoop As Double              '�j��p��

Private blShipped As Boolean         '�O�_�w�g����L Shipped Confirm
Private blSDNConfirm As Boolean      '�O�_�w�g����L SDN Confirm
Private blCanUpdate As Boolean       '�O�_�i�H���� SDN Confirm
Private blRouteT0Change As Boolean   '�O�_������Ӭd��

Private rs_MultiOrder As ADODB.Recordset
Private rs_Tab1_SDN05T As ADODB.Recordset
Private rs_cost As ADODB.Recordset
Private rs_cust As ADODB.Recordset
Private intR, i, j, intC As Integer
Private a, B, C As Double            '�έp�Q��
Private str_DELIVERY_DATE, str_C_ROUTE_NO, str_C_VEHICLE_ID_NO, str_Driver, Str_Receiver, str_ChargeQty As String
Private str_Receivable, str_Payable, str_Premiam, str_Reason, str_AreaStart, str_AreaEnd, str_SDNStatus As String
Private str_ROUTE_NO, str_EXTERN, str_ARRIVE_DATE, str_CUST_NAME, str_SHIP_CS, str_SHIP_CBM, str_SHIP_WT As String
Private route, str_CAR_NOTES, str_SDN_NOTE, str_uom, str_SumReceivable, str_SumPayable, str_C_ROUTE_Time, str_SDN_Date As String
Private str_OnTimeDelivery, str_PODOnTime, str_RejectOrder, str_C_ROUTE_Total, str_SDN_NO, str_SDN_Name, str_CostKind As String
Private rsMain1 As ADODB.Recordset
Private rsRouteT0 As ADODB.Recordset
Private rsOrderT0 As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private DelRecord

Private Sub Ship2TMS(strOrderkey As String)

Call ReDim_Recordset(tmp_Rs)
str_SQL = "select * from sdn03t (nolock) where receipt_no = '" & strOrderkey & "' and ship_qty = 0 "
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'�L���^�Ǹ��
If tmp_Rs.EOF Then tmp_Rs.Close: Exit Sub

str_SQL = "select o.route " & _
        ",o.storerkey " & _
        ",o.orderkey " & _
        ",o.updatesource " & _
        ",o.Externorderkey " & _
        ",od.ExternLineno " & _
        ",od.sku " & _
        ",shippedqty=od.shippedqty + od.qtyallocated " & _
        ",od.editdate " & _
        "from " & strWMSDB & "..orders o (nolock) join " & strWMSDB & "..orderdetail od (nolock) on o.orderkey = od.orderkey " & _
        "and o.status = '9' " & _
        "where len(rtrim(isnull(o.updatesource,''))) > 0 and o.updatesource = '" & strOrderkey & "' " & _
        "and od.shippedqty + od.qtyallocated > 0 "

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockOptimistic

'�L���
If tmp_Rs.EOF Then tmp_Rs.Close: Exit Sub

Dim i As Long
tmp_Rs.MoveFirst

Tran_Level = cn.BeginTrans
Do While Not tmp_Rs.EOF

    str_SQL = "UPDATE TRP03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
             "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
             "and receipt_no ='" & tmp_Rs("updatesource") & "' and product_no = '" & tmp_Rs("sku") & "' and SHIP_QTY = 0 "
    cn.Execute str_SQL ', RowsAffect, adExecuteNoRecords
    
    str_SQL = "UPDATE SDN03T set SHIP_QTY='" & tmp_Rs("shippedqty") & "' " & _
             "where EXTERN='" & RTrim(tmp_Rs("Externorderkey")) & "' and SEQ_NO='" & tmp_Rs("ExternLineno") & "' " & _
             "and receipt_no ='" & tmp_Rs("updatesource") & "' and product_no = '" & tmp_Rs("sku") & "'  and SHIP_QTY = 0 "
    cn.Execute str_SQL ', RowsAffect, adExecuteNoRecords
               
    tmp_Rs.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0
tmp_Rs.Close

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, Me.Name)
End Sub

Private Sub cmb_OneOrder_RBCCode_Click()
'���i�f�D�渹�������i�ƨ��t�έq��G�d��
gd_OneOrder_OrderDetail.Text = cmb_OneOrder_RBCCode.List(cmb_OneOrder_RBCCode.ListIndex)
If Len(Trim(cmb_OneOrder_RBCCode.List(cmb_OneOrder_RBCCode.ListIndex))) > 0 Then
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 10) = Left(cmb_OneOrder_RBCCode.List(cmb_OneOrder_RBCCode.ListIndex), 3)
Else
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 10) = ""
End If
cmb_OneOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_OneOrder_RBCCode_LostFocus()
'���i�f�D�渹�������i�ƨ��t�έq��G�d��
cmb_OneOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_MultiOrder_RBCCode_Click()
'���i�f�D�渹�����h�i�ƨ��t�έq��G�d��
gd_MultiOrder_OrderDetail.Text = cmb_MultiOrder_RBCCode.List(cmb_MultiOrder_RBCCode.ListIndex)
If Len(Trim(cmb_MultiOrder_RBCCode.List(cmb_MultiOrder_RBCCode.ListIndex))) > 0 Then
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 9) = Left(cmb_MultiOrder_RBCCode.List(cmb_MultiOrder_RBCCode.ListIndex), 3)
Else
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 9) = ""
End If
cmb_MultiOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_MultiOrder_RBCCode_LostFocus()
'���i�f�D�渹�����h�i�ƨ��t�έq��G�d��
cmb_MultiOrder_RBCCode.Visible = False
End Sub

Private Sub cmb_OneOrder_RSCCode_Click()
'���i�f�D�渹�������i�ƨ��t�έq��G���`��]
gd_OneOrder_OrderDetail.Text = cmb_OneOrder_RSCCode.List(cmb_OneOrder_RSCCode.ListIndex)
If Len(Trim(cmb_OneOrder_RSCCode.List(cmb_OneOrder_RSCCode.ListIndex))) > 0 Then
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 9) = Left(cmb_OneOrder_RSCCode.List(cmb_OneOrder_RSCCode.ListIndex), 3)
Else
   gd_OneOrder_OrderDetail.TextArray(gd_OneOrder_OrderDetail.Row * gd_OneOrder_OrderDetail.Cols + 9) = ""
End If
cmb_OneOrder_RSCCode.Visible = False
End Sub

Private Sub cmb_OneOrder_RSCCode_LostFocus()
'���i�f�D�渹�������i�ƨ��t�έq��G���`��]
cmb_OneOrder_RSCCode.Visible = False
End Sub

Private Sub cmb_multiOrder_RSCCode_Click()
'���i�f�D�渹�����h�i�ƨ��t�έq��G���`��]
gd_MultiOrder_OrderDetail.Text = cmb_MultiOrder_RSCCode.List(cmb_MultiOrder_RSCCode.ListIndex)
If Len(Trim(cmb_MultiOrder_RSCCode.List(cmb_MultiOrder_RSCCode.ListIndex))) > 0 Then
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 8) = Left(cmb_MultiOrder_RSCCode.List(cmb_MultiOrder_RSCCode.ListIndex), 3)
Else
   gd_MultiOrder_OrderDetail.TextArray(gd_MultiOrder_OrderDetail.Row * gd_MultiOrder_OrderDetail.Cols + 8) = ""
End If
cmb_MultiOrder_RSCCode.Visible = False
End Sub

Private Sub cmb_MultiOrder_RSCCode_LostFocus()
'���i�f�D�渹�����h�i�ƨ��t�έq��G���`��]
cmb_MultiOrder_RSCCode.Visible = False
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
    '���}
    Unload Me
End Sub

Private Sub cmd_MultiOrder_Deliveryok_Click()
'���i�f�D�渹�����h�i�ƨ��t�έq��G���`�q��
If Len(txt_MultiOrder_SignDate.Text) = 0 Then
   msg_text = "��ƿ��~�G����J [�Ȥ�ñ�����]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
   msg_text = "�Ȥ�ñ������G" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_MultiOrder_SignDate.SelStart = 0: txt_MultiOrder_SignDate.SelLength = Len(txt_MultiOrder_SignDate.Text): txt_MultiOrder_SignDate.SetFocus
   Exit Sub
End If

Tran_Level = 0
Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'��s TRP02T
rs_MultiOrder.MoveFirst
Do While Not rs_MultiOrder.EOF
   str_SQL = "Update TRP02T Set CustSignDate = '" & Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate, 5, 2) & "/" & Right(txt_MultiOrder_SignDate, 2) & "'," & _
             "   Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '���`�q��' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("�q��s��").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '��s TRP03T
   str_SQL = "Update TRP03T Set Sign_Qty = Ship_Qty,RSC_Code = '' , RBC_Code = '' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("�q��s��").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   rs_MultiOrder.MoveNext
Loop

cn.CommitTrans
Tran_Level = 0

Call ClearForm
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���X�q��", Me.Caption, "cmd_OnOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_MultiOrder_Expect_Click()
'���i�f�D�渹�����h�i�ƨ��t�έq��G���`�q��
If Len(txt_MultiOrder_SignDate.Text) = 0 Then
   msg_text = "��ƿ��~�G����J [�Ȥ�ñ�����]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
   msg_text = "�Ȥ�ñ������G" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_MultiOrder_SignDate.SelStart = 0: txt_MultiOrder_SignDate.SelLength = Len(txt_MultiOrder_SignDate.Text): txt_MultiOrder_SignDate.SetFocus
   Exit Sub
End If

Dim strRBC As String
Dim strRSC As String
Dim dbSeqNo As Double
Dim dnSignQty As Double

'�ˮ֬O�_��� [���`��]] �P [�d���k��]
With gd_MultiOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 8    '���`�N�X
        If Len(Trim(.Text)) > 0 Then
           .Col = 8: strRSC = strRSC & Trim(.Text)
           .Col = 9: strRBC = strRBC & Trim(.Text)
        End If
     Next iLoop
End With
If strRSC = "" Or strRBC = "" Then
   msg_text = "���`�q�楲����������� [���`��]] �P [�d���k��]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Tran_Level = 0
Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass
rs_MultiOrder.MoveFirst
Do While Not rs_MultiOrder.EOF
   '��s TRP02T
   str_SQL = "Update TRP02T Set CustSignDate = '" & Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate, 5, 2) & "/" & Right(txt_MultiOrder_SignDate, 2) & "'," & _
             "   Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '���`�q��' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("�q��s��").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   rs_MultiOrder.MoveNext
Loop

'��s TRP03T
With gd_MultiOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 1: dbSeqNo = Val(.Text)
        .Col = 5: dnSignQty = Val(.Text)
        .Col = 8: strRSC = Trim(.Text)
        .Col = 9: strRBC = Trim(.Text)
        .Col = 0          '�q��s��
        If strRSC = "" And strRBC = "" Then
           str_SQL = "Update TRP03T Set Sign_Qty = Ship_Qty,RSC_Code = '',RBC_Code = '' " & _
                     "Where Receipt_No = '" & .Text & "' and Seq_No = " & dbSeqNo
        Else
           str_SQL = "Update TRP03T Set Sign_Qty = " & dnSignQty & ",RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "' " & _
                     "Where Receipt_No = '" & .Text & "' and Seq_No = " & dbSeqNo
        End If
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
     Next iLoop
End With

cn.CommitTrans
Tran_Level = 0

Call ClearForm
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���`�q��", Me.Caption, "cmd_MultiOrder_Expect_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_MultiOrder_NoDelivery_Click()
'���i�f�D�渹�����h�i�ƨ��t�έq��G���X�q��

If Len(txt_MultiOrder_SignDate.Text) = 0 Then
   msg_text = "��ƿ��~�G����J [�Ȥ�ñ�����]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
   msg_text = "�Ȥ�ñ������G" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_MultiOrder_SignDate.SelStart = 0: txt_MultiOrder_SignDate.SelLength = Len(txt_MultiOrder_SignDate.Text): txt_MultiOrder_SignDate.SetFocus
   Exit Sub
End If


Dim strRBC As String
Dim strRSC As String

'�ˮ֬O�_��Ĥ@����� [���`��]] �P [�d���k��]
With gd_MultiOrder_OrderDetail
        .Row = 1
        .Col = 9    '���`�N�X
        If Len(Trim(.Text)) = 0 Then
           msg_text = "���X�q��A�Щ�Ӷ��Ĥ@����� [���`��]] �P [�d���k��]"
           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
           Exit Sub
        Else
           .Col = 9: strRSC = .Text
           .Col = 10: strRBC = .Text
        End If
End With

Tran_Level = 0
Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass
'��s TRP02T
rs_MultiOrder.MoveFirst
Do While Not rs_MultiOrder.EOF
   str_SQL = "Update TRP02T Set CustSignDate = '" & Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate, 5, 2) & "/" & Right(txt_MultiOrder_SignDate, 2) & "'," & _
             "   Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '���X�q��' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("�q��s��").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '��s TRP03T
   str_SQL = "Update TRP03T Set Sign_Qty = 0,RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "' " & _
             "Where Receipt_No = '" & rs_MultiOrder.Fields("�q��s��").Value & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   rs_MultiOrder.MoveNext
Loop
cn.CommitTrans
Tran_Level = 0

Call ClearForm
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���X�q��", Me.Caption, "cmd_MultiOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Public Sub cmd_OneOrder_Deliveryok_Click()
On Error GoTo err_Handle
'ñ��O�_�w���@
If Len(Trim(txt_OneOrder_Status)) > 0 Then MsgBox "��ñ��w�g���@�A�L�k�ק�!!", vbOKOnly, Me.Caption: Exit Sub

'�M���S��r��
Call myFormExCharFilter(Me)

Dim strInt As Long, blTmp As Boolean

'�����`�L�k�I�勵�`ñ��
With gd_OneOrder_OrderDetail
        For iLoop = 1 To .Rows - 1
            .Row = iLoop
            '���`�N�X
            .Col = 9: If Len(Trim(.Text)) > 0 Then MsgBox "���Ӧ����@���`�A�L�k������`�q��!!", vbOKOnly, Me.Caption: Exit Sub
            .Col = 10: If Len(Trim(.Text)) > 0 Then MsgBox "���Ӧ����@���`�A�L�k������`�q��!!", vbOKOnly, Me.Caption: Exit Sub
            .Col = 14: If Len(Trim(.Text)) > 0 Then MsgBox "���Ӧ����@���`�A�L�k������`�q��!!", vbOKOnly, Me.Caption: Exit Sub
            .Col = 4: strInt = Val(Trim(.Text))
            .Col = 5: If Val(Trim(.Text)) <> strInt Then blTmp = True
        Next iLoop
End With

If frm_SDNConfirmNotYet.Visible = False Then '�O�_�ֳtñ��T�{
    '�q��q�P�X�f�q����
    If blTmp = True Then
        If MsgBox("�q��q�P�X�f�q���šA�O�_�~��T�{�I", vbYesNo, "���`�q��T�{") <> vbYes Then Exit Sub
    End If
    
    If blTmp = True Then MsgBox "�q��q�P�X�f�q���šA�аȥ��T�{�B�O��ƬO�_���T�I", 64, "���`�q��T�{"
End If

Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'��s SDN01T
str_SQL = "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & txt_C_ROUTE_NO.Text & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'��s SDN02T
str_SQL = "Update SDN02T Set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "', invback = '" & cboInvBack.Text & "'," & _
          " Confirm_UserID = '" & User_id & "',Confirm_Date = '" & Format(Now, "yyyy/mm/dd hh:nn:ss") & "',Confirm_Notes = '���`�q��' , CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' " & _
          ", sdnback = 1 " & _
          ", sdnsenddate = '" & Format(dtpSDNSendDate.Value, "YYYY/MM/DD") & "' " & _
          ", cust_handle = '" & txt_CustHandle.Text & "' " & _
          ", TRP_Handle = '" & txt_TRPHandle.Text & "' " & _
          ", Advance = '" & txt_Advance.Text & "' " & _
          ", INV_Handle = '" & txt_INVHandle.Text & "' " & _
          ", TRP_Cost = '" & txt_TRPCost.Text & "' " & _
          ", Sorting_Cost = '" & txt_SortingCost.Text & "' " & _
          ", Total_Cost = '" & txt_TotalCost.Text & "' " & _
          ", SDN_NOTE = '" & txt_SDNNote.Text & "' " & _
          "Where Receipt_No = '" & RTrim(txt_OneOrder_OrderKey.Text) & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�O�_�ֳtñ��T�{
If frm_SDNConfirmNotYet.Visible = True Then cn.Execute "update sdn02t set sdn_note = '�ֳtñ��T�{' Where Receipt_No = '" & RTrim(txt_OneOrder_OrderKey.Text) & "'", RowsAffect, adExecuteNoRecords

'��s SDN03T
Dim dbSeqNo, dbShipQty, dnSignQty, strRSC, strRBC, strResponsible As String
With gd_OneOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 0: dbSeqNo = .Text
        .Col = 5: dbShipQty = Val(.Text)
        .Col = 6: dnSignQty = Val(.Text)
        .Col = 9: strRSC = Trim(.Text)
        .Col = 10: strRBC = Trim(.Text)
        .Col = 14: strResponsible = Trim(.Text)
        str_SQL = "Update SDN03T Set Ship_Qty = " & dbShipQty & ",Sign_Qty =  " & dbShipQty & ",RSC_Code = '',RBC_Code = '',Responsible = '' " & _
                     "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' and Seq_No = '" & dbSeqNo & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
     Next iLoop
End With

cn.CommitTrans: Tran_Level = 0

'��sñ�檬�A
If rsOrderT0 Is Nothing = False Then
    If rsOrderT0.RecordCount > 0 And rsOrderT0.EOF = False And rsOrderT0("TMS�渹") = txt_OneOrder_OrderKey.Text Then
        rsOrderT0("�禬�渹") = txt_OneOrder_CustomerOrderkey1
        rsOrderT0("���A") = "���`�q��"
    End If
End If

cmdSDNBack.BackColor = vbGreen
cmdSDNBack.Caption = "ñ��w�^"
txt_OneOrder_Status = "���`�q��"

Call cmdCost_Click
'Call cmdSDNBack_Click
Call cmd_OrderQuery_Click

Screen.MousePointer = vbDefault
cmbOrderkey.ListIndex = 0

'�O�_�ֳtñ��T�{
If frm_SDNConfirmNotYet.Visible = False Then
    cmbOrderkey.SetFocus

End If

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���`�q��", Me.Caption, "cmd_OnOrder_Deliveryok_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_OneOrder_Expect_Click()
On Error GoTo err_Handle
'���i�f�D�渹�������i�ƨ��t�έq��G���X�q��
'If Len(dtp_OneOrder_SignDate.Value) = 0 Then
'   msg_text = "��ƿ��~�G����J [�Ȥ�ñ�����]"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'End If
'If Fun_ChkDateFormat(txt_OneOrder_SignDate.Text) = 1 Then
'   msg_text = "�Ȥ�ñ������G" & funRtn_msg
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   txt_OneOrder_SignDate.SelStart = 0: txt_OneOrder_SignDate.SelLength = Len(txt_OneOrder_SignDate.Text): txt_OneOrder_SignDate.SetFocus
'   Exit Sub
'End If

If Len(Trim(txt_OneOrder_Status)) > 0 Then MsgBox "��ñ��w�g���@�A�L�k�ק�!!", vbOKOnly, Me.Caption: Exit Sub

'�M���S��r��
Call myFormExCharFilter(Me)

Dim strRBC As String, strRSC As String, dbSeqNo As String, dnSignQty As Double, dbShipQty As Double, strInt As Double, blTmp As Boolean, strResponsible As String

'�ˮ֬O�_��� [���`��]] �P [�d���k��]
With gd_OneOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 9    '���`�N�X
        If Len(Trim(.Text)) > 0 Then
           .Col = 9: strRSC = strRSC & Trim(.Text)
           .Col = 10: strRBC = strRBC & Trim(.Text)
        End If
            .Col = 4: strInt = Val(Trim(.Text))
            .Col = 5: If Val(Trim(.Text)) <> strInt Then blTmp = True
     Next iLoop
End With

If strRSC = "" Or strRBC = "" Then
   msg_text = "���`�q�楲����������� [���`��]] �P [�d���k��]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'�q��q�P�X�f�q����
If blTmp = True Then
    If MsgBox("�q��q�P�e�f�q���šA�O�_�~��T�{�I", vbYesNo, "���`�q��T�{") <> vbYes Then Exit Sub
End If

If blTmp = True Then MsgBox "�q��q�P�e�f�q���šA�аȥ��T�{�B�O��ƬO�_���T�I", 64, "���`�q��T�{"

Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'��s TRP01T
str_SQL = "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & txt_C_ROUTE_NO.Text & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'��s SDN02T
str_SQL = "Update SDN02T " & _
            "Set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "' " & _
              ", sdnback = 1 " & _
              ", Confirm_UserID = '" & User_id & "' " & _
              ", cust_handle = '" & txt_CustHandle.Text & "' " & _
              ", TRP_Handle = '" & txt_TRPHandle.Text & "' " & _
              ", Advance = '" & txt_Advance.Text & "' " & _
              ", INV_Handle = '" & txt_INVHandle.Text & "' " & _
              ", TRP_Cost = '" & txt_TRPCost.Text & "' " & _
              ", Sorting_Cost = '" & txt_SortingCost.Text & "' " & _
              ", Total_Cost = '" & txt_TotalCost.Text & "' " & _
              ", SDN_NOTE = '" & txt_SDNNote.Text & "' ,invback = '" & cboInvBack.Text & "' " & _
              ",Confirm_Date = '" & Format(Now, "yyyy/mm/dd hh:nn:ss") & "',sdnsenddate = '" & Format(dtpSDNSendDate.Value, "YYYY/MM/DD") & "',Confirm_Notes = '���`�q��' , CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' " & _
              "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'��s SDN03T
With gd_OneOrder_OrderDetail
     For iLoop = 1 To .Rows - 1
        .Row = iLoop
        .Col = 0: dbSeqNo = .Text
        .Col = 5: dbShipQty = Val(.Text)
        .Col = 6: dnSignQty = Val(.Text)
        .Col = 9: strRSC = Trim(.Text)
        .Col = 10: strRBC = Trim(.Text)
        .Col = 14: strResponsible = Trim(.Text)
        If strRSC = "" And strRBC = "" Then '�����`ñ���ƶq=�q��ƶq
           str_SQL = "Update SDN03T set ship_qty = " & dbShipQty & ", Sign_Qty = " & dbShipQty & ",RSC_Code = '',RBC_Code = '',Responsible = '' " & _
                     "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' and Seq_No = '" & dbSeqNo & "' "
        Else
           str_SQL = "Update SDN03T Set Sign_Qty = " & dnSignQty & ",ship_qty = " & dbShipQty & ",RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "',Responsible = '" & strResponsible & "' " & _
                     "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' and Seq_No = '" & dbSeqNo & "' "
        End If
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
     Next iLoop
End With

cn.CommitTrans: Tran_Level = 0

'Ū��ini�ѼơA�O�_�JWMS�t��
Dim objIni As New vbIniFile, strOtherOrder2WMS As String
objIni.FileName = App.Path & "/" & App.title & ".ini"

strOtherOrder2WMS = objIni.ReadData("OPTION", "OtherOrder2WMS", "YES")
Set objIni = Nothing

If UCase(strOtherOrder2WMS) = "YES" And txt_OneOrder_StorerKey <> "LABT01" Then 'WMS�O�_�s�W���ʳ�

    '�g�JWMS���ʳ�
    If RTrim(txt_Priority.Text) = "R" Or RTrim(txt_Priority.Text) = "RC" Or RTrim(txt_Priority.Text) = "A2B" Then '�h�f��B���f�t�e�P���ܤJ�w������WMS���ʳ�
    Else
        '���_�M��K��
        
        Dim rsKeycount As New ADODB.Recordset, strKeycount As String, intLineNumber As Integer
        If RTrim(txt_OneOrder_StorerKey.Text) = "LMBO01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LPSI01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LCHF01" Then
            '�Q�שڵu�������g�JASN�A�����z��n�A��
            If RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Then GoTo NoDo
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select asnkey from " & strWMSDB & "..asn where buyersreference = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '�O�_�w���ͱ��ʳ渹
            
    '            If MsgBox("WMS�O�_���͹w�����ʳ�?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '���t�α��ʳ渹

                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='asn' ", cn
                    '�渹+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'asn'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '�g�J���Y
                    str_SQL = "insert into " & strWMSDB & "..asn (asnKey,StorerKey,externasnkey , sellersreference,BuyersReference,asntype,notes,buyerVAT) " & _
                                "select asnKey = '" & strKeycount & "' , s2.storerkey , rtrim(o.externorderkey) , o.consigneekey , s2.receipt_no , 'A' , description,'" & RTrim(txt_OneOrder_FullName1) & "' " & _
                                "from sdn02t s2 join orders o on s2.c_receipt_no = o.orderkey " & _
                                "Where s2.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "'"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '�g�J��
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    
                    str_SQL = "select s3.product_no , s.descr , s3.storerkey , s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) " & _
                                "from sdn03t s3 join gv_skuxpack s on s.sku = s3.product_no and s3.storerkey = s.storerkey " & _
                                "Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' " & _
                                "group by s3.product_no , s.descr , s3.storerkey, s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    'Add by Gemini @20090303
                    If tmp_Rs.EOF Then cn.RollbackTrans: MsgBox "ñ���q����X�f�q�A�L�ݲ��ͱ��ʳ�I", 64, "ñ����@": GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                    intLineNumber = intLineNumber + 1
            
                    str_SQL = "insert into " & strWMSDB & "..asndetail (asnKey,asnLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered,packkey) " & _
                            "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & RTrim(tmp_Rs("Descr")) & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "�w�s�WWMS�w�����f���ʳ�(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
            End If
        Else
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select pokey from " & strWMSDB & "..po where externpokey = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '�O�_�w���ͱ��ʳ渹
            
    '            If MsgBox("WMS�O�_���͹w�����ʳ�?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '���t�α��ʳ渹
                   
                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
                    '�渹+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '�g�J���Y
                    str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,externpokey , sellername,selleraddress1,BuyersReference,potype,notes) " & _
                                "select poKey = '" & strKeycount & "' , storerkey , receipt_no , consigneekey , cust_name , extern , 'A' , description from sdn02t Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '�g�J��
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    
                    str_SQL = "select s3.product_no , s.descr , s3.storerkey , s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) " & _
                                "from sdn03t s3 join gv_skuxpack s on s.sku = s3.product_no and s3.storerkey = s.storerkey " & _
                                "Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' " & _
                                "group by s3.product_no , s.descr , s3.storerkey, s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    'Add by Gemini @20090303
                    If tmp_Rs.EOF Then cn.RollbackTrans: MsgBox "ñ���q����X�f�q�A�L�ݲ��ͱ��ʳ�I", 64, "ñ����@": GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                    intLineNumber = intLineNumber + 1
            
                    str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered,packkey) " & _
                            "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & RTrim(tmp_Rs("Descr")) & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "�w�s�WWMS�w�����f���ʳ�(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
                End If
        End If
                tmp_Rs.Close
NoDo:
    End If
End If
    
Screen.MousePointer = vbDefault

''���`�O�έp��
'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'str_SQL = "select trp_cost = sum(trp_cost) , sorting_cost = sum(sorting_cost) from gv_ExpectCost Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
'tmp_Rs.Open str_SQL, cn
'
''��sSDN02T
'cn.Execute "update sdn02t set trp_cost = '" & tmp_Rs("trp_cost") & "',sorting_cost = '" & tmp_Rs("sorting_cost") & "',Total_Cost = '" & tmp_Rs("trp_cost") + tmp_Rs("sorting_cost") & "' Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

'��sñ�檬�A
If rsOrderT0 Is Nothing = False Then
    If rsOrderT0.RecordCount > 0 And rsOrderT0.EOF = False And rsOrderT0("TMS�渹") = txt_OneOrder_OrderKey.Text Then
        rsOrderT0("�禬�渹") = txt_OneOrder_CustomerOrderkey1
        rsOrderT0("���A") = "���`�q��"
    End If
End If

txt_OneOrder_Status = "���`�q��"
cmdSDNBack.BackColor = vbGreen
cmdSDNBack.Caption = "ñ��w�^"
Call cmdCost_Click

'Call cmdSDNBack_Click
Call cmd_OrderQuery_Click

cmbOrderkey.ListIndex = 0: cmbOrderkey.SetFocus

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���`�q��", Me.Caption, "cmd_OneOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_OneOrder_NoDelivery_Click()
On Error GoTo err_Handle
'���i�f�D�渹�������i�ƨ��t�έq��G���X�q��
If Len(dtp_OneOrder_SignDate.Value) = 0 Then
   msg_text = "��ƿ��~�G����J [�Ȥ�ñ�����]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
'If Fun_ChkDateFormat(dtp_OneOrder_SignDate.Value) = 1 Then
'   msg_text = "�Ȥ�ñ������G" & funRtn_msg
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   txt_OneOrder_SignDate.SelStart = 0: txt_OneOrder_SignDate.SelLength = Len(txt_OneOrder_SignDate.Text): txt_OneOrder_SignDate.SetFocus
'   Exit Sub
'End If

If Len(Trim(txt_OneOrder_Status)) > 0 Then MsgBox "��ñ��w�g���@�A�L�k�ק�!!", vbOKOnly, Me.Caption: Exit Sub

'�M���S��r��
Call myFormExCharFilter(Me)

Dim strRBC As String, strRSC As String, strResponsible As String

'�ˮ֬O�_��Ĥ@����� [���`��]] �P [�d���k��]
With gd_OneOrder_OrderDetail
        .Row = 1
        .Col = 9: strRSC = .Text    '���`�N�X
        .Col = 10: strRBC = .Text '�d���k��
        .Col = 14: strResponsible = .Text '�d���k�ݤH
End With

If Len(Trim(strRSC)) = 0 Or Len(Trim(strRBC)) = 0 Then MsgBox "�Щ�Ӷ��Ĥ@����� [���`��]] �P [�d���k��]", 64, "���X�q��T�{": Exit Sub

Tran_Level = cn.BeginTrans
Screen.MousePointer = vbHourglass

'��s TRP01T
str_SQL = "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & txt_C_ROUTE_NO.Text & "'"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'��s SDN02T
str_SQL = "Update SDN02T " & _
              "Set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "' " & _
              ", sdnback = 1 " & _
              ", Confirm_UserID = '" & User_id & "' " & _
              ", cust_handle = '" & txt_CustHandle.Text & "' " & _
              ", TRP_Handle = '" & txt_TRPHandle.Text & "' " & _
              ", Advance = '" & txt_Advance.Text & "' " & _
              ", INV_Handle = '" & txt_INVHandle.Text & "' " & _
              ", TRP_Cost = '" & txt_TRPCost.Text & "' " & _
              ", Sorting_Cost = '" & txt_SortingCost.Text & "' " & _
              ", Total_Cost = '" & txt_TotalCost.Text & "' " & _
              ", SDN_NOTE = '" & txt_SDNNote.Text & "' ,invback = '" & cboInvBack.Text & "' " & _
              ",Confirm_Date = '" & Format(Now, "yyyy/mm/dd hh:nn:ss") & "',sdnsenddate = '" & Format(dtpSDNSendDate.Value, "YYYY/MM/DD") & "',Confirm_Notes = '���X�q��' , CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' " & _
              "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
          
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'��s SDN03T
str_SQL = "Update SDN03T Set Sign_Qty = 0,RSC_Code = '" & strRSC & "',RBC_Code = '" & strRBC & "',Responsible = '" & strResponsible & "' " & _
          "Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

''�B�O�p��
'cn.Execute "exec gs_Cost '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords
'
''�p�O�ƶq�אּ 0
'cn.Execute "update sdn05t set chargeqty = 0 , sumreceivable = 0 ,sumpayable = 0 where sdn_no = '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'Ū��ini�ѼơA�O�_�JWMS�t��
Dim objIni As New vbIniFile, strOtherOrder2WMS As String
objIni.FileName = App.Path & "/" & App.title & ".ini"

strOtherOrder2WMS = objIni.ReadData("OPTION", "OtherOrder2WMS", "YES")
Set objIni = Nothing

If UCase(strOtherOrder2WMS) = "YES" And txt_OneOrder_StorerKey <> "LABT01" Then 'WMS�O�_�s�W���ʳ� 1s

    '�g�JWMS���ʳ�
    If RTrim(txt_Priority.Text) <> "R" And RTrim(txt_Priority.Text) <> "RC" And RTrim(txt_Priority.Text) <> "A2B" Then '�h�f��B���f�t�e�P���ܤJ�w������WMS���ʳ�2s
        '���_�M��K��
        
        Dim rsKeycount As New ADODB.Recordset, strKeycount As String, intLineNumber As Integer
        If RTrim(txt_OneOrder_StorerKey.Text) = "LMBO01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LPSI01" Or RTrim(txt_OneOrder_StorerKey.Text) = "LCHF01" Then '3s
            '�Q�שڵu�������g�JASN�A�����z��n�A��
            If RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Then GoTo NoDo
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select asnkey from " & strWMSDB & "..asn where buyersreference = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '�O�_�w���ͱ��ʳ渹 4s
            
    '            If MsgBox("WMS�O�_���͹w�����ʳ�?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '���t�α��ʳ渹

                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='asn' ", cn
                    '�渹+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'asn'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '�g�J���Y
                    str_SQL = "insert into " & strWMSDB & "..asn (asnKey,StorerKey,externasnkey , sellersreference,BuyersReference,asntype,notes,buyerVAT) " & _
                                "select asnKey = '" & strKeycount & "' , s2.store0rkey , rtrim(o.externorderkey) , o.consigneekey , s2.receipt_no , 'A' , description,'" & RTrim(txt_OneOrder_FullName1) & "' " & _
                                "from sdn02t s2 join orders o on s2.c_receipt_no = o.orderkey " & _
                                "Where s2.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "'"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '�g�J��
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    
                    str_SQL = "select s3.product_no , s.descr , s3.storerkey , s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) " & _
                                "from sdn03t s3 join gv_skuxpack s on s.sku = s3.product_no and s3.storerkey = s.storerkey " & _
                                "Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' " & _
                                "group by s3.product_no , s.descr , s3.storerkey, s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    'Add by Gemini @20090303
                    If tmp_Rs.EOF Then cn.RollbackTrans: MsgBox "ñ���q����X�f�q�A�L�ݲ��ͱ��ʳ�I", 64, "ñ����@": GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                    intLineNumber = intLineNumber + 1
            
                    str_SQL = "insert into " & strWMSDB & "..asndetail (asnKey,asnLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered,packkey) " & _
                            "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & RTrim(tmp_Rs("Descr")) & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "')"
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    
                    tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "�w�s�WWMS�w�����f���ʳ�(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
            End If  '4e
        Else    '3m
            Call DB_CheckConnectStatus
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open "select pokey from " & strWMSDB & "..po where externpokey = '" & txt_OneOrder_OrderKey.Text & "' ", cn
            If tmp_Rs.EOF Then '�O�_�w���ͱ��ʳ渹5s
            
                'If MsgBox("WMS�O�_���͹w�����ʳ�?", vbOKCancel, Me.Caption) = vbOK Then
        
                    Tran_Level = cn.BeginTrans
            
                    '���t�α��ʳ渹
                    'Dim rsKeycount As New ADODB.Recordset, strKeycount As String, intLineNumber As Integer
                    rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
                    '�渹+1
                    cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
                    strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                    rsKeycount.Close
            
                    '�g�J���Y
                    str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,externpokey , sellername,selleraddress1,BuyersReference,potype,notes) " & _
                                "select poKey = '" & strKeycount & "' , storerkey , receipt_no , consigneekey , cust_name , extern , 'A' , description from sdn02t Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
                    '�g�J��
                    Call DB_CheckConnectStatus
                    Call ReDim_Recordset(tmp_Rs)
                    str_SQL = "select s3.product_no , s3.storerkey,s.packkey , QtyOrdered=sum(s3.ship_qty - s3.sign_qty) from sdn03t s3 join gv_skuxpack s on s.storerkey = s3.storerkey and s.sku = s3.product_no Where s3.Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' group by s3.product_no , s3.storerkey,s.packkey having sum(s3.ship_qty - s3.sign_qty) > 0 "
                    tmp_Rs.CursorLocation = 3
                    tmp_Rs.Open str_SQL, cn
                    If tmp_Rs.EOF Then cn.RollbackTrans: MsgBox "ñ���q����X�f�q�A�L�ݲ��ͱ��ʳ�I", 64, "ñ����@": GoTo NoDo
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                        intLineNumber = intLineNumber + 1
                
                        str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,StorerKey,QtyOrdered,packkey) " & _
                                "values('" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Format(intLineNumber, "00000") & "','" & tmp_Rs("product_no") & "','" & tmp_Rs("storerkey") & "'," & tmp_Rs("QtyOrdered") & ",'" & tmp_Rs("packkey") & "') "
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        tmp_Rs.MoveNext
            
                    Loop
            
                    cn.CommitTrans: Tran_Level = 0
                    MsgBox "�w�s�WWMS�w�����f���ʳ�(" & strKeycount & ")", vbOKOnly, Me.Caption
    '            End If
                End If '5e
NoDo:
        tmp_Rs.Close
        End If '3e
    End If  '2e
End If  '1e

Screen.MousePointer = vbDefault

'���`�O�έp��
'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'str_SQL = "select trp_cost = sum(trp_cost) , sorting_cost = sum(sorting_cost) from gv_ExpectCost Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' "
'tmp_Rs.Open str_SQL, cn

''��sSDN02T
'cn.Execute "update sdn02t set trp_cost = '" & tmp_Rs("trp_cost") & "',sorting_cost = '" & tmp_Rs("sorting_cost") & "',Total_Cost = '" & tmp_Rs("trp_cost") + tmp_Rs("sorting_cost") & "' Where Receipt_No = '" & txt_OneOrder_OrderKey.Text & "' ", RowsAffect, adExecuteNoRecords

cmbOrderkey.ListIndex = 0: cmbOrderkey.SetFocus

'��sñ�檬�A
If rsOrderT0 Is Nothing = False Then
    If rsOrderT0.RecordCount > 0 And rsOrderT0.EOF = False And rsOrderT0("TMS�渹") = txt_OneOrder_OrderKey.Text Then
        rsOrderT0("�禬�渹") = txt_OneOrder_CustomerOrderkey1
        rsOrderT0("���A") = "���X�q��"
    End If
End If

txt_OneOrder_Status = "���X�q��"
cmdSDNBack.BackColor = vbGreen
cmdSDNBack.Caption = "ñ��w�^"
'Call cmdCost_Click

'Call cmdSDNBack_Click
Call cmd_OrderQuery_Click

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���X�q��", Me.Caption, "cmd_OnOrder_NoDelivery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   
End Sub

Public Sub cmd_OrderQuery_Click()
'�q��d��
If Trim(txt_OrderKey.Text) = "" Then Exit Sub
On Error GoTo err_Handle

Dim strOrderkey As String, strOrderType As String
strOrderkey = Trim(txt_OrderKey.Text)
strOrderType = cmbOrderkey.Text

If cmbOrderkey = "" Then 'Terry 20180907 ���like�d��
    Call ClearForm
    txt_OrderKey.Text = strOrderkey
'    str_SQL = "Select Receipt_No From SDN02T (nolock) Where receipt_no = '" & strOrderkey & "' or receipt_no = '" & Format(strOrderkey, "0000000000") & "' or extern = '" & strOrderkey & "' "
    str_SQL = "Select Receipt_No From SDN02T (nolock) Where receipt_no like '" & strOrderkey & "%' or receipt_no like '" & Format(strOrderkey, "0000000000") & "%' or extern like '" & strOrderkey & "%' "

ElseIf cmbOrderkey.Text = "TMS�渹" Then
    strOrderkey = Format(Trim(txt_OrderKey.Text), "0000000000")
    Call ClearForm
    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
'    str_SQL = "Select Receipt_No From SDN02T (nolock) Where receipt_no = '" & strOrderkey & "' "
    str_SQL = "Select Receipt_No From SDN02T (nolock) Where receipt_no like '" & strOrderkey & "%' "
    
ElseIf cmbOrderkey.Text = "�f�D�渹" Then
    Call ClearForm
    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
'    str_SQL = "Select Receipt_No From SDN02T (nolock) Where extern = '" & strOrderkey & "' "
    str_SQL = "Select Receipt_No From SDN02T (nolock) Where extern like '" & strOrderkey & "%' "
    
Else
    Exit Sub
End If

'�ˬd�渹�O�_�@��h
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.RecordCount = 0 Then
   tmp_Rs.Close
   msg_text = "�d�ߵL���!!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_OrderKey.SelStart = 0:   txt_OrderKey.SelLength = Len(txt_OrderKey.Text): txt_OrderKey.SetFocus
   Exit Sub
   
ElseIf tmp_Rs.RecordCount = 1 Then
   strOrderkey = RTrim(tmp_Rs("Receipt_No"))
   tmp_Rs.Close
   Call Display_OrderData_OneReceipNo(strOrderkey)
Else
tmp_Rs.Close

'�����f�D�渹�����h�i�ƨ��t�έq��
frm_MulitiTMSOrder.Show vbModal

End If
'
''�ɸ��
'str_SQL = "select �����q�N�X=isnull(co.BranchId,' '),�Ȥ�q�����O=isnull(o.externordertype,' '), " & _
'"����=o.cash,�ꦬ=case when isnull(rtrim(cast(o.receiveCash as char)),'') = '0' then ' ' else  isnull(rtrim(cast(o.receiveCash as char)),'') end ," & _
'"�ճ����� = rtrim(isnull(o.B_city,''))" & _
'"from orders o (nolock) left join custorders co on o.orderkey = co.orderkey where o.orderkey = '" & RTrim(txt_OneOrder_OrderKey) & "'"
'Call ReDim_Recordset(tmp_Rs)
'tmp_Rs.CursorLocation = 3
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.RecordCount = 1 Then
'    txt_Externordertype = tmp_Rs.Fields("�Ȥ�q�����O")
'    txt_BranchId = tmp_Rs.Fields("�����q�N�X")
'    txt_ReceiveCash = tmp_Rs.Fields("�ꦬ")
'    txt_Cash = Val(tmp_Rs.Fields("����"))
'    cbx_B_city = tmp_Rs.Fields("�ճ�����")
'Else
'    tmp_Rs.Close
'End If

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "cmd_OrderQuery_Click()")
End Sub

'Public Sub cmd_OrderQuery_Click_del()
''�q��d��
'If Trim(txt_OrderKey.Text) = "" Then Exit Sub
'On Error GoTo err_handle
'
'Dim strOrderkey As String, strOrderType As String
'strOrderkey = Trim(txt_OrderKey.Text)
'strOrderType = cmbOrderkey.Text
'
'If cmbOrderkey.Text = "TMS�渹" Then
'    strOrderkey = Format(Trim(txt_OrderKey.Text), "0000000000")
'    Call ClearForm
'    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
'    str_SQL = "Select Count(Distinct Receipt_No) as OrderCnt From SDN02T Where receipt_no = '" & strOrderkey & "' "
'
'Else
'    Call ClearForm
'    txt_OrderKey.Text = strOrderkey: cmbOrderkey.Text = strOrderType
'    str_SQL = "Select Count(Distinct Receipt_No) as OrderCnt From SDN02T Where extern = '" & strOrderkey & "' "
'End If
'
''�ˬd�f�D�渹�b�ƨ��t�Φ��L�i�� [�q�����]
'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If tmp_Rs.Fields("OrderCnt").Value = 0 Then
'   tmp_Rs.Close
'   msg_text = "�d�ߵL���!!"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   txt_OrderKey.SelStart = 0:   txt_OrderKey.SelLength = Len(txt_OrderKey.Text): txt_OrderKey.SetFocus
'   Exit Sub
'ElseIf tmp_Rs.Fields("OrderCnt").Value = 1 Then
'   '�����f�D�渹�������i�ƨ��t�έq��
'   tmp_Rs.Close
'   Call Display_OrderData_OneReceipNo(strOrderkey)
'Else
'tmp_Rs.Close
'
''�����f�D�渹�����h�i�ƨ��t�έq��
'frm_MulitiTMSOrder.Show vbModal
''   tmp_rs.Close
''   Call Display_OrderData_MultiReceipNo(strOrderKey)
'
'End If
'
'Exit Sub
'err_handle:
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "cmd_OrderQuery_Click()")
'End Sub

Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel "ñ��d��", rsMain1

'..�b���s��EXCEL
With MyXlsApp
  
End With

Set MyXlsApp = Nothing

End Sub

Private Sub cmdCarNOChange_Click()
'If blAdmin = False Then MsgBox "�t�κ޲z���~���v�����榹�@�~!", 64, "�v������": Exit Sub
If Len(RTrim(txt_OneOrder_VehicleID.Text)) = 0 Or Len(RTrim(txt_C_ROUTE_NO.Text)) = 0 Then Exit Sub

intSDNCarChange = 0 '��ñ����@�i�J
frm_SDNCarNOFix.Show vbModal
End Sub

Private Sub cmdDelete_Click()
'Terry 20190328 �p�O�R���\��
If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "��� = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("���") = "V" Then

    strRoute = strRoute & rsRouteT0("�G�����s") & "','"

End If

rsRouteT0.MoveNext
Loop

Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select receipt_no from sdn02t where 1 = 1 and C_ROUTE_NO in ('" & strRoute & "') order by C_ROUTE_NO "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����"
   MsgBox msg_text, vbOKOnly
   Screen.MousePointer = vbDefault
End If

'�R���p�O
str_SQL = "delete from sdn05t where 1 = 1 and C_ROUTE_NO in ('" & strRoute & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

tmp_Rs.Close

MsgBox "���s���p�O��ƬҤw�R��!", vbOKOnly, "�R���p�O"

Screen.MousePointer = vbDefault

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdDeliveryokT0_Click()
On Error GoTo err_Handle

'�O�_�����
If rsOrderT0 Is Nothing Then Exit Sub
If rsOrderT0.RecordCount = 0 Then Exit Sub

rsOrderT0.Filter = "���A = 'V'"

If rsOrderT0.RecordCount = 0 Then rsOrderT0.Filter = "": rsOrderT0.Sort = "�s��": Exit Sub

'�ˬd�ꦬ�����ƶq
rsOrderT0.MoveFirst
Do While Not rsOrderT0.EOF
    If RTrim(rsOrderT0("�ꦬ�N���f��")) <> "" And Val(rsOrderT0("�ꦬ�N���f��")) <> Val(rsOrderT0("�����N���f��")) Then
            DelRecord = MsgBox("�����N���f��<>�ꦬ�N���f�ڡA�нT�{��ƬO�_���T?", vbQuestion + vbYesNo, "�p�O�s��")
            If DelRecord = vbNo Then
                Screen.MousePointer = 0
                rsOrderT0.Filter = "": rsOrderT0.Sort = "�s��"
                Exit Sub
            End If
    End If
    rsOrderT0.MoveNext
Loop

rsOrderT0.MoveFirst
dgOrderT0.Col = 0
'cn.CommandTimeout = 0
cn.Execute "update SDN01T set sdn_Date = getdate() where c_route_no = '" & rsRouteT0("�G�����s") & "'", RowsAffect, adExecuteNoRecords
Screen.MousePointer = 11

Do While Not rsOrderT0.EOF

'mark by Gemini @20140312 �U�C���p�w�L�k����A�G�L�ݦA�ˬd
'    'ñ��O�_�w���@
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    str_SQL = "select receipt_no from sdn02t (nolock) where confirm_notes <> '' and receipt_no = '" & rsOrderT0("TMS�渹") & "' "
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'    If Not tmp_Rs.EOF Then MsgBox "��ñ��w�g���@�L!!", 16, "TMS�渹�G" & rsOrderT0("TMS�渹"): tmp_Rs.Close: Exit Sub
'    tmp_Rs.Close
'
'    '�q��q�P�X�f�q����
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    str_SQL = "select receipt_no from sdn03t (nolock) where order_qty <> ship_qty and receipt_no = '" & rsOrderT0("TMS�渹") & "' "
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'    If Not tmp_Rs.EOF Then MsgBox "�q��q�P�X�f�q����!!", 16, "TMS�渹�G" & rsOrderT0("TMS�渹"): tmp_Rs.Close: Exit Sub
'    tmp_Rs.Close
    
    '��s SDN03T
    str_SQL = "Update SDN03T Set Sign_Qty =  ship_Qty,RSC_Code = '',RBC_Code = '' Where Receipt_No = '" & rsOrderT0("TMS�渹") & "' ; "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '��s SDN02T
    str_SQL = "Update SDN02T Set CustSignDate = isnull(CustSignDate,isnull(SCHEDULEDATE,Arrive_Date)), invback = 'N',sdnback = 1, " & _
              "Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '���`�q��' , CustomerOrderkey1 ='" & rsOrderT0("�禬�渹") & "', Scan = 'N',SDNSendDate = '" & Format(Now, "YYYY/MM/DD") & "' " & _
              "Where Receipt_No = '" & rsOrderT0("TMS�渹") & "' ; "
    
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords


    '��s�^Orders���ꦬ,�ꦬ�ťիh�t�ΥN�w���A�ꦬ���ȫh�ιꦬ
    If RTrim(rsOrderT0("�ꦬ�N���f��")) = "" Then
        '��s�w��=�ꦬ
        If Val(rsOrderT0("�����N���f��")) > 0 Then
            cn.Execute "update o set o.receiveCash = '" & Val(rsOrderT0("�����N���f��")) & "' from orders o join sdn02t s2 on o.orderkey = s2.c_receipt_no where s2.receipt_no = '" & rsOrderT0("TMS�渹") & "' and o.type <> '�R��'", RowsAffect, adExecuteNoRecords
        End If
    Else
        '��s�w��=�ꦬ
        If Val(rsOrderT0("�ꦬ�N���f��")) > 0 Then
            cn.Execute "update o set o.receiveCash = '" & Val(rsOrderT0("�ꦬ�N���f��")) & "' from orders o join sdn02t s2 on o.orderkey = s2.c_receipt_no where s2.receipt_no = '" & rsOrderT0("TMS�渹") & "' and o.type <> '�R��'", RowsAffect, adExecuteNoRecords
        End If
    End If
    
    '�B�O�p��
    str_SQL = "exec gs_Cost '" & rsOrderT0("TMS�渹") & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�������u�p�� Mark by Gemini @20200225 4 ���gs_coat�I�s
'    str_SQL = "exec Es_ARnoDistribution '" & rsOrderT0("TMS�渹") & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'    DoEvents
    '��sñ�檬�A
    rsOrderT0("���A") = "���`�q��"
    
    rsOrderT0.MoveNext

Loop

Screen.MousePointer = 0
rsOrderT0.Filter = "": rsOrderT0.Sort = "�s��"

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "cmdDeliveryokT0_Click;TMS:" & rsOrderT0("TMS�渹") & ";Tran_Level:" & Tran_Level)
End Sub

Private Sub cmdExit_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdOpenOrderT0_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

If dgOrderT0.DataSource Is Nothing Then Exit Sub
If Len(RTrim(rsOrderT0("TMS�渹"))) < 10 Then Exit Sub

cmbOrderkey = "TMS�渹"
txt_OrderKey = rsOrderT0("TMS�渹")
SSTab1.Tab = 3
DoEvents: DoEvents
Call cmd_OrderQuery_Click

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdPremiamAPnew_Click()

If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strCarno As String, strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "��� = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("���") = "V" Then

    If rsRouteT0("���P���X") = "002-34" Or rsRouteT0("���P���X") = "002-29" Or rsRouteT0("���P���X") = "001-97" Or rsRouteT0("���P���X") = "000-31" Or rsRouteT0("���P���X") = "001-36" Or rsRouteT0("���P���X") = "000-70" Or rsRouteT0("���P���X") = "000-67" Or rsRouteT0("���P���X") = "001-23" Then MsgBox "ĳ�����I�p��A�L�k������P���X(" & rsRouteT0("���P���X") & ")!", 16, "�`�N": GoTo EndProc
    If Len(Trim(strCarno)) > 0 And UCase(strCarno) <> UCase(rsRouteT0("���P���X")) Then MsgBox "���P���X���P!", 16, "�`�N": GoTo EndProc
    strCarno = rsRouteT0("���P���X")
    strRoute = strRoute & rsRouteT0("�G�����s") & "','"

End If

rsRouteT0.MoveNext
Loop

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and �G�����s in ('" & strRoute & "')", cn
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'��l�����Excel
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String, dbAP As Double, dbPremiam, dbTRouteNO As String

rsTmp.Filter = "�p�O���O = '���p�O'"
If Not rsTmp.EOF Then
    Recordset2Excel "ĳ�����I���u", rsTmp
    MyXlsApp.Visible = True
    Set MyXlsApp = Nothing
    rsTmp.Filter = ""
    rsTmp.Close
    MsgBox "�����q�楼���@�B�O�Aĳ�����I���u�p��פ�!", 16, Me.Caption
    GoTo EndProc
End If

rsTmp.Filter = ""

'���`���I���B edit by Eric 20140407 ���trp17m�P�_
rsTmp.MoveFirst
Do While Not rsTmp.EOF

If UCase(rsTmp("���I�����u")) <> "1" And rsTmp("�д����O") <> "��B" And Left(rsTmp("�Ƶ�"), 3) <> "�����u" Then
        dbAP = dbAP + rsTmp("���I�`��")
End If

rsTmp.MoveNext
Loop

If dbAP = 0 Then MsgBox "���I�`���B��0�A�L�k�i����u�@�~�A���I���u�פ�I", 16, Me.Caption: GoTo EndProc

dbPremiam = InputBox("���C�JA�q���u����p�U:" & vbCr & vbLf & "1.�p�O�N�X�D�פ������I�����u=1" & vbCr & vbLf & "2.�p�O���O:��B" & vbCr & vbLf & "3.�Ƶ��}�Y:�����u", "�п�Jĳ�����B(���B�ťթΫ������i������u)", "")

If dbPremiam = "" Then 'Edit by Gemini @ 20201208 �ťիh��ĳ��
        Recordset2Excel "ĳ�����I���u", rsTmp
        MyXlsApp.Visible = True
        Set MyXlsApp = Nothing
        GoTo EndProc
End If

dbTRouteNO = InputBox("1.�п�J�@�t���u�s���A�ë��T�{" & vbCr & vbLf & "2.�p�n�M�Ŧ@�t���s�A�п�J�ťըë��T�{" & vbCr & vbLf & "3.�p�������@�t���s�A�Ы�����", "�@�t���u�s��", "")

Tran_Level = cn.BeginTrans

'��s�@�t���s
If Len(dbTRouteNO) > 0 Then cn.Execute "Update sdn02t set T_ROUTE_NO = '" & dbTRouteNO & "' where c_route_no in ('" & strRoute & "') ", RowsAffect, adExecuteNoRecords

'ĳ���k�s
cn.Execute "Update sdn05t set Premiam = 0 where c_route_no in ('" & strRoute & "') ", RowsAffect, adExecuteNoRecords

'�p��ĳ�� edit by Eric 20140407 ���trp17m�P�_
str_SQL = "Update sdn05t " & _
          "Set sdn05t.Premiam = sdn05t.sumpayable / " & dbAP & " * " & dbPremiam & _
          ",sdn05t.note = sdn05t.note + '_" & strCarno & " �M����(" & dbPremiam & ")' " & _
          "from sdn05t(nolock) join sdn02t(nolock) on sdn05t.sdn_no = sdn02t.receipt_no " & _
          "join trp17m(nolock) on trp17m.storerkey = sdn02t.storerkey and trp17m.costcode = sdn05t.costcode " & _
          "where sdn05t.c_route_no in ('" & strRoute & "') and sdn05t.c_route_no <> '' " & _
          "and sdn05t.costkind <> ('��B') " & _
          "and left(isnull(sdn05t.note,''),3) <> '�����u'" & _
          "and trp17m.apnodistribution <> '1'"
          
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and �G�����s in ('" & strRoute & "')", cn
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

 Recordset2Excel "ĳ�����I���u", rsTmp
'�b���s��EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

.Visible = True: End With

Set MyXlsApp = Nothing

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Dim strTmp As String

    '�@���ƨ����u�s��
        str_SQL = "select * from gv_Sdn05tDetail where 1 = 1 "
        
    Dim str_Where As String
    
    '���u�s��
    str_Where = ""
    If Len(RTrim(txtRouteS.Text)) > 0 And Len(RTrim(txtRouteE.Text)) > 0 Then
        str_Where = str_Where & " and ���u�s�� Between '" & RTrim(txtRouteS.Text) & "' and '" & RTrim(txtRouteE.Text) & "' "
    ElseIf Len(RTrim(txtRouteS.Text)) > 0 And Len(RTrim(txtRouteE.Text)) = 0 Then
        str_Where = str_Where & " and ���u�s�� = '" & RTrim(txtRouteS.Text) & "' "
    ElseIf Len(RTrim(txtRouteS.Text)) = 0 And Len(RTrim(txtRouteE.Text)) > 0 Then
        str_Where = str_Where & " and ���u�s�� = '" & RTrim(txtRouteE.Text) & "' "
    End If
    
    '�G�����s
    If Len(RTrim(txt2RouteS.Text)) > 0 And Len(RTrim(txt2RouteE.Text)) > 0 Then
        str_Where = str_Where & " and �G�����s Between '" & RTrim(txt2RouteS.Text) & "' and '" & RTrim(txt2RouteE.Text) & "' "
    ElseIf Len(RTrim(txt2RouteS.Text)) > 0 And Len(RTrim(txt2RouteE.Text)) = 0 Then
        str_Where = str_Where & " and �G�����s = '" & RTrim(txt2RouteS.Text) & "' "
    ElseIf Len(RTrim(txt2RouteS.Text)) = 0 And Len(RTrim(txt2RouteE.Text)) > 0 Then
        str_Where = str_Where & " and �G�����s = '" & RTrim(txt2RouteE.Text) & "' "
    End If

    '�f�D�渹
    If Len(RTrim(txtExternS.Text)) > 0 And Len(RTrim(txtExternE.Text)) > 0 Then
        str_Where = str_Where & " and �f�D�渹 Between '" & RTrim(txtExternS.Text) & "' and '" & RTrim(txtExternE.Text) & "' "
    ElseIf Len(RTrim(txtExternS.Text)) > 0 And Len(RTrim(txtExternE.Text)) = 0 Then
        str_Where = str_Where & " and �f�D�渹 = '" & RTrim(txtExternS.Text) & "' "
    ElseIf Len(RTrim(txtExternS.Text)) = 0 And Len(RTrim(txtExternE.Text)) > 0 Then
        str_Where = str_Where & " and �f�D�渹 = '" & RTrim(txtExternE.Text) & "' "
    End If
        
    'TMS�渹
    If Len(RTrim(txtOrderkeyS.Text)) > 0 And Len(RTrim(txtOrderkeyE.Text)) > 0 Then
        txtOrderkeyS.Text = Format(txtOrderkeyS.Text, "0000000000"): txtOrderkeyE.Text = Format(txtOrderkeyE.Text, "0000000000")
        str_Where = str_Where & " and TMS�渹 Between '" & RTrim(txtOrderkeyS.Text) & "' and '" & RTrim(txtOrderkeyE.Text) & "' "
    ElseIf Len(RTrim(txtOrderkeyS.Text)) > 0 And Len(RTrim(txtOrderkeyE.Text)) = 0 Then
        txtOrderkeyS.Text = Format(txtOrderkeyS.Text, "0000000000")
        str_Where = str_Where & " and TMS�渹 = '" & RTrim(txtOrderkeyS.Text) & "' "
    ElseIf Len(RTrim(txtOrderkeyS.Text)) = 0 And Len(RTrim(txtOrderkeyE.Text)) > 0 Then
        txtOrderkeyS.Text = Format(txtOrderkeyS.Text, "0000000000")
        str_Where = str_Where & " and TMS�渹 = '" & RTrim(txtOrderkeyE.Text) & "' "
    End If
    
    '��f���
    If Len(RTrim(txtDeliveryS.Text)) > 0 And Len(RTrim(txtDeliveryE.Text)) > 0 Then
        str_Where = str_Where & " and ��f�� Between '" & RTrim(txtDeliveryS.Text) & "' and '" & RTrim(txtDeliveryE.Text) & "' "
    ElseIf Len(RTrim(txtDeliveryS.Text)) > 0 And Len(RTrim(txtDeliveryE.Text)) = 0 Then
        str_Where = str_Where & " and ��f�� = '" & RTrim(txtDeliveryS.Text) & "' "
    ElseIf Len(RTrim(txtDeliveryS.Text)) = 0 And Len(RTrim(txtDeliveryE.Text)) > 0 Then
        str_Where = str_Where & " and ��f�� = '" & RTrim(txtDeliveryE.Text) & "' "
    End If
    
    'ñ����
    If Len(RTrim(txtSignDateS.Text)) > 0 And Len(RTrim(txtSignDateE.Text)) > 0 Then
        str_Where = str_Where & " and isnull(convert(varchar(8),ñ���,112),'') Between '" & RTrim(txtSignDateS.Text) & "' and '" & RTrim(txtSignDateE.Text) & "' "
    ElseIf Len(RTrim(txtSignDateS.Text)) > 0 And Len(RTrim(txtSignDateE.Text)) = 0 Then
        str_Where = str_Where & " and isnull(convert(varchar(8),ñ���,112),'') = '" & Len(RTrim(txtSignDateS.Text)) & "' "
    ElseIf Len(RTrim(txtSignDateS.Text)) = 0 And Len(RTrim(txtSignDateE.Text)) > 0 Then
        str_Where = str_Where & " and isnull(convert(varchar(8),ñ���,112),'') = '" & Len(RTrim(txtSignDateE.Text)) & "' "
    End If
    
    '�f�D
    If Len(RTrim(cboStorerKey.Text)) > 0 Then str_Where = str_Where & " and �f�D = '" & RTrim(cboStorerKey.Text) & "' "
    
    '����
    If Len(RTrim(cboCar.Text)) > 0 Then str_Where = str_Where & " and ���� = '" & RTrim(cboCar.Text) & "' "
    
    '���������
    strTmp = ""
    For i = 0 To Car_Num.ListCount - 1
        If Car_Num.Selected(i) Then
                strTmp = strTmp & "'" & Car_Num.List(i) & "',"
        End If
    Next
    If Len(strTmp) > 0 Then str_Where = str_Where & " and ���� in (" & Left(strTmp, Len(strTmp) - 1) & ") "


    '�д����O
    If Len(RTrim(cboCostkind.Text)) > 0 Then str_Where = str_Where & " and �д����O = '" & RTrim(cboCostkind.Text) & "' "
    
    str_SQL = str_SQL & str_Where & " Order by ��f��,�G�����s "
    
    On Error GoTo err_Handle
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Set dgMain1.DataSource = Nothing: Set rsMain1 = Nothing
        Exit Sub
    End If
    
    Call Replication_Recordset(tmp_Rs, rsMain1)
    tmp_Rs.Close
    
    txtAR.Text = 0: txtAP.Text = 0
    
    rsMain1.MoveFirst
    Do While Not rsMain1.EOF
    
        txtAR.Text = txtAR.Text + rsMain1("�����`��")
        txtAP.Text = txtAP.Text + rsMain1("���I�`��")
        rsMain1.MoveNext
    Loop
    
        txtAR.Text = Round(txtAR.Text, 0): txtAP.Text = Round(txtAP.Text, 0)
        txtEarning.Text = txtAR.Text - txtAP.Text
    
    Set dgMain1.DataSource = rsMain1
        
    rsMain1.MoveFirst
    SetDataGridColWidth Me.Caption, dgMain1
    Screen.MousePointer = 0

    Exit Sub
    
err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "�e�fñ��T�{-ñ��d��", Me.Caption, "cmd_Tab1_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_AddCost_Click()
    '�s�W�@��
    dg_Tab2_SDN_Cost.Col = 2
    dg_Tab2_SDN_Cost.Rows = dg_Tab2_SDN_Cost.Rows + 1
    dg_Tab2_SDN_Cost.Row = dg_Tab2_SDN_Cost.Row + 1
    NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
End Sub

Private Sub cmd_Tab2_AddNew_Click()
    Call Clear_CardData
    cmd_Tab2_Save.Enabled = True
    cmd_Tab2_Cancel.Enabled = True
    cmd_Tab2_AddNew.Enabled = False
    cmd_Tab2_Modify.Enabled = False
    cmd_Tab2_Delete.Enabled = False
    txt_Tab02_C_VEHICLE_ID_NO.Enabled = True
    txt_Tab02_Driver.Enabled = True
    txt_Tab02_Receiver.Enabled = True
    txt_Tab02_Delivery_Date.Enabled = True
    cmd_Tab2_SelectCar.Enabled = True
    txt_Tab02_Delivery_Date.SetFocus
End Sub

Private Sub cmd_Tab2_AddOrder_Click()
    '�s�W�@��
    dg_Tab2_SDN_Detail.Col = 2
    dg_Tab2_SDN_Detail.Rows = dg_Tab2_SDN_Detail.Rows + 1
    dg_Tab2_SDN_Detail.Row = dg_Tab2_SDN_Detail.Row + 1
    NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
End Sub

Private Sub cmd_Tab2_Cancel_Click()
    '�d����� >> ����
    Call Clear_CardData
    cmd_Tab2_Cancel.Enabled = False
    cmd_Tab2_Save.Enabled = False
    cmd_Tab2_AddNew.Enabled = True
    cmd_Tab2_Modify.Enabled = True
    cmd_Tab2_Delete.Enabled = True
    cmd_Tab2_SelectCar.Enabled = False
    txt_Tab02_C_VEHICLE_ID_NO.Enabled = False
    txt_Tab02_Driver.Enabled = False
    txt_Tab02_Receiver.Enabled = False
    txt_Tab02_Delivery_Date.Enabled = False
End Sub

Private Sub cmd_Tab2_DelCost_Click()
    '�R���@��
    If dg_Tab2_SDN_Cost.Rows > 2 Then
        dg_Tab2_SDN_Cost.Rows = dg_Tab2_SDN_Cost.Rows - 1
        dg_Tab2_SDN_Cost.Row = dg_Tab2_SDN_Cost.Rows - 1
        NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
    End If
End Sub

Private Sub cmd_Tab2_DelOrder_Click()
    '�R���@��
    If dg_Tab2_SDN_Detail.Rows > 2 Then
        dg_Tab2_SDN_Detail.Rows = dg_Tab2_SDN_Detail.Rows - 1
        dg_Tab2_SDN_Detail.Row = dg_Tab2_SDN_Detail.Rows - 1
        NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
    End If
End Sub

Private Sub cmd_Tab2_Save_Click()
    If Len(Trim(txt_Tab02_Delivery_Date.Text)) = 0 Then
        msg_text = "������J�X����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(txt_Tab02_C_VEHICLE_ID_NO.Text)) = 0 Then
        msg_text = "������J����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(Trim(txt_Tab02_Receiver.Text))) = 0 Then
        msg_text = "������J��ڤH"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    For i = 1 To dg_Tab2_SDN_Detail.Rows - 1
        dg_Tab2_SDN_Detail.Row = i
        dg_Tab2_SDN_Detail.Col = 2: str_EXTERN = Trim(dg_Tab2_SDN_Detail.Text)
        dg_Tab2_SDN_Detail.Col = 3: str_CUST_NAME = Trim(dg_Tab2_SDN_Detail.Text)
'        If Len(Trim(str_EXTERN)) = 0 Or Len(Trim(str_CUST_NAME)) = 0 Then
'            msg_text = "�˸����Ӹ�Ƥ���"
'            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'            Exit Sub
'        End If
        str_SQL = "select * from SDN02T where EXTERN='" & str_EXTERN & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If Not tmp_Rs.EOF Then
            tmp_Rs.Close
            msg_text = "�Ȥ�渹����"
            MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
        tmp_Rs.Close
    End If
    Next
    On Error GoTo err_Handle
    '���o���s
    str_SQL = "select isnull(max(C_Route_No),0) from Logictown.dbo.SDN01T where left(C_Route_No,2)='WD'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    str_C_ROUTE_NO = "WD" & StrPadLeft(Val(Right(Trim(tmp_Rs.Fields(0)), 8)) + 1, 8, 0)
    tmp_Rs.Close
    cn.BeginTrans
        '�s���Y,SDN01T
        str_DELIVERY_DATE = Trim(txt_Tab02_Delivery_Date.Text)
        str_C_VEHICLE_ID_NO = Trim(txt_Tab02_C_VEHICLE_ID_NO.Text)
        str_Driver = Trim(txt_Tab02_Driver.Text)
        Str_Receiver = Trim(txt_Tab02_Receiver.Text)
        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,Receiver,SDNStatus,AddUser)" & _
            "Values ( '" & str_DELIVERY_DATE & "','" & str_C_ROUTE_NO & "','" & str_C_VEHICLE_ID_NO & "','" & str_Driver & "','" & Str_Receiver & "', " & _
            "'" & str_SDNStatus & "','" & User_id & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '�s��:�˸�����,SDN02T
        For i = 1 To dg_Tab2_SDN_Detail.Rows - 1
            dg_Tab2_SDN_Detail.Row = i
            dg_Tab2_SDN_Detail.Col = 2: str_EXTERN = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 3: str_CUST_NAME = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 4: str_SHIP_CS = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 5: str_SHIP_CBM = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 6: str_SHIP_WT = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 7: str_CAR_NOTES = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 8: str_SDN_NOTE = Trim(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 1
            str_SDNStatus = 0
'            If Trim(dg_SDN_Detail.Text) = "��" Then str_SDNStatus = 1
'            dg_SDN_Detail.Col = 0
'            If Trim(dg_SDN_Detail.Text) = "��" Then str_SDNStatus = 2
            If Len(str_EXTERN) = 0 And Len(str_CUST_NAME) = 0 And Len(str_SHIP_CS) = 0 And Len(str_SHIP_CBM) = 0 And Len(str_SHIP_WT) = 0 And Len(str_CAR_NOTES) = 0 And Len(str_SDN_NOTE) = 0 Then
                '�L��Ƥ��s��
            Else
                str_SQL = "Insert into SDN02T (C_ROUTE_NO,ROUTE_NO,EXTERN,ARRIVE_DATE,CUST_NAME,SHIP_CS,SHIP_CBM,SHIP_WT,CAR_NOTES,SDNStatus,SDN_NOTE,RECEIPT_NO) " & _
                    "Values ( '" & str_C_ROUTE_NO & "','" & str_C_ROUTE_NO & "','" & str_EXTERN & "','" & str_DELIVERY_DATE & "','" & str_CUST_NAME & "', " & _
                    "'" & str_SHIP_CS & "','" & str_SHIP_CBM & "','" & str_SHIP_WT & "','" & str_CAR_NOTES & "','" & str_SDNStatus & "','" & str_SDN_NOTE & "','CT" & str_EXTERN & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
            
        Next
        '�s��:�p�O����,SDN05T
        For i = 1 To dg_Tab2_SDN_Cost.Rows - 1
            dg_Tab2_SDN_Cost.Row = i
            dg_Tab2_SDN_Cost.Col = 1: str_SDN_Name = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 2: str_SDN_NO = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 3: str_AreaStart = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 4: str_AreaEnd = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 5: str_uom = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 6: str_ChargeQty = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 7: str_Receivable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 8: str_Payable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 9: str_Premiam = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 10: str_Reason = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 11: str_SumReceivable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 12: str_SumPayable = Trim(dg_Tab2_SDN_Cost.Text)
            dg_Tab2_SDN_Cost.Col = 13: str_CostKind = Trim(dg_Tab2_SDN_Cost.Text)
            If Len(str_SDN_NO) = 0 And Len(str_AreaEnd) = 0 And Len(str_AreaStart) = 0 And Len(str_SumPayable) = 0 And Len(str_SumReceivable) = 0 And Len(str_Reason) = 0 And Len(str_Premiam) = 0 And Len(str_Payable) = 0 And Len(str_Receivable) = 0 And Len(str_ChargeQty) = 0 And Len(str_uom) = 0 Then
                '�L��Ƥ��s��
            Else
                str_SQL = "Insert into SDN05T (C_ROUTE_NO,Uom,ChargeQty,Receivable,Payable,Premiam,Reason,SumReceivable,SumPayable,AreaStart,AreaEnd,SDN_NO,SDN_Name,CostKind) " & _
                    "Values ( '" & str_C_ROUTE_NO & "','" & str_uom & "','" & str_ChargeQty & "','" & str_Receivable & "','" & str_Payable & "', " & _
                    "'" & str_Premiam & "','" & str_Reason & "','" & str_SumReceivable & "','" & str_SumPayable & "','" & str_AreaStart & "','" & str_AreaEnd & "','" & str_SDN_NO & "','" & str_SDN_Name & "','" & str_CostKind & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
        Next
    cn.CommitTrans
    txt_Tab02_C_Route_No.Text = str_C_ROUTE_NO
    cmd_Tab2_Cancel.Enabled = False
    cmd_Tab2_Save.Enabled = False
    cmd_Tab2_AddNew.Enabled = True
    cmd_Tab2_Modify.Enabled = False
    cmd_Tab2_Delete.Enabled = False
    cmd_Tab2_SelectCar.Enabled = False
    txt_Tab02_C_VEHICLE_ID_NO.Enabled = False
    txt_Tab02_Driver.Enabled = False
    txt_Tab02_Receiver.Enabled = False
    txt_Tab02_Delivery_Date.Enabled = False
    Exit Sub
    
err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "�e�fñ��T�{-�s��", Me.Caption, "cmd_Tab0_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_SelectCar_Click()
    Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab2_SelectCar.Name & "2")
End Sub

Private Sub cmdNotYetOrder_Click()
    
    txt_OrderKey = ""
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select * from sdn02t where len(rtrim(isnull(confirm_notes,''))) = 0 "
    tmp_Rs.Open str_SQL, cn
    If tmp_Rs.EOF Then MsgBox "�L�ݽT�{ñ��!!", vbOKOnly, Me.cmdNotYetOrder.Caption: Exit Sub
    
    frm_SDNConfirmNotYet.Show vbModal
    
End Sub

Private Sub cmdCost_Click()

If Len(RTrim(txt_OneOrder_OrderKey.Text)) = 0 Then Exit Sub

If RTrim(txt_ReceiveCash.Text) <> "" Then
    If Val(RTrim(txt_ReceiveCash)) <> Val(RTrim(txt_Cash)) Then
        DelRecord = MsgBox("�����N���f��<>�ꦬ�N���f�ڡA�нT�{��ƬO�_���T?", vbQuestion + vbYesNo, "�p�O�s��")
        If DelRecord = vbNo Then
            Exit Sub
        End If
    End If
End If

'��s�^Orders���ꦬ�A�p�G�O�ťաA�h�ꦬ=�����A���ȫh�ꦬ=�ꦬ
If RTrim(txt_ReceiveCash.Text) = "" Then
    txt_ReceiveCash.Text = Val(RTrim(txt_Cash.Text))
    If Val(txt_ReceiveCash.Text) <> 0 Then
        cn.Execute "update o set o.receiveCash = '" & Val(RTrim(txt_Cash.Text)) & "' from orders o join sdn02t s2 on o.orderkey = s2.c_receipt_no where s2.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "'", RowsAffect, adExecuteNoRecords
    End If
Else
    txt_ReceiveCash.Text = Val(RTrim(txt_ReceiveCash.Text))
    If Val(txt_ReceiveCash.Text) <> 0 Then
        cn.Execute "update o set o.receiveCash = '" & Val(RTrim(txt_ReceiveCash.Text)) & "' from orders o join sdn02t s2 on o.orderkey = s2.c_receipt_no where s2.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "'", RowsAffect, adExecuteNoRecords
    End If
End If

'LABT01�h��s�ճ�����
If RTrim(txt_OneOrder_StorerKey) = "LABT01" Then cn.Execute "update o set o.b_city = '" & RTrim(cbx_B_city.Text) & "' from orders o join sdn02t s2 on o.orderkey = s2.c_receipt_no where s2.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' and o.type <> '�R��' and o.storerkey = 'LABT01'", RowsAffect, adExecuteNoRecords

frm_Cost.Show vbModal
End Sub

Private Sub cmdQueryT0_Click()

txtTotalPLT0 = 0
txtTotalCST0 = 0
txtTotalCubeT0 = 0
txtTotalWgtT0 = 0

On Error GoTo err_Handle
Dim str_Where As String

'�X�����
If Len(RTrim(txtDeliveryDateST0.Text)) > 0 And Len(RTrim(txtDeliveryDateET0.Text)) > 0 Then
    str_Where = str_Where & " and convert(Char(8), t1.delivery_date, 112) Between '" & RTrim(txtDeliveryDateST0.Text) & "' and '" & RTrim(txtDeliveryDateET0.Text) & "' "
ElseIf Len(RTrim(txtDeliveryDateST0.Text)) > 0 And Len(RTrim(txtDeliveryDateET0.Text)) = 0 Then
    str_Where = str_Where & " and convert(Char(8), t1.delivery_date, 112) = '" & RTrim(txtDeliveryDateST0.Text) & "' "
ElseIf Len(RTrim(txtDeliveryDateST0.Text)) = 0 And Len(RTrim(txtDeliveryDateET0.Text)) > 0 Then
    str_Where = str_Where & " and convert(Char(8), t1.delivery_date, 112) = '" & RTrim(txtDeliveryDateET0.Text) & "' "
End If

'�G�����s
If Len(RTrim(txtRouteST0.Text)) > 0 And Len(RTrim(txtRouteET0.Text)) > 0 Then
    str_Where = str_Where & " and t1.c_route_no Between '" & RTrim(txtRouteST0.Text) & "' and '" & RTrim(txtRouteET0.Text) & "' "
ElseIf Len(RTrim(txtRouteST0.Text)) > 0 And Len(RTrim(txtRouteET0.Text)) = 0 Then
    str_Where = str_Where & " and t1.c_route_no = '" & RTrim(txtRouteST0.Text) & "' "
ElseIf Len(RTrim(txtRouteST0.Text)) = 0 And Len(RTrim(txtRouteET0.Text)) > 0 Then
    str_Where = str_Where & " and t1.c_route_no = '" & RTrim(txtRouteET0.Text) & "' "
End If

If Len(RTrim(cboStorerT0)) > 0 Then str_Where = str_Where & "and t2.storerkey = '" & RTrim(cboStorerT0) & "' "
If Len(RTrim(cboCarT0)) > 0 Then str_Where = str_Where & "and t1.c_vehicle_id_no = '" & RTrim(cboCarT0) & "' "

str_SQL = "select distinct ��� = ' ' " & _
        ",�X����� = convert(Char(8), t1.delivery_date, 112) " & _
        ",�G�����s = t1.c_route_no " & _
        ",���P���X = rtrim(t1.c_vehicle_id_no) " & _
        ",�r�p�H = rtrim(t1.driver) " & _
        ",�дڤH = rtrim(isnull(t1.receiver,'')) " & _
        ",�X�f�O�� = round( sum(case when isnull(sp.pallet,0) = 0 then 0 else t3.ship_qty /sp.pallet end) ,3) " & _
        ",�X�f�c�� = sum(case when isnull(sp.casecnt,0) = 0 then 0 else (t3.ship_qty /sp.casecnt) end) " & _
        ",�X�f���n = round( sum(t3.ship_qty * sp.stdcube),3) " & _
        ",�X�f���q = round( sum(t3.ship_qty * sp.stdgrosswgt),3) " & _
        ",�s�W = rtrim(t1.adduser) " & _
        ",�s�W�ɶ� = t1.adddate " & _
        "From sdn01t t1 (nolock) join sdn02t t2 (nolock) on t1.c_route_no = t2.c_route_no " & _
        "join sdn03t t3 (nolock) on t3.receipt_no = t2.receipt_no " & _
        "join gv_skuxpack sp on sp.storerkey = t3.storerkey and t3.product_no = sp.sku " & _
        "where 1 = 1 "
        
str_SQL = str_SQL & str_Where & "group by convert(Char(8), t1.delivery_date, 112),t1.c_route_no,rtrim(t1.c_vehicle_id_no),rtrim(t1.driver),rtrim(isnull(t1.receiver,'')),rtrim(t1.adduser),t1.adddate order by �X�����,�G�����s"

Screen.MousePointer = 11

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧸��s���"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Set dgRouteT0.DataSource = Nothing: Set rsRouteT0 = Nothing
    Set dgOrderT0.DataSource = Nothing: Set rsOrderT0 = Nothing
    Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rsRouteT0)
tmp_Rs.Close: rsRouteT0.MoveFirst

Set dgRouteT0.DataSource = rsRouteT0
'DoEvents

SetDataGridColWidth Me.Caption, dgRouteT0
Screen.MousePointer = 0
blRouteT0Change = True
Call dgRouteT0_RowColChange(1, 1)

Exit Sub
    
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdRecalculate_Click()
'Terry 20180907 �p�O����\��
If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "��� = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("���") = "V" Then

    strRoute = strRoute & rsRouteT0("�G�����s") & "','"

End If

rsRouteT0.MoveNext
Loop

Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select receipt_no from sdn02t where 1 = 1 and C_ROUTE_NO in ('" & strRoute & "') order by C_ROUTE_NO "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����"
   MsgBox msg_text, vbOKOnly
   Screen.MousePointer = vbDefault
End If

'�R���{���p�O
str_SQL = "delete from sdn05t where 1 = 1 and C_ROUTE_NO in ('" & strRoute & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'���� �ΰj��I�s�w�s
Do While Not tmp_Rs.EOF
    str_SQL = "exec gs_cost '" & RTrim(tmp_Rs.Fields("receipt_no").Value) & "'"
         cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
  tmp_Rs.MoveNext
Loop
tmp_Rs.Close

'�ץXExcel
str_SQL = "select �p�O���O = case when s5.Costcode is null then '���p�O' else '' end ,�f�D = rtrim(s2.storerkey) " & _
",ñ��� = isnull(convert(varchar(8),s1.SDN_DATE,112),convert(varchar(8),s2.CONFIRM_DATE,112)) " & _
",�X���� = convert(varchar(8),s1.DELIVERY_DATE,112),�q��� = s2.receipt_DATE " & _
",��f�� = s2.arrive_DATE,�G�����s = s1.C_Route_No,���� = rtrim(s1.C_VEHICLE_ID_NO) " & _
",�r�p = rtrim(s1.driver),�дڤH = rtrim(isnull(s1.receiver,s1.driver)),���u�s�� = s2.Route_No " & _
",�@������ = rtrim(s2.VEHICLE_ID_NO),�f�D�渹 = rtrim(isnull(s2.extern,'')) ,TMS�渹 = rtrim(isnull(s2.receipt_no,'')) " & _
",�Ȥ�s�� = rtrim(s2.consigneekey) ,�Ȥ�W�� = rtrim(isnull(s2.cust_name,'')),��� = rtrim(isnull(s5.Uom,'')) " & _
",�ƶq = isnull(s5.ChargeQty,'0'),������� = isnull(s5.Receivable,'0'),���I��� = isnull(s5.Payable,'0') " & _
",���� = rtrim(isnull(s5.sdn_name,'')),��] = rtrim(isnull(s5.Reason,'')),�����`�� = isnull(s5.SumReceivable,0) " & _
",���I�`�� = isnull(s5.SumPayable,0) " & _
",��I�`�� = case when isnull(s5.Premiam,'0') > 0 then isnull(s5.Premiam,'0') else isnull(s5.SumPayable,0) end " & _
",�_�I = rtrim(isnull(s5.AreaStart,'')),���I = rtrim(isnull(s5.AreaEnd,'')),�Ƶ� = rtrim(isnull(s5.Note,'')) " & _
",�д����O = rtrim(isnull(s5.CostKind,'')),�дڥN�X = rtrim(isnull(s5.Costcode,'')) ,ñ��^�Ǥ�� = isnull(s2.sdnsenddate,'') " & _
",�̲׽T�{ = rtrim(isnull(s2.confirm_userid,'')) ,�T�{�ɶ� = isnull(s2.confirm_date,''),ñ�檬�A = rtrim(s2.confirm_notes) " & _
",ñ��Ƶ� = isnull(s2.sdn_note,''),�p�O�ɶ� = isnull(s2.confirm_date,''),���I�����u = isnull(t17.apnodistribution,'') " & _
",�Ȥ�s�� = rtrim(isnull(t1m.CustGroup,'')),�q�����A = rtrim(isnull(t1m.CHANNEL_TYPE,'')) " & _
"from SDN05T s5 full " & _
"join sdn02t s2 on s5.sdn_no = s2.receipt_no " & _
"left join SDN01T s1 on s1.c_route_no = s2.c_route_no " & _
"left join trp17m t17  on s2.storerkey = t17.storerkey and t17.costcode = s5.costcode " & _
"left join TRP01M t1m on t1m.CONSIGNEEKEY = s2.CONSIGNEEKEY and t1m.STORERKEY = s2.STORERKEY where 1 = 1 and s1.C_Route_No in ('" & strRoute & "')"

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'��Excel
Call Recordset2Excel("Recalculate", rsTmp)

Set MyXlsApp = Nothing
rsTmp.Close: Set rsTmp = Nothing


Screen.MousePointer = vbDefault

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdReceiptDetail_Click()
On Error GoTo err_Handle
If Len(RTrim(txt_OneOrder_StorerOrderKey)) = 0 Then Exit Sub

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)

'���_��K���aexternorderkey
If RTrim(txt_OneOrder_StorerKey.Text) = "LMBO01" Then
    str_SQL = "exec es_SDNReceiptDetail '" & RTrim(txt_OneOrder_StorerOrderKey) & "' "
ElseIf RTrim(txt_OneOrder_StorerKey.Text) = "LLFA01" Then
    str_SQL = "exec es_SDNReceiptDetail '" & RTrim(txt_OneOrder_StorerOrderKey) & "' "
ElseIf RTrim(txt_OneOrder_StorerKey.Text) = "LPSI01" Then
    str_SQL = "exec es_SDNReceiptDetail '" & RTrim(txt_OneOrder_StorerOrderKey) & "' "
ElseIf RTrim(txt_OneOrder_StorerKey.Text) = "LCHF01" Then
    str_SQL = "exec es_SDNReceiptDetail '" & RTrim(txt_OneOrder_StorerOrderKey) & "' "
Else
'��L
    str_SQL = "exec gs_SDNReceiptDetail '" & txt_OneOrder_OrderKey & "' "
End If
tmp_Rs.Open str_SQL, cn
If tmp_Rs.EOF Then MsgBox "�d�L���!", 64, Me.Caption: tmp_Rs.Close: Exit Sub

'��Excel
Recordset2Excel "ñ��d��", tmp_Rs

'..�b���s��EXCEL
With MyXlsApp
  
End With

Set MyXlsApp = Nothing
tmp_Rs.Close

Exit Sub
err_Handle:
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdReset_Click()
Call ClearForm_AllField(Me)
txtDeliveryS = Format(Now - 30, "YYYYMMDD")
txtDeliveryE = Format(Now + 7, "YYYYMMDD")
End Sub

Private Sub cmdSaveToText_Click()
If rsMain1 Is Nothing Then Exit Sub: If rsMain1.EOF Then Exit Sub
End Sub

Private Sub cmdSDNBack_Click()

If Len(RTrim(txt_OneOrder_OrderKey.Text)) = 0 Then Exit Sub

'���\��|���T�{�W�u
''ñ��w�^��Admin�~���i����
'If cmdSDNBack.Caption = "ñ��w�^" Then
'    If blAdmin = True Then
'        If MsgBox("�O�_�A���T�{ñ��w�^?", vbOKCancel, "�t�κ޲z���v��") <> vbOK Then Exit Sub
'    End If
'End If

On Error GoTo err_Handle

Tran_Level = cn.BeginTrans

str_SQL = "update sdn02t set CustSignDate = '" & dtp_OneOrder_SignDate.Value & "', invback = '" & cboInvBack.Text & "' ,sdnback = '1',confirm_date = '" & Format(Now, "yyyy/mm/dd hh:nn:ss") & "',SDNSendDate = '" & Format(dtpSDNSendDate.Value, "YYYY/MM/DD") & "', CustomerOrderkey1 ='" & txt_OneOrder_CustomerOrderkey1.Text & "', Scan = '" & cmbScan.Text & "' where receipt_no = '" & txt_OneOrder_OrderKey & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

cmdSDNBack.BackColor = vbGreen
cmdSDNBack.Caption = "ñ��w�^"

Call cmdCost_Click
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdShipNotes_Click()

If Len(RTrim(txt_OneOrder_StorerKey)) = 0 Then Exit Sub
If MsgBox("�O�_�ɦL�X�f��?", vbOKCancel, "�C�L") <> vbOK Then Exit Sub

On Error GoTo err_Handle

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)

Dim rs_Access As New ADODB.Recordset
Call AccessDB_Connect
strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)
Tran_Level = cnAccess.BeginTrans

If txt_OneOrder_StorerKey = "LVTL01" And Left(txt_C_ROUTE_NO, 1) <> "R" Then

    'VTL�X�f��
    str_SQL = "select * from gv_ReportShipNotesVTL Where �ըƹF�渹 = '" & txt_OneOrder_OrderKey & "' "
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then '�L��ƮɵL���C�L

    str_SQL = "Delete From VTL�X�f��"
    cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
    rs_Access.Open "VTL�X�f��", cnAccess, adOpenStatic, adLockOptimistic
    With tmp_Rs
        .MoveFirst
        Do While Not .EOF
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
           rs_Access.Fields("�Ƶ�").Value = .Fields("�Ƶ�").Value
           rs_Access.Fields("USER").Value = User_Name
           rs_Access.Update
           .MoveNext
        Loop
    
    End With
    cnAccess.CommitTrans: Tran_Level = 0
    Call DB_Disconnect(cnAccess)
    MSAccessAP.DoCmd.OpenReport "VTL�X�f��", acViewPreview
    MSAccessAP.DoCmd.Maximize
    MSAccessAP.Visible = True
    
    End If
    
ElseIf txt_OneOrder_StorerKey = "LYFY09" And Left(txt_C_ROUTE_NO, 1) <> "R" Then
    'YFY P&G�X�f��
    
    'YFY�ɦL�X�f��
    str_SQL = "select * from ev_ReportShipNotesYFY Where TMS�渹 = '" & txt_OneOrder_OrderKey & "' "
    tmp_Rs.Open str_SQL, cn

    If Not tmp_Rs.EOF Then '�L��ƮɵL���C�L

    str_SQL = "Delete From YFY�X�f��"
    cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
    rs_Access.Open "YFY�X�f��", cnAccess, adOpenStatic, adLockOptimistic
    With tmp_Rs
        .MoveFirst
        Do While Not .EOF
           rs_Access.AddNew
               rs_Access.Fields("TMS�渹").Value = .Fields("TMS�渹").Value
               rs_Access.Fields("�q�渹�X").Value = .Fields("�q�渹�X").Value
               rs_Access.Fields("�q��Ӷ�").Value = .Fields("�q��Ӷ�").Value
               rs_Access.Fields("�f�D�W��").Value = .Fields("�f�D�W��").Value
               rs_Access.Fields("�Ȥ�W��").Value = .Fields("�Ȥ�W��").Value
               rs_Access.Fields("�Ƶ�").Value = .Fields("�Ƶ�").Value
               rs_Access.Fields("�a�}").Value = .Fields("�a�}").Value
               rs_Access.Fields("�p���H").Value = .Fields("�p���H").Value
               rs_Access.Fields("�C�L���").Value = .Fields("�C�L���").Value
               rs_Access.Fields("���u�s��").Value = .Fields("���u�s��").Value
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
    
    End With
    cnAccess.CommitTrans: Tran_Level = 0
    Call DB_Disconnect(cnAccess)
    MSAccessAP.DoCmd.OpenReport "YFY�X�f��", acViewPreview
    MSAccessAP.DoCmd.Maximize
    MSAccessAP.Visible = True
    
    End If

ElseIf Left(txt_C_ROUTE_NO, 1) <> "R" Then '��L�X�f��
    

'    str_SQL = "Select * From gv_ReportShipNotes where �ըƹF�渹 = '" & txt_OneOrder_OrderKey & "' "
    
    str_SQL = "select �f�D = RTrim(t1m.storerkey) " & _
    ",�f�D�W�� = (select rtrim(t16.c_name) from trp16m t16 where t16.storerkey = t2.storerkey ) " & _
    ",���u�s��  = rtrim(t2.route_no),�G���ƨ����s = (select isnull(trp01t.c_route_no,'') from trp01t trp01t where t2.route_no = trp01t.route_no) " & _
    ",�X����� = (select convert(varchar(8),trp01t.delivery_date,112) from trp01t trp01t where t2.route_no = trp01t.route_no) " & _
    ",�ݨD��� = convert(varchar(8),t2.arrive_date,112),�f�D�渹 = rtrim(t2.extern),�ըƹF�渹 = t2.receipt_no " & _
    ",���ʽs�� = isnull(o.customerorderkey,''),�Ȥ�W�� = rtrim(t1m.short_name),�Ȥ�a�} = rtrim(t1m.address),�q�� = rtrim(t1m.phone) " & _
    ",�Ȥ�ݨD = cast(isnull(t1m.notes,'') as varchar(300)),�Ƶ� = cast(o.notes as varchar(300)) " & _
    ",�r�p = (select rtrim(isnull(trp05t.driver,'')) from trp05t trp05t where t2.route_no = trp05t.route_no) " & _
    ",���� = rtrim(t2.vehicle_id_no) +  '(�ɦL)',���� = rtrim(isnull(t3.seq_no,'')),�f�� = Rtrim(t3.product_no) " & _
    ",�~�W = Isnull(Rtrim(sp.Descr),''),�ܧO = Isnull(Rtrim(od.lottable06),'') " & _
    ",�X�f�c�� = case when sp.casecnt = 0 then 0 else floor(t3.ship_qty/sp.Casecnt) end,�j�]�� = isnull(rtrim(sp.busr3),'�c') " & _
    ",�X�f�Ӽ� = case when sp.casecnt = 0 then t3.ship_qty else cast(t3.ship_qty as int)%cast(sp.Casecnt as int) end , �p�]�� = isnull(rtrim(sp.busr1),'��') " & _
    ",�`�Ӽ� = t3.ship_qty,��� = case when isnull(t2.otqty,0) = 0 then '' else '�@ '+rtrim(t2.otqty)+'��' end " & _
    "from trp03t t3 join TRP02T t2 on t3.receipt_no = t2.receipt_no " & _
    "join orderdetail od on t3.seq_no = od.orderlinenumber and t2.c_receipt_no = od.orderkey " & _
    "join orders o on o.orderkey = od.orderkey and o.storerkey not in ('LVTL01','LYFY09') " & _
    "join trp01m t1m on t1m.consigneekey = t2.consigneekey and t1m.storerkey = t2.storerkey " & _
    "join gv_skuxpack sp on t3.product_no = sp.sku and sp.storerkey = t2.storerkey where t2.receipt_no = '" & txt_OneOrder_OrderKey & "' "

    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If Not tmp_Rs.EOF Then '�L��ƮɵL���C�L
            str_SQL = "Delete From VLL�X�f��"
            cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Call ReDim_Recordset(rs_Access)
            rs_Access.Open "VLL�X�f��", cnAccess, adOpenStatic, adLockOptimistic
            With tmp_Rs
                .MoveFirst
                Do While Not .EOF
                   rs_Access.AddNew
'                   rs_Access.Fields("�s��").Value = .Fields("�s��").Value
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
                   .MoveNext
                Loop
            End With
            
            cnAccess.CommitTrans: Tran_Level = 0
            Call DB_Disconnect(cnAccess)
            MSAccessAP.DoCmd.OpenReport "VLL�X�f��", acViewPreview
            MSAccessAP.DoCmd.Maximize
            MSAccessAP.Visible = True

    End If
Else '�h�f��

    str_SQL = "select �q�����O = case o2t.priority when 'RC' then '���f�J�w��' when 'A2B' then '���f�t�e��' else case when o2t.storerkey = 'LTKK01' and substring(o2t.extern,3,2) = '12' then '�h�f��(���f)' else '�h�f��' end end " & _
            ", �f�D�W�� =  (select rtrim(t16.c_name) from trp16m t16 where t16.storerkey = o2t.storerkey ) " & _
            ", ���u�s�� = o2t.route_no , �ѦҸ��s = o.ContainerType " & _
            ", �X����� = convert(char(8) , o1t.delivery_date , 112) " & _
            ", ���f��� = convert(char(8) , o2t.arrive_date , 112) " & _
            ", ���� = o2t.vehicle_id_no , �r�p = t9m.driver " & _
            ", TMS�渹 = o2t.receipt_no + '(��)' , �f�D�渹 = o2t.extern " & _
            ", �Ȥ�q�渹�X = o.customerorderkey " & _
            ", �Ȥ�W�� = t1m.short_name , �Ȥ�a�} = t1m.address ,�q�� = t1m.phone, �Ȥ�ݨD = t1m.notes " & _
            ", ��f�Ȥ� = case when o2t.priority in ('R','RC') then '�f�e�G' + rtrim(o.facility) when len(rtrim(o.b_company)) > 0 then '�f�e�G' + rtrim(t1ma.short_name) + '-'+ rtrim(t1ma.address) + ' ' + rtrim(t1ma.phone) else '' end " & _
            ", ���� = rtrim(o3t.seq_no) , �f�� = Rtrim(o3t.Product_No)  " & _
            ", �~�W = sp.descr " & _
            ", �c�� =isnull(case when sp.casecnt = 0 then 0 else floor(o3t.order_qty/sp.Casecnt) end ,0) ,�j�]�� = isnull(rtrim(sp.busr3),'�c') " & _
            ", �Ӽ� =isnull(case when sp.casecnt = 0 then o3t.order_qty else cast(o3t.order_qty as int)%cast(sp.Casecnt as int) end ,0) , �p�]�� = isnull(rtrim(sp.busr1),'��') " & _
            ", �Ƶ� = case when len(cast(o.notes as varchar(1000))) > 0 or len(cast(od.notes as varchar(1000))) > 0 then cast(o.notes as varchar(1000)) + '_' + cast(od.notes as varchar(1000)) else ' ' end  , �`�Ӽ�= o3t.order_qty " & _
            ", �ƨ��� = Case When Isnull(o1t.C_Route_No,'') = '' Then Isnull(Rtrim(o1t.AddWho),'') else Rtrim(o1t.AddWho) End " & _
            "from ort01t o1t join ort02t o2t on o1t.route_no = o2t.route_no " & _
            "join ort03t o3t on o3t.receipt_no = o2t.receipt_no " & _
            "join orders o on o.orderkey = o2t.receipt_no " & _
            "left join trp01m t1m on o2t.consigneekey = t1m.consigneekey and t1m.storerkey = o2t.storerkey " & _
            "left join trp01m t1ma on o.b_company = t1ma.consigneekey and t1ma.storerkey = o.storerkey  " & _
            "left join trp09m t9m on t9m.vehicle_id_no = o2t.vehicle_id_no " & _
            "join orderdetail od on od.orderkey = o.orderkey and od.orderlinenumber = o3t.seq_no  " & _
            "join gv_skuxpack sp on sp.sku = od.sku and sp.storerkey = o2t.storerkey " & _
            "where left(o2t.route_no,1) = 'R' and o2t.receipt_no ='" & txt_OneOrder_OrderKey & "' "
    
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If Not tmp_Rs.EOF Then '�L��ƮɵL���C�L

        cnAccess.Execute "Delete From �h�fñ����", RowsAffect, adExecuteNoRecords
        rs_Access.Open "�h�fñ����", cnAccess, adOpenStatic, adLockOptimistic
        tmp_Rs.MoveFirst
        Do While Not tmp_Rs.EOF
        
            rs_Access.AddNew
            For i = 0 To tmp_Rs.Fields.Count - 1
             rs_Access.Fields(i).Value = RTrim(tmp_Rs.Fields(i).Value)
            Next i
            rs_Access.Update
        
        tmp_Rs.MoveNext
        
        Loop
        
        cnAccess.CommitTrans: Tran_Level = 0
        Call DB_Disconnect(cnAccess)
        MSAccessAP.DoCmd.OpenReport "�h�fñ����", acViewPreview
        MSAccessAP.DoCmd.Maximize
        MSAccessAP.Visible = True
    End If
End If

tmp_Rs.Close

'��s�C�L����
str_SQL = "Update Ort01T Set VLListCount = VLListCount + 1 ,VLListPrintDate = getdate() " & _
          "Where Route_No = '" & txt_OneOrder_RouteNo & "' or C_Route_No = '" & txt_C_ROUTE_NO & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "Update TRP01T Set VLListCount = VLListCount + 1,VLListPrintDate = getdate() " & _
          "Where Route_No = '" & txt_OneOrder_RouteNo & "' or C_Route_No = '" & txt_C_ROUTE_NO & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

Screen.MousePointer = 0
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit:      Set MSAccessAP = Nothing
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�X�f��-�C�L", Me.Caption, "cmdShipNotes_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmdTKPremiamAR_Click()
If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strCarno As String, strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "��� = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("���") = "V" Then

    If rsRouteT0("���P���X") = "000-31" Or rsRouteT0("���P���X") = "001-36" Or rsRouteT0("���P���X") = "000-70" Or rsRouteT0("���P���X") = "000-67" Or rsRouteT0("���P���X") = "001-23" Then MsgBox "TKĳ���������u�p��A�L�k������P���X(" & rsRouteT0("���P���X") & ")!", 16, "�`�N": GoTo EndProc
    If Len(Trim(strCarno)) > 0 And strCarno <> rsRouteT0("���P���X") Then MsgBox "���P���X���P!", 16, "�`�N": GoTo EndProc
    strCarno = rsRouteT0("���P���X")
    strRoute = strRoute & rsRouteT0("�G�����s") & "','"

End If

rsRouteT0.MoveNext
Loop

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and �G�����s in ('" & strRoute & "') and �f�D = 'LTKK01' ", cn
If tmp_Rs.EOF Then MsgBox "�d�L��ƩΩ|�����@�B�O!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'��l�����Excel
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String, dbAR As Double, dbPremiam

rsTmp.Filter = "�p�O�ɶ� = '1900-01-01 00:00:00.000'"
If Not rsTmp.EOF Then
    Recordset2Excel "TKĳ���������u", rsTmp
    MyXlsApp.Visible = True
    Set MyXlsApp = Nothing
    rsTmp.Filter = ""
    rsTmp.Close
    MsgBox "�����q�楼���@�B�O�ATKĳ���������u�p��פ�!", 16, Me.Caption
    GoTo EndProc
End If

rsTmp.Filter = ""

'���`�������B
rsTmp.MoveFirst
Do While Not rsTmp.EOF

If Left(rsTmp("�Ƶ�"), 4) <> "�G���t�e" And Left(rsTmp("�Ƶ�"), 3) <> "�����u" And rsTmp("�дڥN�X") <> "Cancel" And rsTmp("�дڥN�X") <> "I" And rsTmp("�дڥN�X") <> "R" Then dbAR = dbAR + rsTmp("�����`��")

rsTmp.MoveNext
Loop

dbPremiam = InputBox("���C�JTKĳ���������u����p�U�G" & vbCr & vbLf & "1.�N�X�e���X:" & vbCr & vbLf & "2.�p�O�N�X:Cancel,I,R" & vbCr & vbLf & "3.�p�O���O:" & vbCr & vbLf & "4.�Ƶ��}�Y:�����u�A�G���t�e", "�п�Jĳ�����B(��J0���Ϋ������i����p��)", 0, 0)

If Val(dbPremiam) = 0 Then
        Recordset2Excel "TKĳ���������u", rsTmp
        MyXlsApp.Visible = True
        Set MyXlsApp = Nothing
        GoTo EndProc
End If

Tran_Level = cn.BeginTrans

'�p��ĳ��
str_SQL = "Update sdn05t " & _
          "Set sumreceivable = sdn05t.sumreceivable / " & dbAR & " * " & dbPremiam & _
          ",receivable = sdn05t.sumreceivable / " & dbAR & " * " & dbPremiam & " / sdn05t.chargeqty " & _
          ",note = '�M����(" & dbPremiam & ")' + '_' + sdn05t.note " & _
          "from sdn05t join sdn02t s2 on s2.receipt_no = sdn05t.sdn_no and s2.storerkey = 'LTKK01' " & _
          "where sdn05t.c_route_no in ('" & strRoute & "')  " & _
          "and sdn05t.sumreceivable > 0 " & _
          "and sdn05t.costcode not in ('I','R','Cancel') " & _
          "and left(sdn05t.Note,3) <> '�����u' " & _
          "and left(sdn05t.Note,4) <> '�G���t�e' "
          
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'���s�p����v�T���������u

str_SQL = "select receipt_no from sdn02t where c_route_no in ('" & strRoute & "') and storerkey = 'LTKK01'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn

If tmp_Rs.EOF Then
Else
    tmp_Rs.MoveFirst
    Do While Not tmp_Rs.EOF
        str_SQL = "exec Es_ARnoDistribution '" & RTrim(tmp_Rs("receipt_no")) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        tmp_Rs.MoveNext
    Loop
End If

cn.CommitTrans

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and �G�����s in ('" & strRoute & "') and �f�D = 'LTKK01' ", cn
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

 Recordset2Excel "TKĳ���������u", rsTmp
'�b���s��EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

.Visible = True: End With

Set MyXlsApp = Nothing

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdTransCube_Click()

If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strCarno As String, strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "��� = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("���") = "V" Then

    If rsRouteT0("���P���X") = "001-97" Or rsRouteT0("���P���X") = "000-31" Or rsRouteT0("���P���X") = "001-36" Or rsRouteT0("���P���X") = "000-70" Or rsRouteT0("���P���X") = "000-67" Or rsRouteT0("���P���X") = "001-23" Then MsgBox "���I�p��A�L�k������P���X(" & rsRouteT0("���P���X") & ")!", 16, "�`�N": GoTo EndProc
    If Len(Trim(strCarno)) > 0 And UCase(strCarno) <> UCase(rsRouteT0("���P���X")) Then MsgBox "���P���X���P!", 16, "�`�N": GoTo EndProc
    strCarno = rsRouteT0("���P���X")
    strRoute = strRoute & rsRouteT0("�G�����s") & "','"

End If

rsRouteT0.MoveNext
Loop

str_SQL = "select �G�����s = s2.c_route_no, �f�D = rtrim(s2.storerkey),��f��� = s2.arrive_date,�G������ = rtrim(s1.c_vehicle_id_no),�@������ = rtrim(s2.vehicle_id_no),�q�渹�X = rtrim(s2.extern) " & _
            ",TMS�渹 = rtrim(s2.receipt_no),�Ȥ�W�� = rtrim(s2.cust_name),�ϰ� = rtrim(t2m.city),�a�} = rtrim(t1m.address),���~�O = rtrim(isnull(sp.susr1,'')),����=rtrim(s3.seq_no),�~�� = rtrim(s3.product_no) " & _
            ",��ƽT�{ = isnull(t2.otqty,0) + isnull(o2.otqty,0),�X�f�c�� = case when sp.casecnt = 0 then 1 else ceiling(s3.ship_qty/sp.casecnt) end " & _
            ",�X�f�`�Ӽ� = s3.ship_qty,�c�J�� = sp.casecnt,��쭫 = sp.stdgrosswgt,���� = sp.stdcube,�X�f�� = sp.stdgrosswgt*ship_qty,�X�f�� = sp.stdcube*ship_qty,����� = round( " & _
            "case when s3.storerkey = 'LTKK01' and sp.susr1 not in ('�s�i�~') then s3.ship_qty * sp.stdgrosswgt /12.0 " & _
            "when s3.storerkey = 'LVTL01' and sp.susr1 in ('���s','���s') then s3.ship_qty * sp.stdgrosswgt /12.0 " & _
            "when s3.storerkey = 'LAPP01' and sp.casecnt > 0 then case when sp.casecnt = 0 then 1 else ceiling(s3.ship_qty/sp.casecnt) end * 2.2 " & _
            "when s3.storerkey = 'LKAO01' and sp.casecnt > 0 then case when sp.casecnt = 0 then 1 else ceiling(s3.ship_qty/sp.casecnt) end * 0.7 " & _
            "else s3.ship_qty * sp.stdcube end,3), " & _
            "���⭫ = round( " & _
            "case when s3.storerkey = 'LABT01' and sp.casecnt > 0 then case when sp.casecnt = 0 then 1 else ceiling(s3.ship_qty/sp.casecnt) end * 10 " & _
            "when s3.storerkey = 'LITW01'  and sp.casecnt > 0 then case when sp.casecnt = 0 then 1 else ceiling(s3.ship_qty/sp.casecnt) end * 20 " & _
            "else s3.ship_qty * sp.stdgrosswgt end * " & _
            "case when rtrim(t2m.city) in ('��','�]��') then 1.1 " & _
            "when rtrim(t2m.city) in ('�x�_','�s��') then 1 " & _
            "when rtrim(t2m.city) = '���' then 0.8 else 1 end,3) " & _
            "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey not in ('LITW01','LABT01','LMYS01') " & _
            "join sdn01t s1 on s1.c_route_no = s2.c_route_no join gv_skuxpack sp on sp.storerkey = s3.storerkey and s3.product_no = sp.sku " & _
            "join trp01m t1m on t1m.storerkey = s2.storerkey and s2.consigneekey = t1m.consigneekey " & _
            "join trp02m t2m on t1m.zip = t2m.zip " & _
            "left join trp02t t2 on t2.receipt_no = s2.receipt_no left join ort02t o2 on o2.receipt_no = s2.receipt_no where s2.c_route_no in ('" & strRoute & "') "

str_SQL = str_SQL & "union all select �G�����s = s2.c_route_no,�f�D = rtrim(s2.storerkey),��f��� = s2.arrive_date,�G������ = rtrim(s1.c_vehicle_id_no),�@������ = rtrim(s2.vehicle_id_no) " & _
                    ",�q�渹�X = rtrim(s2.extern),TMS�渹 = rtrim(s2.receipt_no),�Ȥ�W�� = rtrim(s2.cust_name),�ϰ� = rtrim(t2m.city),�a�} = rtrim(t1m.address) " & _
                    ",���~�O = '',���� = '',�~�� = '',��ƽT�{ = isnull(t2.otqty,0) + isnull(o2.otqty,0),�X�f�c�� = '',�X�f�`�Ӽ� = '',�c�J�� = '',��쭫 = '',���� = '',�X�f�� = '',�X�f�� = '',����� = round( " & _
                    "case when s2.storerkey = 'LITW01' then (isnull(t2.otqty,0) + isnull(o2.otqty,0)) * 2.0 " & _
                    "when s2.storerkey = 'LABT01' then (isnull(t2.otqty,0) + isnull(o2.otqty,0)) * 1.4 " & _
                    "when s2.storerkey = 'LMYS01' then (isnull(t2.otqty,0) + isnull(o2.otqty,0)) * 1.2 else 0 end,3), " & _
                    "���⭫ = round( " & _
                    "case when s2.storerkey = 'LITW01' then (isnull(t2.otqty,0) + isnull(o2.otqty,0)) * 20 " & _
                    "when s2.storerkey = 'LABT01' then (isnull(t2.otqty,0) + isnull(o2.otqty,0)) * 10 " & _
                    "else 0 end * " & _
                    "case when rtrim(t2m.city) in ('��','�]��') then 1.1 " & _
                    "when rtrim(t2m.city) in ('�x�_','�s��') then 1 " & _
                    "when rtrim(t2m.city) = '���' then 0.8 else 1 end " & _
                    ",3) " & _
                    "from sdn02t s2 join sdn01t s1 on s1.c_route_no = s2.c_route_no and s2.storerkey in ('LITW01','LABT01','LMYS01') " & _
                    "join trp01m t1m on t1m.storerkey = s2.storerkey and s2.consigneekey = t1m.consigneekey " & _
                    "join trp02m t2m on t1m.zip = t2m.zip " & _
                    "left join trp02t t2 on t2.receipt_no = s2.receipt_no left join ort02t o2 on o2.receipt_no = s2.receipt_no where s2.c_route_no in ('" & strRoute & "') " & _
                    "order by s2.c_route_no,rtrim(s2.storerkey),rtrim(s2.extern),rtrim(s3.seq_no) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: GoTo EndProc

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

 Recordset2Excel "�q�����", rsTmp
'�b���s��EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

.Visible = True: End With


'�p���`������P�I��
Dim strAddress As String
txtPointT0 = 0: txtTransCubeT0 = 0
rsTmp.Sort = "�a�}"
rsTmp.MoveFirst
Do While Not rsTmp.EOF

    If rsTmp("�a�}") <> strAddress Then
        txtPointT0 = txtPointT0 + 1
        strAddress = rsTmp("�a�}")
    End If
    txtTransCubeT0 = txtTransCubeT0 + rsTmp("�����")

rsTmp.MoveNext
Loop

rsTmp.Sort = ""
Set MyXlsApp = Nothing

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdUnRouteConfirm_Click()
    On Error GoTo err_Handle
    
    If Left(txt_C_ROUTE_NO.Text, 1) = "N" Then MsgBox "�������ո��s(N�}�Y���s)�A�L�k�����X���T�{�A�ХѥX���T�{�Ҳդ�����!(�w�T�{���s==>���s�R��)", 64, Trim(txt_C_ROUTE_NO) & "==>���X���T�{": Exit Sub
    If Len(RTrim(txt_OneOrder_VehicleID.Text)) = 0 Or Len(RTrim(txt_C_ROUTE_NO.Text)) = 0 Then Exit Sub
    If MsgBox("�����s�N�^�_���X���T�{���A�A�Ӹ��s�Ҧ��q��B�O�Pñ��T�{�N�@�֧R���A�O�_�~��?", vbOKCancel, Trim(txt_C_ROUTE_NO) & "==>���X���T�{") <> vbOK Then Exit Sub
    
    '�T�O���s���s�b(�X���T�{�Ҳդ��A�w�T�{ñ��̨S�Q�R�����s)
    Call cmd_OrderQuery_Click
    
    Tran_Level = cn.BeginTrans
    cn.Execute "exec gs_UnRouteConfirm '" & Trim(txt_C_ROUTE_NO) & "' ", RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans: Tran_Level = 0
    
    Call cmd_OrderQuery_Click
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dg_SDN_Head_Scroll()
    Text1.Visible = False
End Sub

Private Sub cmdPremiamAP_Click()

If rsRouteT0 Is Nothing Then Exit Sub
On Error GoTo err_Handle

Dim strCarno As String, strRoute As String
blRouteT0Change = False

rsRouteT0.Filter = "��� = 'V'"

If rsRouteT0.EOF Then GoTo EndProc
rsRouteT0.MoveFirst
Do While Not rsRouteT0.EOF

If rsRouteT0("���") = "V" Then

    If rsRouteT0("���P���X") = "002-34" Or rsRouteT0("���P���X") = "002-29" Or rsRouteT0("���P���X") = "001-97" Or rsRouteT0("���P���X") = "000-31" Or rsRouteT0("���P���X") = "001-36" Or rsRouteT0("���P���X") = "000-70" Or rsRouteT0("���P���X") = "000-67" Or rsRouteT0("���P���X") = "001-23" Then MsgBox "ĳ�����I�p��A�L�k������P���X(" & rsRouteT0("���P���X") & ")!", 16, "�`�N": GoTo EndProc
    If Len(Trim(strCarno)) > 0 And UCase(strCarno) <> UCase(rsRouteT0("���P���X")) Then MsgBox "���P���X���P!", 16, "�`�N": GoTo EndProc
    strCarno = rsRouteT0("���P���X")
    strRoute = strRoute & rsRouteT0("�G�����s") & "','"

End If

rsRouteT0.MoveNext
Loop

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and �G�����s in ('" & strRoute & "')", cn
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Dim rsTmp As New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

'��l�����Excel
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String, dbAP As Double, dbPremiam

rsTmp.Filter = "�p�O���O = '���p�O'"
If Not rsTmp.EOF Then
    Recordset2Excel "ĳ�����I���u", rsTmp
    MyXlsApp.Visible = True
    Set MyXlsApp = Nothing
    rsTmp.Filter = ""
    rsTmp.Close
    MsgBox "�����q�楼���@�B�O�Aĳ�����I���u�p��פ�!", 16, Me.Caption
    GoTo EndProc
End If

rsTmp.Filter = ""

'���`���I���B
rsTmp.MoveFirst
Do While Not rsTmp.EOF

''�����I�����u
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open "select apnodistribution from trp17m where costcode = '", cn
'If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

If UCase(rsTmp("�дڥN�X")) <> "STAIRS" And UCase(rsTmp("�дڥN�X")) <> "CA-R" And UCase(rsTmp("�дڥN�X")) <> "C-R" And UCase(rsTmp("�дڥN�X")) <> "C1-R" And UCase(rsTmp("�дڥN�X")) <> "FULL_CS-R" And UCase(rsTmp("�дڥN�X")) <> "HT-R" And UCase(rsTmp("�дڥN�X")) <> "HT1-R" And UCase(rsTmp("�дڥN�X")) <> "KL-R" And UCase(rsTmp("�дڥN�X")) <> "KL1-R" And UCase(rsTmp("�дڥN�X")) <> "ML-R" And UCase(rsTmp("�дڥN�X")) <> "ML1-R" And UCase(rsTmp("�дڥN�X")) <> "SA-R" And UCase(rsTmp("�дڥN�X")) <> "TP-R" And UCase(rsTmp("�дڥN�X")) <> "TP1-R" And UCase(rsTmp("�дڥN�X")) <> "TY-R" And UCase(rsTmp("�дڥN�X")) <> "TY1-R" Then
    
    If UCase(rsTmp("�дڥN�X")) <> "W-R" And UCase(rsTmp("�дڥN�X")) <> "FORKLIFT" And UCase(rsTmp("�дڥN�X")) <> "CANCEL" And UCase(rsTmp("�дڥN�X")) <> "REPALLETIS" And Left(rsTmp("�дڥN�X"), 6) <> "002-34" And Left(rsTmp("�дڥN�X"), 6) <> "002-25" And Left(rsTmp("�дڥN�X"), 6) <> "000-31" And Left(rsTmp("�дڥN�X"), 6) <> "001-97" And Left(rsTmp("�дڥN�X"), 6) <> "002-29" And Left(rsTmp("�дڥN�X"), 6) <> "001-36" And Left(rsTmp("�дڥN�X"), 6) <> "000-70" And Left(rsTmp("�дڥN�X"), 6) <> "000-67" And Left(rsTmp("�дڥN�X"), 6) <> "001-23" And UCase(Left(rsTmp("�дڥN�X"), 6)) <> "CANCEL" And rsTmp("�д����O") <> "��B" And Left(rsTmp("�Ƶ�"), 3) <> "�����u" Then
        dbAP = dbAP + rsTmp("���I�`��")
    End If

End If

rsTmp.MoveNext
Loop

If dbAP = 0 Then MsgBox "���I�`���B��0�A�L�k�i����u�@�~�A���I���u�פ�I", 16, Me.Caption: GoTo EndProc

dbPremiam = InputBox("���C�JA�q���u����p�U�G" & vbCr & vbLf & "1.�N�X�e���X:002-34,002-29,001-97,000-31,001-36,000-70,000-67,001-23,002-25,Cancel" & vbCr & vbLf & "2.�p�O�N�X:Stairs,forklift,RePalletIs,CA-R,C-R,C1-R,FULL_CS-R,HT-R,HT1-R,KL-R,KL1-R,ML-R,ML1-R,SA-R,TP-R,TP1-R,TY-R,TY1-R,W-R" & vbCr & vbLf & "3.�p�O���O:��B" & vbCr & vbLf & "4.�Ƶ��}�Y:�����u", "�п�Jĳ�����B(��J0���Ϋ������i����p��)", 0, 0)

If Val(dbPremiam) = 0 Then
        Recordset2Excel "ĳ�����I���u", rsTmp
        MyXlsApp.Visible = True
        Set MyXlsApp = Nothing
        GoTo EndProc
End If

Tran_Level = cn.BeginTrans

'ĳ���k�s
cn.Execute "Update sdn05t set Premiam = 0 where c_route_no in ('" & strRoute & "') ", RowsAffect, adExecuteNoRecords

'�p��ĳ��
str_SQL = "Update sdn05t " & _
          "Set Premiam = sumpayable / " & dbAP & " * " & dbPremiam & _
          ",note = note + '_" & strCarno & " �M����(" & dbPremiam & ")' " & _
          "where c_route_no in ('" & strRoute & "') and c_route_no <> '' " & _
          "and left(costcode,6) not in ('002-34','002-29','001-97','002-25','000-31','001-36','000-70','000-67','001-23','Cancel') " & _
          "and costcode not in ('Stairs','forklift','repalletis','CA-R','C-R','FULL_CS-R','HT-R','KL-R','ML-R','SA-R','TP-R','TY-R','W-R','C1-R','HT1-R','KL1-R','ML1-R','TP1-R','TY1-R') " & _
          "and costkind <> ('��B') " & _
          "and left(isnull(note,''),3) <> '�����u'"
          
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from gv_Sdn05tDetail where 1 = 1 and �G�����s in ('" & strRoute & "')", cn
If tmp_Rs.EOF Then MsgBox "�d�L���!", 16, Me.Caption: tmp_Rs.Close: Exit Sub

Call OffLineRecordset(tmp_Rs, rsTmp)
tmp_Rs.Close

 Recordset2Excel "ĳ�����I���u", rsTmp
'�b���s��EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

.Visible = True: End With

Set MyXlsApp = Nothing

EndProc:
rsRouteT0.Filter = ""
Set dgRouteT0.DataSource = rsRouteT0
SetDataGridColWidth Me.Caption, dgRouteT0
rsRouteT0.MoveFirst
blRouteT0Change = True
Screen.MousePointer = 0
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub




Private Sub dg_Tab2_SDN_Cost_Click()
    If dg_Tab2_SDN_Cost.Col < 13 Then
        NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
    End If
End Sub

Private Sub dg_Tab2_SDN_Detail_Click()
    If dg_Tab2_SDN_Detail.Col = 0 Or dg_Tab2_SDN_Detail.Col = 1 Then
'        If Len(dg_SDN_Detail.Text) = 0 Then
'            dg_Tab2_SDN_Detail.Text = "��"
'        Else
'            dg_Tab2_SDN_Detail.Text = ""
'        End If
    End If
    If dg_Tab2_SDN_Detail.Col = 9 Then
        If Len(dg_Tab2_SDN_Detail.Text) = 0 Then
            dg_Tab2_SDN_Detail.Text = "��"
            dg_Tab2_SDN_Detail.Col = 4
            txt_Tab2_sum_Case.Text = Val(txt_Tab2_sum_Case.Text) + Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 5
            txt_Tab2_sum_CBM.Text = Val(txt_Tab2_sum_CBM.Text) + Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 6
            txt_Tab2_sum_WT.Text = Val(txt_Tab2_sum_WT.Text) + Val(dg_Tab2_SDN_Detail.Text)
        Else
            dg_Tab2_SDN_Detail.Text = ""
            dg_Tab2_SDN_Detail.Col = 4
            txt_Tab2_sum_Case.Text = Val(txt_Tab2_sum_Case.Text) - Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 5
            txt_Tab2_sum_CBM.Text = Val(txt_Tab2_sum_CBM.Text) - Val(dg_Tab2_SDN_Detail.Text)
            dg_Tab2_SDN_Detail.Col = 6
            txt_Tab2_sum_WT.Text = Val(txt_Tab2_sum_WT.Text) - Val(dg_Tab2_SDN_Detail.Text)
        End If
        dg_Tab2_SDN_Detail.Col = 9
    End If
    If dg_Tab2_SDN_Detail.Col > 1 And dg_Tab2_SDN_Detail.Col < 9 Then
        NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
    End If
End Sub

Private Sub dg_Tab2_SDN_Detail_Scroll()
    Text3.Visible = False
End Sub

Private Sub dgMain1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain1

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain1_HeadClick(ByVal ColIndex As Integer)
If dgMain1.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMain1.Sort = dgMain1.Columns(ColIndex).Caption & " DESC"
    dgMain1.ClearSelCols
    intColumnIndex = 255

Else
    rsMain1.Sort = dgMain1.Columns(ColIndex).Caption
    dgMain1.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub





Private Sub dgOrderT0_HeadClick(ByVal ColIndex As Integer)
Dim dg As Object, rs As Object
Set dg = dgOrderT0: Set rs = rsOrderT0

If dg.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rs.Sort = dg.Columns(ColIndex).Caption & " DESC"
    dg.ClearSelCols
    intColumnIndex = 255

Else
    rs.Sort = dg.Columns(ColIndex).Caption
    dg.ClearSelCols
    intColumnIndex = ColIndex

End If
End Sub

Private Sub dgRouteT0_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgRouteT0
If dg Is Nothing Then Exit Sub

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgRouteT0_HeadClick(ByVal ColIndex As Integer)

Dim dg As Object, rs As Object
Set dg = dgRouteT0: Set rs = rsRouteT0

If dg.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rs.Sort = dg.Columns(ColIndex).Caption & " DESC"
    dg.ClearSelCols
    intColumnIndex = 255

Else
    rs.Sort = dg.Columns(ColIndex).Caption
    dg.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgRouteT0_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'�O�_�����
If rsRouteT0 Is Nothing Then Exit Sub
If rsRouteT0.RecordCount = 0 Then Exit Sub

If blRouteT0Change = False Then Exit Sub

'���
If dgRouteT0.Col = 1 Then

    If rsRouteT0("���") = " " Then
    
        rsRouteT0("���") = "V"
        txtTotalPLT0 = Val(txtTotalPLT0) + rsRouteT0("�X�f�O��")
        txtTotalCST0 = Val(txtTotalCST0) + rsRouteT0("�X�f�c��")
        txtTotalCubeT0 = Val(txtTotalCubeT0) + rsRouteT0("�X�f���n")
        txtTotalWgtT0 = Val(txtTotalWgtT0) + rsRouteT0("�X�f���q")
    Else
        rsRouteT0("���") = " "
        txtTotalPLT0 = Val(txtTotalPLT0) - rsRouteT0("�X�f�O��")
        txtTotalCST0 = Val(txtTotalCST0) - rsRouteT0("�X�f�c��")
        txtTotalCubeT0 = Val(txtTotalCubeT0) - rsRouteT0("�X�f���n")
        txtTotalWgtT0 = Val(txtTotalWgtT0) - rsRouteT0("�X�f���q")
    
    End If
    
    dgRouteT0.Col = 0

End If

'�P�@����
If LastRow = Empty Then Exit Sub

Screen.MousePointer = 11

Frame13.Caption = rsRouteT0("�G�����s")
',��=(select top 1 case when rtrim(cod1.InvoicePCode) = 'N' then '��' else ' ' end from custorderdetail cod1 where rtrim(cod1.InvoicePCode) = 'N' and rtrim(cod1.ordertype) + rtrim(cod1.externorderkey) = t02t.extern and adddate > getdate()-90 )

str_SQL = "select ���u�s�� = t02t.Route_No,��f��� = rtrim(t02t.Arrive_Date) ,�q��s�� = Rtrim(t02t.Extern),��=case when t02t.STORERKEY = 'LMBO01' then (select distinct '��' from custorderdetail cod1 where rtrim(cod1.InvoicePCode) = 'N' and rtrim(cod1.ordertype) + rtrim(cod1.externorderkey) = t02t.extern ) else '' end " & _
    ",��� = isnull((select isnull(otqty,0) from trp02t (nolock) where receipt_no = t02t.Receipt_No),0) + isnull((select isnull(otqty,0) from ort02t (nolock) where receipt_no = t02t.Receipt_No),0) " & _
    ",���A = case when len(rtrim(Isnull(Rtrim(t02t.Confirm_Notes),''))) = 0 then (select case when sum(order_qty - ship_qty) <> 0 then '�X�f����' else ' ' end from sdn03t where receipt_no = t02t.Receipt_No) else Isnull(Rtrim(t02t.Confirm_Notes),'') end " & _
    ",�禬�渹=Isnull(rtrim(t02t.CustomerOrderkey1),''),TMS�渹 = rtrim(t02t.Receipt_No),�����N���f�� = isnull(o.cash,''),�ꦬ�N���f�� =  case when isnull(cast(o.receivecash as varchar),'0') = '0' then ' ' else cast(o.receivecash as varchar) end ,�p�O�N�X = rtrim(isnull(t9m.car_type,'')),���P���X = Rtrim(t01t.c_Vehicle_ID_No) " & _
    ",�r�p�H = Rtrim(t01t.driver),�@������ = Rtrim(t02t.Vehicle_ID_No),�f�D =Rtrim(t02t.StorerKey) + '_' + rtrim(t16.short_name),�q�����O = rtrim(t02t.priority) ,�t�e�ܧO = rtrim(o.facility)" & _
    ",�Ȥ�W�� = Rtrim(Isnull(t1m.Short_Name,'')),�q��Ƶ� = Rtrim(Isnull(t02t.Description,'')),�e�f�a�} = Rtrim(Isnull(t1m.Address,'')) " & _
    ",ñ��^�Ǥ��=isnull(t02t.SDNSendDate,getdate()) " & _
    "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no join orders o on o.orderkey = t02t.c_receipt_no " & _
    "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
    "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
    "left join trp09m t9m (nolock) on t9m.vehicle_id_no = t01t.c_vehicle_id_no " & _
    "left join trp08m t8m (nolock) on t8m.company_code = t9m.trp_company_code " & _
    "where 1 = 1 and t02t.c_Route_No = '" & rsRouteT0("�G�����s") & "' "
    
    If Len(RTrim(cboStorerT0)) > 0 Then str_SQL = str_SQL & "and t02t.StorerKey = '" & RTrim(cboStorerT0) & "' "
    If Len(RTrim(cboCarT0)) > 0 Then str_SQL = str_SQL & "and t01t.c_Vehicle_ID_No = '" & RTrim(cboCarT0) & "' "
    str_SQL = str_SQL & "order by t02t.StorerKey,t02t.Extern "

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    Screen.MousePointer = vbDefault
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧭q����"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Set dgOrderT0.DataSource = Nothing: Set rsOrderT0 = Nothing
    Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rsOrderT0)
tmp_Rs.Close: rsOrderT0.MoveFirst

Set dgOrderT0.DataSource = rsOrderT0

SetDataGridColWidth Me.Caption, dgOrderT0

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgOrderT0_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgOrderT0
If dg Is Nothing Then Exit Sub

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgOrderT0_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

txtCustomerOrderkey.Visible = False

''�P�@����
'If LastRow = Empty Then Exit Sub

'�O�_�����
If rsOrderT0 Is Nothing Then Exit Sub
If rsOrderT0.RecordCount = 0 Then Exit Sub
If rsOrderT0.EOF Then Exit Sub
If dgOrderT0.Col = -1 Then Exit Sub

If dgOrderT0.Col <> 0 And rsOrderT0.Fields(dgOrderT0.Col).Name <> "�禬�渹" And rsOrderT0.Fields(dgOrderT0.Col).Name <> "���A" And rsOrderT0.Fields(dgOrderT0.Col).Name <> "�ꦬ�N���f��" Then
    dgOrderT0.Col = 10
End If

If dgOrderT0.Col <> 0 And Len(RTrim(rsOrderT0("���A"))) > 1 Then
    dgOrderT0.Col = 10
    Exit Sub
End If


'Screen.MousePointer = 11
'rsMain.Fields(.Col).Name = "�w����f"

''�ȭ��S�w�f�D
If rsOrderT0.Fields(dgOrderT0.Col).Name = "���A" Then 'And (mySplit(rsOrderT0("�f�D"), "_", 0) = "LTKK01" Or mySplit(rsOrderT0("�f�D"), "_", 0) = "LVTL01" Or mySplit(rsOrderT0("�f�D"), "_", 0) = "LNSL01" Or mySplit(rsOrderT0("�f�D"), "_", 0) = "LABT01") Then
    If rsOrderT0("���A") = "V" Then
        rsOrderT0("���A") = " "
    Else
        rsOrderT0("���A") = "V"
    End If
    dgOrderT0.Col = 0

End If

'�禬�渹
If rsOrderT0.Fields(dgOrderT0.Col).Name = "�禬�渹" Then Call CustomerOrderkey("�禬�渹")

Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub CustomerOrderkey(Str_Field As String)

With dgOrderT0

    If Str_Field = "�禬�渹" Then
        txtCustomerOrderkey.Height = .RowHeight + 10
        If rsOrderT0.Fields(.Col).Name = "�禬�渹" Then
            If .Columns(.Col).Left > 0 Then
                    
                    txtCustomerOrderkey.Move .Left + .Columns(.Col).Left + 15, .Top + .RowTop(.Row), .Columns(.Col).Width
                    If txtCustomerOrderkey.Left + txtCustomerOrderkey.Width > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                        txtCustomerOrderkey.Width = txtCustomerOrderkey.Width + .Left + .Width - txtCustomerOrderkey.Left - txtCustomerOrderkey.Width
                    End If
                    txtCustomerOrderkey.Text = rsOrderT0("�禬�渹")  '��s�x�s�檺��
    
                    txtCustomerOrderkey.Visible = True
            Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
                txtCustomerOrderkey.Visible = False
            End If
        Else
            txtCustomerOrderkey.Visible = False
        End If
    
    End If

End With

End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�e�fñ��T�{"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 10000
dbsrcFormWidth = 15000

Me.Height = 9700: Me.Width = 15000
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
'Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2
'���i�f�D�渹�������i�ƨ��t�έq��
Call SetGridFormat_OneOrder_OrderDetail

'���i�f�D�渹�����h�i�ƨ��t�έq��
Call SetGridFormat_MultiOrder_OrderDetail

'���o�Ҧ� [���p�N�X] From LogicTown.dbo.TRP05M
str_SQL = "SELECT Rtrim(RSC_CODE)+' ' as RSC_CODE,RTRIM(isnull(DESCRIPTION,'')) AS 'Descr' FROM TRP05M Order by RSC_CODE"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then Exit Sub
iLoop = 0
cmb_OneOrder_RSCCode.Clear: cmb_MultiOrder_RSCCode.Clear
cmb_OneOrder_RSCCode.AddItem ""
cmb_MultiOrder_RSCCode.AddItem ""
Do While Not tmp_Rs.EOF
   cmb_OneOrder_RSCCode.AddItem tmp_Rs.Fields("RSC_CODE") & "  " & tmp_Rs.Fields("descr")
   cmb_MultiOrder_RSCCode.AddItem tmp_Rs.Fields("RSC_CODE") & "  " & tmp_Rs.Fields("descr")
   tmp_Rs.MoveNext
   iLoop = iLoop + 1
Loop
tmp_Rs.Close
Call ComboBox_SetWidth(cmb_OneOrder_RSCCode, 30)

'���o�Ҧ� [�d���k�ݥN�X] From LogicTown.dbo.TRP06M
str_SQL = "SELECT Rtrim(RBC_CODE)+' ' as 'RBC_CODE',RTRIM(isnull(Description,'')) AS 'Descr' FROM dbo.TRP06M Order by RBC_CODE"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then Exit Sub
iLoop = 0
cmb_OneOrder_RBCCode.Clear: cmb_OneOrder_RBCCode.AddItem ""
cmb_MultiOrder_RBCCode.Clear: cmb_MultiOrder_RBCCode.AddItem ""
Do While Not tmp_Rs.EOF
   cmb_OneOrder_RBCCode.AddItem tmp_Rs.Fields("RBC_CODE") & "  " & tmp_Rs.Fields("descr")
   cmb_MultiOrder_RBCCode.AddItem tmp_Rs.Fields("RBC_CODE") & "  " & tmp_Rs.Fields("descr")
   tmp_Rs.MoveNext
   iLoop = iLoop + 1
Loop

'�]�wdg_grid���榡
'Call SetGridFormat_SDN_Head
'Call SetGridFormat_SDN_Detail
'Call SetGridFormat_SDN_Cost
Call SetGridFormat_Tab2_SDN_Detail
Call SetGridFormat_Tab2_SDN_Cost
SSTab1.Tab = 3
Op_UnCheck.Visible = False
Op_OnCheck.Value = True
txt_DeliveryDate_Start.Text = Format(Now, "yyyymmdd")

cmbScan.AddItem "Y"
cmbScan.AddItem "N"
cmbScan.ListIndex = 0

cboInvBack.AddItem "Y"
cboInvBack.AddItem "N"

cmbOrderkey.AddItem ""
cmbOrderkey.AddItem "TMS�渹"
cmbOrderkey.AddItem "�f�D�渹"
cmbOrderkey.ListIndex = 0

tmp_Rs.Close

'������
str_SQL = "select distinct vehicle_id_no= rtrim(vehicle_id_no) from trp09m order by vehicle_id_no "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then tmp_Rs.MoveFirst

Do While Not tmp_Rs.EOF
    cboCar.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    cboCarT0.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    Car_Num.AddItem RTrim(tmp_Rs("vehicle_id_no"))
    tmp_Rs.MoveNext
Loop
cboCarT0 = ""
tmp_Rs.Close

'�f�D
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(rtrim(storerkey)) as storerkey from trp16M", cn, adOpenKeyset, adLockPessimistic

If Not tmp_Rs.EOF Then
    
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboStorerKey.AddItem tmp_Rs("storerkey")
        cboStorerT0.AddItem tmp_Rs("storerkey")
        tmp_Rs.MoveNext
    Next
    tmp_Rs.Close: Set tmp_Rs = Nothing
    cboStorerKey = ""
    cboStorerT0 = ""
End If

'���д����O
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select distinct costkind from trp17m order by costkind "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then

    If Not tmp_Rs.EOF Then tmp_Rs.MoveFirst
    
    Do While Not tmp_Rs.EOF
        cboCostkind.AddItem RTrim(tmp_Rs("costkind"))
        tmp_Rs.MoveNext
    Loop

End If

tmp_Rs.Close

txtDeliveryS = Format(Now - 30, "YYYYMMDD")
txtDeliveryE = Format(Now + 7, "YYYYMMDD")
txtDeliveryDateST0 = Format(Now - 1, "YYYYMMDD")
txtDeliveryDateET0 = Format(Now + 2, "YYYYMMDD")

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�d�I��Ӫ����L����ƥ�
'�γ~�G�ϥΪ̫��U Esc �h���Ǧ^������
'�B��������������
mvDate.Visible = False
If KeyCode = vbKeyEscape Then
   
   txt_OneOrder_SignQty.Visible = False
   cmb_OneOrder_RBCCode.Visible = False
   cmb_OneOrder_RSCCode.Visible = False
   
   txt_MultiOrder_SignQty.Visible = False
   cmb_MultiOrder_RBCCode.Visible = False
   cmb_MultiOrder_RSCCode.Visible = False
   
End If

End Sub

Private Sub Form_Resize()
On Error GoTo err_Handle

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��
If SSTab1.Height < 7000 And Me.ScaleHeight < 7000 Then Exit Sub

SSTab1.Height = Me.ScaleHeight: SSTab1.Width = Me.ScaleWidth

Frame10.Width = SSTab1.Width - 240: dgMain1.Width = Frame10.Width - 240
Frame12.Width = SSTab1.Width - 240: dgRouteT0.Width = Frame12.Width - 1380
Frame13.Width = SSTab1.Width - 240: dgOrderT0.Width = Frame13.Width - 240
fra_OneOrder_Detail.Width = SSTab1.Width - 240: gd_OneOrder_OrderDetail.Width = fra_OneOrder_Detail.Width - 240

Frame10.Height = SSTab1.Height - Frame10.Top - 120: dgMain1.Height = Frame10.Height - 360
dgRouteT0.Height = Frame12.Height - 720
Frame13.Height = SSTab1.Height - Frame14.Top - Frame14.Height - Frame12.Height - 120
If Frame13.Height > 840 Then dgOrderT0.Height = Frame13.Height - 840
fra_OneOrder_Detail.Height = SSTab1.Height - fra_OneOrder_Detail.Top - 120: gd_OneOrder_OrderDetail.Height = fra_OneOrder_Detail.Height - 360

Exit Sub
err_Handle:
'Call ErrorMsgbox(Me.Caption & "_Form_Resize", err.Number, err.Description, "")
End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_OP_SDNConfirm = Nothing
Set rsMain1 = Nothing
Set rsRouteT0 = Nothing
Set rsOrderT0 = Nothing
End Sub

Private Sub ClearForm()
'�M���Ҧ�����
Call ClearForm_AllField(Me)
blSDNConfirm = False
blCanUpdate = False
cmdCost.Enabled = False

'���i�f�D�渹�������i�ƨ��t�έq��
cmd_OneOrder_Deliveryok.Enabled = False
cmd_OneOrder_Expect.Enabled = False
cmd_OneOrder_NoDelivery.Enabled = False
Call SetGridFormat_OneOrder_OrderDetail

'���i�f�D�渹�����h�i�ƨ��t�έq��
cmd_MultiOrder_Deliveryok.Enabled = False
cmd_MultiOrder_Expect.Enabled = False
cmd_MultiOrder_NoDelivery.Enabled = False
Call SetGridFormat_MultiOrder_OrderDetail
Set dg_MultiOrder.DataSource = Nothing
Set rs_MultiOrder = Nothing

End Sub

Private Sub Display_OrderData_OneReceipNo(ByVal strExtern As String)
'���i�f�D�渹�������i�ƨ��t�έq��G�q���Ƭd��
Screen.MousePointer = vbHourglass
fra_OneOrder_Header.Visible = True
fra_OneOrder_Detail.Visible = True
fra_MultiOrder_Header.Visible = False
fra_MultiOrder_Detail.Visible = False
cmdCost.Enabled = True
txt_OneOrder_Status.BackColor = "&H80000000"

On Error GoTo err_Handle

'str_SQL = "Select * From SDNConfirm_OrderDate_One Where �q��s�� = '" & strExtern & "'"
str_SQL = "select �G�����s = t02t.c_Route_No " & _
            ",���u�s�� = t02t.Route_No,���P���X = Rtrim(t01t.c_Vehicle_ID_No),�@������ = Rtrim(t02t.Vehicle_ID_No) " & _
            ",�r�p�H = Rtrim(t01t.driver),�f�B���q = Isnull(Rtrim(t8m.Short_Name),'') " & _
            ",�X����� = convert(varchar,t01t.Delivery_Date,112),�f�D = Rtrim(t02t.StorerKey) " & _
            ",�f�D�W�� =rtrim(t16.c_name),���� = Rtrim(Isnull(t02t.Description,'')),�Ȥ�s�� = Rtrim(t02t.ConsigneeKey) " & _
            ",�Ȥ�W�� = Rtrim(Isnull(t1m.Short_Name,'')),�l���ϸ� = Rtrim(Isnull(t1m.zip,'')),�e�f�a�} = Rtrim(Isnull(t1m.Address,'')) " & _
            ",�q�����O = rtrim(t02t.priority),�q��s�� = rtrim(t02t.Receipt_No),C_Receipt_no = rtrim(t02t.C_Receipt_No),�q���� = rtrim(t02t.Receipt_Date) " & _
            ",��f��� = rtrim(t02t.Arrive_Date),�f�D�渹 = Rtrim(t02t.Extern) " & _
            ",ñ���� = isnull(t02t.CustSignDate,isnull(t02t.SCHEDULEDATE,Arrive_Date)),�t�Τ�� = Convert(varchar,Getdate(),112) " & _
            ",���A = Isnull(Rtrim(t02t.Confirm_Notes),''),�Ȥ��禬�渹=Isnull(rtrim(t02t.CustomerOrderkey1),'') " & _
            ",���y=Isnull(rtrim(t02t.Scan),''),ñ��^�Ǥ��=isnull(t02t.SDNSendDate,getdate()),�Ȥ�^�гB�z�覡=Isnull(rtrim(t02t.CUST_Handle),'') " & _
            ",����B�z�覡=Isnull(rtrim(t02t.TRP_Handle),''),�ﵽ�覡=Isnull(rtrim(t02t.Advance),''),�w�s�վ�覡=Isnull(rtrim(t02t.INV_Handle),'') " & _
            ",�t�e�O=Isnull(rtrim(t02t.TRP_Cost),0),�z�f�O=Isnull(rtrim(t02t.Sorting_Cost),0),���`�O�ΦX�p=Isnull(rtrim(t02t.Total_Cost),'') " & _
            ",ñ��Ƶ�=Isnull(rtrim(t02t.sdn_note),''),�J�w����=Isnull(rtrim(ExpectReceiptOK),''),�o���^��=invBack " & _
            ",��f = isnull(ontimedelivery,0),ñ��w�^ = t02t.sdnback,�����q�N�X = rtrim(isnull(co.branchid,'')),�Ȥ��O=rtrim(isnull(co.ordertype,'')),�Ӽh=isnull(co.Stairs,''),�p�O���O=isnull(t9m.car_type,'') " & _
            "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
            "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
            "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
            "left join trp09m t9m (nolock) on t9m.vehicle_id_no = t01t.c_vehicle_id_no " & _
            "left join trp08m t8m (nolock) on t8m.company_code = t9m.trp_company_code " & _
            "left join custorders co (nolock) on co.orderkey = t02t.c_receipt_no " & _
            "where t02t.Receipt_No = '" & strExtern & "' "

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

txt_C_ROUTE_NO.Text = tmp_Rs.Fields("�G�����s").Value
txt_OneOrder_RouteNo.Text = tmp_Rs.Fields("���u�s��").Value
txt_OneOrder_VehicleID.Text = tmp_Rs.Fields("���P���X").Value & "_" & tmp_Rs.Fields("�@������").Value
txt_OneOrder_Driver.Text = tmp_Rs.Fields("�r�p�H").Value & ""
txt_OneOrder_TRPCompany.Text = tmp_Rs.Fields("�f�B���q").Value & ""
txt_OneOrder_DeliveryDate.Text = tmp_Rs.Fields("�X�����").Value & ""
txt_OneOrder_StorerKey.Text = tmp_Rs.Fields("�f�D").Value & ""
txt_Storer.Text = tmp_Rs.Fields("�f�D�W��").Value & ""
txt_OneOrder_Description.Text = tmp_Rs.Fields("����").Value
txt_C_Receipt_No = tmp_Rs.Fields("C_Receipt_no").Value
txt_OneOrder_OrderKey.Text = tmp_Rs.Fields("�q��s��").Value
txt_OneOrder_OrderDate.Text = tmp_Rs.Fields("�q����").Value & ""
txt_OneOrder_ArriveDate.Text = tmp_Rs.Fields("��f���").Value
txt_OneOrder_Status.Text = tmp_Rs.Fields("���A").Value
txt_OneOrder_StorerOrderKey.Text = tmp_Rs.Fields("�f�D�渹").Value
txt_OneOrder_CustomerOrderkey1.Text = tmp_Rs.Fields("�Ȥ��禬�渹").Value & ""
cmbScan.Text = tmp_Rs.Fields("���y").Value: If cmbScan.Text <> "Y" Then cmbScan.Text = "N"
dtpSDNSendDate.Value = tmp_Rs.Fields("ñ��^�Ǥ��").Value
If tmp_Rs("��f") = "5" Then txt_OneOrder_Status.BackColor = "&H00C0C0FF"
If tmp_Rs("��f") = "9" Then txt_OneOrder_Status.BackColor = "&H00C0FFC0"
dtp_OneOrder_SignDate.Value = tmp_Rs.Fields("ñ����").Value
txt_CustHandle.Text = tmp_Rs.Fields("�Ȥ�^�гB�z�覡").Value
txt_TRPHandle.Text = tmp_Rs.Fields("����B�z�覡").Value
txt_Advance.Text = tmp_Rs.Fields("�ﵽ�覡").Value
txt_INVHandle.Text = tmp_Rs.Fields("�w�s�վ�覡").Value
txt_TRPCost.Text = tmp_Rs.Fields("�t�e�O").Value
txt_SortingCost.Text = tmp_Rs.Fields("�z�f�O").Value
txt_TotalCost.Text = tmp_Rs.Fields("���`�O�ΦX�p").Value
txt_SDNNote.Text = tmp_Rs("ñ��Ƶ�")
txt_Priority.Text = tmp_Rs.Fields("�q�����O").Value
cboInvBack.Text = tmp_Rs.Fields("�o���^��").Value
txt_Externordertype.Text = tmp_Rs.Fields("�Ȥ��O").Value
txt_BranchId.Text = tmp_Rs.Fields("�����q�N�X").Value
txt_Stairs = tmp_Rs.Fields("�Ӽh").Value
txt_Cartype.Text = tmp_Rs.Fields("�p�O���O").Value

If Trim(tmp_Rs("�l���ϸ�")) = "" Then MsgBox "�`�N!�l���ϸ��ť�!", 16, Me.Caption

If txt_Priority = "R" Or txt_Priority = "RC" Or txt_Priority = "A2B" Then
    txt_OneOrder_ConsigneeKey.Text = tmp_Rs.Fields("�Ȥ�s��")
    txt_ZIP.Text = tmp_Rs("�l���ϸ�")
    txt_OneOrder_FullName.Text = tmp_Rs.Fields("�Ȥ�W��")
    txt_OneOrder_Address.Text = tmp_Rs.Fields("�e�f�a�}")
    txt_OneOrder_ConsigneeKey1 = ""
    txt_Zip1 = ""
    txt_OneOrder_FullName1 = ""
    txt_OneOrder_Address1 = ""
Else
    txt_OneOrder_ConsigneeKey = ""
    txt_ZIP = ""
    txt_OneOrder_FullName = ""
    txt_OneOrder_Address = ""
    txt_OneOrder_ConsigneeKey1 = tmp_Rs.Fields("�Ȥ�s��")
    txt_Zip1 = tmp_Rs.Fields("�l���ϸ�")
    txt_OneOrder_FullName1 = tmp_Rs.Fields("�Ȥ�W��")
    txt_OneOrder_Address1 = tmp_Rs.Fields("�e�f�a�}")
End If

If tmp_Rs.Fields("ñ��w�^").Value = "1" Then
    cmdSDNBack.BackColor = vbGreen
    cmdSDNBack.Caption = "ñ��w�^"
Else
    cmdSDNBack.BackColor = vbRed
    cmdSDNBack.Caption = "ñ�楼�^"
End If

blShipped = True    '�L�k�P�_�q��z�f�q�O�_�w��s(Ship_Qty)
blCanUpdate = True
blSDNConfirm = False
blCanUpdate = True

tmp_Rs.Close

'���t�e�ܧO
Call ReDim_Recordset(tmp_Rs)
str_SQL = "select " & _
          "UrgentMark = isnull(o.Urgent_Mark,''),ReserveMark = isnull(o.Reserve_Mark,''), " & _
          "����=o.cash,�ꦬ=case when isnull(rtrim(cast(o.receiveCash as char)),'') = '0' then ' ' else  isnull(rtrim(cast(o.receiveCash as char)),'') end ," & _
          "�ճ����� = rtrim(isnull(o.B_city,'')) ," & _
          "priority = rtrim(o.priority),facility=case when o.facility = '' then '�ըƹF�_��' else o.facility end,�Ȥ�s��=rtrim(t1.consigneekey) , �Ȥ�W��=rtrim(t1.full_name) , �l���ϸ�=rtrim(isnull(t1.zip,'')) , �e�f�a�}=rtrim(t1.address) from orders o (nolock) join trp01m t1 (nolock) on t1.storerkey = o.storerkey and t1.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end where o.orderkey = '" & txt_C_Receipt_No & "' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

'�ɸ��
    txt_UrgentMark = tmp_Rs.Fields("UrgentMark")
    txt_ReserveMark = tmp_Rs.Fields("ReserveMark")
    txt_Cash.Text = tmp_Rs.Fields("����")
    txt_ReceiveCash.Text = tmp_Rs.Fields("�ꦬ")
    cbx_B_city.Text = tmp_Rs.Fields("�ճ�����")

If Trim(tmp_Rs("�l���ϸ�")) = "" Then MsgBox "�`�N!�l���ϸ��ť�!", 16, Me.Caption

If tmp_Rs("Priority") = "A2B" Then
    txt_OneOrder_ConsigneeKey1 = tmp_Rs.Fields("�Ȥ�s��")
    txt_OneOrder_FullName1 = tmp_Rs.Fields("�Ȥ�W��")
    txt_Zip1 = tmp_Rs("�l���ϸ�")
    txt_OneOrder_Address1 = tmp_Rs.Fields("�e�f�a�}")
ElseIf tmp_Rs("Priority") = "R" Or tmp_Rs("Priority") = "RC" Then
    txt_OneOrder_ConsigneeKey1 = ""
    txt_Zip1 = ""
    txt_OneOrder_FullName1 = tmp_Rs("facility")
    txt_OneOrder_Address1 = ""
Else
    txt_OneOrder_ConsigneeKey = ""
    txt_ZIP = ""
    txt_OneOrder_FullName = tmp_Rs("facility")
    txt_OneOrder_Address = ""
End If

tmp_Rs.Close

'ñ����@����
If Val(txt_OneOrder_ArriveDate) > lngDueDate Then
    txt_OneOrder_Status.Enabled = True
    cmdCarNOChange.Enabled = True: cmdUnRouteConfirm.Enabled = True: cmdShipNotes.Enabled = True
Else
    txt_OneOrder_Status.Enabled = False
    cmdCarNOChange.Enabled = False: cmdUnRouteConfirm.Enabled = False: cmdShipNotes.Enabled = False
End If

'���i�f�D�渹�������i�ƨ��t�έq��G�q��W��
Call SetGridFormat_OneOrder_OrderDetail
Dim tmpI As Double, bl0Qty As Boolean

bl0Qty = False

str_SQL = "Select Rtrim(t02t.Extern) As �f�D�渹 " & _
    ",rtrim(t02t.receipt_no) as �q�渹�X " & _
    ",t03t.Seq_No as ���� " & _
    ",Rtrim(t03t.Product_No) as �f�� " & _
    ",Rtrim(Isnull(sku.Descr,'')) as �~�W " & _
    ",���=isnull(sku.busr1,'EA') " & _
    ",t03t.Order_Qty as �q��q " & _
    ",Isnull(t03t.Ship_Qty,0) as �e�f�q " & _
    ",Isnull(t03t.Sign_Qty,0) as ñ��q " & _
    ",Isnull(Rtrim(t03t.RSC_Code) + '  ' + Rtrim(t05m.Description),'  ') as ���`��] " & _
    ",Isnull(Rtrim(t03t.RBC_Code) + '  ' + Rtrim(t06m.Description),'  ') as �d���k�� " & _
    ",Isnull(Rtrim(t03t.RSC_Code),'  ') as ���`�X " & _
    ",Isnull(Rtrim(t03t.RBC_Code),'  ') as �d�ݽX " & _
    ",�c�]�ഫ�v = isnull(sku.casecnt,'') " & _
    ",��쭫�q = round(isnull(sku.stdgrosswgt,0),9) " & _
    ",�����n = round(isnull(sku.stdcube,0),9) " & _
    ",�d�ݤH = isnull(responsible,'') " & _
    "From SDN02T t02t (nolock) join SDN03T t03t (nolock) on t03t.Receipt_No = t02t.Receipt_No " & _
    "join gv_SKUxpack sku(nolock) on sku.StorerKey = t03t.StorerKey and sku.SKU = t03t.Product_No " & _
    "Left join TRP05M t05m(nolock) on t05m.RSC_Code = t03t.RSC_Code " & _
    "Left join TRP06M t06m(nolock) on t06m.RBC_Code = t03t.RBC_Code " & _
    "where t02t.receipt_no = '" & strExtern & "' order by t03t.Seq_No "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
tmpI = 1
Do While Not tmp_Rs.EOF
   If tmp_Rs.Fields("�e�f�q").Value = 0 Then bl0Qty = True
   With gd_OneOrder_OrderDetail
        If .Rows < (tmpI + 1) Then .Rows = .Rows + 1
        .Row = tmpI
        .Col = 0: .Text = RTrim(tmp_Rs.Fields("����").Value)
        .Col = 1: .Text = tmp_Rs.Fields("�f��").Value
        .Col = 2: .Text = tmp_Rs.Fields("�~�W").Value
        .Col = 3: .Text = RTrim(tmp_Rs.Fields("���").Value)
        .Col = 4: .Text = tmp_Rs.Fields("�q��q").Value
        .Col = 5: .Text = tmp_Rs.Fields("�e�f�q").Value
        'mark by gemini
'         If blCanUpdate Then    '�|������ SDN Confirmed
'            .Col = 6: .Text = 0   'ñ��q
'            .Col = 7: .Text = ""
'            .Col = 8: .Text = ""
'            .Col = 9: .Text = ""
'            .Col = 10: .Text = ""
'         Else
            .Col = 6: .Text = tmp_Rs.Fields("ñ��q").Value
            .Col = 7: .Text = tmp_Rs.Fields("���`��]").Value
            .Col = 8: .Text = tmp_Rs.Fields("�d���k��").Value
            .Col = 9: .Text = tmp_Rs.Fields("���`�X").Value
            .Col = 10: .Text = tmp_Rs.Fields("�d�ݽX").Value
            .Col = 11: .Text = tmp_Rs.Fields("�c�]�ഫ�v").Value
            .Col = 12: .Text = tmp_Rs.Fields("��쭫�q").Value
            .Col = 13: .Text = tmp_Rs.Fields("�����n").Value
            .Col = 14: .Text = tmp_Rs.Fields("�d�ݤH").Value
            
'        End If
   End With
   tmp_Rs.MoveNext
   tmpI = tmpI + 1
Loop
tmp_Rs.Close

'��s�X�f�q
'If bl0Qty = True Then Call Ship2TMS(strExtern) 'Mark @20130819 4 ��ܪ��X�f�q(�{��O�t��)�w�q����ڥX�f�q(�Ҷq�ܮw�z�f���T�ʡA���i��ܮw�X�h�A�p�O�ݤ֭p)

If blCanUpdate Then
    cmd_OneOrder_Deliveryok.Enabled = True
    cmd_OneOrder_Expect.Enabled = True
    cmd_OneOrder_NoDelivery.Enabled = True
Else
    cmd_OneOrder_Deliveryok.Enabled = False
    cmd_OneOrder_Expect.Enabled = False
    cmd_OneOrder_NoDelivery.Enabled = False
End If

'�S�w�H���i���� SDN Confirm ���s�s��
'�v���]�w�x�s�� CodeLKUP ListName = [SDNRECONDURM]
'  Code�GUser_ID   Short�G�v���]�w 1-�i�H���ư���
'If (Not blCanUpdate) And CheckSDNReConfirm(User_id) Then
'    '���\���ư��� SDN Confirm ���ϥΪ̡A�}�� SDN Confirm ������ק�
'    cmd_OneOrder_Deliveryok.Enabled = True
'    cmd_OneOrder_Expect.Enabled = True
'    cmd_OneOrder_NoDelivery.Enabled = True
'    blCanUpdate = True
'End If

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�q��d��", Me.Caption, "Form ���� SubProgram Display_OrderData_OneReceiptNo", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub SetGridFormat_OneOrder_OrderDetail()
'�W�١GSetGridFormatt_OneOrder_OrderDetail
'���O�G�Ƶ{��
'�\��G�M���ó]�w [SDN Confirm] ��� [���i�f�D�渹�������i�ƨ��t�έq��] �q��W����ܮ榡
'�ѼơG�ǤJ�ȡG�L
Dim sub_var1 As Integer, sub_var2 As Integer
gd_OneOrder_OrderDetail.Visible = False
With gd_OneOrder_OrderDetail
     .Rows = 2: .FixedRows = 1: .Cols = 15
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
     .ColWidth(1) = 1800
     .ColWidth(2) = 3000
     .ColWidth(3) = 500
     .ColWidth(4) = 750
     .ColWidth(5) = 750
     .ColWidth(6) = 750
     .ColWidth(7) = 1500
     .ColWidth(8) = 1000
     .ColWidth(9) = 600
     .ColWidth(10) = 600
     .ColWidth(11) = 500
     .ColWidth(12) = 1200
     .ColWidth(13) = 1200
     .ColWidth(14) = 1000
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "����"
     .Col = 1: .Text = "�f��"
     .Col = 2: .Text = "����~�W"
     .Col = 3: .Text = "���"
     .Col = 4: .Text = "�q��q"
     .Col = 5: .Text = "�e�f�q"
     .Col = 6: .Text = "ñ��q"
     .Col = 7: .Text = "���`��]"
     .Col = 8: .Text = "�d���k��"
     .Col = 9: .Text = "���`�X"
     .Col = 10: .Text = "�d�ݽX"
     .Col = 11: .Text = "�C�c"
     .Col = 12: .Text = "��쭫"
     .Col = 13: .Text = "����"
     .Col = 14: .Text = "�d�ݤH"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignLeftCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .ColAlignment(10) = flexAlignCenterCenter
     .ColAlignment(11) = flexAlignRightCenter
     .ColAlignment(12) = flexAlignLeftCenter
     .ColAlignment(13) = flexAlignLeftCenter
     .ColAlignment(14) = flexAlignLeftCenter
     .Rows = 2
     .Row = 0
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .Text = ""
     Next sub_var1
     
End With
gd_OneOrder_OrderDetail.Visible = True
End Sub

Private Sub HideGridUseObject_OneOrder()
'���i�f�D�渹�������i�ƨ��t�έq��G���� [ñ��q] [���`��]] [�d��] ���
txt_OneOrder_SignQty.Visible = False
cmb_OneOrder_RBCCode.Visible = False
cmb_OneOrder_RSCCode.Visible = False
End Sub

Private Sub gd_MultiOrder_OrderDetail_Click()
'���i�f�D�渹�����h�i�ƨ��t�έq��G�q��W�ӿ��
Dim SelectedCol As Integer, SelectedRow As Integer
If Not blCanUpdate Then Exit Sub
Call HideGridUseObject_MultiOrder
On Error Resume Next
With gd_MultiOrder_OrderDetail
     SelectedCol = .Col: SelectedRow = .Row
     Select Case SelectedCol
       Case 5      'ñ��q
           txt_MultiOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_MultiOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_MultiOrder_SignQty.Left = txt_MultiOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_MultiOrder_SignQty.Top = txt_MultiOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_MultiOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_MultiOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_MultiOrder_SignQty.Text = .Text
           txt_MultiOrder_SignQty.Visible = True
           txt_MultiOrder_SignQty.SelStart = 0: txt_MultiOrder_SignQty.SelLength = Len(txt_MultiOrder_SignQty.Text)
           txt_MultiOrder_SignQty.SetFocus
           
       Case 6      '���`��]
           cmb_MultiOrder_RSCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_MultiOrder_RSCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_MultiOrder_RSCCode.Left = cmb_MultiOrder_RSCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_MultiOrder_RSCCode.Top = cmb_MultiOrder_RSCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_MultiOrder_RSCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_MultiOrder_RSCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_MultiOrder_RSCCode.ListCount - 1
                  If Left(cmb_MultiOrder_RSCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_MultiOrder_RSCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_MultiOrder_RSCCode.Visible = True
           cmb_MultiOrder_RSCCode.SetFocus
           SendKeys "%{DOWN}"
           
       Case 7      '�v�d�Ϥ��G�d��
           cmb_MultiOrder_RBCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_MultiOrder_RBCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_MultiOrder_RBCCode.Left = cmb_MultiOrder_RBCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_MultiOrder_RBCCode.Top = cmb_MultiOrder_RBCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_MultiOrder_RBCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_MultiOrder_RBCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_MultiOrder_RBCCode.ListCount - 1
                  If Left(cmb_MultiOrder_RBCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_MultiOrder_RBCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_MultiOrder_RBCCode.Visible = True
           cmb_MultiOrder_RBCCode.SetFocus
           SendKeys "%{DOWN}"
           
     End Select
End With

End Sub

Private Sub gd_OneOrder_OrderDetail_Click()
'���i�f�D�渹�������i�ƨ��t�έq��G�q��W�ӿ��
Dim SelectedCol As Integer, SelectedRow As Integer
If Not blCanUpdate Then Exit Sub
cmb_OneOrder_RSCCode.Visible = False: cmb_OneOrder_RSCCode.Visible = False
Call HideGridUseObject_OneOrder
On Error Resume Next
With gd_OneOrder_OrderDetail
     SelectedCol = .Col: SelectedRow = .Row
     Select Case SelectedCol
        Case 5      '�X�f�q
           txt_OneOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_OneOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_OneOrder_SignQty.Left = txt_OneOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_OneOrder_SignQty.Top = txt_OneOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_OneOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_OneOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_OneOrder_SignQty.Text = .Text
           txt_OneOrder_SignQty.Visible = True
           txt_OneOrder_SignQty.SelStart = 0: txt_OneOrder_SignQty.SelLength = Len(txt_OneOrder_SignQty.Text)
           txt_OneOrder_SignQty.SetFocus
       Case 6      'ñ��q
           txt_OneOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_OneOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_OneOrder_SignQty.Left = txt_OneOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_OneOrder_SignQty.Top = txt_OneOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_OneOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_OneOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_OneOrder_SignQty.Text = .Text
           txt_OneOrder_SignQty.Visible = True
           txt_OneOrder_SignQty.SelStart = 0: txt_OneOrder_SignQty.SelLength = Len(txt_OneOrder_SignQty.Text)
           txt_OneOrder_SignQty.SetFocus
           
       Case 7      '���`��]
       DoEvents: DoEvents
           cmb_OneOrder_RSCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_OneOrder_RSCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_OneOrder_RSCCode.Left = cmb_OneOrder_RSCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_OneOrder_RSCCode.Top = cmb_OneOrder_RSCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_OneOrder_RSCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_OneOrder_RSCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_OneOrder_RSCCode.ListCount - 1
                  If Left(cmb_OneOrder_RSCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_OneOrder_RSCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_OneOrder_RSCCode.Visible = True
           cmb_OneOrder_RSCCode.SetFocus
           SendKeys "%{DOWN}"
           
       Case 8      '�v�d�Ϥ��G�d��
       DoEvents: DoEvents
           cmb_OneOrder_RBCCode.Left = .Left + .ColPos(SelectedCol)
           cmb_OneOrder_RBCCode.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              cmb_OneOrder_RBCCode.Left = cmb_OneOrder_RBCCode.Left + 2 * Screen.TwipsPerPixelX
              cmb_OneOrder_RBCCode.Top = cmb_OneOrder_RBCCode.Top + 2 * Screen.TwipsPerPixelY
           End If
           cmb_OneOrder_RBCCode.Width = .ColWidth(SelectedCol)
           If Len(Trim(.Text)) = 0 Then
              cmb_OneOrder_RBCCode.ListIndex = 0
           Else
              For iLoop = 0 To cmb_OneOrder_RBCCode.ListCount - 1
                  If Left(cmb_OneOrder_RBCCode.List(iLoop), 2) = Left(.Text, 2) Then
                     cmb_OneOrder_RBCCode.ListIndex = iLoop
                     Exit For
                  End If
              Next iLoop
           End If
           cmb_OneOrder_RBCCode.Visible = True
           cmb_OneOrder_RBCCode.SetFocus
           SendKeys "%{DOWN}"
           
       Case 14      '�v�d�Ϥ��G�d�ݤH
           txt_OneOrder_SignQty.Left = .Left + .ColPos(SelectedCol)
           txt_OneOrder_SignQty.Top = .Top + .RowPos(SelectedRow)
           If .Appearance = 1 Then
              txt_OneOrder_SignQty.Left = txt_OneOrder_SignQty.Left + 2 * Screen.TwipsPerPixelX
              txt_OneOrder_SignQty.Top = txt_OneOrder_SignQty.Top + 2 * Screen.TwipsPerPixelY
           End If
           txt_OneOrder_SignQty.Width = .ColWidth(SelectedCol)
           txt_OneOrder_SignQty.Height = .RowHeight(SelectedRow)
           txt_OneOrder_SignQty.Text = .Text
           txt_OneOrder_SignQty.Visible = True
           txt_OneOrder_SignQty.SelStart = 0: txt_OneOrder_SignQty.SelLength = Len(txt_OneOrder_SignQty.Text)
           txt_OneOrder_SignQty.SetFocus
           
     End Select
End With

End Sub

Private Function CheckSDNReConfirm(ByVal strUserID As String) As Boolean
'�S�w�H���i���� SDN Confirm ���s�s��
'�v���]�w�x�s�� CodeLKUP ListName = [SDNRECONDURM]
'  Code�GUser_ID   Short�G�v���]�w 1-�i�H���ư���
CheckSDNReConfirm = False
str_SQL = "Select Isnull(Rtrim(Short),'0') as CheckFlag From CodeLKUP " & _
          "Where ListName = 'SDNRECONFIRM' AND Code = '" & strUserID & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   If tmp_Rs.Fields("CheckFlag").Value = "1" Then
      CheckSDNReConfirm = True
   End If
End If
tmp_Rs.Close
End Function

'Private Sub Lb_Route_Change()
'    cmdOTQtyFix.Enabled = False
'    If Left(Lb_Route.Caption, 1) = "R" Then cmdOTQtyFix.Enabled = True
'End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")

'������
Select Case mvDate.Tag
    Case "�Ȥ�ñ�����E�h��"
         txt_MultiOrder_SignDate.Text = Format(mvDate.Value, "yyyymmdd")
'    Case "�Ȥ�ñ�����E�@��"
'         txt_OneOrder_SignDate.Text = Format(mvDate.Value, "yyyymmdd") by gemini
    Case "Tab0-�X����_"
         txt_DeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
    Case "Tab0-�X���騴"
         txt_DeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
    Case "Tab2-�X����"
         txt_Tab02_Delivery_Date.Text = Format(mvDate.Value, "yyyymmdd")

End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

'Private Sub Op_CBM_Click()
'    dg_SDN_Detail.Row = 1
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix((Val(Trim(txt_Tab0_srcTotal_Volumn.Text)) * 10) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "���n"
'    dg_SDN_Cost.Col = 2: dg_SDN_Detail.Col = 2: dg_SDN_Cost.Text = Trim(dg_SDN_Detail.Text)
'    Call Cost_SumAll
'End Sub
'
'Private Sub Op_CS_Click()
'    dg_SDN_Detail.Row = 1
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(txt_Tab0_srcTotal_Case.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "�c��"
'    dg_SDN_Cost.Col = 2: dg_SDN_Detail.Col = 2: dg_SDN_Cost.Text = Trim(dg_SDN_Detail.Text) 'Fix((a * b) + 0.5)
'    Call Cost_SumAll
'End Sub

Private Sub Op_OnCheck_Click()
    ck_confirm.Visible = True
    ck_confirm.Value = 1
    ck_back.Visible = True
  
End Sub

'Private Sub Op_SumCBM_Click()
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix((Val(Trim(Me.txt_Tab0_sum_CBM.Text)) * 10) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "���n"
'    dg_SDN_Cost.Col = 2: dg_SDN_Cost.Text = Trim(txt_Tab0_Route_No.Text)
'    Call Cost_SumAll
'End Sub
'
'Private Sub Op_SumCS_Click()
'
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(Me.txt_Tab0_sum_Case.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "�c��"
'    dg_SDN_Cost.Col = 2: dg_SDN_Cost.Text = Trim(txt_Tab0_Route_No.Text)
'    Call Cost_SumAll
'End Sub
'
'
'Private Sub Op_SumWT_Click()
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(Me.txt_Tab0_sum_WT.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "���q"
'    dg_SDN_Cost.Col = 2: dg_SDN_Cost.Text = Trim(txt_Tab0_Route_No.Text)
'    Call Cost_SumAll
'End Sub

Private Sub Op_Tab2_CBM_Click()
    dg_Tab2_SDN_Cost.Row = 1
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_srcTotal_Volumn.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "���n"
    dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Detail.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(dg_Tab2_SDN_Detail.Text)
    Call Cost_Tab2_SumAll
End Sub

Private Sub Op_Tab2_CS_Click()
    dg_Tab2_SDN_Cost.Row = 1
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_srcTotal_Case.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "�c��"
    dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Detail.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(dg_Tab2_SDN_Detail.Text)
    Call Cost_Tab2_SumAll
End Sub

Private Sub Op_Tab2_SumCBM_Click()
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(Me.txt_Tab2_sum_CBM.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "���n"
    'dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_Route_NO.Text)
    Call Cost_SumAll
End Sub

Private Sub Op_Tab2_SumCS_Click()
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(Me.txt_Tab2_sum_Case.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "�c��"
    'dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_Route_NO.Text)
    Call Cost_SumAll
End Sub

Private Sub Op_Tab2_SumWT_Click()
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(Me.txt_Tab0_sum_WT.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "���q"
    'dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_Route_NO.Text)
    Call Cost_SumAll
End Sub

Private Sub Op_Tab2_WT_Click()
    dg_Tab2_SDN_Cost.Row = 1
    dg_Tab2_SDN_Cost.Col = 6
    dg_Tab2_SDN_Cost.Text = Trim(txt_Tab2_srcTotal_Weight.Text)
    dg_Tab2_SDN_Cost.Col = 5
    dg_Tab2_SDN_Cost.Text = "���q"
    dg_Tab2_SDN_Cost.Col = 2: dg_Tab2_SDN_Detail.Col = 2: dg_Tab2_SDN_Cost.Text = Trim(dg_Tab2_SDN_Detail.Text)
    Call Cost_Tab2_SumAll
End Sub

Private Sub Op_UnCheck_Click()
    ck_confirm.Visible = False
    ck_confirm.Value = 0
    ck_back.Visible = False
    ck_back.Value = 0
End Sub

'Private Sub OpWT_Click()
'    dg_SDN_Detail.Row = 1
'    dg_SDN_Cost.Col = 6
'    dg_SDN_Cost.Text = Fix(Trim(txt_Tab0_srcTotal_Weight.Text) + 0.5)
'    dg_SDN_Cost.Col = 5
'    dg_SDN_Cost.Text = "���q"
'    dg_SDN_Cost.Col = 2: dg_SDN_Detail.Col = 2: dg_SDN_Cost.Text = Trim(dg_SDN_Detail.Text)
'    Call Cost_SumAll
'End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.mvDate.Visible = False
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_DeliveryDate_End_Click()
    'Tab0-�X���騴
    If Trim(txt_DeliveryDate_End.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_DeliveryDate_End.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_DeliveryDate_End.Text, 4) & "/" & Mid(txt_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_DeliveryDate_End.Text, 2))
       End If
    End If
    mvDate.Left = txt_DeliveryDate_End.Left + txt_DeliveryDate_End.Left
    mvDate.Top = txt_DeliveryDate_End.Top + txt_DeliveryDate_End.Top + txt_DeliveryDate_End.Height
    mvDate.Tag = "Tab0-�X���騴"
    mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_Start_Click()
    'Tab0-�X����_
    If Trim(txt_DeliveryDate_Start.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_DeliveryDate_Start.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_DeliveryDate_Start.Text, 2))
       End If
    End If
    mvDate.Left = txt_DeliveryDate_Start.Left + txt_DeliveryDate_Start.Left
    mvDate.Top = txt_DeliveryDate_Start.Top + txt_DeliveryDate_Start.Top + txt_DeliveryDate_Start.Height
    mvDate.Tag = "Tab0-�X����_"
    mvDate.Visible = True
End Sub

Private Sub txt_ExternOrderKey_KeyPress(KeyAscii As Integer)
'�f�D�渹
Select Case KeyAscii
       Case vbKeyReturn
            cmd_OrderQuery.SetFocus
End Select
End Sub


Private Sub txt_MultiOrder_SignDate_Click()
'�Ȥ�ñ�����G���i�f�D�渹�����h�i�ƨ��t�έq��
If Trim(txt_MultiOrder_SignDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_MultiOrder_SignDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_MultiOrder_SignDate.Text, 4) & "/" & Mid(txt_MultiOrder_SignDate.Text, 5, 2) & "/" & Right(txt_MultiOrder_SignDate.Text, 2))
   End If
End If
mvDate.Left = fra_MultiOrder_Header.Left + txt_MultiOrder_SignDate.Left - (mvDate.Width - txt_MultiOrder_SignDate.Width)
mvDate.Top = fra_MultiOrder_Header.Top + txt_MultiOrder_SignDate.Top + txt_MultiOrder_SignDate.Height
mvDate.Tag = "�Ȥ�ñ�����E�h��"
mvDate.Visible = True
End Sub

Private Sub txt_OneOrder_OrderKey_KeyPress(KeyAscii As Integer)
'�q��s���G���i�f�D�渹�����@�i�ƨ��t�έq��
KeyAscii = 0
End Sub

Private Sub txt_OneOrder_SignQty_Change()
'���i�f�D�渹�������i�ƨ��t�έq��Gñ��q
gd_OneOrder_OrderDetail.Text = txt_OneOrder_SignQty.Text
End Sub

Private Sub txt_OneOrder_SignQty_KeyDown(KeyCode As Integer, Shift As Integer)
'���i�f�D�渹�������i�ƨ��t�έq��Gñ��q
If KeyCode = vbKeyReturn Then
   txt_OneOrder_SignQty.Visible = False
End If
End Sub

Private Sub txt_OneOrder_SignQty_KeyPress(KeyAscii As Integer)
'���i�f�D�渹�������i�ƨ��t�έq��Gñ��q

If gd_OneOrder_OrderDetail.Col <> 14 Then
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
   End Select
End If
End Sub

Private Sub txt_OneOrder_SignQty_LostFocus()
'���i�f�D�渹�������i�ƨ��t�έq��Gñ��q
   txt_OneOrder_SignQty.Visible = False
End Sub

Private Sub txt_MultiOrder_SignQty_Change()
'���i�f�D�渹�����h�i�ƨ��t�έq��Gñ��q
gd_MultiOrder_OrderDetail.Text = txt_MultiOrder_SignQty.Text
End Sub

Private Sub txt_multiOrder_SignQty_KeyDown(KeyCode As Integer, Shift As Integer)
'���i�f�D�渹�����h�i�ƨ��t�έq��Gñ��q
If KeyCode = vbKeyReturn Then
   txt_MultiOrder_SignQty.Visible = False
End If
End Sub

Private Sub txt_multiOrder_SignQty_KeyPress(KeyAscii As Integer)
'���i�f�D�渹�����h�i�ƨ��t�έq��Gñ��q
   Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
   End Select
End Sub

Private Sub txt_multiOrder_SignQty_LostFocus()
'���i�f�D�渹�����h�i�ƨ��t�έq��Gñ��q
   txt_MultiOrder_SignQty.Visible = False
End Sub

Private Sub Display_OrderData_MultiReceipNo(ByVal strExtern As String)
'���i�f�D�渹�����h�i�ƨ��t�έq��G�q���Ƭd��
Screen.MousePointer = vbHourglass
fra_OneOrder_Header.Visible = False
fra_OneOrder_Detail.Visible = False
fra_MultiOrder_Header.Visible = True
fra_MultiOrder_Detail.Visible = True

'���o�ƨ��t�έq�� TRP02T
On Error GoTo err_Handle
str_SQL = "Select ���u�s��,���P���X,�r�p�H,�f�B���q,�X�����,�Ȥ�s��,�Ȥ�W��,�e�f�a�},�f�D,�q��s��,�q����,�X�f���,ñ����,�t�Τ��,���A " & _
          "From SDNConfirm_OrderDate_Multi Where �f�D�渹 = '" & strExtern & "'"
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
txt_MultiOrder_StorerKey.Text = tmp_Rs.Fields("�f�D").Value
txt_MultiOrder_ConsigneeKey.Text = tmp_Rs.Fields("�Ȥ�s��").Value
txt_MultiOrder_FullName.Text = tmp_Rs.Fields("�Ȥ�W��").Value
txt_MultiOrder_Address.Text = tmp_Rs.Fields("�e�f�a�}").Value
txt_MultiOrder_OrderDate.Text = tmp_Rs.Fields("�q����").Value
txt_MultiOrder_ArriveDate.Text = tmp_Rs.Fields("�X�f���").Value
txt_MultiOrder_Status.Text = tmp_Rs.Fields("���A").Value
blShipped = True    '�L�k�P�_�q��z�f�q�O�_�w��s(Ship_Qty)
blCanUpdate = True
If Len(Trim(tmp_Rs.Fields("ñ����").Value)) > 0 Then
   txt_MultiOrder_SignDate.Text = tmp_Rs.Fields("ñ����").Value
   blSDNConfirm = True
   blCanUpdate = False
Else
   txt_MultiOrder_SignDate.Text = tmp_Rs.Fields("�X�f���").Value
   blSDNConfirm = False
   blCanUpdate = True
End If
'Reset �q��s��-���s�C��
Call CreateRS_MultiOrder_RouteDate
Do While Not tmp_Rs.EOF
   rs_MultiOrder.AddNew
   rs_MultiOrder.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
   rs_MultiOrder.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
   rs_MultiOrder.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
   rs_MultiOrder.Fields("�f�B���q").Value = tmp_Rs.Fields("�f�B���q").Value
   rs_MultiOrder.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
   rs_MultiOrder.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close

With dg_MultiOrder
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_MultiOrder.MoveFirst
Set dg_MultiOrder.DataSource = rs_MultiOrder
With dg_MultiOrder
    .RowHeight = 250
    .Columns(0).Width = 1000         '���u�s��
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000         '���P���X
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 800          '�r�p�H
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000         '�f�B���q
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 1100         '�q��s��
    .Columns(4).Alignment = dbgCenter
End With

'���i�f�D�渹�����h�i�ƨ��t�έq��G�q��W��
Call SetGridFormat_MultiOrder_OrderDetail
Dim tmpI As Double
str_SQL = "Select ����,�f��,�~�W,�q��q,�e�f�q,ñ��q,���`��],�d���k��,���`�X,�d�ݽX,�q��s�� " & _
          "From SDNConfirm_OrderDetail_MultiOrder " & _
          "Where �f�D�渹 = '" & strExtern & "' Order by ����"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
tmpI = 1
Do While Not tmp_Rs.EOF
   With gd_MultiOrder_OrderDetail
        If .Rows < (tmpI + 1) Then .Rows = .Rows + 1
        .Row = tmpI
        .Col = 0: .Text = tmp_Rs.Fields("�q��s��").Value
        .Col = 1: .Text = tmp_Rs.Fields("����").Value
        .Col = 2: .Text = tmp_Rs.Fields("�f��").Value
        .Col = 3: .Text = tmp_Rs.Fields("�~�W").Value
        .Col = 4: .Text = tmp_Rs.Fields("�e�f�q").Value
         If blCanUpdate Then      '�|������ SDN Confirmed
            .Col = 5: .Text = 0   'ñ��q
            .Col = 6: .Text = ""  '���`��]
            .Col = 7: .Text = ""  '�d���k��
            .Col = 8: .Text = ""  '���`��]�N�X
            .Col = 9: .Text = ""  '�d���k�ݥN�X
         Else
            .Col = 5: .Text = tmp_Rs.Fields("ñ��q").Value
            .Col = 6: .Text = tmp_Rs.Fields("���`��]").Value
            .Col = 7: .Text = tmp_Rs.Fields("�d���k��").Value
            .Col = 8: .Text = tmp_Rs.Fields("���`�X").Value
            .Col = 9: .Text = tmp_Rs.Fields("�d�ݽX").Value
        End If
   End With
   tmp_Rs.MoveNext
   tmpI = tmpI + 1
Loop
tmp_Rs.Close

If blCanUpdate Then
    cmd_MultiOrder_Deliveryok.Enabled = True
    cmd_MultiOrder_Expect.Enabled = True
    cmd_MultiOrder_NoDelivery.Enabled = True
Else
    cmd_MultiOrder_Deliveryok.Enabled = False
    cmd_MultiOrder_Expect.Enabled = False
    cmd_MultiOrder_NoDelivery.Enabled = False
End If

'�S�w�H���i���� SDN Confirm ���s�s��
'�v���]�w�x�s�� CodeLKUP ListName = [SDNRECONDURM]
'  Code�GUser_ID   Short�G�v���]�w 1-�i�H���ư���
'If (Not blCanUpdate) And CheckSDNReConfirm(User_id) Then
'    '���\���ư��� SDN Confirm ���ϥΪ̡A�}�� SDN Confirm ������ק�
'    cmd_MultiOrder_Deliveryok.Enabled = True
'    cmd_MultiOrder_Expect.Enabled = True
'    cmd_MultiOrder_NoDelivery.Enabled = True
'    blCanUpdate = True
'End If

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��d��", Me.Caption, "Form ���� SubProgram Display_OrderData_MultiReceiptNo", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub CreateRS_MultiOrder_RouteDate()
Call ReDim_Recordset(rs_MultiOrder)
With rs_MultiOrder
     .Fields.Append "���u�s��", adVarChar, 10
     .Fields.Append "���P���X", adVarChar, 20
     .Fields.Append "�r�p�H", adVarChar, 60
     .Fields.Append "�f�B���q", adVarChar, 60
     .Fields.Append "�q��s��", adVarChar, 120
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With

End Sub

Private Sub SetGridFormat_MultiOrder_OrderDetail()
'�W�١GSetGridFormatt_MultiOrder_OrderDetail
'���O�G�Ƶ{��
'�\��G�M���ó]�w [SDN Confirm] ��� [���i�f�D�渹�����h�i�ƨ��t�έq��] �q��W����ܮ榡
'�ѼơG�ǤJ�ȡG�L
Dim sub_var1 As Integer, sub_var2 As Integer
gd_MultiOrder_OrderDetail.Visible = False
With gd_MultiOrder_OrderDetail
     .Rows = 2: .FixedRows = 1: .Cols = 11
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
     .ColWidth(0) = 1200
     .ColWidth(1) = 500
     .ColWidth(2) = 1000
     .ColWidth(3) = 2700
     .ColWidth(4) = 750
     .ColWidth(5) = 750
     .ColWidth(6) = 1850
     .ColWidth(7) = 1400
     .ColWidth(8) = 1000
     .ColWidth(9) = 1000
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "�q��s��"
     .Col = 1: .Text = "����"
     .Col = 2: .Text = "�f��"
     .Col = 3: .Text = "����~�W"
     .Col = 4: .Text = "�e�f�q"
     .Col = 5: .Text = "ñ��q"
     .Col = 6: .Text = "���`��]"
     .Col = 7: .Text = "�d���k��"
     .Col = 8: .Text = "���`�X"
     .Col = 9: .Text = "�d�ݽX"
     '�]�w�C����r���
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .Rows = 2
     .Row = 0
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .Text = ""
     Next sub_var1
     
End With
gd_MultiOrder_OrderDetail.Visible = True
End Sub

Private Sub HideGridUseObject_MultiOrder()
'���i�f�D�渹�����h�i�ƨ��t�έq��G���� [ñ��q] [���`��]] [�d��] ���
txt_MultiOrder_SignQty.Visible = False
cmb_MultiOrder_RBCCode.Visible = False
cmb_MultiOrder_RSCCode.Visible = False
End Sub


Private Sub SetGridFormat_Tab2_SDN_Detail()
'�^�Ǳ��]�w�����u�s�����
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Tab2_SDN_Detail.Visible = False
With dg_Tab2_SDN_Detail
     .Rows = 2: .Cols = 10
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
     .ColWidth(0) = 450
     .ColWidth(1) = 450
     .ColWidth(2) = 1500
     .ColWidth(3) = 2000
     .ColWidth(4) = 700
     .ColWidth(5) = 700
     .ColWidth(6) = 700
     .ColWidth(7) = 900
     .ColWidth(8) = 3500
     .ColWidth(9) = 450

     '�]�w�C�����D:�e�f��,���u�s��,���� ,�r�p�H,���q,��ڤH,�������,���I���,��L���B,��], �_�I,���I
     '�G���ƨ�,���u�s��,�Ȥ�渹,���,���e�Ȥ�,�c��,���n,���q,�h��
     .Row = 0
     .Col = 0: .Text = "�T�{"
     .Col = 1: .Text = "�^��"
     .Col = 2: .Text = "�Ȥ�渹"
     .Col = 3: .Text = "���e�Ȥ�"
     .Col = 4: .Text = "�c��"
     .Col = 5: .Text = "���n"
     .Col = 6: .Text = "���q"
     .Col = 7: .Text = "�h��"
     .Col = 8: .Text = "�Ƶ�"
     .Col = 9: .Text = "�p�p"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter


     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Tab2_SDN_Detail.Visible = True
End Sub


Private Sub SetGridFormat_Tab2_SDN_Cost()
'�^�Ǳ��]�w�����u�s�����
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Tab2_SDN_Cost.Visible = False
With dg_Tab2_SDN_Cost
     .Rows = 2: .Cols = 14
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
     .ColWidth(0) = 800
     .ColWidth(1) = 1500
     .ColWidth(2) = 1000
     .ColWidth(3) = 700
     .ColWidth(4) = 700
     .ColWidth(5) = 600
     .ColWidth(6) = 800
     .ColWidth(7) = 700
     .ColWidth(8) = 700
     .ColWidth(9) = 800
     .ColWidth(10) = 1500
     .ColWidth(11) = 700
     .ColWidth(12) = 700
     .ColWidth(13) = 1000
     '�]�w�C�����D:�e�f��,���u�s��,���� ,�r�p�H,���q,��ڤH,�p�O�ƶq,�������,���I���,��L���B,��], �_�I,���I
     .Row = 0
     .Col = 0: .Text = "�N�X"
     .Col = 1: .Text = "�Ȥ�"
     .Col = 2: .Text = "�渹"
     .Col = 3: .Text = "�_�I"
     .Col = 4: .Text = "���I"
     .Col = 5: .Text = "���"
     .Col = 6: .Text = "�p�O�ƶq"
     .Col = 7: .Text = "������"
     .Col = 8: .Text = "���I��"
     .Col = 9: .Text = "��L���B"
     .Col = 10: .Text = "��]"
     .Col = 11: .Text = "�ꦬ"
     .Col = 12: .Text = "��I"
     .Col = 13: .Text = "�д����O"
     '�]�w�C����r���
     
     .ColAlignment(0) = flexAlignLeftCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignLeftCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignRightCenter
     .ColAlignment(10) = flexAlignLeftCenter
     .ColAlignment(11) = flexAlignRightCenter
     .ColAlignment(12) = flexAlignRightCenter
     .ColAlignment(13) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Tab2_SDN_Cost.Visible = True
End Sub

Public Sub NextPositionTab2Detail(ByVal r As Integer, ByVal C As Integer)     '���ʤ�r���
    On Error GoTo NextError
    Text3.Width = dg_Tab2_SDN_Detail.CellWidth                     '�e��
    Text3.Height = dg_Tab2_SDN_Detail.CellHeight                   '����
    Text3.Left = dg_Tab2_SDN_Detail.Left + dg_Tab2_SDN_Detail.ColPos(C) + 30 '����
    Text3.Top = dg_Tab2_SDN_Detail.Top + dg_Tab2_SDN_Detail.RowPos(r)     '�W��
    Text3.Text = dg_Tab2_SDN_Detail.Text       '�NMSFlexGrid�ثe�@���x�s�椺�e��m���r���
    Text3.Visible = True                '�N��r�����ܩ�e���W
    Text3.SetFocus                      '�N��в��ܤ�r���
    Exit Sub
NextError:
    MsgBox err.Description
End Sub

Public Sub NextPositionTab2Cost(ByVal r As Integer, ByVal C As Integer)     '���ʤ�r���
    On Error GoTo NextError
    Text4.Width = dg_Tab2_SDN_Cost.CellWidth                     '�e��
    Text4.Height = dg_Tab2_SDN_Cost.CellHeight                   '����
    Text4.Left = dg_Tab2_SDN_Cost.Left + dg_Tab2_SDN_Cost.ColPos(C) + 30 '����
    Text4.Top = dg_Tab2_SDN_Cost.Top + dg_Tab2_SDN_Cost.RowPos(r)     '�W��
    Text4.Text = dg_Tab2_SDN_Cost.Text       '�NMSFlexGrid�ثe�@���x�s�椺�e��m���r���
    Text4.Visible = True                '�N��r�����ܩ�e���W
    Text4.SetFocus                      '�N��в��ܤ�r���
    Exit Sub
NextError:
    MsgBox err.Description
End Sub

Private Sub Text1_LostFocus()
    On Error GoTo TextError
        Text1.Visible = False
    Exit Sub
TextError:
    MsgBox err.Description
End Sub


Private Sub Text3_LostFocus()
    On Error GoTo TextError
        Text3.Visible = False
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Text3_Change()  '�N��r������e�g�ܹ����x�s��
    On Error GoTo TextError
    dg_Tab2_SDN_Detail.Text = Text3.Text   '�N��r������e�g�ܹ����x�s��
    If dg_Tab2_SDN_Detail.Col = 4 Or dg_Tab2_SDN_Detail.Col = 5 Or dg_Tab2_SDN_Detail.Col = 6 Then
        Call Tab2Detail_Sum
    End If
    Exit Sub
 
TextError:
    MsgBox err.Description
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    On Error GoTo TextError
    If KeyAscii = vbKeyReturn Then                '�b���UEnter�ɡA�M�w�U��grid����m
        If dg_Tab2_SDN_Detail.Col < 8 Then
            dg_Tab2_SDN_Detail.Col = dg_Tab2_SDN_Detail.Col + 1
            NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
        End If

    End If
    'i = KeyAscii
    If KeyAscii = 1 Then 'Ctrl+A
        Call cmd_Tab2_AddOrder_Click
    End If
    If KeyAscii = 4 Then 'Ctrl+D
        Call cmd_Tab2_DelOrder_Click
    End If
    If KeyAscii = 26 Then 'Ctrl+Z
        dg_Tab2_SDN_Cost.Row = 1
        dg_Tab2_SDN_Cost.Col = 0
        NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
    End If
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Text4_LostFocus()
    On Error GoTo TextError
        Text4.Visible = False
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Text4_Change()  '�N��r������e�g�ܹ����x�s��
    On Error GoTo TextError
    dg_Tab2_SDN_Cost.Text = Text4.Text   '�N��r������e�g�ܹ����x�s��
    Call Cost_Tab2_Sum
    Exit Sub
 
TextError:
    MsgBox err.Description
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    On Error GoTo TextError
    If KeyAscii = vbKeyReturn Then                '�b���UEnter�ɡA�M�w�U��grid����m
        If dg_Tab2_SDN_Cost.Col = 0 Then
            If Len(Trim(dg_Tab2_SDN_Cost.Text)) > 0 Then
                Call Confirm_Recordset_Closed(tmp_Rs)
                str_SQL = "SELECT RTRIM(CostCode) AS �N�X,CostName as �Ȥ�W��,Receivable as �������,Payable as ���I���,AreaStart as �_�I,AreaEnd as ���I,CostKind as �д����O  " & _
                          "From TRP17M where CostCode='" & Trim(Text4.Text) & "'"
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If Not tmp_Rs.EOF Then
                    dg_Tab2_SDN_Cost.Col = 1: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(1).Value)
                    dg_Tab2_SDN_Cost.Col = 3: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(4).Value)
                    dg_Tab2_SDN_Cost.Col = 4: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(5).Value)
                    dg_Tab2_SDN_Cost.Col = 7: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(2).Value)
                    dg_Tab2_SDN_Cost.Col = 8: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(3).Value)
                    dg_Tab2_SDN_Cost.Col = 13: dg_Tab2_SDN_Cost.Text = Trim(tmp_Rs.Fields(6).Value)
                    dg_Tab2_SDN_Cost.Col = 1
                    Call Cost_Tab2_SumAll
                End If
                tmp_Rs.Close
            End If
        End If
        If dg_Tab2_SDN_Cost.Col < 12 Then
            dg_Tab2_SDN_Cost.Col = dg_Tab2_SDN_Cost.Col + 1
            NextPositionTab2Cost dg_Tab2_SDN_Cost.Row, dg_Tab2_SDN_Cost.Col
        End If
    End If
    'i = KeyAscii
    If KeyAscii = 1 Then 'Ctrl+A
        Call cmd_Tab2_AddCost_Click
    End If
    If KeyAscii = 4 Then 'Ctrl+D
        Call cmd_Tab2_DelCost_Click
    End If
    If KeyAscii = 26 Then 'Ctrl+Z
    
    End If
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Cost_Tab2_Sum()  '�έp�ꦬ�P��I
    intR = dg_Tab2_SDN_Cost.Col
    '�έp�ꦬ
    If dg_Tab2_SDN_Cost.Col = 6 Or dg_Tab2_SDN_Cost.Col = 7 Then
        dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 7: B = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 11: dg_Tab2_SDN_Cost.Text = Round(a * B)
    End If
    dg_Tab2_SDN_Cost.Col = intR
    '�έp��I
    If dg_Tab2_SDN_Cost.Col = 6 Or dg_Tab2_SDN_Cost.Col = 8 Or dg_Tab2_SDN_Cost.Col = 9 Then
        dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 8: B = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 9: C = Val(dg_Tab2_SDN_Cost.Text)
        dg_Tab2_SDN_Cost.Col = 12: dg_Tab2_SDN_Cost.Text = Round(a * B + C, 0)
    End If
    dg_Tab2_SDN_Cost.Col = intR
End Sub

Private Sub Cost_SumAll()  '�έp�ꦬ�P��I
'    intR = dg_SDN_Cost.Col
'    '�έp�ꦬ
'    dg_SDN_Cost.Col = 6: a = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 7: B = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 11: dg_SDN_Cost.Text = Round(a * B, 0)
'    dg_SDN_Cost.Col = intR
'    '�έp��I
'    dg_SDN_Cost.Col = 6: a = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 8: B = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 9: C = Val(dg_SDN_Cost.Text)
'    dg_SDN_Cost.Col = 12: dg_SDN_Cost.Text = Round(a * B + C, 0)
'    dg_SDN_Cost.Col = intR
End Sub

Private Sub Cost_Tab2_SumAll()  '�έp�ꦬ�P��I
'    intR = dg_SDN_Cost.Col
'    '�έp�ꦬ
'    dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 7: B = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 11: dg_Tab2_SDN_Cost.Text = Round(a * B, 0)
'    dg_Tab2_SDN_Cost.Col = intR
'    '�έp��I
'    dg_Tab2_SDN_Cost.Col = 6: a = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 8: B = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 9: C = Val(dg_Tab2_SDN_Cost.Text)
'    dg_Tab2_SDN_Cost.Col = 12: dg_Tab2_SDN_Cost.Text = Round(a * B + C, 0)
'    dg_Tab2_SDN_Cost.Col = intR
End Sub

Private Sub Tab2Detail_Sum()  '�έp�ꦬ�P��I
    intC = dg_Tab2_SDN_Detail.Col
    intR = dg_Tab2_SDN_Detail.Row
    txt_Tab2_srcTotal_Case.Text = 0
    txt_Tab2_srcTotal_Volumn.Text = 0
    txt_Tab2_srcTotal_Weight.Text = 0
    For i = 1 To dg_Tab2_SDN_Detail.Rows - 1
        dg_Tab2_SDN_Detail.Row = i
        dg_Tab2_SDN_Detail.Col = 4
        txt_Tab2_srcTotal_Case.Text = Val(dg_Tab2_SDN_Detail.Text) + Val(txt_Tab2_srcTotal_Case.Text)
        dg_Tab2_SDN_Detail.Col = 5
        txt_Tab2_srcTotal_Volumn.Text = Val(dg_Tab2_SDN_Detail.Text) + Val(txt_Tab2_srcTotal_Volumn.Text)
        dg_Tab2_SDN_Detail.Col = 6
        txt_Tab2_srcTotal_Weight.Text = Val(dg_Tab2_SDN_Detail.Text) + Val(txt_Tab2_srcTotal_Weight.Text)
    Next
    dg_Tab2_SDN_Detail.Col = intC
    dg_Tab2_SDN_Detail.Row = intR
End Sub

Private Sub txt_Tab02_C_VEHICLE_ID_NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '�b���UEnter�ɡA�M�w�U��grid����m
        txt_Tab02_C_VEHICLE_ID_NO.SetFocus
    End If
End Sub

Private Sub txt_Tab02_Delivery_Date_Click()
    'Tab2-�X����_
    If Trim(txt_Tab02_Delivery_Date.Text) = "" Then
       mvDate.Value = Now
    Else
       If Fun_ChkDateFormat(txt_Tab02_Delivery_Date.Text) = 1 Then
          mvDate.Value = Now
       Else
          mvDate.Value = CDate(Left(txt_Tab02_Delivery_Date.Text, 4) & "/" & Mid(txt_Tab02_Delivery_Date.Text, 5, 2) & "/" & Right(txt_Tab02_Delivery_Date.Text, 2))
       End If
    End If
    mvDate.Left = txt_Tab02_Delivery_Date.Left + txt_Tab02_Delivery_Date.Left
    mvDate.Top = txt_Tab02_Delivery_Date.Top + txt_Tab02_Delivery_Date.Top + txt_Tab02_Delivery_Date.Height
    mvDate.Tag = "Tab2-�X����"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab02_Delivery_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '�b���UEnter�ɡA�M�w�U��grid����m
        txt_Tab02_C_VEHICLE_ID_NO.SetFocus
    End If
End Sub

Private Sub txt_Tab02_Driver_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '�b���UEnter�ɡA�M�w�U��grid����m
        txt_Tab02_Receiver.SetFocus
    End If
End Sub

Private Sub txt_Tab02_Receiver_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then '�b���UEnter�ɡA�M�w�U��grid����m
        dg_Tab2_SDN_Detail.Col = 2
        dg_Tab2_SDN_Detail.Row = 1
        NextPositionTab2Detail dg_Tab2_SDN_Detail.Row, dg_Tab2_SDN_Detail.Col
    End If
    
End Sub

Private Sub Clear_CardData()
    txt_Tab02_Delivery_Date.Text = ""
    txt_Tab02_C_VEHICLE_ID_NO.Text = ""
    txt_Tab02_Driver.Text = ""
    txt_Tab02_Receiver.Text = ""
    txt_Tab02_C_Route_No = ""
    txt_Tab2_srcTotal_Case.Text = ""
    txt_Tab2_srcTotal_Volumn.Text = ""
    txt_Tab2_srcTotal_Weight.Text = ""
    txt_Tab2_sum_Case.Text = ""
    txt_Tab2_sum_CBM.Text = ""
    txt_Tab2_sum_WT.Text = ""
    dg_Tab2_SDN_Detail.Rows = 2
    dg_Tab2_SDN_Detail.Row = 1
    For i = 0 To dg_Tab2_SDN_Detail.Cols - 1
        dg_Tab2_SDN_Detail.Col = i
        dg_Tab2_SDN_Detail.Text = ""
    Next
    dg_Tab2_SDN_Cost.Rows = 2
    dg_Tab2_SDN_Cost.Row = 1
    For i = 0 To dg_Tab2_SDN_Cost.Cols - 1
        dg_Tab2_SDN_Cost.Col = i
        dg_Tab2_SDN_Cost.Text = ""
    Next
End Sub

Private Sub Tab0_SumQty(cr As String)
    str_SQL = "select isnull(sum(ChargeQty),0) from SDN05T where C_ROUTE_NO ='" & cr & "' and Uom in ('���q','���n')"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    txt_Tab0_SumQty.Text = tmp_Rs.Fields(0).Value
    tmp_Rs.Close
End Sub

Private Sub txtCustomerOrderkey_Change()
    If txtCustomerOrderkey.Visible = False Then Exit Sub
    rsOrderT0("�禬�渹") = txtCustomerOrderkey.Text
End Sub

Private Sub txtCustomerOrderkey_Click()
    txtCustomerOrderkey.SetFocus: txtCustomerOrderkey.SelStart = 0: txtCustomerOrderkey.SelLength = Len(txtCustomerOrderkey.Text)
End Sub

Private Sub txtCustomerOrderkey_GotFocus()
    txtCustomerOrderkey.SelStart = 0: txtCustomerOrderkey.SelLength = Len(txtCustomerOrderkey.Text)
End Sub

Private Sub txtCustomerOrderkey_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   txtCustomerOrderkey.Visible = False
End If
End Sub

Private Sub txtCustomerOrderkey_LostFocus()
   txtCustomerOrderkey.Visible = False
   dgOrderT0.Col = 0
End Sub

Private Sub txtDeliveryE_Click()
    Set objMvdateTarget = txtDeliveryE
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryS_Click()
    Set objMvdateTarget = txtDeliveryS
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtSignDateS_Click()
    Set objMvdateTarget = txtSignDateS
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub txtSignDateE_Click()
    Set objMvdateTarget = txtSignDateE
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtSignDateS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txtSignDateE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryDateST0_Click()
    Set objMvdateTarget = txtDeliveryDateST0
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryDateET0_Click()
    Set objMvdateTarget = txtDeliveryDateET0
    mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
    mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryDateST0_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryDateET0_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mvDate.Visible = False
End Sub

Private Sub txt_OrderKey_KeyPress(KeyAscii As Integer)
'�K�X
If KeyAscii = vbKeyReturn Then Call cmd_OrderQuery_Click
'   KeyAscii = 0
'   cmd_Login.SetFocus
'End If
End Sub
