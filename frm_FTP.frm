VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_FTP 
   Caption         =   "�q�汵��"
   ClientHeight    =   8775
   ClientLeft      =   210
   ClientTop       =   750
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11280
   WindowState     =   2  '�̤j��
   Begin InetCtlsObjects.Inet ITC 
      Left            =   240
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8160
      Left            =   120
      TabIndex        =   0
      Top             =   15
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   14393
      _Version        =   393216
      Tabs            =   23
      Tab             =   18
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm_FTP.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmd_Import"
      Tab(0).Control(1)=   "fraControls"
      Tab(0).Control(2)=   "fraRemoteFiles"
      Tab(0).Control(3)=   "fraLoginInfo"
      Tab(0).Control(4)=   "fraStatus"
      Tab(0).Control(5)=   "fraLocalFiles"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frm_FTP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dg_CustInv"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Vitalon�q��פJ"
      TabPicture(2)   =   "frm_FTP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "dgMainT2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frm_FTP.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "int3"
      Tab(3).Control(2)=   "dgMainT3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "LKAO01"
      TabPicture(4)   =   "frm_FTP.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dgMainT4"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frm_FTP.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dgMainT5"
      Tab(5).Control(1)=   "Frame8"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "�¤�q��פJ"
      TabPicture(6)   =   "frm_FTP.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "dgMainT6"
      Tab(6).Control(1)=   "Frame9"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "�ʨ�I�q��פJ"
      TabPicture(7)   =   "frm_FTP.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "dgMainT7"
      Tab(7).Control(1)=   "Frame10"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   " "
      TabPicture(8)   =   "frm_FTP.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame11"
      Tab(8).Control(1)=   "dgMainT8"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   " "
      TabPicture(9)   =   "frm_FTP.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame12"
      Tab(9).Control(1)=   "dgMainT9"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   " "
      TabPicture(10)  =   "frm_FTP.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame13"
      Tab(10).Control(1)=   "dgMainT10"
      Tab(10).ControlCount=   2
      TabCaption(11)  =   " "
      TabPicture(11)  =   "frm_FTP.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "dgMainT11"
      Tab(11).Control(1)=   "Frame3"
      Tab(11).ControlCount=   2
      TabCaption(12)  =   "�ʨ�RC�q��פJ"
      TabPicture(12)  =   "frm_FTP.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "dgMainT12"
      Tab(12).Control(1)=   "Frame4"
      Tab(12).ControlCount=   2
      TabCaption(13)  =   "--��L�f�D�q��"
      TabPicture(13)  =   "frm_FTP.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "dgMainT13"
      Tab(13).Control(1)=   "Frame5"
      Tab(13).ControlCount=   2
      TabCaption(14)  =   "Excel�q��פJ"
      TabPicture(14)  =   "frm_FTP.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "dgMainT14"
      Tab(14).Control(1)=   "Frame6"
      Tab(14).ControlCount=   2
      TabCaption(15)  =   "�����q��פJ"
      TabPicture(15)  =   "frm_FTP.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "dgMainT15"
      Tab(15).Control(1)=   "Frame14"
      Tab(15).ControlCount=   2
      TabCaption(16)  =   "���_�q��פJ"
      TabPicture(16)  =   "frm_FTP.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Frame15"
      Tab(16).Control(1)=   "SSTab2"
      Tab(16).ControlCount=   2
      TabCaption(17)  =   " "
      TabPicture(17)  =   "frm_FTP.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Frame16"
      Tab(17).Control(1)=   "SSTab3"
      Tab(17).ControlCount=   2
      TabCaption(18)  =   "�ʨ�A2B�BC��פJ"
      TabPicture(18)  =   "frm_FTP.frx":01F8
      Tab(18).ControlEnabled=   -1  'True
      Tab(18).Control(0)=   "dgMainT18"
      Tab(18).Control(0).Enabled=   0   'False
      Tab(18).Control(1)=   "Frame17"
      Tab(18).Control(1).Enabled=   0   'False
      Tab(18).ControlCount=   2
      TabCaption(19)  =   " "
      TabPicture(19)  =   "frm_FTP.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "Frame18"
      Tab(19).Control(1)=   "dgMainT19"
      Tab(19).ControlCount=   2
      TabCaption(20)  =   " �S�O�έq��פJ"
      TabPicture(20)  =   "frm_FTP.frx":0230
      Tab(20).ControlEnabled=   0   'False
      Tab(20).Control(0)=   "Frame19"
      Tab(20).Control(1)=   "dgMainT20"
      Tab(20).ControlCount=   2
      TabCaption(21)  =   "����RC�q��פJ"
      TabPicture(21)  =   "frm_FTP.frx":024C
      Tab(21).ControlEnabled=   0   'False
      Tab(21).Control(0)=   "Frame20"
      Tab(21).Control(1)=   "dgMainT21"
      Tab(21).ControlCount=   2
      TabCaption(22)  =   " "
      TabPicture(22)  =   "frm_FTP.frx":0268
      Tab(22).ControlEnabled=   0   'False
      Tab(22).Control(0)=   "Frame21"
      Tab(22).Control(1)=   "dgMainT22"
      Tab(22).ControlCount=   2
      Begin VB.Frame Frame21 
         Caption         =   "LYFY09�h�f�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   194
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton cmdImportT22 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   199
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT22 
            Height          =   300
            Left            =   135
            TabIndex        =   198
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT22 
            Height          =   1560
            Left            =   135
            TabIndex        =   197
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT22 
            Height          =   1530
            Left            =   4560
            Pattern         =   "PG�h�f*.xls"
            TabIndex        =   196
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT22 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   195
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   200
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "����RC�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   190
         Top             =   1320
         Width           =   9840
         Begin VB.ComboBox cboSheetT21 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   210
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT21 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   209
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT21 
            Height          =   1560
            Left            =   120
            TabIndex        =   208
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT21 
            Height          =   300
            Left            =   120
            TabIndex        =   207
            ToolTipText     =   "Local Drive List"
            Top             =   360
            Width           =   2040
         End
         Begin VB.CommandButton cmdOpenFilesT21 
            BackColor       =   &H0080FFFF&
            Caption         =   "�}��"
            Height          =   375
            Left            =   3480
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   206
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdImportT21 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   2400
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   191
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   193
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "�S�O�έq��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   186
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton cmdOpenFilesT20 
            BackColor       =   &H0080FFFF&
            Caption         =   "�}��"
            Height          =   375
            Left            =   3360
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   215
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboSheetT20 
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
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   214
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT20 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   213
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT20 
            Height          =   1560
            Left            =   120
            TabIndex        =   212
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmdImportT20 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   2280
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   188
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT20 
            Height          =   300
            Left            =   135
            TabIndex        =   187
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4440
            TabIndex        =   211
            Top             =   300
            Width           =   720
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   181
         Top             =   3720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "�q����(Format)"
         TabPicture(0)   =   "frm_FTP.frx":0284
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dgMainT17"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "�X�f�Ƶ�(RawHerder)"
         TabPicture(1)   =   "frm_FTP.frx":02A0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgMainT17_1"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgMainT17 
            Height          =   3855
            Left            =   120
            TabIndex        =   182
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT17_1 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   183
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
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
      Begin VB.Frame Frame18 
         Caption         =   "LAPP01-�h�f"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   172
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton cmdImportT19 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   177
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT19 
            Height          =   300
            Left            =   135
            TabIndex        =   176
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT19 
            Height          =   1560
            Left            =   135
            TabIndex        =   175
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT19 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   174
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT19 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   173
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   178
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "�ʨ�A2B�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   120
         TabIndex        =   164
         Top             =   1320
         Width           =   9840
         Begin VB.ComboBox cboSheetT18 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   169
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT18 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   168
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT18 
            Height          =   1560
            Left            =   135
            TabIndex        =   167
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT18 
            Height          =   300
            Left            =   135
            TabIndex        =   166
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT18 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   165
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   170
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "LAPP01-�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   -74760
         TabIndex        =   158
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   202
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.ComboBox cboSheetT17 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   184
            Top             =   240
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.CommandButton CmbStartT17 
            Caption         =   "�}���ɮ�"
            Height          =   375
            Left            =   4560
            TabIndex        =   180
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdImportT17 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   162
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT17 
            Height          =   300
            Left            =   135
            TabIndex        =   161
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT17 
            Height          =   1560
            Left            =   135
            TabIndex        =   160
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT17 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   159
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   185
            Top             =   300
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   163
            Top             =   300
            Width           =   720
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   149
         Top             =   3780
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "�q��D��"
         TabPicture(0)   =   "frm_FTP.frx":02BC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dgMainT16"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "�q�������"
         TabPicture(1)   =   "frm_FTP.frx":02D8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgMainT16_1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "frm_FTP.frx":02F4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dgMainT16_2"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   " "
         TabPicture(3)   =   "frm_FTP.frx":0310
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "dgMainT16_3"
         Tab(3).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgMainT16 
            Height          =   3975
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT16_1 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   151
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT16_2 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   152
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT16_3 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   153
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
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
      Begin VB.Frame Frame15 
         Caption         =   "LMBO01-�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74760
         TabIndex        =   144
         Top             =   1260
         Width           =   9840
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   205
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Cmb_Import2 
            Caption         =   "�P�f�פJ"
            Height          =   375
            Left            =   5760
            TabIndex        =   157
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Cmb_Import4 
            Caption         =   "�H�w�P�f�פJ"
            Height          =   375
            Left            =   8160
            TabIndex        =   156
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Cmb_Import3 
            Caption         =   "�༷�פJ"
            Height          =   375
            Left            =   6960
            TabIndex        =   155
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Cmb_Import1 
            Caption         =   "�}���ɮ�"
            Height          =   375
            Left            =   4560
            TabIndex        =   154
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Cmd_Impdata 
            BackColor       =   &H0080FFFF&
            Caption         =   "�q��פJ"
            Height          =   375
            Left            =   3480
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   148
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT16 
            Enabled         =   0   'False
            Height          =   300
            Left            =   120
            TabIndex        =   147
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT16 
            Enabled         =   0   'False
            Height          =   1560
            Left            =   135
            TabIndex        =   146
            ToolTipText     =   "Local Directory"
            Top             =   720
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT16 
            Enabled         =   0   'False
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   145
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Visible         =   0   'False
            Width           =   5190
         End
         Begin VB.Label lab_Orderdetail 
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
            Height          =   375
            Left            =   6960
            TabIndex        =   204
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lab_Orders 
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
            Height          =   375
            Left            =   5640
            TabIndex        =   203
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "�����q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   136
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdOpenFilesT15 
            BackColor       =   &H0080FFFF&
            Caption         =   "�}��"
            Height          =   375
            Left            =   3480
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   216
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboSheetT15 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   141
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT15 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   140
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT15 
            Height          =   1560
            Left            =   135
            TabIndex        =   139
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT15 
            Height          =   300
            Left            =   135
            TabIndex        =   138
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT15 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   2400
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   137
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   142
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Excel�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   128
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT14 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   133
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT14 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   132
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT14 
            Height          =   1560
            Left            =   120
            TabIndex        =   131
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT14 
            Height          =   300
            Left            =   135
            TabIndex        =   130
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT14 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   129
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   134
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "�f�D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   118
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboStorerkeyT13 
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
            Left            =   120
            TabIndex        =   126
            Text            =   "cboStorerkeyT13"
            Top             =   240
            Width           =   1605
         End
         Begin VB.CommandButton cmdImportT13 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3480
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   123
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT13 
            Height          =   300
            Left            =   1815
            TabIndex        =   122
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   1560
         End
         Begin VB.DirListBox dirLocalDirT13 
            Height          =   1560
            Left            =   135
            TabIndex        =   121
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT13 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   120
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT13 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   119
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   124
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "�ʨ�RC���f�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   110
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT12 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   115
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT12 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   114
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT12 
            Height          =   1560
            Left            =   135
            TabIndex        =   113
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT12 
            Height          =   300
            Left            =   135
            TabIndex        =   112
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT12 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3480
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   111
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   116
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "�ߨ��q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   100
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT11 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   105
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT11 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   104
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT11 
            Height          =   1560
            Left            =   135
            TabIndex        =   103
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT11 
            Height          =   300
            Left            =   135
            TabIndex        =   102
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT11 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   101
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   106
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "LNSL01-PX�h�f"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   88
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdImportT10 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   93
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT10 
            Height          =   300
            Left            =   135
            TabIndex        =   92
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT10 
            Height          =   1560
            Left            =   135
            TabIndex        =   91
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT10 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   90
            ToolTipText     =   "����� ""*.xls"" �����ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT10 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   89
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   94
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "LNSL01-PX�q��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   80
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT9 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   85
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT9 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   84
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT9 
            Height          =   1560
            Left            =   135
            TabIndex        =   83
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT9 
            Height          =   300
            Left            =   135
            TabIndex        =   82
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT9 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   81
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   86
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "LNSL01-�@��h�f"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   72
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT8 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   79
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT8 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   76
            ToolTipText     =   "����� ""*.xls"" �����ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT8 
            Height          =   1560
            Left            =   135
            TabIndex        =   75
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT8 
            Height          =   300
            Left            =   135
            TabIndex        =   74
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT8 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   73
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   77
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "LPSI01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   64
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton Command2 
            BackColor       =   &H0080FFFF&
            Caption         =   "��}��[������"
            Height          =   495
            Left            =   2160
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   127
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdImportT7 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   69
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT7 
            Height          =   300
            Left            =   135
            TabIndex        =   68
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT7 
            Height          =   1560
            Left            =   135
            TabIndex        =   67
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT7 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   66
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT7 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   65
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   70
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "LSJR01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   56
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT6 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   61
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT6 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   60
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT6 
            Height          =   1560
            Left            =   135
            TabIndex        =   59
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT6 
            Height          =   300
            Left            =   135
            TabIndex        =   58
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT6 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   57
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   62
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "LKAO01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   48
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmd2Excel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "��Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   109
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdImportT4 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3480
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT4 
            Height          =   300
            Left            =   135
            TabIndex        =   52
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT4 
            Height          =   1560
            Left            =   135
            TabIndex        =   51
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT4 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   50
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT4 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   49
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   54
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "LNIP01-�h�f"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   40
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdImportT5 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT5 
            Height          =   300
            Left            =   135
            TabIndex        =   44
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT5 
            Height          =   1560
            Left            =   135
            TabIndex        =   43
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT5 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   42
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT5 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   41
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   46
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "LVTL01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   35
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdImportT2 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT2 
            Height          =   300
            Left            =   135
            TabIndex        =   38
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT2 
            Height          =   1560
            Left            =   135
            TabIndex        =   37
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT2 
            Height          =   2070
            Left            =   4560
            Pattern         =   "*.txt"
            TabIndex        =   36
            ToolTipText     =   "����� ""*.txt"" �ɮ�"
            Top             =   240
            Width           =   5190
         End
      End
      Begin VB.CommandButton cmd_Import 
         BackColor       =   &H0080FFFF&
         Caption         =   "�פJ"
         Height          =   375
         Left            =   -71400
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
         ToolTipText     =   "�q��פJ"
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fraControls 
         Caption         =   "�W�U��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -71550
         TabIndex        =   31
         Top             =   4920
         Width           =   6750
         Begin VB.CommandButton Command1 
            Caption         =   "�U��"
            Height          =   375
            Left            =   3840
            TabIndex        =   32
            Top             =   225
            Width           =   615
         End
         Begin VB.Image imgSendFile 
            DragIcon        =   "frm_FTP.frx":032C
            Enabled         =   0   'False
            Height          =   345
            Left            =   135
            Picture         =   "frm_FTP.frx":076E
            Stretch         =   -1  'True
            ToolTipText     =   "Send Selected File"
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblSendFile 
            AutoSize        =   -1  'True
            Caption         =   "�W���ɮ�"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   555
            TabIndex        =   34
            Top             =   360
            Width           =   735
         End
         Begin VB.Image imgReceiveFile 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1665
            Picture         =   "frm_FTP.frx":0BB0
            Stretch         =   -1  'True
            ToolTipText     =   "Recieve Selected File"
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblReceiveFile 
            AutoSize        =   -1  'True
            Caption         =   "Alc�U���öפJ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2085
            TabIndex        =   33
            Top             =   360
            Width           =   1185
         End
      End
      Begin VB.Frame fraRemoteFiles 
         Caption         =   "FTP �ɮ�"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   -71550
         TabIndex        =   23
         Top             =   1800
         Width           =   6750
         Begin VB.ListBox lstRemoteFile 
            Enabled         =   0   'False
            Height          =   2040
            ItemData        =   "frm_FTP.frx":0FF2
            Left            =   120
            List            =   "frm_FTP.frx":0FF4
            TabIndex        =   30
            ToolTipText     =   "Remote Files"
            Top             =   840
            Width           =   6450
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "�R��"
            Enabled         =   0   'False
            Height          =   320
            Left            =   1800
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   29
            ToolTipText     =   "Delete"
            Top             =   315
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton cmdUpFolder 
            Enabled         =   0   'False
            Height          =   320
            Left            =   2775
            Picture         =   "frm_FTP.frx":0FF6
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   28
            ToolTipText     =   "Move Up One Folder"
            Top             =   315
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton cmdNewFolder 
            Caption         =   "�s�W��Ƨ�"
            Enabled         =   0   'False
            Height          =   320
            Left            =   3720
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   27
            ToolTipText     =   "New Folder"
            Top             =   315
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.CommandButton cmd_shp 
            BackColor       =   &H00FF8080&
            Caption         =   "SHP"
            Height          =   320
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   26
            Top             =   315
            Width           =   495
         End
         Begin VB.CommandButton cmd_Alc 
            BackColor       =   &H00FF8080&
            Caption         =   "ALC"
            Height          =   320
            Left            =   660
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   25
            Top             =   315
            Width           =   495
         End
         Begin VB.CommandButton cmd_CFM 
            BackColor       =   &H00FF8080&
            Caption         =   "CFM"
            Height          =   320
            Left            =   1200
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
            Top             =   315
            Width           =   495
         End
      End
      Begin VB.Frame fraLoginInfo 
         Caption         =   "�ըƹF���y"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   -74880
         TabIndex        =   10
         Top             =   1800
         Width           =   3210
         Begin VB.TextBox txtServer 
            Height          =   285
            Left            =   180
            TabIndex        =   15
            ToolTipText     =   "FTP Server Name"
            Top             =   765
            Width           =   2805
         End
         Begin VB.TextBox txtUserName 
            Height          =   285
            Left            =   180
            TabIndex        =   14
            ToolTipText     =   "User Name"
            Top             =   1395
            Width           =   2805
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  '�Ȥ�
            Left            =   180
            PasswordChar    =   "*"
            TabIndex        =   13
            ToolTipText     =   "Password"
            Top             =   2025
            Width           =   2805
         End
         Begin VB.CommandButton cmdLogOn 
            BackColor       =   &H0000FF00&
            Caption         =   "�n  �J"
            Height          =   420
            Left            =   225
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   12
            ToolTipText     =   "Log On"
            Top             =   2475
            Width           =   1320
         End
         Begin VB.CommandButton cmdLogOff 
            BackColor       =   &H008080FF&
            Caption         =   "�n  �X"
            Enabled         =   0   'False
            Height          =   420
            Left            =   1665
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   11
            ToolTipText     =   "Log Off"
            Top             =   2475
            Width           =   1320
         End
         Begin VB.Label lblServer 
            AutoSize        =   -1  'True
            Caption         =   "Server:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   495
            Width           =   630
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            Caption         =   "User Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label lblPassword 
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   1755
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�{���q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   3
         Top             =   1380
         Width           =   9885
         Begin VB.ComboBox cboSheetT3 
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
            Left            =   5400
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   98
            Top             =   240
            Width           =   4365
         End
         Begin VB.CommandButton cmdImportT3 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT3 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT3 
            Height          =   1560
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT3 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   4
            ToolTipText     =   "Local Files"
            Top             =   720
            Width           =   5190
         End
         Begin VB.Label Label1 
            Alignment       =   2  '�m�����
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�u�@��"
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
            Left            =   4560
            TabIndex        =   99
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame fraStatus 
         Caption         =   "���A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74880
         TabIndex        =   1
         Top             =   4920
         Width           =   3210
         Begin VB.Label lblStatus 
            Alignment       =   2  '�m�����
            Caption         =   "Not Connected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   240
            TabIndex        =   2
            ToolTipText     =   "Connection Status"
            Top             =   360
            Width           =   2895
         End
      End
      Begin MSDataGridLib.DataGrid dg_CustInv 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   8
         Top             =   1440
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   11668
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
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
      Begin InetCtlsObjects.Inet int3 
         Left            =   -65520
         Top             =   4560
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Frame fraLocalFiles 
         Caption         =   "TK�q��פJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   19
         Top             =   5760
         Width           =   10080
         Begin VB.CommandButton cmdOpenFile 
            BackColor       =   &H0080FFFF&
            Caption         =   "�}���ɮ�"
            Height          =   375
            Left            =   2280
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   108
            ToolTipText     =   "�}�Ҩ�L���M��"
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDrive 
            Height          =   300
            Left            =   135
            TabIndex        =   22
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDir 
            Height          =   1560
            Left            =   135
            TabIndex        =   21
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFile 
            Height          =   2070
            Left            =   4560
            TabIndex        =   20
            ToolTipText     =   "Local Files"
            Top             =   240
            Width           =   5415
         End
      End
      Begin MSDataGridLib.DataGrid dgMainT5 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   47
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT4 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   55
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT6 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   63
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT7 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   71
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT8 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   78
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT9 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   87
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT10 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   95
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT2 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   96
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid dgMainT3 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   97
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid dgMainT11 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   107
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT12 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   117
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT13 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   125
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT14 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   135
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT15 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   143
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT18 
         Height          =   4335
         Left            =   120
         TabIndex        =   171
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT19 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   179
         Top             =   3900
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT20 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   189
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT21 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   192
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT22 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   201
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private strTranFileName As String
Private str_file As String
Private strOrderNo As String
Private str_Orderkey As String
Private str_CustomerID As String
Private str_CSKU, str_Note, str_DESCR As String
Private str_ExternOrderkey As String
'Private i As Double

Private RecievingSize As Boolean
Private rs_Src As ADODB.Recordset           '��l�q����
Private rs_Head As ADODB.Recordset          '���Ϋᤧ�q����Y���
Private rs_Detail As ADODB.Recordset        '���Ϋᤧ�q����Ӹ��
'Private int_Repeat As Integer
'Private int_Order As Integer
'Private int_OrderLine As Integer
Private cn_Self As ADODB.Connection
Private fso As Scripting.FileSystemObject
Private rsMainTK As ADODB.Recordset
Private rsMainT2 As ADODB.Recordset
Private rsMainT3 As ADODB.Recordset
Private rsMainT4 As ADODB.Recordset
Private rsMainT5 As ADODB.Recordset
Private rsMainT6 As ADODB.Recordset
Private rsMainT7 As ADODB.Recordset
Private rsMainT8 As ADODB.Recordset
Private rsMainT9 As ADODB.Recordset
Private rsMainT10 As ADODB.Recordset
Private rsMainT11 As ADODB.Recordset
Private rsMainT12 As ADODB.Recordset
Private rsMainT13 As ADODB.Recordset
Private rsMainT14 As ADODB.Recordset
Private rsMainT15 As ADODB.Recordset
Private rsMainT16 As ADODB.Recordset
Private rsMainT16_1 As ADODB.Recordset
Private rsMainT16_2 As ADODB.Recordset
Private rsMainT16_3 As ADODB.Recordset
Private rsMainT17 As ADODB.Recordset
Private rsMainT17_1 As ADODB.Recordset
Private rsMainT18 As ADODB.Recordset
Private rsMainT19 As ADODB.Recordset
Private rsMainT20 As ADODB.Recordset
Private rsMainT21 As ADODB.Recordset
Private rsMainT22 As ADODB.Recordset
Private strTranFileNameT4 As String
Private blDo As Boolean
Private ConfirmYN
Private Str_updatesource1 As String '�D��
Private Str_updatesource2 As String '����
Private strFileHeader As String
Private strFileDetail As String
Private arrTmp


Private Sub cboSheetT15_Click()
'On Error GoTo err_Handle
'
'Dim strFilePath As String, strFieldName As String
'If Right(filLocalFileT15.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT15.Path & "\"
'Else
'    strFilePath = filLocalFileT15.Path
'End If
'
'
'Set rsMainT15 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'
'Set dgMainT15.DataSource = rsMainT15
'
'If rsMainT15 Is Nothing Then
'
'    MsgBox "�d�L���!", 64, "Excel2Recordset"
''
'Else
'
'rsMainT15.Sort = "�X�f�渹"
'
'    SetDataGridColWidth Me.Caption, dgMainT15
'    MsgBox "���u�@��@ " & rsMainT15.RecordCount & "������", 64, "Excel2Recordset"
'
'End If
'
'Exit Sub
'err_Handle:
'
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cboSheetT17_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String
On Error GoTo err_Handle

'�T�{���|�O�_�a"\"
If Right(filLocalFileT17.Path, 1) = "\" Then
    strFilePath = filLocalFileT17.Path
Else
    strFilePath = filLocalFileT17.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "DC�N��" & Chr(9) & "�Ȥ�N��" & Chr(9) & "EXE���s��" & Chr(9) & "SAP DN NO." & Chr(9) & "������O" & Chr(9) & "��ڲ��ͤ�" & Chr(9) & "�w�p�X�f��" & Chr(9) & "����" & Chr(9) & "�ӫ~�s�X" & Chr(9) & "�ӫ~�q�ʼƶq" & Chr(9) & "�P��O" & Chr(9) & "�ӫ~�̤p�ƶq" & Chr(9) & "�Ȥ�i��" & Chr(9) & _
              "�����i�B" & Chr(9) & "��ڥX�f�ƶq" & Chr(9) & "��ڥX�f�ܧO" & Chr(9) & "�|�O" & Chr(9) & "�妸" & Chr(9) & "����˳f���" & Chr(9) & "�f�D" & Chr(9) & "�e�f�a�}" & Chr(9) & "�P���´" & Chr(9) & "��~��" & Chr(9) & "�~�Ȳժ�" & Chr(9) & "���w�渹" & Chr(9) & "SAP�q�f���" & Chr(9) & "��]" & Chr(9) & "�Ȥ�W��" & Chr(9) & "�ƪ`" & Chr(9) & "�Ȥ�q���O" & Chr(9)

'"DC�N��" & Chr(9) & "�Ȥ�N��" & Chr(9) & "SAP DN NO." & Chr(9) & "��ڲ��ͤ�" & Chr(9) & "�w�p�X�f��" & Chr(9) & "����" & Chr(9) & "�ӫ~�s�X" & Chr(9) & "�ӫ~�q�ʼƶq" & Chr(9) & "�e�f�a�}" & Chr(9) & "�Ȥ�W��" & Chr(9) & "�ƪ`" & Chr(9) & "�Ȥ�q���O" & Chr(9)
If Right(filLocalFileT17.Path, 1) <> "\" Then
    strFilePath = filLocalFileT17.Path & "\"
Else
    strFilePath = filLocalFileT17.Path
End If

Set rsMainT17 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT17.FileName)   '���}���|
    
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (cboSheetT17) Then .Sheets(i).Select: Exit For '��w�u�@��
    Next

    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "DC�N��" Then k = i: Exit For
    Next i
    
    If Trim(.Cells(i, 1)) <> "DC�N��" Then MsgBox "�䤣��""DC�N��""���W�١A�ɮ׸��J�פ�!", 64, "�����@�q��פJ": GoTo endsub
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT17 = Nothing: GoTo endsub
    
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT17.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT17.CursorType = adOpenKeyset
    rsMainT17.LockType = adLockOptimistic
    rsMainT17.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT17.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT17(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT17.Update
    k = k + 1
    Loop
    
    If rsMainT17.RecordCount > 0 Then rsMainT17.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
End With
'''''''''''''''''''''''''''
Set dgMainT17.DataSource = rsMainT17

If rsMainT17 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT17
    MsgBox "���u�@��@ " & rsMainT17.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

endsub:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Exit Sub

err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "�����@�q��פJ")

End Sub

Private Sub cboSheetT18_Click()
On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT18.Path, 1) <> "\" Then
    strFilePath = filLocalFileT18.Path & "\"
Else
    strFilePath = filLocalFileT18.Path
End If

'strFieldName = "Delivery" & Chr(9) & "Sold-to.Pt" & Chr(9) & "Name.of.sold-to" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Name.of.the.ship-to.Party" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "Route" & Chr(9) & "Deliv.date" & Chr(9) & "Qty(stckpg.unit)" & Chr(9) & "BUn" & Chr(9) & "Delivery.qty" & Chr(9) & "SU" & Chr(9) & "order.no" & Chr(9) & "po.no" & Chr(9) & "remarks" & Chr(9)

Set rsMainT18 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT18.FileName, cboSheetT18, strFieldName, rsMainT18)

Set dgMainT18.DataSource = rsMainT18

If rsMainT18 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
'
Else

rsMainT18.Sort = "��f"

    SetDataGridColWidth Me.Caption, dgMainT18
    MsgBox "���u�@��@ " & rsMainT18.RecordCount & "������", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub


Private Sub cboSheetT19_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String
On Error GoTo err_Handle

'�T�{���|�O�_�a"\"
If Right(filLocalFileT19.Path, 1) = "\" Then
    strFilePath = filLocalFileT19.Path
Else
    strFilePath = filLocalFileT19.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "��f�渹" & Chr(9) & "������O" & Chr(9) & "���f�ܧO" & Chr(9) & "���f�w�s�a" & Chr(9) & "����" & Chr(9) & "�~��" & Chr(9) & "�̤p���" & Chr(9) & "�ӫ~�q�ʼƶq" & Chr(9) & "�w�p��f���" & Chr(9) & "�P���´" & Chr(9) & "��~��" & Chr(9) & "�Ȥ�N��" & Chr(9) & "SAP���" & Chr(9) & "���f�a�}" & Chr(9) & "�Ȥ�W��" & Chr(9) & "�q��" & Chr(9)
'��f�渹    ������O    ���f�ܧO    ���f�w�s�a  ����    �~��    �̤p���     �ӫ~�q�ʼƶq   �w�p��f���    �P���´    ��~��  �Ȥ�N��    SAP��� ���f�a�}    �Ȥ�W��    �q��


If Right(filLocalFileT19.Path, 1) <> "\" Then
    strFilePath = filLocalFileT19.Path & "\"
Else
    strFilePath = filLocalFileT19.Path
End If

Set rsMainT19 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT19.FileName)   '���}���|
    
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (cboSheetT19) Then .Sheets(i).Select: Exit For '��w�u�@��
    Next

    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "��f�渹" Then k = i: Exit For
    Next i
    
    If Trim(.Cells(i, 1)) <> "��f�渹" Then MsgBox "�䤣��""��f�渹""���W�١A�ɮ׸��J�פ�!", 64, "�����@��h�f�פJ": GoTo endsub
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT19 = Nothing: GoTo endsub
    
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT19.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT19.CursorType = adOpenKeyset
    rsMainT19.LockType = adLockOptimistic
    rsMainT19.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT19.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT19(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT19.Update
    k = k + 1
    Loop
    
    If rsMainT19.RecordCount > 0 Then rsMainT19.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
End With
'''''''''''''''''''''''''''
Set dgMainT19.DataSource = rsMainT19

rsMainT19.Sort = "��f�渹"

If rsMainT19 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT19
    MsgBox "���u�@��@ " & rsMainT19.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

endsub:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Exit Sub

err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "�����@�h�f�פJ")

End Sub

Private Sub cboSheetT20_Change()
'On Error GoTo err_Handle
'
'Dim strFilePath As String, strFieldName As String
'If Right(filLocalFileT20.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT20.Path & "\"
'Else
'    strFilePath = filLocalFileT20.Path
'End If
'
'
'Set rsMainT20 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT20.FileName, cboSheetT20, strFieldName, rsMainT20)
'
'Set dgMainT20.DataSource = rsMainT20
'
'If rsMainT20 Is Nothing Then
'
'    MsgBox "�d�L���!", 64, "Excel2Recordset"
''
'Else
'
'rsMainT20.Sort = "�X�f�渹"
'
'    SetDataGridColWidth Me.Caption, dgMainT20
'    MsgBox "���u�@��@ " & rsMainT20.RecordCount & "������", 64, "Excel2Recordset"
'
'End If
'
'Exit Sub
'err_Handle:
'
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cboSheetT21_Click()
'Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
'Dim str As String, strFieldName As String, strFilePath As String
'On Error GoTo err_Handle
'
''�T�{���|�O�_�a"\"
'If Right(filLocalFileT21.Path, 1) = "\" Then
'    strFilePath = filLocalFileT21.Path
'Else
'    strFilePath = filLocalFileT21.Path & "\"
'End If
'
''�إ����W�ٰ}�C
'strFieldName = "�ռ���" & Chr(9) & "���~�N��" & Chr(9) & "�~�W" & Chr(9) & "�ƶq" & Chr(9) & "�ռ��渹" & Chr(9) & "���X�ܮw�W��" & Chr(9) & "���J�ܮw�W��" & Chr(9) & "�Ƶ�" & Chr(9) & "�U" & Chr(9) & "��" & Chr(9) & "�ƶq" & Chr(9)
''�q��渹    �w�p��f��  �Ȥ�N��    �Ȥ�W��    �橱�N��    �Ƶ�    �~��    �c  �U  ��  �ƶq
'
'If Right(filLocalFileT21.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT21.Path & "\"
'Else
'    strFilePath = filLocalFileT21.Path
'End If
'
'Set rsMainT21 = New ADODB.Recordset
''''''''''''''''''''''''''''
'Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
'    .Workbooks.Open (strFilePath & filLocalFileT21.FileName)   '���}���|
'
'    '�M����w�u�@��
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = (cboSheetT21) Then .Sheets(i).Select: Exit For '��w�u�@��
'    Next
'
'    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
'
'    '�Y�L�ӷ����W��
'    If strFieldName = "" Then
'        '�����W��
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '�ѲĤG�C�}�l�פJ
'    End If
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "�q��渹" Then k = i: Exit For
'    Next i
'
'    '�q��渹    �w�p��f��  �Ȥ�N��    �Ȥ�W��    �橱�N��    �Ƶ�    �~��    �c  �U  ��  �ƶq   '��key��@�w�n���������~
'    If Trim(.Cells(i, 1)) <> "�ռ���" Then MsgBox "�䤣��""�ռ���""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'    If Trim(.Cells(i, 2)) <> "���~�N��" Then MsgBox "�䤣��""���~�N��""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'    If Trim(.Cells(i, 3)) <> "�~�W" Then MsgBox "�䤣��""�~�W""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'    If Trim(.Cells(i, 4)) <> "�ƶq" Then MsgBox "�䤣��""�ƶq""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'    If Trim(.Cells(i, 5)) <> "�ռ��渹" Then MsgBox "�䤣��""�ռ��渹""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'    If Trim(.Cells(i, 6)) <> "���X�ܮw�W��" Then MsgBox "�䤣��""���X�ܮw�W��""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'    If Trim(.Cells(i, 7)) <> "���J�ܮw�W��" Then MsgBox "�䤣��""���J�ܮw�W��""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'    If Trim(.Cells(i, 8)) <> "�Ƶ�" Then MsgBox "�䤣��""�Ƶ�""���W�١A�ɮ׸��J�פ�!", 64, "����RC�q��פJ": GoTo endsub
'
'
'    '�������W��
'    arrTmp = Split(strFieldName, Chr(9))
'
'
'    If UBound(arrTmp) < 1 Then Set rsMainT21 = Nothing: GoTo endsub
'
'    '�إ�Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT21.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT21.CursorType = adOpenKeyset
'    rsMainT21.LockType = adLockOptimistic
'    rsMainT21.Open
'
'    '�g�JRecordset  '�q�o��}�l���U�g
'    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
'    rsMainT21.AddNew
'        For j = 1 To UBound(arrTmp)
'            rsMainT21(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'        Next j
'    rsMainT21.Update
'    k = k + 1
'    Loop
'
'    If rsMainT21.RecordCount > 0 Then rsMainT21.MoveFirst
'
'
'End With
''''''''''''''''''''''''''''
'Set dgMainT21.DataSource = rsMainT21
'
'rsMainT21.Sort = "�ռ��渹"
'
'
'If rsMainT21 Is Nothing Then
'
'    MsgBox "�d�L���!", 64, "Excel2Recordset"
'
'Else
'    SetDataGridColWidth Me.Caption, dgMainT21
'    MsgBox "���u�@��@ " & rsMainT21.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
'
'End If
'
'endsub:
'MyXlsApp.Quit: Set MyXlsApp = Nothing
'Exit Sub
'
'err_Handle:
'
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "����RC�q��פJ")
End Sub


Private Sub cboSheetT22_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String
On Error GoTo err_Handle

'�T�{���|�O�_�a"\"
If Right(filLocalFileT22.Path, 1) = "\" Then
    strFilePath = filLocalFileT22.Path
Else
    strFilePath = filLocalFileT22.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "�h�f�渹" & Chr(9) & "�w�p���h��" & Chr(9) & "�Ȥ�N��" & Chr(9) & "�Ȥ�W��" & Chr(9) & "�橱�N��" & Chr(9) & "�Ƶ�" & Chr(9) & "�~��" & Chr(9) & "�c" & Chr(9) & "�U" & Chr(9) & "��" & Chr(9) & "�ƶq" & Chr(9)
'�h�f�渹    �w�p���h��  �Ȥ�N��    �Ȥ�W��    �橱�N��    �Ƶ�    �~��    �c  �U  ��  �ƶq

If Right(filLocalFileT22.Path, 1) <> "\" Then
    strFilePath = filLocalFileT22.Path & "\"
Else
    strFilePath = filLocalFileT22.Path
End If

Set rsMainT22 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT22.FileName)   '���}���|
    
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (cboSheetT22) Then .Sheets(i).Select: Exit For '��w�u�@��
    Next

    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "�h�f�渹" Then k = i: Exit For
    Next i

    '�h�f�渹    �w�p���h��  �Ȥ�N��    �Ȥ�W��    �橱�N��    �Ƶ�    �~��    �c  �U  ��  �ƶq   '��key��@�w�n���������~
    If Trim(.Cells(i, 1)) <> "�h�f�渹" Then MsgBox "�䤣��""�h�f�渹""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 2)) <> "�w�p���h��" Then MsgBox "�䤣��""�w�p���h��""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 3)) <> "�Ȥ�N��" Then MsgBox "�䤣��""�Ȥ�N��""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 4)) <> "�Ȥ�W��" Then MsgBox "�䤣��""�Ȥ�W��""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 5)) <> "�橱�N��" Then MsgBox "�䤣��""�橱�N��""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 6)) <> "�Ƶ�" Then MsgBox "�䤣��""�Ƶ�""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 7)) <> "�~��" Then MsgBox "�䤣��""�~��""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 8)) <> "�c" Then MsgBox "�䤣��""�c""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 9)) <> "�U" Then MsgBox "�䤣��""�U""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 10)) <> "��" Then MsgBox "�䤣��""��""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    If Trim(.Cells(i, 11)) <> "�ƶq" Then MsgBox "�䤣��""�ƶq""���W�١A�ɮ׸��J�פ�!", 64, "���׾lP&G�q��פJ": GoTo endsub
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT19 = Nothing: GoTo endsub
    
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT22.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT22.CursorType = adOpenKeyset
    rsMainT22.LockType = adLockOptimistic
    rsMainT22.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT22.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT22(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT22.Update
    k = k + 1
    Loop
    
    If rsMainT22.RecordCount > 0 Then rsMainT22.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
End With
'''''''''''''''''''''''''''
Set dgMainT22.DataSource = rsMainT22

rsMainT22.Sort = "�h�f�渹"


If rsMainT22 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT22
    MsgBox "���u�@��@ " & rsMainT22.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

endsub:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Exit Sub

err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "���׾lP&G�q��פJ")
End Sub


Private Sub Cmb_Import1_Click()
 
Dim strFileName As String, strFieldName As String, Str_Filter As String
Str_Filter = ""

''�_�}�ثe��
'LCDisConnect "bestprepares\IPC$"
'LCDisConnect "192.168.200.200\IPC$"
''���s�s�u
'LCConnect "192.168.200.200", "LMBO01", "34245356"

On Error GoTo err_Handle

'MsgBox "�Ш̷ӤU�C�覡�}�l����q��}�ҧ@�~:" & Chr(13) & "1.���I��q��D��(ST�}�Y)�i��q��}��" & Chr(13) & "2.�A�I��q�����(SD�}�Y)�i����Ӷ}��" & Chr(13) & "3.�Ы��T�w��}�l�i��^_^", vbOKOnly + vbInformation, "���_�q��}��"

'�פJ�q��D��
With dlgCommonDialog
    .DialogTitle = "���_�q��D�ɶפJ"
    .CancelError = True
    '.InitDir = App.Path
    .InitDir = "\\192.168.200.200\ftp$\LMBO01\to_Best"
    '.InitDir = "ftp://LMBO01:34245356@192.168.2.202"
    'ToDo: �]�w�q�ι�ܤ��������X�Ф��ݩ�
    .Filter = "ST*.txt|ST*.txt"
    '.Filter = "rtb*.txt|rtb*.txt"
    .ShowOpen
    strFileName = .FileName
    
    If err.Number = cdlCancel Then strFileName = "": Exit Sub
    
    If Len(strFileName) = 0 Then Exit Sub

End With

strFileHeader = strFileName
arrTmp = Split(strFileName, "\")
strFileName = arrTmp(UBound(arrTmp)) '���ocount
Str_Filter = Mid(strFileName, 3, Len(strFileName))

If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "���_�q��}��": Exit Sub '�䤣���ɮ�

Call FixedLenghtText2Recordset(strFileName, "�����q�N��,�q�渹�X,�q����,�o�����X,�o�����X�ˬd�X,�o�����,�Ȥ�s��,�Ȥ�W��,�~�N�N��,�U�f���{,�e�f�a�},�p��,�Τ@�s��,�������B,�ƶq�������B,�S�O�������B,�{������,�f��,�|�e���B,�|�B,�Ƶ�," & _
                                            "�Ȥ�q��s��,�H�f���o���X,�H�f���q��X,�p�⪫�y�O,�e�f�_,�q�����,�ꦬ�q�B�zMARK,�s���H,�q��,�~�N�m�W,�D�ީm�W,���e�Ȥ�,�w�p���,�B�O,�I�ڤ覡,�~�Ȥ��,�O�_���q�l�o��,�`���q,�H�d��4�X,�N���f��,�o���C�L�覡,�q��2,�έp��H," & _
                                            "�����O,��F��,�Ӽh,�V�w�q��,���f��,�|��/�|�v,�Ȥ�²��,�q�浡�f,���p�q�渹�X", "1,8,7,10,2,7,8,50,3,2,70,1,8,8,8,8,8,10,10,8,70,25,1,1,1,1,2,1,12,20,12,12,50,7,8,1,20,1,6,4,10,1,20,8,3,3,2,1,12,10,40,10,10", rsMainT16)

'���f�ܨ�����T�X add by Gemini @ 20160425
rsMainT16.MoveFirst
Do While Not rsMainT16.EOF
    rsMainT16("���f��") = Left(rsMainT16("���f��"), 3)
    rsMainT16.MoveNext
Loop

Set dgMainT16.DataSource = rsMainT16

'Recordset2Excel "TEST", rsMainT16


'�����ɦW
Str_updatesource1 = strFileName

''�פJ�q�������
'With dlgCommonDialog
'    .DialogTitle = "���_�q������ɶפJ"
'    .CancelError = True
'    '.InitDir = App.Path
'    .InitDir = "\\192.168.200.200\ftp$\LMBO01\to_Best"
'    'ToDo: �]�w�q�ι�ܤ��������X�Ф��ݩ�
'    '.Filter = "SD*.txt|SD*.txt"
'    .Filter = "SD" & Str_Filter & "|" & "SD" & Str_Filter
''    .Filter = "rdb*.txt|rdb*.txt"
'    .ShowOpen
'    strFileName = .FileName
'
'    If err.Number = cdlCancel Then strFileName = "": Exit Sub
'
'    If Len(strFileName) = 0 Then Exit Sub
'
'End With

strFileDetail = Replace(strFileHeader, strFileName, "SD" & Str_Filter)

strFileName = "SD" & Str_Filter

'arrTmp = Split(strFileName, "\")
'strFileName = arrTmp(UBound(arrTmp)) '���ocount


If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "���_�q������ɶפJ": Exit Sub '�䤣���ɮ�

Call FixedLenghtText2Recordset(strFileName, "�q�渹�X,���~�s��,���~�W��,�q�f�q,���(���|),�q�f���B(���|),���(�t�|),�q�f���B(�t�|),�q�f�q-�ꦬ�q,��ڱ��X,�渹,���,�q�����,�o�����ӦC�L�_,������", "8,16,60,10,8,10,8,10,10,25,7,2,2,1,20", rsMainT16_1)

Set dgMainT16_1.DataSource = rsMainT16_1

'MsgBox "���_�q��D�ɶ}��:" & rsMainT16.RecordCount & "���A�нT�{���ƬO�_���T!", vbOKOnly + vbInformation, "���_�q��}��"

MsgBox "���_�q��D�ɶ}��:" & rsMainT16.RecordCount & "���A�q������ɶ}��:" & rsMainT16_1.RecordCount & "���A�нT�{���ƬO�_���T!", vbOKOnly + vbInformation, "���_�q������ɶ}��"

Str_updatesource2 = strFileName

rsMainT16.Sort = "�q�渹�X,�q�����"
rsMainT16_1.Sort = "�q�渹�X,�q�����,�渹"


lab_Orders.Caption = "�q��:" & rsMainT16.RecordCount & "��":
lab_Orderdetail.Caption = "����:" & rsMainT16_1.RecordCount & "��":

'Recordset2Excel "TEST", rsMainT16_1

'�Ƨ�


'���ɤs���{���X
'datagrid1=�X�f�`��;�P�f���Ӫ�;�༷���Ӫ�;�H�w�P�f���Ӫ�
'Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
'Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String, S As String, Str_Sku As String
'Dim bl_Check As Boolean '�ˬd�פJ�����~���L�X�{�b�`��~�S���Nstop
'bl_Check = True
'S = "": Str_Sku = ""
''S�O���W�@�����f�渹,Str_sku�O���W�@���~��,�p�G�ťիh�a�W�@��
'
'On Error GoTo err_Handle
'SSTab2.Tab = 0: SSTab2.Enabled = False: Cmb_Import1.Enabled = False
'
'Call DB_Connect_Self(cn_string) '�إ߷s�s�u
'
''�T�{���|�O�_�a"\"
'If Right(filLocalFileT16.Path, 1) = "\" Then
'    strFilePath = filLocalFileT16.Path
'Else
'    strFilePath = filLocalFileT16.Path & "\"
'End If
'
''�إ����W�ٰ}�C
'strFieldName = "��ڦW��" & Chr(9) & "�P�f��O" & Chr(9) & "�P�f�渹" & Chr(9) & "��ڤ��" & Chr(9) & "���w���" & Chr(9) & "�Ȥ�N��" & Chr(9) & "�Ȥ�²��" & Chr(9) & "���" & Chr(9) & "�e�f�a�}" & Chr(9) & "�Ƶ�" & Chr(9) '�X�f�`��
'If Right(filLocalFileT16.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT16.Path & "\"
'Else
'    strFilePath = filLocalFileT16.Path
'End If
'
'Set rsMainT16 = New ADODB.Recordset
'
'Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
'    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
'    '�M����w�u�@��
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "�X�f�`��" Then .Sheets(i).Select: Exit For '��w�u�@��
'    Next
'
'    If (.ActiveSheet.Name) <> "�X�f�`��" Then MsgBox "�䤣��X�f�`��u�@��!!", 16, "�}���ɮפ���": GoTo endsub
'
'    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
'
'    '�Y�L�ӷ����W��
'    If strFieldName = "" Then
'        '�����W��
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '�ѲĤG�C�}�l�פJ
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "��ڦW��" Then k = i: Exit For
'    Next i
'
'    '�������W��
'    arrTmp = Split(strFieldName, Chr(9))
'
'    'Dim rsMainT15 As New ADODB.Recordset
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16 = Nothing: GoTo endsub
'
'    '�إ�Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�X�f�`��u�@��A�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16.CursorType = adOpenKeyset
'    rsMainT16.LockType = adLockOptimistic
'    rsMainT16.Open
'
'    '�g�JRecordset  '�q�o��}�l���U�g
'    Do While Len(RTrim(.Cells(k + 1, 8))) > 0
'    rsMainT16.AddNew
'        For j = 1 To UBound(arrTmp)
'            rsMainT16(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'        Next j
'    rsMainT16.Update
'    k = k + 1
'    Loop
'
'    If rsMainT16.RecordCount > 0 Then rsMainT16.MoveFirst
'
'End With
'''�H�U�Ndatagrid����Ʀs�J���stable��
''    '�s�Wtabel
''    cn_Self.Execute "if object_id ('tempdb..##all_data') is not null drop table tempdb..##all_data ", RowsAffect, adExecuteNoRecords
''    str_TmpSQL = "CREATE TABLE tempdb..##all_data(��ڦW�� varchar(30),�P�f��O varchar(30),�P�f�渹 varchar(30),��ڤ�� varchar(30),���w��� varchar(30),�Ȥ�N�� varchar(30),�Ȥ�²�� varchar(30),��� varchar(30),�e�f�a�} varchar(80),�Ƶ� varchar(60))"
''    Call Confirm_Recordset_Closed(tmp_Rs)
''    tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
''    '��Jdatagrid��ƨ�table
''    dgMainT16.Visible = False
''    Do While Not rsMainT16.EOF
''        '�N�פJ��excel��� �s�J�Ȧs����ƪ�##SkuCompare��
''        str_TmpSQL = "INSERT INTO tempdb..##all_data (��ڦW��,�P�f��O,�P�f�渹,��ڤ��,���w���,�Ȥ�N��,�Ȥ�²��,���,�e�f�a�},�Ƶ�) " & _
''                     "VALUES ('" & Trim(rsMainT16("��ڦW��").Value) & "','" & Trim(rsMainT16("�P�f��O").Value) & "','" & Trim(rsMainT16("�P�f�渹").Value) & "','" & _
''                     "" & Trim(rsMainT16("��ڤ��").Value) & "','" & Trim(rsMainT16("���w���").Value) & "','" & Trim(rsMainT16("�Ȥ�N��").Value) & "','" & Trim(rsMainT16("�Ȥ�²��").Value) & "','" & _
''                     "" & Trim(rsMainT16("���").Value) & "','" & Trim(rsMainT16("�e�f�a�}").Value) & "','" & Trim(rsMainT16("�Ƶ�").Value) & "')"
''
''        Call Confirm_Recordset_Closed(tmp_Rs)
''        tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
''        rsMainT16.MoveNext
''    Loop
''    dgMainT16.Visible = True
'
'Set dgMainT16.DataSource = rsMainT16
'
'If rsMainT16 Is Nothing Then
'
'    MsgBox "�d�L���!", 64, "Excel2Recordset"
'
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16 '�]�w��e
'End If
'
''�p�G�S���X�f�`��h����
'If rsMainT16 Is Nothing Then MsgBox "�䤣��X�f�`��u�@��!!", 16, "�}���ɮפ���": SSTab2.Enabled = True: Cmb_Import1.Enabled = True: Exit Sub
'If rsMainT16.EOF Then MsgBox "�䤣��X�f�`��u�@��!!", 16, "�}���ɮפ���": SSTab2.Enabled = True: Cmb_Import1.Enabled = True: Exit Sub
'
''/////////////////////////////////////////////////////////////////////////////////�פJ�P�f������/////////////////////////////////////////////////////////////////////////////////
'SSTab2.Tab = 1
'strFieldName = "�P�f�渹" & Chr(9) & "�P�f���" & Chr(9) & "���w���" & Chr(9) & "�Ȥ�N��" & Chr(9) & "�Ȥ�²��" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�P�f�ƶq" & Chr(9) & "��/�ƫ~�q" & Chr(9) & "���" & Chr(9) & "�帹" & Chr(9) & "�Ƶ�" & Chr(9) & "���" & Chr(9) & "�e�f�a�}" & Chr(9)  '�P�f���Ӫ�
'
'Set rsMainT16_1 = New ADODB.Recordset
'
''Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
''    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
'    '�M����w�u�@��
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "�P�f���Ӫ�" Then .Sheets(i).Select: Exit For '��w�u�@��
'    Next
'
'    If (.ActiveSheet.Name) <> "�P�f���Ӫ�" Then MsgBox "�䤣��P�f���Ӫ�u�@��!!", 16, "�}���ɮפ���": GoTo endsub
'
'    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
'
'    '�Y�L�ӷ����W��
'    If strFieldName = "" Then
'        '�����W��
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '�ѲĤG�C�}�l�פJ
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "�P�f�渹" Then k = i: Exit For
'    Next i
'
'    '�������W��
'    arrTmp = Split(strFieldName, Chr(9))
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16 = Nothing: GoTo endsub
'
'    '�إ�Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�P�f���Ӫ�u�@��A�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16_1.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16_1.CursorType = adOpenKeyset
'    rsMainT16_1.LockType = adLockOptimistic
'    rsMainT16_1.Open
'    rsMainT16.MoveFirst: S = ""
'
'    '�g�JRecordset  '�q�o��}�l���U�g
'    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '�D�Ǽƶq���ŭȫh����
''    If RTrim(.Cells(k + 1, 6)) = "60400119" Then '�ư��B�O
''    Else
'        rsMainT16_1.AddNew
'            For j = 1 To UBound(arrTmp)
'                If Len(Trim(rsMainT16_1("�P�f�渹").Value)) = 0 Then rsMainT16_1("�P�f�渹").Value = S
'                If Len(Trim(rsMainT16_1("�~��").Value)) = 0 Then rsMainT16_1("�~��").Value = Str_Sku
'                If j = UBound(arrTmp) Then    '�e�f�a�}���
'                rsMainT16.MoveFirst
'                    Do While Not rsMainT16.EOF
'                        Str_Sku = Trim(rsMainT16_1("�~��").Value)
'                        If Trim(rsMainT16_1("�P�f�渹").Value) = Trim(rsMainT16("�P�f��O").Value) & "-" & Trim(rsMainT16("�P�f�渹").Value) Then rsMainT16_1("�e�f�a�}").Value = Trim(rsMainT16("�e�f�a�}").Value):  rsMainT16_1("���w���").Value = Trim(rsMainT16("���w���").Value): rsMainT16_1("�Ƶ�").Value = Trim(rsMainT16("�Ƶ�").Value): rsMainT16_1("���").Value = Trim(rsMainT16("���").Value): S = Trim(rsMainT16_1("�P�f�渹").Value): bl_Check = False: Exit Do
'                        rsMainT16.MoveNext
'                    Loop
'                    If bl_Check = True Then MsgBox "�X�f�`��d�L:" & Trim(rsMainT16_1("�P�f�渹").Value) & "���!", 64, "�P�f���Ӫ�פJ����": SSTab2.Enabled = True: Cmb_Import1.Enabled = True: GoTo endsub
'                Else
'                    rsMainT16_1(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'                End If
'            Next j
'        rsMainT16_1.Update
''    End If
'    k = k + 1
'
'    Loop
'
'    If rsMainT16_1.RecordCount > 0 Then rsMainT16_1.MoveFirst
'
'End With
'
'Set dgMainT16_1.DataSource = rsMainT16_1
'
'If rsMainT16_1 Is Nothing Then
'    MsgBox "�P�f���Ӫ�d�L���!", 64, "Excel2Recordset"
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16_1
'End If
''////////////////////////////////////////////////////////////////////////////////�פJ�༷���Ӫ�/////////////////////////////////////////////////////////////////////////////////////
'SSTab2.Tab = 2: S = "": Str_Sku = ""
'
''�إ����W�ٰ}�C
'strFieldName = "��O-�渹" & Chr(9) & "��ڤ��" & Chr(9) & "��J�w�O" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�༷�ƶq" & Chr(9) & "���" & Chr(9) & "�帹" & Chr(9) & "�Ƶ�" & Chr(9) & "���w���" & Chr(9) & "���" & Chr(9) & "�e�f�a�}" & Chr(9) '�༷���Ӫ�
'If Right(filLocalFileT16.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT16.Path & "\"
'Else
'    strFilePath = filLocalFileT16.Path
'End If
'
'Set rsMainT16_2 = New ADODB.Recordset
'
''Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
''    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
'    '�M����w�u�@��
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "�༷���Ӫ�" Then .Sheets(i).Select: Exit For '��w�u�@��
'    Next
'
'    If (.ActiveSheet.Name) <> "�༷���Ӫ�" Then MsgBox "�䤣���༷���Ӫ�u�@��!!", 16, "�}���ɮפ���": GoTo endsub
'
'    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
'
'    '�Y�L�ӷ����W��
'    If strFieldName = "" Then
'        '�����W��
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '�ѲĤG�C�}�l�פJ
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "��O-�渹" Then k = i: Exit For
'    Next i
'
'    '�������W��
'    arrTmp = Split(strFieldName, Chr(9))
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16_2 = Nothing: GoTo endsub
'
'    '�إ�Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�༷���Ӫ�u�@��A�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16_2.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16_2.CursorType = adOpenKeyset
'    rsMainT16_2.LockType = adLockOptimistic
'    rsMainT16_2.Open
'    rsMainT16.MoveFirst: S = ""
'    '�g�JRecordset  '�q�o��}�l���U�g
'    Do While Len(RTrim(.Cells(k + 1, 6))) > 0   '�D�Ǽƶq���ŭȫh����
'    rsMainT16_2.AddNew
'            For j = 1 To UBound(arrTmp)
'                If Len(Trim(rsMainT16_2("��O-�渹").Value)) = 0 Then rsMainT16_2("��O-�渹").Value = S
'                If Len(Trim(rsMainT16_2("�~��").Value)) = 0 Then rsMainT16_2("�~��").Value = Str_Sku
'                If j = UBound(arrTmp) Then    '�e�f�a�}���
'                    rsMainT16.MoveFirst
'                    Do While Not rsMainT16.EOF
'                        Str_Sku = Trim(rsMainT16_2("�~��").Value)
'                        If Trim(rsMainT16_2("��O-�渹").Value) = Trim(rsMainT16("�P�f��O").Value) & "-" & Trim(rsMainT16("�P�f�渹").Value) Then
'                            rsMainT16_2("�e�f�a�}").Value = Trim(rsMainT16("�e�f�a�}").Value): rsMainT16_2("���").Value = Trim(rsMainT16("���").Value): rsMainT16_2("���w���").Value = Trim(rsMainT16("���w���").Value): rsMainT16_2("�Ƶ�").Value = Trim(rsMainT16("�Ƶ�").Value): rsMainT16_2("��ڤ��").Value = Trim(rsMainT16("��ڤ��").Value): S = Trim(rsMainT16_2("��O-�渹").Value): bl_Check = False: Exit Do
'                        End If
'                        rsMainT16.MoveNext
'                    Loop
'                    If bl_Check = True Then MsgBox "�X�f�`��d�L:" & Trim(rsMainT16_2("��O-�渹").Value) & "���!", 64, "�༷���Ӫ�פJ����":  SSTab2.Enabled = True: Cmb_Import1.Enabled = True: GoTo endsub
'                Else
'                    rsMainT16_2(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'                End If
'                 bl_Check = True
'            Next j
'    rsMainT16_2.Update
'    k = k + 1
'    Loop
'
'    If rsMainT16_2.RecordCount > 0 Then rsMainT16_2.MoveFirst
'
'End With
'
'Set dgMainT16_2.DataSource = rsMainT16_2
'
'If rsMainT16_2 Is Nothing Then
'    MsgBox "�༷���Ӫ�d�L���!", 64, "Excel2Recordset"
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16_2
'End If
'
''///////////////////////////////////////////////////////////////////////////////////�פJ�H�w�P�f���Ӫ�//////////////////////////////////////////////////////////////////////////////////////////
'SSTab2.Tab = 3: S = "": Str_Sku = ""
'
''�إ����W�ٰ}�C
'strFieldName = "�P�f�渹" & Chr(9) & "�P�f���" & Chr(9) & "���w���" & Chr(9) & "�Ȥ�N��" & Chr(9) & "�Ȥ�²��" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�P�f�ƶq" & Chr(9) & "��/�ƫ~�q" & Chr(9) & "���" & Chr(9) & "�帹" & Chr(9) & "�Ƶ�" & Chr(9) & "���" & Chr(9) & "�e�f�a�}" & Chr(9) '�H�w�P�f���Ӫ�
'If Right(filLocalFileT16.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT16.Path & "\"
'Else
'    strFilePath = filLocalFileT16.Path
'End If
'
'Set rsMainT16_3 = New ADODB.Recordset
'
''Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
''    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
'
'    '�M����w�u�@��
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "�H�w�P�f���Ӫ�" Then .Sheets(i).Select: Exit For '��w�u�@��
'    Next
'
'    If (.ActiveSheet.Name) <> "�H�w�P�f���Ӫ�" Then MsgBox "�䤣��H�w�P�f���Ӫ�u�@��!!", 16, "�}���ɮפ���": GoTo endsub
'
'    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
'
'    '�Y�L�ӷ����W��
'    If strFieldName = "" Then
'        '�����W��
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '�ѲĤG�C�}�l�פJ
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "�P�f�渹" Then k = i: Exit For
'    Next i
'
'    '�������W��
'    arrTmp = Split(strFieldName, Chr(9))
'
'    'Dim rsMainT15 As New ADODB.Recordset
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16_3 = Nothing: GoTo endsub
'    '�إ�Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�H�w�P�f���Ӫ�u�@��A�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16_3.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16_3.CursorType = adOpenKeyset
'    rsMainT16_3.LockType = adLockOptimistic
'    rsMainT16_3.Open
'    rsMainT16.MoveFirst: S = ""
'
'    '�g�JRecordset  '�q�o��}�l���U�g
'    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '�D�Ǽƶq���ŭȫh����
'    rsMainT16_3.AddNew
'
'            For j = 1 To UBound(arrTmp)
'                If Len(Trim(rsMainT16_3("�P�f�渹").Value)) = 0 Then rsMainT16_3("�P�f�渹").Value = S
'                If Len(Trim(rsMainT16_3("�~��").Value)) = 0 Then rsMainT16_3("�~��").Value = Str_Sku
'                If j = UBound(arrTmp) Then    '�e�f�a�}���
'                    If Trim(rsMainT16_3("�Ȥ�N��")) = "1201011001" Or Trim(rsMainT16_3("�Ȥ�N��")) = "1201011002" Or Trim(rsMainT16_3("�Ȥ�N��")) = "1201011003" Or Trim(rsMainT16_3("�Ȥ�N��")) = "1201011006" Or Trim(rsMainT16_3("�Ȥ�N��")) = "1201011009" Or Trim(rsMainT16_3("�Ȥ�N��")) = "1201011004" Or Trim(rsMainT16_3("�Ȥ�N��")) = "1201011010" Or Trim(rsMainT16_3("�Ȥ�N��")) = "1201011011" Then
'                        S = Trim(rsMainT16_3("�P�f�渹").Value): Str_Sku = Trim(rsMainT16_3("�~��").Value)
'
'                        '�z�L�Ȥ�s���d�X�t�Τ�����f�a�},�Ȥ�W�� ;�]���D�ɨS���H�w�P�f���Ӫ��Ӷ�
'                        Call Confirm_Recordset_Closed(tmp_Rs)
'                        str_SQL = "select full_name,address from trp01m where storerkey = 'LMYS01' and consigneekey = '" & Trim(rsMainT16_3("�Ȥ�N��").Value) & "'"
'                        tmp_Rs.Open str_SQL, cn
'
'                        If Not tmp_Rs.EOF Then rsMainT16_3("�Ȥ�²��") = Trim(tmp_Rs("full_name")): rsMainT16_3("�e�f�a�}") = Trim(tmp_Rs("address"))
'                        tmp_Rs.Close
'
'                    Else
'                        MsgBox "�o�{�D���ݫȤ�N��: " & Trim(rsMainT16_3("�Ȥ�N��")) & " �нT�{�O�_���n�Ʀh�B�O�_��Ȥ�D�ɷs�W��ơB�ýгq����T���ק�{��!", 64, "�H�w�P�f���Ӫ�פJ�פ�": GoTo endsub   '�o�{6�ӫ��w���Ȥ�N���H�~������,�h����פJ
'                    End If
'               Else
'                    rsMainT16_3(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j))):
'                End If
'            Next j
'    rsMainT16_3.Update
'    k = k + 1
'    Loop
'
'    If rsMainT16_3.RecordCount > 0 Then rsMainT16_3.MoveFirst
'
'End With
'
'SSTab2.Enabled = True: SSTab2.Tab = 0: Cmb_Import1.Enabled = True
'Set dgMainT16_3.DataSource = rsMainT16_3
'
'If rsMainT16_3 Is Nothing Then
'    MsgBox "�H�w�P�f���Ӫ�d�L���!", 64, "���ɤs�q��}��"
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16_3
'    MsgBox "���q���ɦ@�פJ:" & Chr(13) & "���f�`��:" & rsMainT16.RecordCount & "������" & Chr(13) & "" & _
'                                          "�P�f���Ӫ�:" & rsMainT16_1.RecordCount & "������" & Chr(13) & "" & _
'                                          "�༷���Ӫ�:" & rsMainT16_2.RecordCount & "������" & Chr(13) & "" & _
'                                          "�H�w�P�f���Ӫ�:" & rsMainT16_3.RecordCount & "������" & Chr(13) & "" & _
'                                          "�нT�{���ƬO�_���T!", 64, "���ɤs�q��}��"
'End If
'
''�p�G���X�f�`��A��L�T�Ӥu�@��S����ƫh���ܡA������
'If rsMainT16_1.RecordCount = 0 And rsMainT16_2.RecordCount = 0 And rsMainT16_3.RecordCount = 0 Then MsgBox "���q��L�Ӷ���ơA�нT�{���q��O�_���T!", vbCritical, "���ɤs�q��}��"
'
'endsub:
'SSTab2.Enabled = True: Cmb_Import1.Enabled = True
'MyXlsApp.Quit: Set MyXlsApp = Nothing

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
Set rsMainT16_1 = Nothing: Set rsMainT16 = Nothing
SSTab2.Enabled = True: Cmb_Import1.Enabled = True
End Sub

Private Sub Cmb_Import2_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String, S As String

On Error GoTo err_Handle
SSTab2.Tab = 1
Call DB_Connect_Self(cn_string) '�إ߷s�s�u
'�T�{���|�O�_�a"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "�P�f�渹" & Chr(9) & "�P�f���" & Chr(9) & "���w���" & Chr(9) & "�Ȥ�N��" & Chr(9) & "�Ȥ�²��" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�P�f�ƶq" & Chr(9) & "��/�ƫ~�q" & Chr(9) & "���" & Chr(9) & "�帹" & Chr(9) & "�Ƶ�" & Chr(9) & "���" & Chr(9) & "�e�f�a�}" & Chr(9)  '�P�f���Ӫ�
If Right(filLocalFileT16.Path, 1) <> "\" Then
    strFilePath = filLocalFileT16.Path & "\"
Else
    strFilePath = filLocalFileT16.Path
End If

Set rsMainT16_1 = New ADODB.Recordset

Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "�P�f���Ӫ�" Then .Sheets(i).Select: Exit For '��w�u�@��
    Next

    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "�P�f�渹" Then k = i: Exit For
    Next i
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT16 = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT16_1.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT16_1.CursorType = adOpenKeyset
    rsMainT16_1.LockType = adLockOptimistic
    rsMainT16_1.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '�D�Ǽƶq���ŭȫh����
    If RTrim(.Cells(k + 1, 7)) = "�B�O" Then
    Else
        rsMainT16_1.AddNew
            For j = 1 To UBound(arrTmp)
                rsMainT16_1(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
            Next j
        rsMainT16_1.Update
    End If
    k = k + 1
    Loop
    
    If rsMainT16_1.RecordCount > 0 Then rsMainT16_1.MoveFirst

endsub:
End With
    
Set dgMainT17.DataSource = rsMainT16_1

If rsMainT16_1 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT17
    MsgBox "���u�@��@ " & rsMainT16_1.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If
 Call Cmb_Import3_Click:
Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Cmb_Import3_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String

On Error GoTo err_Handle
Call DB_Connect_Self(cn_string) '�إ߷s�s�u
SSTab2.Tab = 2
'�T�{���|�O�_�a"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "��O-�渹" & Chr(9) & "��ڤ��" & Chr(9) & "��J�w�O" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�༷�ƶq" & Chr(9) & "���" & Chr(9) & "�帹" & Chr(9) & "�Ƶ�" & Chr(9) '�༷���Ӫ�
If Right(filLocalFileT16.Path, 1) <> "\" Then
    strFilePath = filLocalFileT16.Path & "\"
Else
    strFilePath = filLocalFileT16.Path
End If

Set rsMainT16_2 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "�༷���Ӫ�" Then .Sheets(i).Select: Exit For '��w�u�@��
    Next

    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "��O-�渹" Then k = i: Exit For
    Next i
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT16_2 = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT16_2.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT16_2.CursorType = adOpenKeyset
    rsMainT16_2.LockType = adLockOptimistic
    rsMainT16_2.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 6))) > 0   '�D�Ǽƶq���ŭȫh����
    rsMainT16_2.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT16_2(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT16_2.Update
    k = k + 1
    Loop
    
    If rsMainT16_2.RecordCount > 0 Then rsMainT16_2.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
endsub:
End With

''�H�U�Ndatagrid����Ʀs�J���stable��
'    '�s�Wtabel
'    cn_Self.Execute "if object_id ('tempdb..##data2') is not null drop table tempdb..##data2 ", RowsAffect, adExecuteNoRecords
'    str_TmpSQL = "CREATE TABLE tempdb..##data2(��O�渹 varchar(50),��ڤ�� varchar(50),��J�w�O varchar(50),�~�� varchar(50),�~�W varchar(80),�༷�ƶq varchar(50),��� varchar(80),�帹 varchar(50),�Ƶ� varchar(80))"
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'    '��Jdatagrid��ƨ�table
'    dgMainT18.Visible = False
'    Do While Not rsMainT16_2.EOF
'        '�N�פJ��excel��� �s�J�Ȧs����ƪ�##SkuCompare��
'        str_TmpSQL = "INSERT INTO tempdb..##data2 (��O�渹,��ڤ��,��J�w�O,�~��,�~�W,�༷�ƶq,���,�帹,�Ƶ�) " & _
'                     "VALUES ('" & Trim(rsMainT16_2("��O-�渹").Value) & "','" & Trim(rsMainT16_2("��ڤ��").Value) & "','" & Trim(rsMainT16_2("��J�w�O").Value) & "','" & _
'                     "" & Trim(rsMainT16_2("�~��").Value) & "','" & Trim(rsMainT16_2("�~�W").Value) & "','" & Trim(rsMainT16_2("�༷�ƶq").Value) & "','" & Trim(rsMainT16_2("���").Value) & "','" & _
'                     "" & Trim(rsMainT16_2("�帹").Value) & "','" & Trim(rsMainT16_2("�Ƶ�").Value) & "')"
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'        rsMainT16_2.MoveNext
'    Loop
'    dgMainT18.Visible = True
    
Set dgMainT16_2.DataSource = rsMainT16_2

If rsMainT16_2 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT16_2
    MsgBox "���u�@��@ " & rsMainT16_2.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If
 Call Cmb_Import4_Click:
Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub Cmb_Import4_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String
On Error GoTo err_Handle
Call DB_Connect_Self(cn_string) '�إ߷s�s�u
SSTab2.Tab = 3
'�T�{���|�O�_�a"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "�P�f�渹" & Chr(9) & "�P�f���" & Chr(9) & "���w���" & Chr(9) & "�Ȥ�N��" & Chr(9) & "�Ȥ�²��" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�P�f�ƶq" & Chr(9) & "��/�ƫ~�q" & Chr(9) & "���" & Chr(9) & "�帹" & Chr(9) & "�Ƶ�" & Chr(9) & "���" & Chr(9) '�H�w�P�f���Ӫ�
If Right(filLocalFileT16.Path, 1) <> "\" Then
    strFilePath = filLocalFileT16.Path & "\"
Else
    strFilePath = filLocalFileT16.Path
End If

Set rsMainT16_3 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "�H�w�P�f���Ӫ�" Then .Sheets(i).Select: Exit For '��w�u�@��
    Next

    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "�P�f�渹" Then k = i: Exit For
    Next i
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT16_3 = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT16_3.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT16_3.CursorType = adOpenKeyset
    rsMainT16_3.LockType = adLockOptimistic
    rsMainT16_3.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '�D�Ǽƶq���ŭȫh����
    rsMainT16_3.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT16_3(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT16_3.Update
    k = k + 1
    Loop
    
    If rsMainT16_3.RecordCount > 0 Then rsMainT16_3.MoveFirst
  
endsub:
End With


''�H�U�Ndatagrid����Ʀs�J���stable��
'    '�s�Wtabel
'    cn_Self.Execute "if object_id ('tempdb..##data3') is not null drop table tempdb..##data3 ", RowsAffect, adExecuteNoRecords
'    str_TmpSQL = "CREATE TABLE tempdb..##data3(�P�f�渹 varchar(50),�P�f��� varchar(50),���w��� varchar(50),�Ȥ�N�� varchar(50),�Ȥ�²�� varchar(50),�~�� varchar(50),�~�W varchar(80),�P�f�ƶq varchar(50),�سƫ~�q varchar(50),��� varchar(50),�帹 varchar(50),�Ƶ� varchar(80),��� varchar(50))"
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'    '��Jdatagrid��ƨ�table
'    dgMainT19.Visible = False
'    Do While Not rsMainT16_3.EOF
'        '�N�פJ��excel��� �s�J�Ȧs����ƪ�##SkuCompare��
'        str_TmpSQL = "INSERT INTO tempdb..##data3 (�P�f�渹,�P�f���,���w���,�Ȥ�N��,�Ȥ�²��,�~��,�~�W,�P�f�ƶq,�سƫ~�q,���,�帹,�Ƶ�,���) " & _
'                     "VALUES ('" & Trim(rsMainT16_3("�P�f�渹").Value) & "','" & Trim(rsMainT16_3("�P�f���").Value) & "','" & Trim(rsMainT16_3("���w���").Value) & "','" & _
'                     "" & Trim(rsMainT16_3("�Ȥ�N��").Value) & "','" & Trim(rsMainT16_3("�Ȥ�²��").Value) & "','" & Trim(rsMainT16_3("�~��").Value) & "','" & Trim(rsMainT16_3("�~�W").Value) & "','" & _
'                     "" & Trim(rsMainT16_3("�P�f�ƶq").Value) & "','" & Trim(rsMainT16_3("��/�ƫ~�q").Value) & "','" & Trim(rsMainT16_3("���").Value) & "','" & Trim(rsMainT16_3("�帹").Value) & "','" & Trim(rsMainT16_3("�Ƶ�").Value) & "','" & Trim(rsMainT16_3("���").Value) & "')"
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'        rsMainT16_3.MoveNext
'    Loop
'    dgMainT19.Visible = True
    
Set dgMainT16_3.DataSource = rsMainT16_3

If rsMainT16_3 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT16_3
    MsgBox "���u�@��@ " & rsMainT16_3.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub



Private Sub Cmd_Impdata_Click()
Dim bl_OrderCheck As Boolean '�ˬd�q��O�_�X�{�b���Ӥ�
Dim str_Storerkey As String, str_Priority As String, Str_Address As String, Str_Lot05 As String
Dim Int_RC As Integer: Dim Int_C As Integer: Dim Int_i As Integer: Dim Int_otqty As Integer
Dim Str_AllOrderkey As String
Str_AllOrderkey = ""
Int_RC = 0: Int_C = 0: Int_i = 0 '�p��q�����O������
Int_otqty = 0 '�p��q����
str_Storerkey = "LMBO01"

bl_OrderCheck = False
If rsMainT16 Is Nothing Then Exit Sub
If rsMainT16.EOF Then Exit Sub
If rsMainT16_1 Is Nothing Then Exit Sub
If rsMainT16_1.EOF Then Exit Sub

'GoTo copy:
On Error GoTo err_Handle

SSTab2.Enabled = False: Cmd_Impdata.Enabled = False

'�������--�P�_�ɮ׬O�_�w��J

Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where storerkey = '" & str_Storerkey & "' and rtrim(updatesource)='" & Str_updatesource1 & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

'�ˬd�D�ɤ����q�渹�X�A���L�X�{�b�q����Ӥ�
rsMainT16.MoveFirst: rsMainT16_1.MoveFirst
Do While Not rsMainT16.EOF
        Do While Not rsMainT16_1.EOF
            If RTrim(rsMainT16.Fields("�q�渹�X")) + RTrim(rsMainT16.Fields("�q�����")) = RTrim(rsMainT16_1.Fields("�q�渹�X")) + RTrim(rsMainT16_1.Fields("�q�����")) Then
            '���T�A���X�{��ơC
                bl_OrderCheck = True
                Exit Do
            End If
            rsMainT16_1.MoveNext
        Loop
    If bl_OrderCheck = False Then MsgBox "�q�渹�X+�q�����:" & RTrim(rsMainT16.Fields("�q�渹�X")) & RTrim(rsMainT16.Fields("�q�����")) & " ���X�{�b�����ɤ��A�нT�{�q���ɸ�ƬO�_���T�A�q����J����", vbOKOnly + vbCritical, "���_�q��פJ": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: bl_OrderCheck = False: Exit Sub
    bl_OrderCheck = False
    rsMainT16.MoveNext
    rsMainT16_1.MoveFirst
Loop

rsMainT16_1.MoveFirst
Do While Not rsMainT16_1.EOF
    If RTrim(rsMainT16_1.Fields("������")) = "1" Or RTrim(rsMainT16_1.Fields("������")) = "0" Then
    '���������i�H��0��1
    MsgBox "�q�渹�X:" & Trim(rsMainT16_1("�q�渹�X")) & "�A������=" & Trim(rsMainT16("������")) & "�A���i��1��0�ȡA�q��פJ����", vbCritical + vbOKOnly
    Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
    End If
    rsMainT16_1.MoveNext
Loop
rsMainT16_1.MoveFirst

'�ˬd�w�p������i�p�󤵤�
    rsMainT16.MoveFirst
    Do While Not rsMainT16.EOF
        If Val(Left(Trim(rsMainT16("�w�p���")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("�w�p���")), 4, 2) & "/" & Right(Trim(rsMainT16("�w�p���")), 2) < Format(Now, "YYYY/MM/DD") Then
            MsgBox "�w�p���:" & Trim(rsMainT16("�w�p���")) & "�p�󤵤�A�нT�{�w�p����O�_���~�A�q��פJ����", vbCritical + vbOKOnly
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16.MoveNext
    Loop
    
'�ˬd���f�ܤ����� 154,160,161,162,163,165
    rsMainT16.MoveFirst
    Do While Not rsMainT16.EOF
        If Trim(rsMainT16("���f��")) <> "154" And Trim(rsMainT16("���f��")) <> "160" And Trim(rsMainT16("���f��")) <> "161" And Trim(rsMainT16("���f��")) <> "162" And Trim(rsMainT16("���f��")) <> "163" And Trim(rsMainT16("���f��")) <> "165" Then
            MsgBox "���f��:" & Trim(rsMainT16("���f��")) & "�A��������_�ϥέܧO�A�нT�{�榡���L���D!", vbCritical + vbOKOnly
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16.MoveNext
    Loop
    rsMainT16.MoveFirst

'�ˬd���f��162�A��O������CO�MSC��
    Do While Not rsMainT16.EOF
        If (Trim(rsMainT16("���f��")) <> "162" And Trim(rsMainT16("�q�����")) = "SC") Or (Trim(rsMainT16("���f��")) <> "162" And Trim(rsMainT16("�q�����")) = "CO") Then
            MsgBox "�h�f�q�����:" & Trim(rsMainT16("�q�����")) & "�����f��:" & Trim(rsMainT16("���f��")) & "������162�A�нT�{��Ʀ��L���D!", vbCritical + vbOKOnly
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16.MoveNext
    Loop
    rsMainT16.MoveFirst
'�ˬd�~���O�_�s�b
    rsMainT16_1.MoveFirst
    Do While Not rsMainT16_1.EOF
        '�ˬd�O�_�����~��
        str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where storerkey = '" & str_Storerkey & "' and sku = '" & Trim(rsMainT16_1("���~�s��")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        If tmp_Rs.EOF Then
            '�L���~��
            MsgBox "�t�Χ䤣��~��:" & Trim(rsMainT16_1("���~�s��")) & "����ơA�Х��إ߰ӫ~�D�ɸ�ơA�פJ����", vbCritical + vbOKOnly, "�~���ˬd"
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16_1.MoveNext
    Loop
    
rsMainT16.MoveFirst: rsMainT16_1.MoveFirst

Tran_Level = cn.BeginTrans: Cmd_Impdata.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strInvoiceDate As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String

'���̫�Ȥ�s��
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & str_Storerkey & "' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close
'�}�l�פJ
Do While Not rsMainT16.EOF
    DoEvents: DoEvents
        '�s�W�q����
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        Str_Address = RTrim(rsMainT16.Fields("�e�f�a�}"))
        
        '�ˬd�O�_�����Ȥ�s��
        str_SQL = "select top 1 consigneekey from trp01m where storerkey = '" & str_Storerkey & "' and rtrim(consigneekey) = '" & RTrim(rsMainT16.Fields("�Ȥ�s��")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        If tmp_Rs.EOF Then
            '�L���Ȥ�s���h�s�W
            '�s�W�Ȥ�D��
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
            " values('" & str_Storerkey & "','','" & Trim(rsMainT16("�Ȥ�s��")) & "','" & myExCharFilter(Trim(rsMainT16("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT16("�Ȥ�²��"))) & "','" & myExCharFilter(Trim(rsMainT16("�s���H"))) & "','" & myExCharFilter(Trim(rsMainT16("�q��"))) & "','" & Str_Address & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
            '�����s�W���Ȥ�s��
            strNewConsigneekey = strNewConsigneekey & Trim(rsMainT16("�Ȥ�s��")) & "','"
            strConsigneeKey = Trim(rsMainT16("�Ȥ�s��"))
        Else
            '���Ȥ�W�١A²�ٻP��f�a�}�O�_�۲�
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select consigneekey from trp01m(nolock) " & _
                        "where storerkey = '" & str_Storerkey & "' and full_name = '" & myExCharFilter(Trim(rsMainT16("�Ȥ�W��"))) & "' and short_name = '" & myExCharFilter(Trim(rsMainT16("�Ȥ�²��"))) & "' " & _
                        "and rtrim(address) = '" & Str_Address & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn

                If rsTmp.EOF Then
                    '�p���H�B�q�ܻP��f�a�}����
                    intTmp = intTmp + 1
                    strConsigneeKey = "BEST" & Format(intTmp, "000000") '�ݽT�{BEST

                    '�s�W�Ȥ�D��
                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
                    " values('" & str_Storerkey & "','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT16("�Ȥ�²��"))) & "','" & myExCharFilter(Trim(rsMainT16("�s���H"))) & "','" & myExCharFilter(Trim(rsMainT16("�q��"))) & "','" & Str_Address & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
            
                    '�����s�W���Ȥ�s��
                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
                Else '�۲Ūu���«Ƚs
                    strConsigneeKey = Trim(rsTmp("consigneekey"))
                    blCustomerMatch = True

                End If
            rsTmp.Close
        End If
        tmp_Rs.Close

        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select ExternOrderKey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16("�q�����"))) & myExCharFilter(Trim(rsMainT16("�q�渹�X"))) & "' and externordertype = '" & myExCharFilter(Trim(rsMainT16("�q�����"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            '�����Ҧ�orderkey,�̫�@���妸��spackkey
            Str_AllOrderkey = Str_AllOrderkey & "'" & str_Orderkey & "',"
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"

            If Len(Trim(rsMainT16("�w�p���"))) = 0 Then    '�p�G�S�����w���,�h�a�j��@��
                strDate = Format(Now + 1, "YYYY/MM/DD")
            Else
                strDate = Val(Left(Trim(rsMainT16("�w�p���")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("�w�p���")), 4, 2) & "/" & Right(Trim(rsMainT16("�w�p���")), 2)
            End If
            
            strOrderDate = Val(Left(Trim(rsMainT16("�q����")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("�q����")), 4, 2) & "/" & Right(Trim(rsMainT16("�q����")), 2)
            

            
            Dim intPointer As Integer
            intPointer = 1
            
            '�q�����O�ഫ
            If myExCharFilter(Trim(rsMainT16("���f��"))) = "154" Then
                str_Priority = "A2B"
            Else
                If myExCharFilter(Trim(rsMainT16("�q�����"))) = "CO" Or myExCharFilter(Trim(rsMainT16("�q�����"))) = "SC" Then
                    str_Priority = "R"
                Else
                    '�ݥδ��f���d��P�_�O�_��I�άOA2B
                    str_Priority = "I"
                End If
            End If
            '�g�J���_�M��Table�A�����q��Ҧ���T
            If Len(Trim(rsMainT16("�o�����"))) = 0 Then
                 str_SQL = "insert CustOrders(Storerkey,orderkey,BranchId,ExternOrderkey,OrderDate,Invoice,InvoiceCheck,InvoiceDate,Consigneekey, " & _
                        "Full_Name,SalesCode,COD,Address,Coupled,VAT,Allowance,QuantityAllowance,SpecialAllowance, " & _
                        "CashAllowance,Amount,NetAmount,Tax,Notes,CustOrderkey,InvoiceCode,OrderCode, " & _
                        "LogisticsCode , DeliveryCode, OrderType, PaidMARK, Contact, Phone1, SalesName, LeaderName, " & _
                        "Address2,DeliveryDate,Freight,Payment,SalesPhone,EInvoiceMark,TotalWeight,Credit_Last4,Cash, " & _
                        "InvoicePrint , Phone2, ExternNumber, City, Administration, Stairs, CrossCode, Storage, InvoiceArea,Short_name,keyinuser,ConnectOrderkey,addwho,updatesource) " & _
                        "values ('LMBO01','" & str_Orderkey & "','" & Trim(rsMainT16("�����q�N��")) & "','" & Trim(rsMainT16("�q�渹�X")) & "','" & strOrderDate & "','" & _
                        Trim(rsMainT16("�o�����X")) & "','" & Trim(rsMainT16("�o�����X�ˬd�X")) & "',null,'" & _
                        Trim(rsMainT16("�Ȥ�s��")) & "','" & Trim(rsMainT16("�Ȥ�W��")) & "','" & Trim(rsMainT16("�~�N�N��")) & "','" & _
                        Trim(rsMainT16("�U�f���{")) & "','" & Trim(rsMainT16("�e�f�a�}")) & "','" & Trim(rsMainT16("�p��")) & "','" & _
                        Trim(rsMainT16("�Τ@�s��")) & "','" & Trim(rsMainT16("�������B")) & "','" & Trim(rsMainT16("�ƶq�������B")) & "','" & _
                        Trim(rsMainT16("�S�O�������B")) & "','" & Trim(rsMainT16("�{������")) & "','" & Trim(rsMainT16("�f��")) & "','" & _
                        Trim(rsMainT16("�|�e���B")) & "','" & Trim(rsMainT16("�|�B")) & "','" & Trim(rsMainT16("�Ƶ�")) & "','" & _
                        Trim(rsMainT16("�Ȥ�q��s��")) & "','" & Trim(rsMainT16("�H�f���o���X")) & "','" & Trim(rsMainT16("�H�f���q��X")) & "','" & _
                        Trim(rsMainT16("�p�⪫�y�O")) & "','" & Trim(rsMainT16("�e�f�_")) & "','" & Trim(rsMainT16("�q�����")) & "','" & _
                        Trim(rsMainT16("�ꦬ�q�B�zMARK")) & "','" & Trim(rsMainT16("�s���H")) & "','" & Trim(rsMainT16("�q��")) & "','" & _
                        Trim(rsMainT16("�~�N�m�W")) & "','" & Trim(rsMainT16("�D�ީm�W")) & "','" & Trim(rsMainT16("���e�Ȥ�")) & "','" & _
                        strDate & "','" & Trim(rsMainT16("�B�O")) & "','" & Trim(rsMainT16("�I�ڤ覡")) & "','" & Trim(rsMainT16("�~�Ȥ��")) & "','" & _
                        Trim(rsMainT16("�O�_���q�l�o��")) & "','" & Trim(rsMainT16("�`���q")) & "','" & Trim(rsMainT16("�H�d��4�X")) & "','" & _
                        Trim(rsMainT16("�N���f��")) & "','" & Trim(rsMainT16("�o���C�L�覡")) & "','" & Trim(rsMainT16("�q��2")) & "','" & Trim(rsMainT16("�έp��H")) & "','" & _
                        Trim(rsMainT16("�����O")) & "','" & Trim(rsMainT16("��F��")) & "','" & Trim(rsMainT16("�Ӽh")) & "','" & Trim(rsMainT16("�V�w�q��")) & "','" & Trim(rsMainT16("���f��")) & "','" & Trim(rsMainT16("�|��/�|�v")) & "','" & Trim(rsMainT16("�Ȥ�²��")) & "','" & Trim(rsMainT16("�q�浡�f")) & "','" & Trim(rsMainT16("���p�q�渹�X")) & "','" & User_id & "','" & Str_updatesource1 & "')"

            Else
                strInvoiceDate = Val(Left(Trim(rsMainT16("�o�����")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("�o�����")), 4, 2) & "/" & Right(Trim(rsMainT16("�o�����")), 2)
                str_SQL = "insert CustOrders(Storerkey,orderkey,BranchId,ExternOrderkey,OrderDate,Invoice,InvoiceCheck,InvoiceDate,Consigneekey, " & _
                        "Full_Name,SalesCode,COD,Address,Coupled,VAT,Allowance,QuantityAllowance,SpecialAllowance, " & _
                        "CashAllowance,Amount,NetAmount,Tax,Notes,CustOrderkey,InvoiceCode,OrderCode, " & _
                        "LogisticsCode , DeliveryCode, OrderType, PaidMARK, Contact, Phone1, SalesName, LeaderName, " & _
                        "Address2,DeliveryDate,Freight,Payment,SalesPhone,EInvoiceMark,TotalWeight,Credit_Last4,Cash, " & _
                        "InvoicePrint , Phone2, ExternNumber, City, Administration, Stairs, CrossCode, Storage, InvoiceArea,Short_name,keyinuser,ConnectOrderkey,addwho,updatesource) " & _
                        "values ('LMBO01','" & str_Orderkey & "','" & Trim(rsMainT16("�����q�N��")) & "','" & Trim(rsMainT16("�q�渹�X")) & "','" & strOrderDate & "','" & _
                        Trim(rsMainT16("�o�����X")) & "','" & Trim(rsMainT16("�o�����X�ˬd�X")) & "','" & strInvoiceDate & "','" & _
                        Trim(rsMainT16("�Ȥ�s��")) & "','" & Trim(rsMainT16("�Ȥ�W��")) & "','" & Trim(rsMainT16("�~�N�N��")) & "','" & _
                        Trim(rsMainT16("�U�f���{")) & "','" & Trim(rsMainT16("�e�f�a�}")) & "','" & Trim(rsMainT16("�p��")) & "','" & _
                        Trim(rsMainT16("�Τ@�s��")) & "','" & Trim(rsMainT16("�������B")) & "','" & Trim(rsMainT16("�ƶq�������B")) & "','" & _
                        Trim(rsMainT16("�S�O�������B")) & "','" & Trim(rsMainT16("�{������")) & "','" & Trim(rsMainT16("�f��")) & "','" & _
                        Trim(rsMainT16("�|�e���B")) & "','" & Trim(rsMainT16("�|�B")) & "','" & Trim(rsMainT16("�Ƶ�")) & "','" & _
                        Trim(rsMainT16("�Ȥ�q��s��")) & "','" & Trim(rsMainT16("�H�f���o���X")) & "','" & Trim(rsMainT16("�H�f���q��X")) & "','" & _
                        Trim(rsMainT16("�p�⪫�y�O")) & "','" & Trim(rsMainT16("�e�f�_")) & "','" & Trim(rsMainT16("�q�����")) & "','" & _
                        Trim(rsMainT16("�ꦬ�q�B�zMARK")) & "','" & Trim(rsMainT16("�s���H")) & "','" & Trim(rsMainT16("�q��")) & "','" & _
                        Trim(rsMainT16("�~�N�m�W")) & "','" & Trim(rsMainT16("�D�ީm�W")) & "','" & Trim(rsMainT16("���e�Ȥ�")) & "','" & _
                        strDate & "','" & Trim(rsMainT16("�B�O")) & "','" & Trim(rsMainT16("�I�ڤ覡")) & "','" & Trim(rsMainT16("�~�Ȥ��")) & "','" & _
                        Trim(rsMainT16("�O�_���q�l�o��")) & "','" & Trim(rsMainT16("�`���q")) & "','" & Trim(rsMainT16("�H�d��4�X")) & "','" & _
                        Trim(rsMainT16("�N���f��")) & "','" & Trim(rsMainT16("�o���C�L�覡")) & "','" & Trim(rsMainT16("�q��2")) & "','" & Trim(rsMainT16("�έp��H")) & "','" & _
                        Trim(rsMainT16("�����O")) & "','" & Trim(rsMainT16("��F��")) & "','" & Trim(rsMainT16("�Ӽh")) & "','" & Trim(rsMainT16("�V�w�q��")) & "','" & Trim(rsMainT16("���f��")) & "','" & Trim(rsMainT16("�|��/�|�v")) & "','" & Trim(rsMainT16("�Ȥ�²��")) & "','" & Trim(rsMainT16("�q�浡�f")) & "','" & Trim(rsMainT16("���p�q�渹�X")) & "','" & User_id & "','" & Str_updatesource1 & "')"
    
            End If
           
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            '�g�JOrders
            If str_Priority = "A2B" Then
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,b_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externordertype,cash) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT16("�q�����")) & myExCharFilter(Trim(rsMainT16("�q�渹�X"))) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strDate & "','" & strFacility & "','LMBO01-154" & _
                "','" & myExCharFilter(Trim(rsMainT16("�Ȥ�W��"))) & "','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16("�s���H"))) & "','','','" & myExCharFilter(Trim(rsMainT16("�q��"))) & "','','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16("�Ƶ�"))) & "','" & Str_updatesource1 & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT16("�o�����X"))) & "','" & myExCharFilter(Trim(rsMainT16("�q�����"))) & "','" & myExCharFilter(Trim(rsMainT16("�N���f��"))) & "') "
            Else
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externordertype,cash) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT16("�q�����")) & myExCharFilter(Trim(rsMainT16("�q�渹�X"))) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
                strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT16("�s���H"))) & "','','','" & myExCharFilter(Trim(rsMainT16("�q��"))) & "','','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16("�Ƶ�"))) & "','" & Str_updatesource1 & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT16("�o�����X"))) & "','" & myExCharFilter(Trim(rsMainT16("�q�����"))) & "','" & myExCharFilter(Trim(rsMainT16("�N���f��"))) & "') "
            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1


            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            'If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT16("�q�����")) & myExCharFilter(Trim(rsMainT16("�q�渹�X"))) & "','" 'Trim(rsMainT16("�q�����"))
            blDuplicationOrder = True

        End If
        
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
            rsMainT16_1.Filter = "�q�渹�X = '" & rsMainT16.Fields("�q�渹�X") & "' and �q����� = '" & rsMainT16.Fields("�q�����") & "'"
            rsMainT16_1.Sort = "�渹"
            rsMainT16_1.MoveFirst
            Do While Not rsMainT16_1.EOF
                '�W�[����
                int_orderlinenuber = int_orderlinenuber + 1
                lngCasecnt = 1

                '�Ĵ����� lot05 = ��f��+(������X���ĤѼ�)
                    '���c�]�ഫ�v
'                    str_SQL = "select susr2=isnull(susr2,0) from " & strWMSDB & "..sku where sku = '" & myExCharFilter(Trim(rsMainT16_1("���~�s��"))) & "'"
'
'                    Call Confirm_Recordset_Closed(tmp_Rs)
'                    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'                    Str_Lot05 = Format(strDate + (Val(rsMainT16_1.Fields("������")) * Val(tmp_Rs.Fields("susr2"))), "YYYYMMDD")
'                    tmp_Rs.Close
                    
                intQTY = Abs(Val(rsMainT16_1("�q�f�q")))
                strLot06 = RTrim(rsMainT16.Fields("���f��"))
                
                '�������_�M�έq�����
                str_SQL = "insert CustOrderdetail(Storerkey,orderkey,orderlinenumber,ExternOrderkey,Sku,Descr,OriginalQty,UnitNetPrice,NetPrice,UnitGrossPrice,GrossPrice,RefusalQty,BarCode,Externlineno,UOM,Ordertype,InvoicePCode,Acceptance,addwho) " & _
                "values('LMBO01','" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT16_1.Fields("�q�渹�X")) & "','" & Trim(rsMainT16_1.Fields("���~�s��")) & "','" & Trim(rsMainT16_1.Fields("���~�W��")) & _
                "','" & Trim(rsMainT16_1.Fields("�q�f�q")) & "','" & Trim(rsMainT16_1.Fields("���(���|)")) & "','" & Trim(rsMainT16_1.Fields("�q�f���B(���|)")) & _
                "','" & Trim(rsMainT16_1.Fields("���(�t�|)")) & "','" & Trim(rsMainT16_1.Fields("�q�f���B(�t�|)")) & "','" & Trim(rsMainT16_1.Fields("�q�f�q-�ꦬ�q")) & _
                "','" & Trim(rsMainT16_1.Fields("��ڱ��X")) & "','" & Trim(rsMainT16_1.Fields("�渹")) & "','" & Trim(rsMainT16_1.Fields("���")) & _
                "','" & Trim(rsMainT16_1.Fields("�q�����")) & "','" & Trim(rsMainT16_1.Fields("�o�����ӦC�L�_")) & "','" & Trim(rsMainT16_1.Fields("������")) & "','" & User_id & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '�q����Ӹ�Ʒs�W
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice)" & _
                "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_1("�渹"))) & "','" & Trim(rsMainT16("�q�����")) & myExCharFilter(Trim(rsMainT16_1("�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT16_1("���~�s��"))) & "','" & str_Storerkey & "'," & _
                "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_1("���"))) & "','0')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
                int_OrderLine = int_OrderLine + 1
                rsMainT16_1.MoveNext
            Loop
        End If

        rsMainT16.MoveNext
        rsMainT16_1.MoveFirst
Loop

'�妸��spackkey
If Str_AllOrderkey <> "" Then '�[�JStr_AllOrderkey <> "" �P�_ by Gemini @20160704
    str_SQL = "update orderdetail " & _
    "Set orderdetail.packkey = sku.packkey " & _
    "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
    "where orderkey in (" & Mid(Str_AllOrderkey, 1, Len(Str_AllOrderkey) - 1) & ") "

    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

End If

cn.Execute "exec gs_ordersupdate 'LMBO01'", RowsAffect, adExecuteNoRecords
cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ " & int_OrderLine & " ������" & Chr(13) & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & Chr(13)
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�A�ɮ� " & Str_updatesource1)

'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & Str_updatesource1 & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "'"

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption

    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing

End If

copy:

'�ƥ��ɮר쥻��
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource1) = "" Then
    FileCopy strFileHeader, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource1
    FileCopy strFileDetail, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource2
    FileCopy strFileHeader, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource1
    FileCopy strFileDetail, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource2
Else
    FileCopy strFileHeader, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource1, ".", 0) & "." & mySplit(Str_updatesource1, ".", -1)
    FileCopy strFileDetail, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource2, ".", 0) & "." & mySplit(Str_updatesource2, ".", -1)
    FileCopy strFileHeader, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource1, ".", 0) & "." & mySplit(Str_updatesource1, ".", -1)
    FileCopy strFileDetail, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource2, ".", 0) & "." & mySplit(Str_updatesource2, ".", -1)
End If

Kill strFileHeader
Kill strFileDetail

''�_�}�ثe�s�u�s�u
'LCDisConnect "192.168.200.200\IPC$"
'LCDisConnect "bestprepares\IPC$"
'
''���s�s�u���@��Ƨ�
'LCConnect "bestprepares", "share", "share"
''LCConnect "192.168.200.200", "share", "share"

filLocalFileT16.Refresh:
Screen.MousePointer = 0: Cmd_Impdata.Enabled = True: SSTab2.Enabled = True
Exit Sub




'�H�U�����ɤs�q��פJ���{���X
'Dim Int_RC As Integer: Dim Int_C As Integer: Dim Int_I As Integer: Dim Int_otqty As Integer
'Int_RC = 0: Int_C = 0: Int_I = 0 '�p��q�����O������
'Int_otqty = 0 '�p��q����
'
'If rsMainT16 Is Nothing Then Exit Sub
'If rsMainT16.EOF Then Exit Sub
'
'On Error GoTo err_Handle
'SSTab2.Enabled = False: Cmd_Impdata.Enabled = False
'strTranFileName = filLocalFileT16.Path & "\" & filLocalFileT16.FileName
'
''�������--�P�_�ɮ׬O�_�w��J
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT16.FileName & "' "
'
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF = False Then SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
'tmp_Rs.Close
'
'Dim arrTmp
'
'If rsMainT16_1.RecordCount = 0 Or rsMainT16_1 Is Nothing Then
'Else
'rsMainT16_1.MoveFirst
'Do While Not rsMainT16_1.EOF
'    '��f����ˬd
'    If Len(Trim(rsMainT16_1("���w���"))) = 0 Then
'    Else
'        arrTmp = Split(Trim(rsMainT16_1("���w���")), "/")
'        If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "���w����p�󤵤�A�q����J�פ�!", 16, Me.Caption: SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'    End If
'
'    '�ƶq�ˬd
'    If Val(rsMainT16_1("�P�f�ƶq")) + Val(rsMainT16_1("��/�ƫ~�q")) < 1 Then
'        MsgBox "�q��ƶq�p��1�A" & Trim(rsMainT16_1("�P�f�渹")) & "-�~���G" & Trim(rsMainT16_1("�~��")) & "(" & Trim(rsMainT16_1("�~�W")) & ")�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'        Exit Sub
'    End If
'
''    ������� --�P�_SKU�O�_�s�b
'    If Trim(rsMainT16_1("�~��")) <> "60400119" Then
'        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT16_1("�~��")) & "' and Storerkey = 'LMYS01' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then
'            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT16_1("�~��")) & " ) " & Trim(rsMainT16_1("�~�W")) & "�A�q����J�פ�!!": Cmd_Impdata.Enabled = True: Screen.MousePointer = 0
'            SSTab2.Enabled = True: Cmd_Impdata.Enabled = True
'            Exit Sub
'        End If
'
'    End If
'
'    rsMainT16_1.MoveNext
'Loop
'rsMainT16_1.MoveFirst
'End If
'
'If rsMainT16_2.RecordCount = 0 Or rsMainT16_2 Is Nothing Then
'Else
'rsMainT16_2.MoveFirst
'Do While Not rsMainT16_2.EOF
'    '��f����ˬd
'    If Len(Trim(rsMainT16_2("���w���"))) = 0 Then
'    Else
'        arrTmp = Split(Trim(rsMainT16_2("���w���")), "/")
'        If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "���w����p�󤵤�A�q����J�פ�!", 16, Me.Caption: SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'    End If
'
'    '�ƶq�ˬd
'    If Val(rsMainT16_2("�༷�ƶq")) < 1 Then
'        MsgBox "�q��ƶq�p��1�A" & Trim(rsMainT16_2("��O-�渹")) & "-�~���G" & Trim(rsMainT16_2("�~��")) & "(" & Trim(rsMainT16_2("�~�W")) & ")�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'        Exit Sub
'    End If
'
''    ������� --�P�_SKU�O�_�s�b
'    If Trim(rsMainT16_2("�~��")) <> "60400119" Then
'        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT16_2("�~��")) & "' and Storerkey = 'LMYS01' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then
'            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT16_2("�~��")) & " ) " & Trim(rsMainT16_2("�~�W")) & "�A�q����J�פ�!!": Cmd_Impdata.Enabled = True: Screen.MousePointer = 0
'            SSTab2.Enabled = True: Cmd_Impdata.Enabled = True
'            Exit Sub
'        End If
'    End If
'
'    rsMainT16_2.MoveNext
'Loop
'rsMainT16_2.MoveFirst
'End If
'
'If rsMainT16_3.RecordCount = 0 Or rsMainT16_3 Is Nothing Then
'Else
'rsMainT16_3.MoveFirst
'Do While Not rsMainT16_3.EOF
'    '��f����ˬd
'    If Len(Trim(rsMainT16_3("���w���"))) = 0 Then
'    Else
'        arrTmp = Split(Trim(rsMainT16_3("���w���")), "/")
'        If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "���w����p�󤵤�A�q����J�פ�!", 16, Me.Caption: SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'    End If
'
'    '�ƶq�ˬd
'    If Val(rsMainT16_3("�P�f�ƶq")) + Val(rsMainT16_3("��/�ƫ~�q")) < 1 Then
'        MsgBox "�q��ƶq�p��1�A" & Trim(rsMainT16_3("�P�f�渹")) & "-�~���G" & Trim(rsMainT16_3("�~��")) & "(" & Trim(rsMainT16_3("�~�W")) & ")�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'        Exit Sub
'    End If
'
''    ������� --�P�_SKU�O�_�s�b
'    If Trim(rsMainT16_3("�~��")) <> "60400119" Then
'        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT16_3("�~��")) & "' and Storerkey = 'LMYS01' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then
'            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT16_3("�~��")) & " ) " & Trim(rsMainT16_3("�~�W")) & "�A�q����J�פ�!!": Cmd_Impdata.Enabled = True: Screen.MousePointer = 0
'            SSTab2.Enabled = True: Cmd_Impdata.Enabled = True
'            Exit Sub
'        End If
'    End If
'
'    rsMainT16_3.MoveNext
'Loop
'rsMainT16_3.MoveFirst
'End If
'
'Tran_Level = cn.BeginTrans: Cmd_Impdata.Enabled = False
'
'Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long
'Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
'Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
'Dim strDate As String, strOrderType As String, strSku As String
'
''�}�l�פJ �P�f���Ӫ�
'If rsMainT16_1 Is Nothing Then GoTo next18
'If rsMainT16_1.RecordCount = 0 Then GoTo next18
'
''���̫�Ȥ�s��
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LMYS01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close
'
'Do While Not rsMainT16_1.EOF
'    DoEvents: DoEvents
'
''    �������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If Trim(rsMainT16_1("�~��")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow17
'
'    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
'    If strOrderNo <> UCase(Trim(rsMainT16_1("�P�f�渹"))) Then
'        strOrderNo = UCase(Trim(rsMainT16_1("�P�f�渹")))
'        int_orderlinenuber = 0
'        blDuplicationOrder = False
'
'        '�ˬd�O�_�����Ȥ�s��
'        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LMYS01' and consigneekey = '" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�N��"))) & "' order by len(consigneekey),consigneekey "
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '�L���Ȥ�s���h�s�W
'            intTmp = intTmp + 1
'            strConsigneeKey = "BEST" & Format(intTmp, "000000")
'            'strConsigneeKey = myExCharFilter(Trim(rsMainT16_1("�Ȥ�N��")))
'
'            '�s�W�Ȥ�D��
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�²��"))) & "','" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�²��"))) & "','','','" & myExCharFilter(Trim(rsMainT16_1("�e�f�a�}"))) & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'            '�����s�W���Ȥ�s��
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
'            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LMYS01' and full_name = '" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�²��"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT16_1("�e�f�a�}"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'                If rsTmp.EOF Then
'                    '�p���H�B�q�ܻP��f�a�}����
'                    intTmp = intTmp + 1
'                    strConsigneeKey = "BEST" & Format(intTmp, "000000") '�ݽT�{BEST
'
'                    '�s�W�Ȥ�D��
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                    " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�²��"))) & "','" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�²��"))) & "','','','" & myExCharFilter(Trim(rsMainT16_1("�e�f�a�}"))) & "','','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'
'                    '�����s�W���Ȥ�s��
'                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'                Else '�۲Ūu���«Ƚs
'                    strConsigneeKey = Trim(rsTmp("consigneekey"))
'                    blCustomerMatch = True
'
'                End If
'            rsTmp.Close
'        End If
'        tmp_Rs.Close
'
'        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16_1("�P�f�渹"))) & "' and storerkey = 'LMYS01' and isnull(type,'') <> '�R��' "
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_Rs.EOF Then
'
'            '���q�渹�X
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'            tmp_Rs.Close
'
'            '�t�e�ܧO�P�_
''           strFacility = Trim(rsMainT16_1("�ܮw"))          �ܮw�ݽT�{
'            strFacility = "�ըƹF�_��"
'
'            arrTmp = Split(Trim(rsMainT16_1("���w���")), "/")
'            If Len(Trim(rsMainT16_1("���w���"))) = 0 Then    '�p�G�S�����w���,�h�a�j��@��
'                strDate = Format(Now + 1, "YYYY/MM/DD")
'            Else
'                strDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            End If
'
'            arrTmp = Split(Trim(rsMainT16_1("�P�f���")), "/")
'            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            Dim intPointer As Integer
'            intPointer = 1
'            Int_otqty = Int_otqty + Val(Trim(rsMainT16_1("���")))
'            'updatesource
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,otqty,externconsigneekey) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT16_1("�P�f�渹"))) & "','C','LMYS01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
'            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�²��"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT16_1("�e�f�a�}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT16_1("�e�f�a�}"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16_1("�Ƶ�"))) & "','" & filLocalFileT16.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT16_1("���"))) & "','" & myExCharFilter(Trim(rsMainT16_1("�Ȥ�N��"))) & "') "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_Order = int_Order + 1
'            Int_C = Int_C + 1
'
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LMYS01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
'        Else
'            '�q�歫��
'            Call FTPlog("�q�歫��" & str_SQL)
'            '��������
'            strReOrderkey = strReOrderkey & Trim(rsMainT16_1("�P�f�渹")) & "','"
'            blDuplicationOrder = True
'
'        End If
'    End If
'
'        '�q�歫���ˬd
'        If blDuplicationOrder = False Then
'
'            '�W�[����
'            int_orderlinenuber = int_orderlinenuber + 1
'
'            lngCasecnt = 1
'
'            '��촫��
'            If Left(myExCharFilter(Trim(rsMainT16_1("���"))), 1) = "�c" Then
'
'                '���c�]�ഫ�v
'                str_SQL = "select casecnt = case when p.casecnt = 0 then 1 else p.casecnt end from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey " & _
'                "where s.storerkey = 'LMYS01' and s.sku = '" & myExCharFilter(Trim(rsMainT16_1("�~��"))) & "' "
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'                lngCasecnt = tmp_Rs("casecnt")
'                tmp_Rs.Close
'
'            End If
'
'            intQTY = (Val(rsMainT16_1("�P�f�ƶq")) + Val(rsMainT16_1("��/�ƫ~�q"))) * lngCasecnt
'            strLot06 = "R01" '�w�]R01 , ����R01-C
'
'            '�q����Ӹ�Ʒs�W
'            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
'            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_1("�P�f�渹"))) & "','" & myExCharFilter(Trim(rsMainT16_1("�~��"))) & "','LMYS01'," & _
'            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_1("���"))) & "','0','" & myExCharFilter(Trim(rsMainT16_1("�Ƶ�"))) & "')"
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            '��spackkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            int_OrderLine = int_OrderLine + 1
'        End If
'
'nextRow17:
'        rsMainT16_1.MoveNext
'Loop
'
'next18:
'
'If rsMainT16_2 Is Nothing Then GoTo next19
'If rsMainT16_2.RecordCount = 0 Then GoTo next19
'
''���̫�Ȥ�s��
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LMYS01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close
'
''�}�l�פJ �༷���Ӫ�
'strOrderNo = ""
'Do While Not rsMainT16_2.EOF
'
'    '�������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If Trim(rsMainT16_2("�~��")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow18
'
'    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
'    If strOrderNo <> UCase(Trim(rsMainT16_2("��O-�渹"))) Then
'        strOrderNo = UCase(Trim(rsMainT16_2("��O-�渹")))
'        int_orderlinenuber = 0
'        strLot06 = ""
'        blDuplicationOrder = False
'
'        '�ˬd�O�_�����Ȥ�W��
'        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LMYS01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT16_2("��J�w�O"))) & "' order by len(consigneekey),consigneekey "
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '�L���Ȥ�W�٫h�s�W
'            intTmp = intTmp + 1
'            strConsigneeKey = "BEST" & Format(intTmp, "000000") '�ݽT�{
'
'            '�s�W�Ȥ�D�� �ݽT�{,updatesource�n�a�ƻ�
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_2("��J�w�O"))) & "','" & myExCharFilter(Trim(rsMainT16_2("��J�w�O"))) & "','','','" & myExCharFilter(Trim(rsMainT16_2("�e�f�a�}"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'
'            '�����s�W���Ȥ�s��
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
'            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LMYS01' and full_name = '" & myExCharFilter(Trim(rsMainT16_2("��J�w�O"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT16_2("�e�f�a�}"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'            If rsTmp.EOF Then
'                '�p���H�B�q�ܻP��f�a�}����
'                intTmp = intTmp + 1
'                strConsigneeKey = "BEST" & Format(intTmp, "000000") '�ݽT�{
'
'                '�s�W�Ȥ�D��
'                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_2("��J�w�O"))) & "','" & myExCharFilter(Trim(rsMainT16_2("��J�w�O"))) & "','','','" & myExCharFilter(Trim(rsMainT16_2("�e�f�a�}"))) & "','" & myExCharFilter(Trim(rsMainT16_2("�Ƶ�"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'
'                '�����s�W���Ȥ�s��
'                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'            Else '�۲Ūu���«Ƚs
'                strConsigneeKey = Trim(rsTmp("consigneekey"))
'                blCustomerMatch = True
'
'            End If
'            rsTmp.Close
'        End If
'        tmp_Rs.Close
'
'        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16_2("��O-�渹"))) & "' and storerkey = 'LMYS01' and isnull(type,'') <> '�R��' "
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_Rs.EOF Then
'
'            '���q�渹�X
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'            tmp_Rs.Close
'
'            '�t�e�ܧO�P�_
'            'strFacility = Trim(rsMainT16_2("��J�w�O"))
'            strFacility = "�ըƹF�_��"
'
'            arrTmp = Split(Trim(rsMainT16_2("���w���")), "/")
'            If Len(Trim(rsMainT16_2("���w���"))) = 0 Then    '�p�G�S�����w���,�h�a�j��@��
'                strDate = Format(Now + 1, "YYYY/MM/DD")
'            Else
'                strDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            End If
'
'            arrTmp = Split(Trim(rsMainT16_2("��ڤ��")), "/")
'            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            intPointer = 1
'
'            strOrderType = "C"
'            If Trim(rsMainT16_2("��J�w�O")) = "���y��/�ըƹF�_��" Or Trim(rsMainT16_2("��J�w�O")) = "���y��/�ըƹF����" Then strOrderType = "RC"
'            If strOrderType = "C" Then
'                Int_C = Int_C + 1
'            Else
'                Int_RC = Int_RC + 1
'            End If
'            strLot06 = IIf(UCase(Trim(rsMainT16_2("��J�w�O"))) = "���y��/�ըƹF����", "R01-C", "R01")
'            strFacility = IIf(UCase(Trim(rsMainT16_2("��J�w�O"))) = "���y��/�ըƹF����", "�ըƹF����", "�ըƹF�_��")
'
'            If UCase(Trim(rsMainT16_2("��J�w�O"))) = "���y��/�ըƹF�_��" Or UCase(Trim(rsMainT16_2("��J�w�O"))) = "���y��/�ըƹF����" Then
'                'RC���q�椣�֥[���
'            Else
'                Int_otqty = Int_otqty + Val(Trim(rsMainT16_2("���")))
'            End If
'
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,otqty) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT16_2("��O-�渹"))) & "','" & strOrderType & "','LMYS01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
'            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_2("��J�w�O"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT16_2("�e�f�a�}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT16_2("�e�f�a�}"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16_2("�Ƶ�"))) & "','" & filLocalFileT16.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT16_2("���"))) & "')"
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_Order = int_Order + 1
'
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LMYS01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
'        Else
'            '�q�歫��
'            Call FTPlog("�q�歫��" & str_SQL)
'            '��������
'            strReOrderkey = strReOrderkey & Trim(rsMainT16_2("��O-�渹")) & "','"
'            blDuplicationOrder = True
'
'        End If
'    End If
'
'        '�q�歫���ˬd
'        If blDuplicationOrder = False Then
'            '�W�[����
'            int_orderlinenuber = int_orderlinenuber + 1
'
''            lngCasecnt = 1
''            '��촫��
''            If Left(myExCharFilter(Trim(rsMainT16_2("���"))), 1) = "�c" Then
''
''                '���c�]�ഫ�v
''                str_SQL = "select casecnt = case when p.casecnt = 0 then 1 else p.casecnt end from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey " & _
''                "where s.storerkey = 'LMYS01' and s.sku = '" & myExCharFilter(Trim(rsMainT16_2("�~��"))) & "' "
''
''                Call Confirm_Recordset_Closed(tmp_Rs)
''                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
''                lngCasecnt = tmp_Rs("casecnt")
''                tmp_Rs.Close
''            End If
'
'            intQTY = Val(rsMainT16_2("�༷�ƶq")) '* lngCasecnt
'
'            '�q����Ӹ�Ʒs�W
'            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
'            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_2("��O-�渹"))) & "','" & myExCharFilter(Trim(rsMainT16_2("�~��"))) & "','LMYS01'," & _
'            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_2("���"))) & "','0','" & myExCharFilter(Trim(rsMainT16_2("�Ƶ�"))) & "')"
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            '��spackkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            int_OrderLine = int_OrderLine + 1
'        End If
'
'nextRow18:
'        rsMainT16_2.MoveNext
'Loop
'
'next19:
'
''�}�l�פJ�H�w�P�f���Ӫ�
'If rsMainT16_3 Is Nothing Then GoTo nextend
'If rsMainT16_3.RecordCount = 0 Then GoTo nextend
'
''���̫�Ȥ�s��
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LMYS01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close
'
'Do While Not rsMainT16_3.EOF
'
'    '�������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If Trim(rsMainT16_3("�~��")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow19
'
'    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
'    If strOrderNo <> UCase(Trim(rsMainT16_3("�P�f�渹"))) Then
'        strOrderNo = UCase(Trim(rsMainT16_3("�P�f�渹")))
'        int_orderlinenuber = 0
'        blDuplicationOrder = False
'
'        '�ˬd�O�_�����Ȥ�W��
'        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LMYS01' and rtrim(consigneekey) = '" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�N��"))) & "' order by len(consigneekey),consigneekey "
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '�L���Ȥ�W�٫h�s�W
'            intTmp = intTmp + 1
'            'strConsigneeKey = "BEST" & Format(intTmp, "000000")
'            strConsigneeKey = myExCharFilter(Trim(rsMainT16_3("�Ȥ�N��")))
'
'            '�s�W�Ȥ�D��
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�²��"))) & "','" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�²��"))) & "','','','" & myExCharFilter(Trim(rsMainT16_3("�e�f�a�}"))) & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'            '�����s�W���Ȥ�s��
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
'            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LMYS01' and full_name = '" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�²��"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT16_3("�e�f�a�}"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'                If rsTmp.EOF Then
'                    '�p���H�B�P��f�a�}����
'                    intTmp = intTmp + 1
'                    strConsigneeKey = "BEST" & Format(intTmp, "000000") '�ݽT�{BEST
'
'                    '�s�W�Ȥ�D��
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                    " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�²��"))) & "','" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�²��"))) & "','','','" & myExCharFilter(Trim(rsMainT16_3("�e�f�a�}"))) & "','','" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�N��"))) & "','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'
'                    '�����s�W���Ȥ�s��
'                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'                Else '�۲Ūu���«Ƚs
'                    strConsigneeKey = Trim(rsTmp("consigneekey"))
'                    blCustomerMatch = True
'
'                End If
'            rsTmp.Close
'        End If
'        tmp_Rs.Close
'
'        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16_3("�P�f�渹"))) & "' and storerkey = 'LMYS01' and isnull(type,'') <> '�R��' "
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_Rs.EOF Then
'
'            '���q�渹�X
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'            tmp_Rs.Close
'
'            '�t�e�ܧO�P�_ 20130404�ק�
'            'strFacility = IIf(UCase(Trim(rsMainT16_3("�Ȥ�N��"))) = "1201011004", "�ըƹF����", "�ըƹF�_��")
'            strFacility = "�ըƹF�_��"
'            arrTmp = Split(Trim(rsMainT16_3("���w���")), "/")
'            If Len(Trim(rsMainT16_3("���w���"))) = 0 Then    '�p�G�S�����w���,�h�a�j��@��
'                strDate = Format(Now + 1, "YYYY/MM/DD")
'            Else
'                strDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            End If
'
'            arrTmp = Split(Trim(rsMainT16_3("�P�f���")), "/")
'            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            intPointer = 1
'            'Int_otqty = Int_otqty + Val(Trim(rsMainT16_3("���"))) I���q�椣�[�`���
'
'            'updatesource �n�afilLocalFileT11.FileName �٬O �Ȥ�N�� �ݽT�{
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,otqty,externconsigneekey) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT16_3("�P�f�渹"))) & "','I','LMYS01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
'            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�²��"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT16_3("�e�f�a�}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT16_3("�e�f�a�}"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16_3("�Ƶ�"))) & "','" & filLocalFileT16.FileName & "','','" & User_id & "','" & User_id & "','','" & Val(myExCharFilter(Trim(rsMainT16_3("���")))) & "','" & myExCharFilter(Trim(rsMainT16_3("�Ȥ�N��"))) & "') "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_Order = int_Order + 1
'            Int_I = Int_I + 1
'
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LMYS01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
'        Else
'            '�q�歫��
'            Call FTPlog("�q�歫��" & str_SQL)
'            '��������
'            strReOrderkey = strReOrderkey & Trim(rsMainT16_3("�P�f�渹")) & "','"
'            blDuplicationOrder = True
'
'        End If
'    End If
'
'        '�q�歫���ˬd
'        If blDuplicationOrder = False Then
'            '�W�[����
'            int_orderlinenuber = int_orderlinenuber + 1
'
''            lngCasecnt = 1
''
''            '��촫��
''            If Left(myExCharFilter(Trim(rsMainT16_3("���"))), 1) = "�c" Then
''
''                '���c�]�ഫ�v
''                str_SQL = "select casecnt = case when p.casecnt = 0 then 1 else p.casecnt end from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey " & _
''                "where s.storerkey = 'LMYS01' and s.sku = '" & myExCharFilter(Trim(rsMainT16_3("�~��"))) & "' "
''
''                Call Confirm_Recordset_Closed(tmp_Rs)
''                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
''                lngCasecnt = tmp_Rs("casecnt")
''                tmp_Rs.Close
''            End If
'
'            intQTY = (Val(rsMainT16_3("�P�f�ƶq")) + Val(rsMainT16_3("��/�ƫ~�q"))) '* lngCasecnt
'
'            strLot06 = "R01" '�w�]R01 , ����R01-C ,20130408�ק� �Ҧ���R01,�ըƹF�_��
'
''            If Trim(rsMainT16_3("�Ȥ�N��")) = "1201011004" Then strLot06 = "R01-C"
''            strLot06 = IIf(UCase(Trim(rsMainT16_3("�Ȥ�N��"))) = "1201011004", "R01-C", "R01")
'
'            '�q����Ӹ�Ʒs�W
'            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
'            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_3("�P�f�渹"))) & "','" & myExCharFilter(Trim(rsMainT16_3("�~��"))) & "','LMYS01'," & _
'            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_3("���"))) & "','0','" & myExCharFilter(Trim(rsMainT16_3("�Ƶ�"))) & "')"
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            '��spackkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            int_OrderLine = int_OrderLine + 1
'        End If
'
'nextRow19:
'        rsMainT16_3.MoveNext
'Loop
'
'nextend:
'
''�[�`C���q�� �s�W�@��A2B�q��
'        Dim ExternOrderKey As String
'        strDate = Format(Now, "YYYYMMDD") '�ثe�ɶ�
'        int_orderlinenuber = 0
'        int_orderlinenuber = int_orderlinenuber + 1
'        ExternOrderKey = "A2B" & strDate
'        '�ˬdexternorderkey�O�_�w�g���ۦP�� , �P�@�ѶפJ�⦸�H�W�q����
'
'        str_SQL = "select top 1 externorderkey from orders where storerkey = 'LMYS01' and externorderkey like '" & ExternOrderKey & "%' order by externorderkey desc"
'        Call Confirm_Recordset_Closed(rsTmp)
'        rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If rsTmp.EOF Then
'            GoTo A2B
'        End If
'
'        If ExternOrderKey = Left(Trim(rsTmp("externorderkey").Value), 11) And Mid(Trim(rsTmp("externorderkey").Value), 12, 1) = "-" Then
'                ExternOrderKey = ExternOrderKey & "-" & mySplit(Trim(rsTmp("externorderkey").Value), "-", 1) + 1
'        Else
'                ExternOrderKey = ExternOrderKey & "-1"
'        End If
'
'A2B:
'        '���q�渹�X
'        str_SQL = "select isnull(max(orderkey),0) from orders"
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'        If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'        tmp_Rs.Close: rsTmp.Close
'        '���Ȥ�D�ɸ��
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from trp01m where storerkey = 'LMYS01' and consigneekey = 'BEST000016' " 'BEST000016=���ɤs���� ;BEST000017 = ���y��\�ըƹF�_��
'        tmp_Rs.Open str_SQL, cn
'        '�g�Jorders
'        str_SQL = "INSERT orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,Deliverydate,Priority,ConsigneeKey,c_contact1,c_company,c_address1,c_zip,c_phone1,b_company,UpdateSource,type,door,route,stop,Notes,adddate,addwho,editdate,editwho,doroute,CustomerOrderkey,externconsigneekey,otqty) " & _
'                  "values('" & str_Orderkey & "','LMYS01','" & ExternOrderKey & "','" & strDate & "','" & strDate & "','A2B','" & Trim(tmp_Rs("ConsigneeKey").Value) & "','" & Trim(tmp_Rs("Contact").Value) & "','" & Trim(tmp_Rs("full_name").Value) & "','" & Trim(tmp_Rs("address").Value) & "'," & _
'                         "'" & Trim(tmp_Rs("zip").Value) & "','" & Trim(tmp_Rs("phone").Value) & "','BEST000017','" & filLocalFileT16.FileName & "','','99','99','99','�������ɤs���f���[���� �@�p" & Int_otqty & "��',getdate(),'" & User_id & "',getdate(),'" & User_id & "','Y','" & ExternOrderKey & "','" & Trim(tmp_Rs("consigneekey").Value) & "','" & Int_otqty & "')"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'        '�g�Jorderdetail
'        str_SQL = "insert into orderdetail (orderkey,orderlinenumber,externorderkey,sku,storerkey,originalqty,openqty,uom,packkey,status,adddate,addwho,editwho,lottable06) " & _
'                  "values ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "0000") & "','" & ExternOrderKey & "','OT','LMYS01','" & Int_otqty & "','" & Int_otqty & "','EA','OT','0',getdate(),'" & User_id & "','" & User_id & "','R01') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'        tmp_Rs.Close
''-----------------------
'
'cn.CommitTrans: Tran_Level = 0
'
''�T�����
'    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "��B�X�f(C): " & Int_C & " �� : ���f�J�w(RC): " & Int_RC & " �� : �@��X�f(I) : " & Int_I & " ��" & Chr(13) & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & Chr(13) & "�פJ �D�ըƹF�q�� " & intNotBest & " ������" & Chr(13) & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & Chr(13) & Chr(13) & "�t�β���1��A2B�q��Ω��ӡA�q�渹�X:" & str_Orderkey
'    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "��B�X�f(C): " & Int_C & " �� : ���f�J�w(RC): " & Int_RC & " �� : �@��X�f(I) : " & Int_I & " ��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT16.FileName)
'
''�q�歫�����
'If Len(strReOrderkey & strRePoOrderkey) > 0 Then
'
'    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT16.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
'        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = 'LMYS01'"
'
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
'
'    Call Recordset2Excel("�q�歫��", tmp_Rs)
'    If Dir("C:\LMYS01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LMYS01\�q�歫��"
'    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
'    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LMYS01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'    Set MyXlsApp = Nothing
'
'End If
'
''�ƥ���FTP
'If Dir("O:\LMYS01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LMYS01\OrdersBackup"
'FileCopy strTranFileName, "O:\LMYS01\OrdersBackup\" & filLocalFileT16.FileName
'
''�ƥ��ɮ�
'If Dir("C:\BEST\LMYS01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LMYS01\Orders\Backup"
'If Dir("C:\BEST\LMYS01\Orders\Backup\" & filLocalFileT16.FileName) = "" Then
'    FileCopy strTranFileName, "C:\BEST\LMYS01\Orders\Backup\" & filLocalFileT16.FileName
'Else
'    FileCopy strTranFileName, "C:\BEST\LMYS01\Orders\Backup\" & mySplit(filLocalFileT16.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT16.FileName, ".", -1)
'End If
'
'Kill strTranFileName
'
'filLocalFileT16.Refresh:
'Screen.MousePointer = 0: Cmd_Impdata.Enabled = True: SSTab2.Enabled = True
'Exit Sub

err_Handle:
    Set dgMainT16.DataSource = Nothing: Set dgMainT16_1.DataSource = Nothing
    Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: SSTab2.Enabled = True
    Call ErrorMsgbox(App.title, err.Number, err.Description, "�D���ɮצW�١G " & Str_updatesource1 & "�����ɮצW�١G" & Str_updatesource2)
End Sub

Private Sub cmdImportT15_Click()
Dim bl_Error As Boolean '�O�����p���I���X��
Dim str_Error As String '�O�����p���I���~�����
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT15 Is Nothing Then Exit Sub
If rsMainT15.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT15.Enabled = False: cmdImportT15.Enabled = False
strTranFileName = filLocalFileT15.Path & "\" & filLocalFileT15.FileName

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT15.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT15.Enabled = True: dgMainT15.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT15.RecordCount = 0 Or rsMainT15 Is Nothing Then
Else
rsMainT15.MoveFirst
str_Storerkey = "LCHF01"
Do While Not rsMainT15.EOF
    '��f����ˬd
    If Len(Trim(rsMainT15("�q��w���"))) = 0 Then
        If Trim(rsMainT15("��O")) = "�X�h" Or Trim(rsMainT15("��O")) = "�N�h" Then
        Else
            MsgBox "�q�渹�X:" & Trim(rsMainT15("�X�f�渹")) & "����f�鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
        End If
    ElseIf Len(Trim(rsMainT15("�q��w���"))) > 0 And Len(Trim(rsMainT15("�q��w���"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT15("�X�f�渹")) & "����f��:" & Trim(rsMainT15("�q��w���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
    Else
        '�ˬd��f�餣�i�p�󤵤�
        If Trim(rsMainT15("�q��w���")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '�̰��v�����ˬd��f��
                 x = MsgBox("��f��p�󤵤�A�A�T�w�n�~���?", vbQuestion + vbYesNo, "�̰��v����f���ˬd") '�������U���O�T�w�άO����
                    If x = 6 Then
                        '�~��
                    Else
                        '���}
                         dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
'    '�f�D�ˬd
'    If Len(Trim(rsMainT15("�f�D"))) = 0 Or Trim(rsMainT15("�f�D")) <> "LCHF01" Then
'        MsgBox "�q�渹�X:" & Trim(rsMainT15("�X�f�渹")) & "���f�D���~�A�q����J�פ�!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
'    End If
    
    
    '�q����ˬd
    If Len(Trim(rsMainT15("���"))) = 0 Then
        MsgBox "�q�渹�X:" & Trim(rsMainT15("�X�f�渹")) & "���q��鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT15("���"))) > 0 And Len(Trim(rsMainT15("���"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT15("�X�f�渹")) & "���q���:" & Trim(rsMainT15("���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
    Else
        If Trim(rsMainT15("��O")) = "�X�h" Or Trim(rsMainT15("��O")) = "�N�h" Then
        Else
            If Trim(rsMainT15("���")) > Trim(rsMainT15("�q��w���")) Then MsgBox "�q�渹�X:" & Trim(rsMainT15("�X�f�渹")) & "���q���:" & Trim(rsMainT15("���")) & "�A�j���f��A�q����J�פ�!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
        End If
    End If
    
    '�ƶq�ˬd
    If Val(rsMainT15("�ƶq")) < 1 Then
        If Trim(rsMainT15("��O")) = "�X�h" Or Trim(rsMainT15("��O")) = "�N�h" Then
        Else
            MsgBox "�ƶq�p��1�A" & Trim(rsMainT15("�X�f�渹")) & "-�~���G" & Trim(rsMainT15("���~�N��")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
            Exit Sub
        End If
    End If
    
        '������� --�P�_SKU�O�_�s�b
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT15("���~�N��")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT15("���~�N��")) & ")�A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
            dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
        '�ˬdA2B�q��H�~���Ȥ�s���O�_�s�b
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT15("���f�Ȥ�")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT15("���f�Ȥ�")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
            dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
        '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT15("�ƶq")), ".") <> 0 Then
            str_Error = "�q�渹�X:" & Trim(rsMainT15("�X�f�渹")) & "�A�~��:" & Trim(rsMainT15("���~�N��")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
'        '�ˬd�f�D--mark by Gemini @20160602
'        If UCase(Trim(rsMainT15("�f�D"))) <> "LABT01" And UCase(Trim(rsMainT15("�f�D"))) <> "LLFA01" Then
'            MsgBox "�q��o�{�D�Ȱ����f�D: " & Trim(rsMainT15("�f�D")) & " )�A���פJ�{���ȨѶפJ�Ȱ��ΧQ�׭q��A�нT�{��A�פJ�A�q����J�פ�!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
'            dgMainT15.Enabled = True: cmdImportT15.Enabled = True
'            Exit Sub
'        End If

        '�P�_��O
        If Trim(rsMainT15("��O")) = "A2B" Then
            MsgBox "��O��A2B:" & Trim(rsMainT15("��O")) & "�AA2B�q��ХѤ���EXCEL�q��פJ�A�q����J�פ�!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
                dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
        If Trim(rsMainT15("��O")) = "�X�f" Or Trim(rsMainT15("��O")) = "�X�h" Or Trim(rsMainT15("��O")) = "�N�h" Or Trim(rsMainT15("��O")) = "�N�P" Then
        Else
            MsgBox "�t�εL����O:" & Trim(rsMainT15("��O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
                dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
    rsMainT15.MoveNext
Loop
rsMainT15.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT15.Enabled = True: cmdImportT15.Enabled = True
                Exit Sub
End If


Tran_Level = cn.BeginTrans: cmdImportT15.Enabled = False: dgMainT15.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String


'�}�l�פJ
Do While Not rsMainT15.EOF
    DoEvents: DoEvents
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT15("�X�f�渹"))) Then
        strOrderNo = UCase(Trim(rsMainT15("�X�f�渹")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�w�g�ˬd�L�A�h�����A������X�Ȥ�D�ɤ��� �Ȥ�s���A�l���ϸ��A�s���H�A�q�ܡA�a�}
        '��O��A2B�h�A�촣�f�Ƚs�A�DA2B�h���f�Ƚs
        str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = '" & Trim(rsMainT15("���f�Ȥ�")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        Dim str_Priority As String
        '�۲Ūu���«Ƚs
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        
        '�N��O�N��I or R
        If myExCharFilter(Trim(rsMainT15("��O"))) = "�X�h" Or myExCharFilter(Trim(rsMainT15("��O"))) = "�N�h" Then
            str_Priority = "R"
        Else
            str_Priority = "I"
        End If
        
        blCustomerMatch = True
        tmp_Rs.Close

        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT15("�X�f�渹"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            tmp_Rs.Close
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
'            If UCase(Right(Trim(rsMainT15("�ܧO")), 2)) = "-C" Then
'                strFacility = "�ըƹF����"
'            ElseIf UCase(Right(Trim(rsMainT15("�ܧO")), 2)) = "-S" Then
'                strFacility = "�ըƹF�n��"
'            Else
            strFacility = "�ըƹF�_��"
'            End If
            
'            If Trim(rsMainT15("�ܧO")) = "" Then strFacility = ""

            strOrderDate = Trim(rsMainT15("���"))
            Dim intPointer As Integer
            intPointer = 1
            
            Call Confirm_Recordset_Closed(tmp_Rs)
            
            If (Trim(rsMainT15("��O")) = "�X�h" Or Trim(rsMainT15("��O")) = "�N�h") And Len(Trim(rsMainT15("�q��w���"))) = 0 Then
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT15("�X�f�渹")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
                Trim(rsMainT15("���f�Ȥ�")) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','" & Trim(rsMainT15(["�Ȥ�q��(�q��)"])) & "','" & Trim(rsMainT15("��ڳƵ�")) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
            Else
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT15("�X�f�渹")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & Trim(rsMainT15("�q��w���")) & "','" & strFacility & "','" & _
                Trim(rsMainT15("���f�Ȥ�")) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','" & Trim(rsMainT15(["�Ȥ�q��(�q��)"])) & "','" & Trim(rsMainT15("��ڳƵ�")) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
            End If
            
            
'            If (Trim(rsMainT15("��O")) = "�X�h" Or Trim(rsMainT15("��O")) = "�N�h") And Len(Trim(rsMainT15("�q��w���"))) = 0 Then
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT15("�X�f�渹"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT15("���f�Ȥ�"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT15("��ڳƵ�"))) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            Else
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT15("�X�f�渹"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT15("�q��w���"))) & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT15("���f�Ȥ�"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT15("��ڳƵ�"))) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            End If
            
            
            
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT15("�X�f�渹")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Abs(Val(rsMainT15("�ƶq")))
            
            If Trim(rsMainT15("��O")) = "�X�h" Or Trim(rsMainT15("��O")) = "�N�h" Then
                '�q����Ӹ�Ʒs�W �h�f
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM) " & _
                "values ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT15("�X�f�渹")) & "','" & Trim(rsMainT15("���~�N��")) & "','LCHF01'," & _
                "'" & intQTY & "','" & intQTY & "','R01','" & strFacility & "','" & Trim(rsMainT15("�ƶq���")) & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            Else
                '�q����Ӹ�Ʒs�W
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Lottable03)" & _
                "select  '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT15("�X�f�渹")) & "','" & Trim(rsMainT15("���~�N��")) & "','LCHF01'," & intQTY & " * p.casecnt ," & intQTY & " * p.casecnt " & _
                ",'R01','" & strFacility & "','" & Trim(rsMainT15("�ƶq���")) & "','" & Trim(rsMainT15("�帹")) & "'" & _
                "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT15("���~�N��")) & "' and s.storerkey = 'LCHF01'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT15.MoveNext
Loop

'��s������
cn.Execute "exec gs_ordersupdate 'LCHF01'", RowsAffect, adExecuteNoRecords

'�ˬd���L���`�q��  �ܧO�w�O���~ es_Checklot06_by_storer '�f�D','�q���ɦW'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT18.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '���^�ǿ��~���q���ơA����excel
'    Recordset2Excel "�t�e�ܧO�P���ӭܧO���Ū��q����", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT15.Enabled = True


'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�ɮ� " & filLocalFileT15.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT15.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ��ɮ�
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT15.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT15.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT15.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT15.FileName, ".", -1)
End If

'�ƥ���FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT15.FileName


Kill strTranFileName
    
filLocalFileT15.Refresh:
Screen.MousePointer = 0: cmdImportT15.Enabled = True: dgMainT15.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT15.Enabled = True: Screen.MousePointer = 0: dgMainT15.Enabled = True

End Sub



Private Sub cmdImportT17_Click()
Dim bl_Error As Boolean '�O�����p���I���X��
Dim str_Error As String '�O�����p���I���~�����
Dim x As Integer '�O�����U�T�w�άO�����A�M�w�O�_��s���q��urgent_mark
Dim Str_packkey As String '����packkey
bl_Error = False: str_Error = "": Str_packkey = ""

If rsMainT17 Is Nothing Then Exit Sub
If rsMainT17.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT17.Enabled = False: cmdImportT17.Enabled = False
strTranFileName = filLocalFileT17.Path & "\" & filLocalFileT17.FileName

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT17.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT17.Enabled = True: cmdImportT17.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT17.RecordCount = 0 Or rsMainT17 Is Nothing Then
Else
rsMainT17.MoveFirst
Do While Not rsMainT17.EOF
    '��f����ˬd
    If Len(Trim(rsMainT17("�w�p�X�f��"))) = 0 Then
    Else
        arrTmp = Split(Trim(rsMainT17("�w�p�X�f��")), "/")
        If Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "�w�p�X�f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: dgMainT17.Enabled = True: cmdImportT17.Enabled = True: Exit Sub
    End If
    
    '�ƶq�ˬd
    If Val(rsMainT17("�ӫ~�q�ʼƶq")) < 1 Then
        MsgBox "�ӫ~�q�ʼƶq�p��1�A" & Trim(rsMainT17("SAP DN NO.")) & "-�~���G" & Trim(rsMainT17("�ӫ~�s�X")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT17.Enabled = True: cmdImportT17.Enabled = True: Exit Sub
        Exit Sub
    End If
    
    '������� --�P�_SKU�O�_�s�b
        str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where Storerkey = 'LAPP01' and sku='" & Trim(rsMainT17("�ӫ~�s�X")) & "'"
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT17("�ӫ~�s�X")) & ")�A�q����J�פ�!!": cmdImportT17.Enabled = True: Screen.MousePointer = 0
            dgMainT17.Enabled = True: cmdImportT17.Enabled = True
            Exit Sub
        End If
    '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT17("�ӫ~�q�ʼƶq")), ".") <> 0 Then
            str_Error = "�q�渹�X:" & Trim(rsMainT17("SAP DN NO.")) & "�A�~��:" & Trim(rsMainT17("�ӫ~�s�X")) & Chr(13) & str_Error
            bl_Error = True
        End If
    rsMainT17.MoveNext
Loop
rsMainT17.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT17.Enabled = True: cmdImportT17.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT17.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String

'�}�l�פJ
If rsMainT17 Is Nothing Then GoTo next18
If rsMainT17.RecordCount = 0 Then GoTo next18

'���̫�Ȥ�s��
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LAPP01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close


Do While Not rsMainT17.EOF
    DoEvents: DoEvents
    
'    �������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If Trim(rsMainT17("�~��")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow17
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT17("SAP DN NO."))) Then
        strOrderNo = UCase(Trim(rsMainT17("SAP DN NO.")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�ˬd�O�_�����Ȥ���
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LAPP01' and consigneekey = '" & Right(myExCharFilter(Trim(rsMainT17("�Ȥ�N��"))), 7) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '�L���Ȥ�W�٫h�s�W
            strConsigneeKey = Right(myExCharFilter(Trim(rsMainT17("�Ȥ�N��"))), 7)
            
            '�s�W�Ȥ�D��
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho,channel) " & _
            " values('LAPP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT17("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT17("�Ȥ�W��"))) & "','','','" & myExCharFilter(Trim(rsMainT17("�e�f�a�}"))) & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT17("�Ȥ�q���O"))) & "' ) ", RowsAffect, adExecuteNoRecords
            '�����s�W���Ȥ�s��
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
'            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LAPP01' and full_name = '" & myExCharFilter(Trim(rsMainT17("�Ȥ�W��"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT17("�e�f�a�}"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'                If rsTmp.EOF Then
'                    '�p���H�B�q�ܻP��f�a�}����
'                strConsigneekey = myExCharFilter(Trim(rsMainT17("�Ȥ�N��")))
'
'                    '�s�W�Ȥ�D��
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho,channel) " & _
'                    " values('LAPP01','','" & strConsigneekey & "','" & myExCharFilter(Trim(rsMainT17("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT17("�Ȥ�W��"))) & "','','','" & myExCharFilter(Trim(rsMainT17("�e�f�a�}"))) & "','','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT17("�Ȥ�q���O"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
'                    '�����s�W���Ȥ�s��
'                    strNewConsigneekey = strNewConsigneekey & strConsigneekey & "','"
'                Else
                    
                    '�۲Ūu���«Ƚs
                    strConsigneeKey = Trim(tmp_Rs("consigneekey"))
                    blCustomerMatch = True
'
'                End If
'            rsTmp.Close
        End If
        tmp_Rs.Close
    
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select orderkey from orders where storerkey = 'LAPP01' and isnull(type,'') <> '�R��' and rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT17("SAP DN NO."))) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            If Right(UCase(Trim(rsMainT17("DC�N��"))), 1) = "C" Then strFacility = "�ըƹF����"
            If Right(UCase(Trim(rsMainT17("DC�N��"))), 1) = "S" Then strFacility = "�ըƹF�n��"

            
            arrTmp = Split(Trim(rsMainT17("�w�p�X�f��")), "/")
            If Len(Trim(rsMainT17("�w�p�X�f��"))) = 0 Then
                cn.RollbackTrans: Tran_Level = 0
                msg_text = "�q�渹�X:" & Trim(rsMainT17("SAP DN NO.")) & "�A�~��:" & Trim(rsMainT17("�ӫ~�s�X")) & Chr(13) & "��ƨS���w�p�X�f��I�ЦV�t�ӽT�{"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT17.Enabled = True: cmdImportT17.Enabled = True
                Exit Sub
            Else
                strDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            End If
            
            arrTmp = Split(Trim(rsMainT17("��ڲ��ͤ�")), "/")
            strOrderDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT17("SAP DN NO."))) & "','I','LAPP01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT17("�Ȥ�W��"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT17("�e�f�a�}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT17("�e�f�a�}"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT17("�ƪ`"))) & "','" & filLocalFileT17.FileName & "','','" & User_id & "','" & User_id & "','','" & Right(myExCharFilter(Trim(rsMainT17("�Ȥ�N��"))), 7) & "') "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            'Mark by Eric�]��gs_ordersupdate�N�|��szip�F
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LAPP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT17("SAP DN NO.")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            
            lngCasecnt = 1
            
'            '��촫��
'            If Left(myExCharFilter(Trim(rsMainT17("���"))), 1) = "�c" Then
'
                '���c�]�ഫ�v
                str_SQL = "select p.casecnt, p.innerpack,p.packkey from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey where s.storerkey = 'LAPP01' and s.sku = '" & myExCharFilter(Trim(rsMainT17("�ӫ~�s�X"))) & "'"
                
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                lngCasecnt = tmp_Rs("casecnt")      '�j���J��
                lngInnerpack = tmp_Rs("innerpack")  '�����J��
                Str_packkey = tmp_Rs("packkey")  'packkey
                tmp_Rs.Close
'
'            End If
            
            
'            If UCase(rsMainT17("SAP�q�f���")) = "BDL" Then intQTY = Val(rsMainT17("�ӫ~�q�ʼƶq")) * lngInnerpack  '�����J�ơA��Xpack��ƪ�
'            If UCase(rsMainT17("SAP�q�f���")) = "KAR" Then intQTY = Val(rsMainT17("�ӫ~�q�ʼƶq")) * lngCasecnt  '�j���J�ơA

            
            intQTY = Val(rsMainT17("�ӫ~�q�ʼƶq"))

            '��trp19m�ܧO��Ӫ�
            str_SQL = "select bestlot06 from trp19m(nolock) where storerkey = 'LAPP01' and storerlot06 = '" & rsMainT17("DC�N��") & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

            If tmp_Rs.EOF Then
            '�S�����������ܧO
            '�u�έq��W���ܧO
                strLot06 = rsMainT17("DC�N��") '�ݽT�{
            Else
            '�����������ܧO
                strLot06 = Trim(tmp_Rs("bestlot06"))
            End If

            tmp_Rs.Close
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,Packkey)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT17("SAP DN NO."))) & "','" & myExCharFilter(Trim(rsMainT17("�ӫ~�s�X"))) & "','LAPP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT17("SAP�q�f���"))) & "','0','" & Str_packkey & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            'Mark by Eric 20141216����c�J�Ʈɶ��K���packkey�g�J
'            '��spackkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT17.MoveNext
Loop

next18:

'����gs_ordersupdate   �ΫȤ�D�ɧ�s�q����

cn.Execute "exec gs_ordersupdate 'LAPP01'", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'�N���q��аO�borders..urgent_mark���
'�ˬd�O�_�����q��?
str_SQL = "select orderkey " & _
"From orders(nolock) " & _
"where storerkey = 'LAPP01' and priority = 'I' and updatesource = '" & filLocalFileT17.FileName & "' and " & _
"((convert(varchar(8),adddate,114) > '17:00:00' and convert(varchar(8),deliverydate,112) = convert(varchar(8),getdate()+1,112)) or " & _
"(convert(varchar(8),adddate,114) > '17:30:00' and convert(varchar(8),deliverydate,112) = convert(varchar(8),getdate()+2,112) ) or " & _
"(convert(varchar(8),adddate,112) = convert(varchar(8),deliverydate,112)) " & _
") "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '���^��
    x = MsgBox("�o�{���q��A�O�_�۰ʱN�q���s�����q��?", vbQuestion + vbYesNo, "APP�q��פJ") '���U���O�T�w6�άO����
    If x = 6 Then
           '��surgent_mark���V:���q��
           cn.Execute "exec es_update_urgent_mark 'LAPP01','" & filLocalFileT17.FileName & "'", RowsAffect, adExecuteNoRecords
    End If
End If

'�ˬd���L���`�q��  �ܧO�w�O���~ es_Checklot06_by_storer '�f�D','�q���ɦW'
str_SQL = "exec es_Checklot06_by_storer 'LAPP01','" & filLocalFileT17.FileName & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '���^�ǿ��~���q���ơA����excel
    Recordset2Excel "�t�e�ܧO�P���ӭܧO���Ū��q����", tmp_Rs
End If

tmp_Rs.Close


'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�פJ �D�ըƹF�q�� " & intNotBest & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT17.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT17.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = 'LAPP01' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LAPP01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LAPP01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ���FTP
If Dir("O:\LAPP01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LAPP01\OrdersBackup"
FileCopy strTranFileName, "O:\LAPP01\OrdersBackup\" & filLocalFileT17.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LAPP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\Orders\Backup"
If Dir("C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT17.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT17.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & mySplit(filLocalFileT17.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT17.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT17.Refresh:
Screen.MousePointer = 0: cmdImportT17.Enabled = True: dgMainT17.Enabled = True
Exit Sub

err_Handle:
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT17.Enabled = True: Screen.MousePointer = 0: dgMainT17.Enabled = True
End Sub

Private Sub cmdImportT18_Click()
Dim bl_Error As Boolean '�O�����p���I���X��
Dim str_Error As String '�O�����p���I���~�����
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT18 Is Nothing Then Exit Sub
If rsMainT18.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT18.Enabled = False: cmdImportT18.Enabled = False
strTranFileName = filLocalFileT18.Path & "\" & filLocalFileT18.FileName

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT18.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT18.RecordCount = 0 Or rsMainT18 Is Nothing Then
Else
rsMainT18.MoveFirst
str_Storerkey = "LPSI01"

Do While Not rsMainT18.EOF
    '��f����ˬd
    If Len(Trim(rsMainT18("��f���"))) = 0 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT18("��f")) & "����f�鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT18("��f���"))) > 0 And Len(Trim(rsMainT18("��f���"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT18("��f")) & "����f��:" & Trim(rsMainT18("��f���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT18("��f���")), 4) + "/" + Mid(Trim(rsMainT18("��f���")), 6, 2) + "/" + Right(Trim(rsMainT18("��f���")), 2)) = False Then
         MsgBox "�q�渹�X:" & Trim(rsMainT18("��f")) & "����f��:" & Trim(rsMainT18("��f���")) & "�A���O�@�ӥ��`����A�q����J�פ�!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    Else
        '�ˬd��f�餣�i�p�󤵤�
        If Trim(rsMainT18("��f���")) < Format(Now, "YYYY.MM.DD") Then
            If blAdmin = True Then
            
            '�̰��v�����ˬd��f��
                 x = MsgBox("��f��p�󤵤�A�A�T�w�n�~���?", vbQuestion + vbYesNo, "�̰��v����f���ˬd") '�������U���O�T�w�άO����
                    If x = 6 Then
                        '�~��
                    Else
                        '���}
                         dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If

    '�q����ˬd
    If Len(Trim(rsMainT18("�ѳf���"))) = 0 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT18("��f")) & "���q��鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT18("�ѳf���"))) > 0 And Len(Trim(rsMainT18("�ѳf���"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT18("��f")) & "���q���:" & Trim(rsMainT18("�ѳf���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT18("�ѳf���")), 4) + "/" + Mid(Trim(rsMainT18("�ѳf���")), 6, 2) + "/" + Right(Trim(rsMainT18("�ѳf���")), 2)) = False Then
         MsgBox "�q�渹�X:" & Trim(rsMainT18("��f")) & "���q���:" & Trim(rsMainT18("�ѳf���")) & "�A���O�@�ӥ��`����A�q����J�פ�!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub

    Else
        If Trim(rsMainT18("�ѳf���")) > Trim(rsMainT18("��f���")) Then MsgBox "�q�渹�X:" & Trim(rsMainT18("��f")) & "���q���:" & Trim(rsMainT18("�ѳf���")) & "�A�j���f��A�q����J�פ�!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    End If
    
    '�ƶq�ˬd
    If Val(rsMainT18("��f�ƶq")) < 1 Then
        MsgBox "�ƶq�p��1�A" & Trim(rsMainT18("��f")) & "-�~���G" & Trim(rsMainT18("����")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '������� --�P�_SKU�O�_�s�b
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT18("����")) & "' and Storerkey = 'LPSI01' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT18("����")) & ")�A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": cmdImportT18.Enabled = True: Screen.MousePointer = 0
            dgMainT18.Enabled = True: cmdImportT18.Enabled = True
            Exit Sub
        End If
        
        '�ˬd���f�Ȥ�s���O�_�s�b
            str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT18("�u�t")) & "' and Storerkey = 'LPSI01' "
        
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then  '���s���ǭn��
                MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT18("�u�t")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT18.Enabled = True: Screen.MousePointer = 0
                dgMainT18.Enabled = True: cmdImportT18.Enabled = True
                Exit Sub
            End If
        
        '�ˬd��f�Ȥ�s���O�_�s�b
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT18("���f�H")) & "' and Storerkey = 'LPSI01' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT18("���f�H")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT18.Enabled = True: Screen.MousePointer = 0
            dgMainT18.Enabled = True: cmdImportT18.Enabled = True
            Exit Sub
        End If
        
        '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT18("��f�ƶq")), ".") <> 0 Then
            str_Error = "�q�渹�X:" & Trim(rsMainT18("��f")) & "�A�~��:" & Trim(rsMainT18("����")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
    rsMainT18.MoveNext
Loop
rsMainT18.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT18.Enabled = True: cmdImportT18.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT18.Enabled = False: dgMainT18.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'�}�l�פJ
Do While Not rsMainT18.EOF
    DoEvents: DoEvents
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT18("��f"))) Then
        strOrderNo = UCase(Trim(rsMainT18("��f")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�w�g�ˬd�L�A�h�����A������X�Ȥ�D�ɤ��� �Ȥ�s���A�l���ϸ��A�s���H�A�q�ܡA�a�}
        If myExCharFilter(Trim(rsMainT18("�X�f��m"))) = "�_��" Then
            '��O��C
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = 'LPSI01' and consigneekey = '" & myExCharFilter(Trim(rsMainT18("���f�H"))) & "'"
        Else
            '��O��A2B�h�A�촣�f�Ƚs�A�DA2B�h���f�Ƚs
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = 'LPSI01' and consigneekey = '" & myExCharFilter(Trim(rsMainT18("�u�t"))) & "'"
        End If
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        '�۲Ūu���«Ƚs
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close
        
        strFacility = strShort_name
        
        'C����t�e�ܧO
        If myExCharFilter(Trim(rsMainT18("�X�f��m"))) = "�_��" Then
'            str_SQL = "select short_name=isnull(short_name,'') from trp01m where storerkey = 'LPSI01' and consigneekey = '" & myExCharFilter(Trim(rsMainT18("�u�t"))) & "'"
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
        
        strFacility = "�ըƹF�_��"
        
'        tmp_Rs.Close
        End If

        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT18("��f"))) & "' and storerkey = 'LPSI01' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
                        
            strOrderDate = Trim(rsMainT18("�ѳf���"))
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            
            If myExCharFilter(Trim(rsMainT18("�X�f��m"))) = "�_��" Then
            
                '�_�ϥ�C��
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT18("��f"))) & "','C','LPSI01','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT18("��f���"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT18("���f�H"))) & "','','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT18("���ʤ��")) & "','','" & filLocalFileT18.FileName & "','','" & User_id & "','" & User_id & "','','','','0') "
            
            Else
                '���n�ϥ�A2B�A�h�����@��B�I���ȽsB_company
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT18("��f"))) & "','A2B','LPSI01','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT18("��f���"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT18("�u�t"))) & "','" & myExCharFilter(Trim(rsMainT18("���f�H"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT18("���ʤ��")) & "','','" & filLocalFileT18.FileName & "','','" & User_id & "','" & User_id & "','','','','0') "
            End If
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT18("��f")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Val(rsMainT18("��f�ƶq"))
            
            
            '�q����Ӹ�Ʒs�W
'            If Trim(rsMainT18("���W��")) = "�c" Or Trim(rsMainT18("���W��")) = "CS" Or Trim(rsMainT18("���W��")) = "CASE" Then
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
            " select '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT18("��f"))) & "','" & myExCharFilter(Trim(rsMainT18("����"))) & "','LPSI01'," & _
            "'" & intQTY & "' * p.casecnt ,'" & intQTY & "' * p.casecnt,'','" & strFacility & "',''" & _
            "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT18("����")) & "' and s.storerkey = 'LPSI01' "
'            Else
'                 str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
'                "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT18("�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT18("�~��"))) & "','" & myExCharFilter(Trim(rsMainT18("�f�D"))) & "'," & _
'                "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT18("�ܧO"))) & "','','')"
'           End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT18.MoveNext
Loop

'����gs_ordersupdate   �ΫȤ�D�ɧ�s�q����

'cn.Execute "exec gs_ordersupdate 'LPSI01'", RowsAffect, adExecuteNoRecords

'�ˬd���L���`�q��  �ܧO�w�O���~ es_Checklot06_by_storer '�f�D','�q���ɦW'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT18.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '���^�ǿ��~���q���ơA����excel
'    Recordset2Excel "�t�e�ܧO�P���ӭܧO���Ū��q����", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT18.Enabled = True


'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�ɮ� " & filLocalFileT18.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT18.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ��ɮ�
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT18.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT18.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT18.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT18.FileName, ".", -1)
End If

'�ƥ���FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT18.FileName



Kill strTranFileName
    
filLocalFileT18.Refresh:
Screen.MousePointer = 0: cmdImportT18.Enabled = True: dgMainT18.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT18.Enabled = True: Screen.MousePointer = 0: dgMainT18.Enabled = True

End Sub

Private Sub cmdImportT19_Click()
'
'APP�h�f�q�檺�פJ�{�� create by Eric 20130613
'�}�Y�ˬd�פJ���q��:
'1.��f��O�_�j�󤵤� 2.�ƶq�O�_<0 3.�ƶq�O�_���p���I 4.�~���O�_�s�b
'�ˬd�q��Ȥ�s�� , �t�άO�_�s�b, �s�b�h�a�X�t�ΫȤ�D�ɤ������, ���s�b�h�ϥέq��W���Ȥ�s���i��s�W (�����Ȥ�D�ɸ�ơA�i�歫�s)
'�{�������� , �|����ordersupdate, �t�ΫȤ�D��, ��s�q����

Dim str_Error As String
Dim bl_Error As Boolean
On Error GoTo err_Handle
strTranFileName = filLocalFileT19.Path & "\" & filLocalFileT19.FileName
If Len(RTrim(cboSheetT19)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT19.EOF Or rsMainT19 Is Nothing Then Exit Sub

Screen.MousePointer = 11: SSTab2.Enabled = False: cmdImportT19.Enabled = False


'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT19.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT19.MoveFirst

Do While Not rsMainT19.EOF

    '��f����ˬd
    arrTmp = Split(Trim(rsMainT19("�w�p��f���")), ".")
    If Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then:  MsgBox "�w�p��f����p�󤵤�A�q����J�פ�!", 16, Me.Caption: cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0: Exit Sub
    
    '�ƶq�ˬd
    If Trim(rsMainT19("�ӫ~�q�ʼƶq")) < 1 Then
        MsgBox "�o�{�ӫ~�q�ʼƶq�p��1�A" & "��f�渹:" & Trim(rsMainT19("��f�渹")) & "-�Ȥ�N��:" & Trim(rsMainT19("�Ȥ�N��")) & "-�Ȥ�W��:" & Trim(rsMainT19("�Ȥ�W��")) & "-�ӫ~�q�ʼƶq:" & Trim(rsMainT19("�ӫ~�q�ʼƶq")) & ")�A�q����J�פ�!!", , "�h�f��פJ": cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0: Exit Sub
        Exit Sub
    End If
    
    '�ˬd�ƶq���L�p���I
    If InStr(Trim(rsMainT19("�ӫ~�q�ʼƶq")), ".") <> 0 Then
        str_Error = "�q�渹�X:" & Trim(rsMainT19("��f�渹")) & "�A�~��:" & Trim(rsMainT19("�~��")) & Chr(13) & str_Error
        bl_Error = True
    End If
    
    '������� --�P�_SKU�O�_�s�b
    str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where sku='" & Trim(rsMainT19("�~��")) & "' and Storerkey = 'LAPP01' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    If tmp_Rs.EOF Then  '���s���ǭn��
        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT19("�~��")) & ")�A�q����J�פ�!!":
        cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0:
        Exit Sub
    End If
    
    rsMainT19.MoveNext
Loop

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0:
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT19.Enabled = False: dgMainT19.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim str_Fullname As String, str_Contact As String, str_Phone As String, Str_Address As String, str_Channel As String
            
'���̫�Ȥ�s��
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m(nolock) where storerkey = 'LAPP01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT19.MoveFirst
Do While Not rsMainT19.EOF
    DoEvents: DoEvents
    
'    '�������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If UCase(Trim(rsMainT19("�ܮw"))) = "" Then
''        MsgBox "�Ȥ�渹�G" & Trim(rsMainT4("�P�f�渹")) & "( " & Trim(rsMainT4("�ܮw")) & " )" & vbCrLf & "�гq���Ȥ�A�T�{�ӵ��q��O�_���~!?", vbOKOnly, "�D�ըƹF���q�椣��J"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '�������--�P�_SKU�O�_�s�b
'    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT19("�~��")) & "' and Storerkey = 'LAPP01' "
'
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.EOF Then
'        cn.RollbackTrans: Tran_Level = 0
'        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT19("�~��")) & " ) " & Trim(rsMainT19("�~�W")) & "�A�q����J�פ�!!": cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0
'        Exit Sub
'    End If

'    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT19("��f�渹"))) Then
        strOrderNo = UCase(Trim(rsMainT19("��f�渹")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�ˬd�O�_�����Ȥ�s���A(�a�X�t�� �Ȥ�W�١B�p���H�B�q�ܡB�a�}�B�q��)
        str_SQL = "select consigneekey,full_name,contact=isnull(contact,''),phone=isnull(phone,''),address,channel=isnull(channel,'') from trp01m(nolock) where  storerkey = 'LAPP01'  and consigneekey = '" & myExCharFilter(Trim(rsMainT19("�Ȥ�N��"))) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        
        If tmp_Rs.EOF Then
            '�L���Ȥ�W�٫h�s�W
            intTmp = intTmp + 1
            strConsigneeKey = myExCharFilter(Trim(rsMainT19("�Ȥ�N��")))
            
            '�s�W�Ȥ�D��
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho,channel) " & _
            " values('LAPP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT19("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT19("�Ȥ�W��"))) & "','','','" & myExCharFilter(Trim(rsMainT19("���f�a�}"))) & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT19("�q��"))) & "') ", RowsAffect, adExecuteNoRecords
'
            '�����s�W���Ȥ�s��
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
'            '�T�{�{���O�_�n�έq��a�}�A�n���ܫh���s�Ȥ�D�ɡA�_�h�h�����ϥΨt�ΫȤ�s�������
'            '��� (�q�ܡB��f�a�}) �O�_�۲�
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LAPP01' and full_name = '" & myExCharFilter(Trim(rsMainT19("�Ȥ�W��"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT19("���f�a�}"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'            If rsTmp.EOF Then
'                '�q�ܻP�a�}���ūh���s
'                intTmp = intTmp + 1
'                strConsigneeKey = "BEST" & Format(intTmp, "000000")
'
'                '�s�W�Ȥ�D��
'                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                " values('LAPP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT19("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT19("�Ȥ�W��"))) & "','','','" & myExCharFilter(Trim(rsMainT19("���f�a�}"))) & "','','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'
'                '�����s�W���Ȥ�s��
'                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'            Else '�۲Ūu���«Ƚs
'
'                str_Fullname = Trim(tmp_Rs("full_name").Value)
'                str_Contact = Trim(tmp_Rs("contact").Value)
'                str_Phone = Trim(tmp_Rs("phone").Value)
'                str_Address = Trim(tmp_Rs("address").Value)
'                str_Channel = Trim(tmp_Rs("channel").Value)
'                strConsigneeKey = Trim(rsTmp("consigneekey"))
'                blCustomerMatch = True
'
'            End If
'                rsTmp.Close
                str_Fullname = Trim(tmp_Rs("full_name").Value)
                str_Contact = Trim(tmp_Rs("contact").Value)
                str_Phone = Trim(tmp_Rs("phone").Value)
                Str_Address = Trim(tmp_Rs("address").Value)
                str_Channel = Trim(tmp_Rs("channel").Value)
                strConsigneeKey = Trim(tmp_Rs("consigneekey"))
                blCustomerMatch = True

        End If
        tmp_Rs.Close
    
'        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select orderkey from orders where storerkey = 'LAPP01' and isnull(type,'') <> '�R��' and rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT19("��f�渹"))) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '���f�w�s�a  --�T�{��}��
            strFacility = "�ըƹF�_��"
            If Right(myExCharFilter(Trim(rsMainT19("���f�w�s�a"))), 1) = "C" Then strFacility = "�ըƹF����"
            If Right(myExCharFilter(Trim(rsMainT19("���f�w�s�a"))), 1) = "S" Then strFacility = "�ըƹF�n��"
            
            arrTmp = Split(Trim(rsMainT19("�w�p��f���")), ".")
            strDeliveryDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)   '��f���
            strOrderDate = Format(Now, "YYYY/MM/DD")    '�q����
            Dim intPointer As Integer
            intPointer = 1
            
            '�q�����ݽT�{
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT19("��f�渹"))) & "','R','LAPP01','" & strOrderDate & "','" & strDeliveryDate & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & str_Fullname & "','" & str_Contact & "','','','" & str_Phone & "','','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 45))) & "','','','" & filLocalFileT19.FileName & "','','" & User_id & "','" & User_id & "','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            'If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LAPP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT19("��f�渹")) & "','"
            blDuplicationOrder = True

        End If
    End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Abs(Trim(rsMainT19("�ӫ~�q�ʼƶq")))
            strLot06 = (Trim(rsMainT19("���f�w�s�a")))
            
            '��trp19m�ܧO��Ӫ�
            str_SQL = "select bestlot06 from trp19m(nolock) where storerkey = 'LAPP01' and storerlot06 = '" & Trim(rsMainT19("���f�w�s�a")) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

            If tmp_Rs.EOF Then
            '�S�����������ܧO
            '�u�έq��W���ܧO
                strLot06 = (Trim(rsMainT19("���f�w�s�a")))
            Else
            '�����������ܧO
                strLot06 = Trim(tmp_Rs("bestlot06"))
            End If

            tmp_Rs.Close
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,addwho,editwho)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT19("��f�渹"))) & "','" & myExCharFilter(Trim(rsMainT19("�~��"))) & "','LAPP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','','0','" & User_id & "','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT19.MoveNext
Loop

'����gs_ordersupdate   �ΫȤ�D�ɧ�s�q����

cn.Execute "exec gs_ordersupdate 'LAPP01'", RowsAffect, adExecuteNoRecords

cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0:

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT19.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT19.FileName & " �ƥ��� C:\BEST\LAPP01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT19.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT19.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = 'LAPP01' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\LAPP01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LAPP01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ���FTP
If Dir("O:\LAPP01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LAPP01\OrdersBackup"
FileCopy strTranFileName, "O:\LAPP01\OrdersBackup\" & filLocalFileT19.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LAPP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\Orders\Backup"
If Dir("C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT19.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT19.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & mySplit(filLocalFileT19.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT19.FileName, ".", -1)
End If

filLocalFileT19.Refresh: cboSheetT19.Clear

Kill strTranFileName

Screen.MousePointer = 0: cmdImportT19.Enabled = True: dgMainT19.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    cmdImportT19.Enabled = True: Screen.MousePointer = 0: dgMainT19.Enabled = True
    Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
End Sub

Private Sub cmdImportT20_Click() 'Terry 20180825 �S�O�έq��פJ
Dim bl_Error As Boolean '�O�����p���I���X��
Dim str_Error As String '�O�����p���I���~�����
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT20 Is Nothing Then Exit Sub
If rsMainT20.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT20.Enabled = False: cmdImportT20.Enabled = False
strTranFileName = filLocalFileT20.Path & "\" & filLocalFileT20.FileName

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT20.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT20.Enabled = True: dgMainT20.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT20.RecordCount = 0 Or rsMainT20 Is Nothing Then
Else
rsMainT20.MoveFirst
str_Storerkey = "LTRI03"
Do While Not rsMainT20.EOF
    '��f����ˬd
    If Len(Trim(rsMainT20("��f��"))) = 0 Then
    
        'Terry �ݽT�{��O��� �ק襤
        If Trim(rsMainT20("�q�����O")) = "R" Then
        Else
            MsgBox "�q�渹�X:" & Trim(rsMainT20("�q�渹�X")) & "����f�鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
        End If
    ElseIf Len(Trim(rsMainT20("��f��"))) > 0 And Len(Trim(rsMainT20("��f��"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT20("�q�渹�X")) & "����f��:" & Trim(rsMainT20("��f��")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    Else
        '�ˬd��f�餣�i�p�󤵤�
        If Trim(rsMainT20("��f��")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '�̰��v�����ˬd��f��
                 x = MsgBox("��f��p�󤵤�A�A�T�w�n�~���?", vbQuestion + vbYesNo, "�̰��v����f���ˬd") '�������U���O�T�w�άO����
                    If x = 6 Then
                        '�~��
                    Else
                        '���}
                         dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
'    '�f�D�ˬd
'    If Len(Trim(rsMainT20("�f�D"))) = 0 Or Trim(rsMainT20("�f�D")) <> "LCHF01" Then
'        MsgBox "�q�渹�X:" & Trim(rsMainT20("�q�渹�X")) & "���f�D���~�A�q����J�פ�!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
'    End If
    
    
    '�q����ˬd
    If Len(Trim(rsMainT20("�q���"))) = 0 Then
        MsgBox "�q�渹�X:" & Trim(rsMainT20("�q�渹�X")) & "���q��鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT20("�q���"))) > 0 And Len(Trim(rsMainT20("�q���"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT20("�q�渹�X")) & "���q���:" & Trim(rsMainT20("�q���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    ElseIf Trim(rsMainT20("�q���")) > Trim(rsMainT20("��f��")) Then
        MsgBox "�q�渹�X:" & Trim(rsMainT20("�q�渹�X")) & "���q���:" & Trim(rsMainT20("�q���")) & "�A�j���f��A�q����J�פ�!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    End If
    
    '�ƶq�ˬd
    If Val(rsMainT20("�ƶq")) < 1 Then
        MsgBox "�ƶq�p��1�A" & Trim(rsMainT20("�q�渹�X")) & "-�~���G" & Trim(rsMainT20("�~��")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '������� --�P�_SKU�O�_�s�b
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT20("�~��")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT20("�~��")) & ")�A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
            dgMainT20.Enabled = True: cmdImportT20.Enabled = True
            Exit Sub
        End If
        
        'Terry �S�O�ΨS���Ȥ�D��
'        '�ˬdA2B�q��H�~���Ȥ�s���O�_�s�b
'        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT20("��f�Ȥ�s��")) & "' and Storerkey = '" & str_Storerkey & "' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
'        If tmp_Rs.EOF Then  '���s���ǭn��
'            MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT20("��f�Ȥ�s��")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
'            dgMainT20.Enabled = True: cmdImportT20.Enabled = True
'            Exit Sub
'        End If
        
        '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT20("�ƶq")), ".") <> 0 Then
            str_Error = "�q�渹�X:" & Trim(rsMainT20("�q�渹�X")) & "�A�~��:" & Trim(rsMainT20("�~��")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
        
        'Terry �ݽT�{��O��� �ק襤
'        '�P�_��O
'        If Trim(rsMainT20("�q�����O")) = "A2B" Then
'            MsgBox "�q�����O��A2B:" & Trim(rsMainT20("�q�����O")) & "�AA2B�q��ХѤ���EXCEL�q��פJ�A�q����J�פ�!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
'                dgMainT20.Enabled = True: cmdImportT20.Enabled = True
'            Exit Sub
'        End If
'
'
'        If Trim(rsMainT20("�q�����O")) = "�X�f" Or Trim(rsMainT20("�q�����O")) = "�X�h" Or Trim(rsMainT20("�q�����O")) = "�N�h" Or Trim(rsMainT20("�q�����O")) = "�N�P" Then
'        Else
'            MsgBox "�t�εL����O:" & Trim(rsMainT20("�q�����O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
'                dgMainT20.Enabled = True: cmdImportT20.Enabled = True
'            Exit Sub
'        End If
        
    rsMainT20.MoveNext
Loop
rsMainT20.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT20.Enabled = True: cmdImportT20.Enabled = True
                Exit Sub
End If


Tran_Level = cn.BeginTrans: cmdImportT20.Enabled = False: dgMainT20.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String


'�}�l�פJ
Do While Not rsMainT20.EOF
    DoEvents: DoEvents
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT20("�q�渹�X"))) Then
        strOrderNo = UCase(Trim(rsMainT20("�q�渹�X")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        
        'Terry �S�O�ΨS���Ȥ�D��
        '�w�g�ˬd�L�A�h�����A������X�Ȥ�D�ɤ��� �Ȥ�s���A�l���ϸ��A�s���H�A�q�ܡA�a�}
        '��O��A2B�h�A�촣�f�Ƚs�A�DA2B�h���f�Ƚs
        str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = 'LTRI03'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        Dim str_Priority As String
        '�۲Ūu���«Ƚs
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        
'        'Terry �ȮɥN��
'        strConsigneeKey = ""
'        strZip = ""
'        strContact = ""
'        strPhone = ""
'        strAddress = ""
'        strShort_name = ""


        
        'Terry �ݽT�{��O��� �ק襤
'        '�N��O�N��I or R
'        If myExCharFilter(Trim(rsMainT20("�q�����O"))) = "�X�h" Or myExCharFilter(Trim(rsMainT20("�q�����O"))) = "�N�h" Then
'            str_Priority = "R"
'        Else
'            str_Priority = "I"
'        End If
        
'        blCustomerMatch = True
        tmp_Rs.Close

        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT20("�q�渹�X"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            tmp_Rs.Close
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
'            If UCase(Right(Trim(rsMainT20("�ܧO")), 2)) = "-C" Then
'                strFacility = "�ըƹF����"
'            ElseIf UCase(Right(Trim(rsMainT20("�ܧO")), 2)) = "-S" Then
'                strFacility = "�ըƹF�n��"
'            Else
            strFacility = "�ըƹF�_��"
'            End If
            
'            If Trim(rsMainT20("�ܧO")) = "" Then strFacility = ""

            strOrderDate = Trim(rsMainT20("�q���"))
            Dim intPointer As Integer
            intPointer = 1
            
            Call Confirm_Recordset_Closed(tmp_Rs)
            
            
'            'Terry �ݽT�{��O��� �ק襤
'            If (Trim(rsMainT20("�q�����O")) = "�X�h" Or Trim(rsMainT20("�q�����O")) = "�N�h") And Len(Trim(rsMainT20("��f��"))) = 0 Then
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT20("�q�渹�X")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
'                Trim(rsMainT20("��f�Ȥ�s��")) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','" & Trim(rsMainT20(["�Ȥ�q��(�q��)"])) & "','" & Trim(rsMainT20("�q��Ƶ�")) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            Else
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT20("�q�渹�X")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & Trim(rsMainT20("��f��")) & "','" & strFacility & "','" & _
                "LTRI03','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','','" & Trim(rsMainT20("�q��Ƶ�")) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            End If
            
            
'            If (Trim(rsMainT20("�q�����O")) = "�X�h" Or Trim(rsMainT20("�q�����O")) = "�N�h") And Len(Trim(rsMainT20("��f��"))) = 0 Then
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT20("�q�渹�X"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT20("��f�Ȥ�s��"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT20("�q��Ƶ�"))) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            Else
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT20("�q�渹�X"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT20("��f��"))) & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT20("��f�Ȥ�s��"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT20("�q��Ƶ�"))) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            End If
            
            
            
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT20("�q�渹�X")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Abs(Val(rsMainT20("�ƶq")))
            
            '�q����Ӹ�Ʒs�W
            If RTrim(rsMainT20("���W��")) = "�c" Or RTrim(rsMainT20("���W��")) = "CS" Then
                '�c��
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Lottable03)" & _
                "select  '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT20("�q�渹�X")) & "','" & Trim(rsMainT20("�~��")) & "','LTRI03'," & intQTY & " * p.casecnt ," & intQTY & " * p.casecnt " & _
                ",'R01','" & strFacility & "','" & Trim(rsMainT20("���W��")) & "','" & Trim(rsMainT20("�ت��x��")) & "' " & _
                "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT20("�~��")) & "' and s.storerkey = 'LTRI03'"
            Else
                '�Ӽ�
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Lottable03) " & _
                "values ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT20("�q�渹�X")) & "','" & Trim(rsMainT20("�~��")) & "','LTRI03'," & _
                "'" & intQTY & "','" & intQTY & "','R01','" & strFacility & "','" & Trim(rsMainT20("���W��")) & "','" & Trim(rsMainT20("�ت��x��")) & "')"
            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT20.MoveNext
Loop

'��s������
'cn.Execute "exec gs_ordersupdate 'LTRI03'", RowsAffect, adExecuteNoRecords

'�ˬd���L���`�q��  �ܧO�w�O���~ es_Checklot06_by_storer '�f�D','�q���ɦW'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT14.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '���^�ǿ��~���q���ơA����excel
'    Recordset2Excel "�t�e�ܧO�P���ӭܧO���Ū��q����", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT20.Enabled = True


'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�ɮ� " & filLocalFileT20.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT20.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ��ɮ�
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT20.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT20.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT20.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT20.FileName, ".", -1)
End If


'�ƥ���FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT20.FileName


Kill strTranFileName
    
filLocalFileT20.Refresh:
Screen.MousePointer = 0: cmdImportT20.Enabled = True: dgMainT20.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT20.Enabled = True: Screen.MousePointer = 0: dgMainT20.Enabled = True
End Sub

Private Sub cmdImportT21_Click()
    If rsMainT21 Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    dgMainT21.Enabled = False: cmdImportT21.Enabled = False
    Dim str_ASNkey As String, int_orderlinenuber As Integer, str_Storerkey As String, strKeycount As String, str_Lottable06 As String
    Dim rsKeycount As New ADODB.Recordset
    Dim bl_Error As Boolean '�O�����p���I���X��
    Dim str_Error As String '�O�����p���I���~�����
    Dim x As Long
    bl_Error = False: str_Error = ""
    
    str_Storerkey = "LCHF01"
    str_ASNkey = ""
    dgMainT21.Enabled = False: cmdImportT21.Enabled = False
    strTranFileName = filLocalFileT21.Path & "\" & filLocalFileT21.FileName
    
    Do While Not rsMainT21.EOF
        str_SQL = "select externorderkey from orders(nolock) where storerkey = '" & str_Storerkey & "' and externorderkey = '" & Trim(rsMainT21("�ռ��渹")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn
        If Not tmp_Rs.EOF Then
            msg_text = "�q�歫��:" & Trim(rsMainT21("�ռ��渹")) & "�A�нT�{��ơA���¡C"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            dgMainT21.Enabled = True: cmdImportT21.Enabled = True
            tmp_Rs.Close
            Exit Sub
        End If
        rsMainT21.MoveNext
    Loop
    
    rsMainT21.MoveFirst
    
    Do While Not rsMainT21.EOF
        '��f����ˬd
        If Len(Trim(rsMainT21("�ռ���"))) = 0 Then
             MsgBox "����渹:" & Trim(rsMainT21("�ռ��渹")) & "���ѳf������ťաA�q����J�פ�!", 16, Me.Caption: dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
        Else
            '�ˬd��f�餣�i�p�󤵤�
            If Format(Trim(rsMainT21("�ռ���")), "YYYYMMDD") < Format(Now, "YYYYMMDD") Then
                If blAdmin = True Then
                
                '�̰��v�����ˬd��f��
                     x = MsgBox("�ռ���p�󤵤�A�A�T�w�n�~���?", vbQuestion + vbYesNo, "�̰��v����f���ˬd") '�������U���O�T�w�άO����
                        If x = 6 Then
                            '�~��
                        Else
                            '���}
                             dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
                        End If
                Else
                
                End If
            End If
        End If
        
        '�q����ˬd
        If Len(Trim(rsMainT21("�ռ���"))) = 0 Then
             MsgBox "�ѳf������ťաA�q����J�פ�!", 16, Me.Caption: dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
        End If
        
        '�ƶq�ˬd
        If Val(rsMainT21("�ƶq")) < 1 Then
            MsgBox "�ƶq�p��1�A" & Trim(rsMainT21("�ռ��渹")) & "-�~���G" & Trim(rsMainT21("���~�N��")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
            Exit Sub
        End If
        
        '������� --�P�_SKU�O�_�s�b
        If InStr(1, Trim(rsMainT21("�~�W")), "�̪O") Then
        Else
            str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT21("���~�N��")) & "' and Storerkey = '" & str_Storerkey & "'"
            
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then
                MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT21("���~�N��")) & ")�A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": cmdImportT21.Enabled = True: Screen.MousePointer = 0
                dgMainT21.Enabled = True: cmdImportT21.Enabled = True
                Exit Sub
            End If
        End If
        '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT21("�ƶq")), ".") <> 0 Then
            str_Error = "����渹:" & Trim(rsMainT21("�ռ��渹")) & "�A�~��:" & Trim(rsMainT21("���~�N��")) & Chr(13) & str_Error
            bl_Error = True
        End If

            
        rsMainT21.MoveNext
    Loop
    rsMainT21.MoveFirst
    
    If bl_Error = True Then
                    msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                    dgMainT21.Enabled = True: cmdImportT21.Enabled = True
                    Exit Sub
    End If
    
    rsMainT21.MoveFirst
    Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long, IntCasecnt As Integer
    Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim strDate As String, strOrderType As String, strSku As String
    Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String
    Tran_Level = cn.BeginTrans: cmdImportT12.Enabled = False: dgMainT12.Enabled = False
    Do While Not rsMainT21.EOF
        If InStr(1, Trim(rsMainT21("�~�W")), "�̪O") Then
            GoTo NextRow1
        End If
        If str_ASNkey <> Trim(rsMainT21("�ռ��渹")) Then
            str_ASNkey = Trim(rsMainT21("�ռ��渹"))
            int_orderlinenuber = 0
            
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = 'LCHF01'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
    '        End If
            '�۲Ūu���«Ƚs
            strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
            strZip = myExCharFilter(Trim(tmp_Rs("zip")))
            strContact = myExCharFilter(Trim(tmp_Rs("contact")))
            strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
            strAddress = myExCharFilter(Trim(tmp_Rs("address")))
            strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
            tmp_Rs.Close
            
            '�������--�P�_�q��O�_���ơA���Ƥ��W�[
            Call Confirm_Recordset_Closed(tmp_Rs)
            str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT21("�ռ��渹"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '�R��' "
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF Then
    
                '���q�渹�X
                str_SQL = "select isnull(max(orderkey),0) from orders"
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
                If strOrderKeyS = "" Then
                    strOrderKeyS = str_Orderkey
                End If
                
                tmp_Rs.Close
                
                Dim intPointer As Integer
                intPointer = 1
                
            
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                            "VALUES ('" & str_Orderkey & "','" & RTrim(rsMainT21("�ռ��渹")) & "','" & "RC" & "','" & str_Storerkey & "',convert(char(10),'" & Trim(rsMainT21("�ռ���")) & "',111)," & _
                            "convert(char(10),'" & Trim(rsMainT21("�ռ���")) & "',111) ,'" & strFacility & "','LCHF01','" & strShort_name & "','" & strContact & "','','','" & strPhone & "'," & _
                            "'" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & _
                            "','','" & "" & "','" & filLocalFileT12.FileName & "','','" & User_id & "','" & User_id & "','','','" & "RC" & "','" & "" & "') "
    
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                int_Order = int_Order + 1
                int_orderlinenuber = 1
            Else
                tmp_Rs.Close
                
                '�q�歫��
                Call FTPlog("�q�歫��" & str_SQL)
                '��������
                strReOrderkey = strReOrderkey & Trim(rsMainT21("�ռ��渹")) & "','"
                GoTo NextRow1
                
            End If
        End If

        If Trim(rsMainT21("���X�ܮw�W��")) = "�}�~��" Then
            str_Lottable06 = "R01"
        End If
        '�g�J��
        str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,otherUOM)" & _
                    "select '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Format(int_orderlinenuber, "00000") & "','" & RTrim(rsMainT21("�ռ��渹")) & "','" & Trim(rsMainT21("���~�N��")) & "','" & str_Storerkey & "'," & _
                    "cast('" & Trim(rsMainT21("�ƶq")) & "' as Int) * p.casecnt,cast('" & Trim(rsMainT21("�ƶq")) & "' as Int) * p.casecnt,CONVERT(CHAR(8),CONVERT(DATETIME,'" & Trim(rsMainT21("�Ƶ�")) & "',111),112),'R01','" & strFacility & "','',''" & _
                    "FROM " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey where s.sku = '" & Trim(rsMainT21("���~�N��")) & "' and s.storerkey = '" & str_Storerkey & "'"

        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '��spackkey
        str_SQL = "update orderdetail " & _
        "Set orderdetail.packkey = sku.packkey " & _
        "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
        "where orderkey = '" & str_Orderkey & "' "
        
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        int_orderlinenuber = int_orderlinenuber + 1
        
        int_OrderLine = int_OrderLine + 1
NextRow1:

        rsMainT21.MoveNext

        Loop
        'Close RecordSet
        rsMainT21.Close: Set rsMainT21 = Nothing
                
                
cn.Execute "exec gs_ordersupdate 'LCHF01'", RowsAffect, adExecuteNoRecords


cn.CommitTrans: Tran_Level = 0: dgMainT21.Enabled = True


'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�ɮ� " & filLocalFileT12.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT21.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ��ɮ�
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT21.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT21.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT21.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT21.FileName, ".", -1)
End If

'�ƥ���FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT21.FileName


Kill strTranFileName
    
filLocalFileT21.Refresh:
Screen.MousePointer = 0: cmdImportT21.Enabled = True: dgMainT21.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT21.Enabled = True: Screen.MousePointer = 0: dgMainT21.Enabled = True

End Sub


Private Sub cmdImportT22_Click()
Dim bl_Error As Boolean '�O�����p���I���X��
Dim str_Error As String '�O�����p���I���~�����
Dim x As Integer '�O�����U�T�w�άO�����A�M�w�O�_��s���q��urgent_mark
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

str_Storerkey = "LYFY09"    '�f�D

If rsMainT22 Is Nothing Then Exit Sub
If rsMainT22.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT22.Enabled = False: cmdImportT22.Enabled = False
strTranFileName = filLocalFileT22.Path & "\" & filLocalFileT22.FileName

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT22.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT22.Enabled = True: cmdImportT22.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT22.RecordCount = 0 Or rsMainT22 Is Nothing Then
Else
rsMainT22.MoveFirst
Do While Not rsMainT22.EOF
    '��f����ˬd
    If Len(Trim(rsMainT22("�w�p���h��"))) = 0 Then
    Else
        arrTmp = Split(Trim(rsMainT22("�w�p���h��")), "/")
        If Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "�w�p���h��p�󤵤�A�q����J�פ�!", 16, Me.Caption: dgMainT22.Enabled = True: cmdImportT22.Enabled = True: Exit Sub
    End If
    
    '�ƶq�ˬd
    If Val(rsMainT22("�ƶq")) < 1 Then
        MsgBox "�ƶq�p��1�A" & Trim(rsMainT22("�h�f�渹")) & "-�~���G" & Trim(rsMainT22("�~��")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT22.Enabled = True: cmdImportT22.Enabled = True: Exit Sub
        Exit Sub
    End If
    
    '������� --�P�_SKU�O�_�s�b
        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT22("�~��")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT22("�~��")) & ")�A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": cmdImportT22.Enabled = True: Screen.MousePointer = 0
            dgMainT22.Enabled = True: cmdImportT22.Enabled = True
            Exit Sub
        End If
        
       '������� --�P�_consigneekey�O�_�s�b
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT22("�橱�N��")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT22("�橱�N��")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT22.Enabled = True: Screen.MousePointer = 0
            dgMainT22.Enabled = True: cmdImportT22.Enabled = True
            Exit Sub
        End If
        
    '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT22("�ƶq")), ".") <> 0 Then
            str_Error = "�q�渹�X:" & Trim(rsMainT22("�h�f�渹")) & "�A�~��:" & Trim(rsMainT22("�~��")) & Chr(13) & str_Error
            bl_Error = True
        End If
    rsMainT22.MoveNext
Loop
rsMainT22.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT22.Enabled = True: cmdImportT22.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT22.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'�}�l�פJ
If rsMainT22 Is Nothing Then GoTo next18
If rsMainT22.RecordCount = 0 Then GoTo next18

'���̫�Ȥ�s��
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LAPP01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close


Do While Not rsMainT22.EOF
    DoEvents: DoEvents
    
'    �������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If Trim(rsMainT22("�~��")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow17
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT22("�h�f�渹"))) Then
        strOrderNo = UCase(Trim(rsMainT22("�h�f�渹")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�w�g�ˬd�L�A�h�����A������X�Ȥ�D�ɤ��� �Ȥ�s���A�l���ϸ��A�s���H�A�q�ܡA�a�}
        str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT22("�橱�N��"))) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '�L���Ȥ�W�٫h�s�W
'            strConsigneeKey = myExCharFilter(Trim(rsMainT22("SINGLE_SHOP_CODE")))
'
'            '�s�W�Ȥ�D��
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('" & Str_Storerkey & "','" & myExCharFilter(Trim(rsMainT22("POSTAL_CODE"))) & "','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT22("CUSTOMER_NAME"))) & "','" & myExCharFilter(Trim(rsMainT22("SUPPLIER_CODE"))) & "','" & myExCharFilter(Trim(rsMainT22("CONTACT"))) & "','" & myExCharFilter(Trim(rsMainT22("TELEPHONE"))) & "','" & myExCharFilter(Trim(rsMainT22("REQUEST_ADDRESS"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'            '�����s�W���Ȥ�s��
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
''            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = '" & Str_Storerkey & "' and full_name = '" & myExCharFilter(Trim(rsMainT22("CUSTOMER_NAME"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT22("REQUEST_ADDRESS"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
                    

'                If rsTmp.EOF Then
'                    '�p���H�B�q�ܻP��f�a�}����
'                    strConsigneeKey = myExCharFilter(Trim(rsMainT22("SINGLE_SHOP_CODE")))
'
'                    '�s�W�Ȥ�D��
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'                    " values('" & Str_Storerkey & "','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT22("SINGLE_SHOP_CODE"))) & "','" & myExCharFilter(Trim(rsMainT22("CUSTOMER_NAME"))) & "','" & myExCharFilter(Trim(rsMainT22("SUPPLIER_CODE"))) & "','" & myExCharFilter(Trim(rsMainT22("CONTACT"))) & "','" & myExCharFilter(Trim(rsMainT22("TELEPHONE"))) & "','" & myExCharFilter(Trim(rsMainT22("REQUEST_ADDRESS"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'                    '�����s�W���Ȥ�s��
'                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'                Else
'
'                    '�۲Ūu���«Ƚs
'                    strConsigneeKey = myExCharFilter(Trim(rsMainT22("POSTAL_CODE")))
'                    blCustomerMatch = True
''
'                End If
''            rsTmp.Close
'        End If

        '�۲Ūu���«Ƚs
        
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close

        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT22("�h�f�渹"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
'            If Right(UCase(Trim(rsMainT22("DC�N��"))), 1) = "C" Then strFacility = "�ըƹF����"
'            If Right(UCase(Trim(rsMainT22("DC�N��"))), 1) = "S" Then strFacility = "�ըƹF�n��"

            
            arrTmp = Split(Trim(rsMainT22("�w�p���h��")), "/")
            If Len(Trim(rsMainT22("�w�p���h��"))) = 0 Then
                cn.RollbackTrans: Tran_Level = 0
                msg_text = "�q�渹�X:" & Trim(rsMainT22("�h�f�渹")) & "�A�~��:" & Trim(rsMainT22("�~��")) & Chr(13) & "��ƨS���w�p���h��I�ЦV�t�ӽT�{"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT22.Enabled = True: cmdImportT22.Enabled = True
                Exit Sub
            Else
                strDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            End If
            
'            arrTmp = Split(Trim(rsMainT22("ORDERED_DATE")), "/")
'            strOrderDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            
            strOrderDate = Format(Now, "YYYY/MM/DD")
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT22("�h�f�渹"))) & "','R','" & str_Storerkey & "','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT22("�Ƶ�"))) & "','" & filLocalFileT22.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT22("�h�f�渹"))) & "') "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT22("�h�f�渹")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            
'            lngCasecnt = 1
            
'            '��촫��
'            If Left(myExCharFilter(Trim(rsMainT22("���"))), 1) = "�c" Then
'
                '���c�]�ഫ�v
'                str_SQL = "select p.casecnt, p.innerpack from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey where s.storerkey = 'LAPP01' and s.sku = '" & myExCharFilter(Trim(rsMainT22("�ӫ~�s�X"))) & "'"
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'                lngCasecnt = tmp_Rs("casecnt")      '�j���J��
'                lngInnerpack = tmp_Rs("innerpack")  '�����J��
'                tmp_Rs.Close
'
'            End If
            
            
'            If UCase(rsMainT22("SAP�q�f���")) = "BDL" Then intQTY = Val(rsMainT22("�ӫ~�q�ʼƶq")) * lngInnerpack  '�����J�ơA��Xpack��ƪ�
'            If UCase(rsMainT22("SAP�q�f���")) = "KAR" Then intQTY = Val(rsMainT22("�ӫ~�q�ʼƶq")) * lngCasecnt  '�j���J�ơA

            
            intQTY = Val(rsMainT22("�ƶq"))
            
            strLot06 = "R01"
            
'            '��trp19m�ܧO��Ӫ�
'            str_SQL = "select bestlot06 from trp19m where storerkey = 'LAPP01' and storerlot06 = '" & rsMainT22("DC�N��") & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'            If tmp_Rs.EOF Then
'            '�S�����������ܧO
'            '�u�έq��W���ܧO
'                strLot06 = rsMainT22("DC�N��") '�ݽT�{
'            Else
'            '�����������ܧO
'                strLot06 = Trim(tmp_Rs("bestlot06"))
'            End If
'
'            tmp_Rs.Close
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT22("�h�f�渹"))) & "','" & myExCharFilter(Trim(rsMainT22("�~��"))) & "','" & str_Storerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT22.MoveNext
Loop

next18:

'����gs_ordersupdate   �ΫȤ�D�ɧ�s�q����

cn.Execute "exec gs_ordersupdate 'LYFY09'", RowsAffect, adExecuteNoRecords

'�ˬd���L���`�q��  �ܧO�w�O���~ es_Checklot06_by_storer '�f�D','�q���ɦW'
str_SQL = "exec es_Checklot06_by_storer '" & str_Storerkey & "','" & filLocalFileT22.FileName & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '���^�ǿ��~���q���ơA����excel
    Recordset2Excel "�t�e�ܧO�P���ӭܧO���Ū��q����", tmp_Rs
End If

tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�ɮ� " & filLocalFileT22.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT22.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ���FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT22.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT22.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT22.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT22.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT22.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT22.Refresh:
Screen.MousePointer = 0: cmdImportT22.Enabled = True: dgMainT22.Enabled = True
Exit Sub

err_Handle:
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT22.Enabled = True: Screen.MousePointer = 0: dgMainT22.Enabled = True
End Sub

Private Sub cmdOpenFile_Click()

'Dim strFullFileName As String, strFileName As String, ExecuteDOSCommand
'
'cmdOpenFile.Enabled = False
'strFullFileName = filLocalFile.Path & "\" & filLocalFile.FileName
'If Len(Trim(filLocalFile.FileName)) = 0 Then Exit Sub
'strFileName = filLocalFile.FileName
'If UCase(Left(strFullFileName, 1)) <> "T" Then cmdOpenFile.Enabled = True: MsgBox "�Х�T:�Ϻо��פJ!", 64, Me.Caption: Exit Sub
'
'If strFullFileName = "" Then Exit Sub
'
'On Error GoTo err_Handle
'
''�ƻs�ɮ�
'If Dir("C:\LTKK01\Document", vbDirectory) = "" Then MkDirs "C:\LTKK01\Document"
'FileCopy strFullFileName, "C:\LTKK01\Document\" & strFileName
'
''�ƥ���FTP
'If Dir("O:\Kirin\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\Kirin\OrdersBackup"
'FileCopy strFullFileName, "O:\Kirin\OrdersBackup\" & strFileName
'
'MsgBox "�ɮ�(" & strFileName & ")" & vbCrLf & "�w�ƻs��G" & vbCrLf & "C:\LTKK01\Document\" & vbCrLf & "O:\LTKK01\OrdersBackup\", 64, Me.Caption
'
''��Ʈw�O��
'str_SQL = "insert into gt_filelog(storerkey,filename,filedate,filelen,addwho) values('LTKK01','" & strFileName & "','" & Format(FileDateTime(strFullFileName), "YYYYMMDD hh:mm:ss") & "','" & FileLen(strFullFileName) & "','" & User_id & "')"
'cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String
''�����ɸ��
'Call ReDim_Recordset(tmp_Rs)
'Call Confirm_Recordset_Closed(tmp_Rs)
'
'str_SQL = "select �ɮ׮ɶ� = filename , �ɮ׮ɶ� = filedate, ���ɮɶ� = gettime, �t���ɶ� = convert(char(20),gettime - filedate,20) from gv_FileTime where filename = '" & strFileName & "' "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then
'    strTextbody = "�t�Ψ��ɡG" & strFileName & "-�ɮ׮ɶ��G" & tmp_Rs("�ɮ׮ɶ�") & " �ɮפj�p�G" & FileLen(strFullFileName) & " �ɶ��t�G" & ((Mid(tmp_Rs("�t���ɶ�"), 9, 2) - 1) * 24) + Mid(tmp_Rs("�t���ɶ�"), 12, 2) & Mid(tmp_Rs("�t���ɶ�"), 14, 6)
'Else
'    strTextbody = "�t�Ψ��ɡG" & strFileName & "-�ɮ׮ɶ��G�L �ɮפj�p�G" & FileLen(strFullFileName) & " �ɶ��t�G�L"
'End If
'
'tmp_Rs.Close
'
''LTKK01�����ɦ۰� Mail �q��
''�������w
'strFrom = "Tkedi@bestlog.com.tw"
'strTo = "jack@mail.kirin.com.tw,irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;simon@mail.kirin.com.tw"
'strCC = "Tkedi@bestlog.com.tw"
''strTo = "eric_huang@bestlog.com.tw"
''strCC = ""
'strBCC = strBCC
'strSubject = "���ɳq��(" & strFileName & ")"
'strTextbody = strTextbody
'strEmailID = "tkedi"
'strEmailPW = "tkedibl01"
'strAlways = "NO"
'
''�ǰe�l��
'Dim objEmail As Object
'Set objEmail = CreateObject("CDO.Message")
'
'objEmail.From = strFrom
'objEmail.To = strTo
'objEmail.CC = strCC   ' �ƥ�
'objEmail.BCC = strBCC ' �K��ƥ�
'objEmail.Subject = RTrim(strSubject)
'objEmail.TextBody = strTextbody
'objEmail.AddAttachment strAddAttachment
'
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
''SMTP ���A���ݭn���Ү�
'If Len(RTrim(strEmailID)) > 0 Then
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
'End If
'objEmail.Configuration.Fields.Update
'objEmail.Send
'
'Set objEmail = Nothing
'
'MsgBox "���ɳq��Email����", 64, strFileName
'
'ExecuteDOSCommand = Shell("cmd /c start C:\LTKK01\Document\" & strFileName, 0)
'
'
''�R���ӷ��ɮ�
'Kill strFullFileName
'filLocalFile.Refresh
'cmdOpenFile.Enabled = True

Exit Sub

err_Handle:
cmdOpenFile.Enabled = True
Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)

End Sub

Private Sub CmbStartT17_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String, S As String, Str_Sku As String
Dim bl_Check As Boolean '�ˬd�פJ�����~���L�X�{�b�`��~�S���Nstop
bl_Check = True
S = "": Str_Sku = ""
'S�O���W�@�����f�渹,Str_sku�O���W�@���~��,�p�G�ťիh�a�W�@��

On Error GoTo err_Handle
SSTab3.Tab = 1: SSTab3.Enabled = False: CmbStartT17.Enabled = False: cmdImportT17.Enabled = False

Call DB_Connect_Self(cn_string) '�إ߷s�s�u

'�T�{���|�O�_�a"\"
If Right(filLocalFileT17.Path, 1) = "\" Then
    strFilePath = filLocalFileT17.Path
Else
    strFilePath = filLocalFileT17.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "��f�渹" & Chr(9) & "�X�f�Ƶ�" & Chr(9)

If Right(filLocalFileT17.Path, 1) <> "\" Then
    strFilePath = filLocalFileT17.Path & "\"
Else
    strFilePath = filLocalFileT17.Path
End If

Set rsMainT17_1 = New ADODB.Recordset

Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT17.FileName)   '���}���|
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "RawHeader" Then .Sheets(i).Select: Exit For '��w�u�@��
    Next
    
    If (.ActiveSheet.Name) <> "RawHeader" Then MsgBox "�䤣��RawHeader�u�@��!!", 16, "�}���ɮפ���": GoTo endsub
    
    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "��f�渹" Then k = i: Exit For
    Next i
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT17_1 = Nothing: GoTo endsub
    
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "RawHeader�u�@��A�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT17_1.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT17_1.CursorType = adOpenKeyset
    rsMainT17_1.LockType = adLockOptimistic
    rsMainT17_1.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT17_1.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT17_1(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT17_1.Update
    k = k + 1
    Loop
    
    If rsMainT17_1.RecordCount > 0 Then rsMainT17_1.MoveFirst

End With


Set dgMainT17_1.DataSource = rsMainT17_1

'If rsMainT17_1 Is Nothing Then
'
'    MsgBox "�d�L���!", 64, "Excel2Recordset"
'
'Else
'    SetDataGridColWidth Me.Caption, rsMainT17_1 '�]�w��e
'End If

'�p�G�S���X�f�`��h����
If rsMainT17_1 Is Nothing Then MsgBox "�䤣��RawHeader�u�@��!!", 16, "�}���ɮפ���": SSTab3.Enabled = True: CmbStartT17.Enabled = True: Exit Sub
If rsMainT17_1.EOF Then MsgBox "�䤣��RawHeader�u�@��!!", 16, "�}���ɮפ���": SSTab3.Enabled = True: CmbStartT17.Enabled = True: Exit Sub

'/////////////////////////////////////////////////////////////////////////////////�פJFormat////////////////////////////////////////////////////////////////////////////////
SSTab3.Tab = 0
'�إ����W�ٰ}�C
strFieldName = "DC�N��" & Chr(9) & "�Ȥ�N��" & Chr(9) & "EXE���s��" & Chr(9) & "SAP DN NO." & Chr(9) & "������O" & Chr(9) & "��ڲ��ͤ�" & Chr(9) & "�w�p�X�f��" & Chr(9) & "����" & Chr(9) & "�ӫ~�s�X" & Chr(9) & "�ӫ~�q�ʼƶq" & Chr(9) & "�P��O" & Chr(9) & "�ӫ~�̤p�ƶq" & Chr(9) & "�Ȥ�i��" & Chr(9) & _
              "�����i�B" & Chr(9) & "��ڥX�f�ƶq" & Chr(9) & "��ڥX�f�ܧO" & Chr(9) & "�|�O" & Chr(9) & "�妸" & Chr(9) & "����˳f���" & Chr(9) & "�f�D" & Chr(9) & "�e�f�a�}" & Chr(9) & "�P���´" & Chr(9) & "��~��" & Chr(9) & "�~�Ȳժ�" & Chr(9) & "���w�渹" & Chr(9) & "SAP�q�f���" & Chr(9) & "��]" & Chr(9) & "�Ȥ�W��" & Chr(9) & "�ƪ`" & Chr(9) & "�Ȥ�q���O" & Chr(9)

Set rsMainT17 = New ADODB.Recordset

'Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
'    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '���}���|
    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "Format" Then .Sheets(i).Select: Exit For '��w�u�@��
    Next
    
    If (.ActiveSheet.Name) <> "Format" Then MsgBox "�䤣��Format�u�@��!!", 16, "�}���ɮפ���": GoTo endsub

    'k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '�ѲĤG�C�}�l�פJ
    End If
    
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "DC�N��" Then k = i: Exit For
    Next i
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    If UBound(arrTmp) < 1 Then Set rsMainT17 = Nothing: GoTo endsub
    
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "Format�u�@��A�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsMainT17.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT17.CursorType = adOpenKeyset
    rsMainT17.LockType = adLockOptimistic
    rsMainT17.Open
    
    rsMainT17_1.MoveFirst
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0   '�D�Ǽƶq���ŭȫh����
'    If RTrim(.Cells(k + 1, 6)) = "60400119" Then '�ư��B�O
'    Else
        rsMainT17.AddNew
            For j = 1 To UBound(arrTmp)
                If j = UBound(arrTmp) Then    '�ƪ`���
                    bl_Check = True
                    rsMainT17_1.MoveFirst
                    Do While Not rsMainT17_1.EOF
                        If Trim(rsMainT17_1("��f�渹").Value) = Trim(rsMainT17("SAP DN NO.").Value) Then rsMainT17("�ƪ`").Value = Trim(rsMainT17_1("�X�f�Ƶ�").Value): bl_Check = False: Exit Do
                        rsMainT17_1.MoveNext
                    Loop
                    rsMainT17(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
                    If bl_Check = True Then MsgBox "RawHeader�d�L�q�渹�X:" & Trim(rsMainT17("SAP DN NO.").Value) & "���X�f�Ƶ����!", 64, "Format�פJ����": SSTab3.Enabled = True: CmbStartT17.Enabled = True: GoTo endsub
                Else
                    rsMainT17(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
                End If
            Next j
        rsMainT17.Update
'    End If
    k = k + 1
    
    Loop
    
    If rsMainT17.RecordCount > 0 Then rsMainT17.MoveFirst

End With
    
Set dgMainT17.DataSource = rsMainT17

rsMainT17.Sort = "[SAP DN NO.]"

If rsMainT17 Is Nothing Then
    MsgBox "Format�d�L���!", 64, "Excel2Recordset"
Else
    SetDataGridColWidth Me.Caption, dgMainT17
    MsgBox "���q���ɦ@�פJ:" & Chr(13) & "Format:" & rsMainT17.RecordCount & "������" & Chr(13) & "" & _
                                          "RawHeader:" & rsMainT17_1.RecordCount & "������" & Chr(13) & "" & _
                                          "�нT�{���ƬO�_���T!", 64, "�����@�q��}��"
    cmdImportT17.Enabled = True
End If

'�p�G���X�f�`��A��L�T�Ӥu�@��S����ƫh���ܡA������
'If rsMainT16_1.RecordCount = 0 And rsMainT16_2.RecordCount = 0 And rsMainT16_3.RecordCount = 0 Then MsgBox "���q��L�Ӷ���ơA�нT�{���q��O�_���T!", vbCritical, "���ɤs�q��}��"

endsub:
SSTab3.Enabled = True: CmbStartT17.Enabled = True:
MyXlsApp.Quit: Set MyXlsApp = Nothing

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
SSTab3.Enabled = True: CmbStartT17.Enabled = True
End Sub


Private Sub cmdOpenFilesT15_Click()
If Trim(cboSheetT15) = "" Then Exit Sub
On Error GoTo err_Handle
dgMainT15.Enabled = False: cmdImportT15.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, Str_Sku As String
'�T�{���|�O�_�a"\"
If Right(filLocalFileT15.Path, 1) = "\" Then
    strFilePath = filLocalFileT15.Path
Else
    strFilePath = filLocalFileT15.Path & "\"
End If
'�إ����W�ٰ}�C
strFieldName = ""
If Right(filLocalFileT15.Path, 1) <> "\" Then
    strFilePath = filLocalFileT15.Path & "\"
Else
    strFilePath = filLocalFileT15.Path
End If
Set rsMainT15 = New ADODB.Recordset
strSheetName = cboSheetT15
Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, strSheetName, strFieldName, rsMainT15)
rsMainT15.MoveFirst

Set dgMainT15.DataSource = rsMainT15

If rsMainT15 Is Nothing Then
    MsgBox "�d�L���!", 64, "Excel2Recordset"
Else
    'SetDataGridColWidth Me.Caption, dgMainT15
    MsgBox "���u�@��@ " & rsMainT15.RecordCount & "����ơA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    cmdImportT15.Enabled = True
End If
rsMainT15.Sort = "�X�f�渹"
rsMainT15.MoveFirst
dgMainT15.Enabled = True: cmdImportT15.Enabled = True

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdOpenFilesT20_Click()
If Trim(cboSheetT20) = "" Then Exit Sub
On Error GoTo err_Handle
dgMainT20.Enabled = False: cmdImportT20.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, Str_Sku As String
'�T�{���|�O�_�a"\"
If Right(filLocalFileT20.Path, 1) = "\" Then
    strFilePath = filLocalFileT20.Path
Else
    strFilePath = filLocalFileT20.Path & "\"
End If
'�إ����W�ٰ}�C
strFieldName = ""
If Right(filLocalFileT20.Path, 1) <> "\" Then
    strFilePath = filLocalFileT20.Path & "\"
Else
    strFilePath = filLocalFileT20.Path
End If
Set rsMainT20 = New ADODB.Recordset
strSheetName = cboSheetT20
Call Excel2Recordset(strFilePath & filLocalFileT20.FileName, strSheetName, strFieldName, rsMainT20)
rsMainT20.MoveFirst

Set dgMainT20.DataSource = rsMainT20

If rsMainT20 Is Nothing Then
    MsgBox "�d�L���!", 64, "Excel2Recordset"
Else
    'SetDataGridColWidth Me.Caption, dgMainT20
    MsgBox "���u�@��@ " & rsMainT20.RecordCount & "����ơA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    cmdImportT20.Enabled = True
End If

rsMainT20.Sort = "�q�渹�X"
rsMainT20.MoveFirst
dgMainT20.Enabled = True: cmdImportT20.Enabled = True

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdOpenFilesT21_Click()
If Trim(cboSheetT21) = "" Then Exit Sub
On Error GoTo err_Handle
dgMainT21.Enabled = False: cmdImportT21.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, Str_Sku As String
'�T�{���|�O�_�a"\"
If Right(filLocalFileT21.Path, 1) = "\" Then
    strFilePath = filLocalFileT21.Path
Else
    strFilePath = filLocalFileT21.Path & "\"
End If
'�إ����W�ٰ}�C
strFieldName = ""
If Right(filLocalFileT21.Path, 1) <> "\" Then
    strFilePath = filLocalFileT21.Path & "\"
Else
    strFilePath = filLocalFileT21.Path
End If
Set rsMainT21 = New ADODB.Recordset
strSheetName = cboSheetT21
Call Excel2Recordset(strFilePath & filLocalFileT21.FileName, strSheetName, strFieldName, rsMainT21)
rsMainT21.MoveFirst

Set dgMainT21.DataSource = rsMainT21

If rsMainT21 Is Nothing Then
    MsgBox "�d�L���!", 64, "Excel2Recordset"
Else
    'SetDataGridColWidth Me.Caption, dgMainT21
    MsgBox "���u�@��@ " & rsMainT21.RecordCount & "����ơA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    cmdImportT21.Enabled = True
End If
rsMainT21.Sort = "�ռ��渹"
rsMainT21.MoveFirst

dgMainT21.Enabled = True: cmdImportT21.Enabled = True

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

'Private Sub cmdUpload3_Click()
'On Error GoTo err_Handle
'
'If cmdLogOff3.Enabled = False Then MsgBox "�Х��n�J���A���I", 64, Me.Caption: Exit Sub
'
'cmdUpload3.Enabled = False
'
'If int3Ready(True) = True Then
'
'    int3.Execute , "Put " & Chr(34) & filLocal3.Path & "\" & filLocal3.FileName & Chr(34) & " ""XRSLUPL.TXT"""
'    lblStatus3 = "�W�Ǥ��еy��...."
'End If
'
'    Do While int3.StillExecuting
'        DoEvents: DoEvents: DoEvents
'    Loop
'
'    lstRemoteFile3.Clear
'    int3.Execute , "DIR"
'
'    Do While int3.StillExecuting
'    DoEvents: DoEvents: DoEvents
'    Loop
'
'    '�W�ǽT�{
'    For i = 0 To lstRemoteFile3.ListCount - 1
'        If lstRemoteFile3.List(i) = "XRSLUPL.TXT" Then
'            '�W�ǧ����R�������W���ɮ�
'            Kill filLocal3.Path & "\" & "XRSLUPL.TXT"
'            lblStatus3 = "�ɮפW�ǧ����I"
'            filLocal3.Refresh
'            GoTo Step1
'        End If
'    Next i
'
'    lblStatus3 = "�W�ǥ��ѡI"
'    MsgBox "�ɮפW�ǥ��ѡA�Э��s�W�ǡI", 64, "Error"
'Step1:
'
'cmdUpload3.Enabled = True
'Exit Sub
'
'err_Handle:
'    Dim tmpString As String
'    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'    CreateErrorLog Me.Name & "-�W��", Me.Caption, "cmdUpload3_Click", tmpString
'    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'End Sub

Private Sub Command1_Click()
    Dim i As Double
    '�Ȥ��ϥ�
    If ITCReady(True) = True Then
        'Check that they are not recieving a folder
        If Right(lstRemoteFile.Text, 1) = "/" Then
            MsgBox lstRemoteFile.Text & " is a folder and cannot be sent.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        'Check that the file does not already exist on the computer, if it does exit sub
        For i = 0 To filLocalFile.ListCount
            If lstRemoteFile.Text = filLocalFile.List(i) Then
                MsgBox "�ɮ� " & Right(lstRemoteFile.Text, 18) & " �w�s�b.", vbInformation + vbOKOnly, "Recieve"
                Exit Sub
            End If
        Next i
        str_file = Trim(Right(lstRemoteFile.Text, 18))
        ITC.Execute , "GET " & Chr(34) & str_file & Chr(34) & " " & Chr(34) & filLocalFile.Path & "\" & str_file & Chr(34)
        lblStatus = "�U�����еy��...."
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        filLocalFile.Refresh
        lblStatus = "�w�s�u"
    End If
    
End Sub

Private Sub cmd_import_Click()
    'TK�q��@���q��2�����Ӫ��Ƶ��p�G���P�A�|������i�q��i��פJ
    
    Dim strTranFileName As String, strFileName As String, strRePoOrderkey As String
    '�}�l�פJ�ɮ�
    
    strTranFileName = filLocalFile.Path & "\" & filLocalFile.FileName
    If Len(Trim(filLocalFile.FileName)) = 0 Then Exit Sub
    If UCase(Left(strTranFileName, 1)) <> "T" Then MsgBox "�Х�T�Ϻо��פJ!", 64, Me.Caption: Exit Sub
    strFileName = filLocalFile.FileName
    If strTranFileName = "" Then Exit Sub
    
    On Error GoTo err_Handle
    If FileLen(strTranFileName) = 0 Then MsgBox "�ɮפj�p = 0 , �ɦW: " & filLocalFile.FileName, vbOKOnly + vbInformation, Me.Caption: Exit Sub
      
    cmd_Import.Enabled = False: Screen.MousePointer = 11: dg_CustInv.Enabled = False
    Set dg_CustInv.DataSource = Nothing
    DoEvents: DoEvents
    
    Dim strRow As String    'Ū���C�@���r
    Dim strField As String  'Ū���C�ӰϹj�����
    Dim intPointer As Integer
    Set rs_Src = New Recordset
    
    With rs_Src
        .Fields.Append "�q�渹�X", adChar, 30, adFldUpdatable
        .Fields.Append "�q�涵��", adChar, 10, adFldUpdatable
        .Fields.Append "��f�渹", adChar, 20, adFldUpdatable
        .Fields.Append "�q����", adChar, 8, adFldUpdatable
        .Fields.Append "�q�����O", adChar, 10, adFldUpdatable
        .Fields.Append "�a�}�O", adChar, 30, adFldUpdatable
        .Fields.Append "�Ƹ�", adChar, 35, adFldUpdatable
        .Fields.Append "����W��", adChar, 100, adFldUpdatable
        .Fields.Append "�̤p���ƶq", adDouble, adFldUpdatable
        .Fields.Append "�q��ƶq", adDouble, adFldUpdatable
        .Fields.Append "�q����", adChar, 10, adFldUpdatable
        .Fields.Append "���", adDouble, adFldUpdatable
        .Fields.Append "��f���", adChar, 8, adFldUpdatable
        .Fields.Append "�Ȥ�渹", adChar, 60, adFldUpdatable
        .Fields.Append "�Ƶ�", adChar, 255, adFldUpdatable
        .Fields.Append "�K��", adChar, 10, adFldUpdatable
        .Fields.Append "�ܧO", adChar, 18, adFldUpdatable
        .Fields.Append "�x��", adChar, 18, adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        
        Dim arrTmp, arrTmp1, strLineTmp As String, intArrline As Integer, i As Double
        
        '�}���ɮ�
        Open strTranFileName For Input As #1
        Do While Not EOF(1)
        '�פJ�ɮ�
            Line Input #1, strLineTmp '����Ʀ�
            arrTmp = Split(strLineTmp, Chr(10)) '��������
            If UBound(arrTmp) = -1 Then GoTo NextStep
            If UBound(arrTmp) > 0 Then
                For intArrline = 0 To UBound(arrTmp) - 1
                    '�����ƦA�����
                    arrTmp1 = Split(arrTmp(intArrline), ",")
                    .AddNew
                        For i = 0 To .Fields.Count - 1
                            .Fields(i) = Trim(arrTmp1(i))
                        Next i
                    .Update
                    
                Next intArrline
            Else '���������
                    arrTmp1 = Split(arrTmp(intArrline), ",") '�����
                    .AddNew
                        For i = 0 To .Fields.Count - 1
                            If i = 15 Then
                                .Fields(i) = GetWord(Trim(arrTmp1(i)), 1, 10) & "" '�K��od.updatesource��10�X
                            Else
                                .Fields(i) = Trim(arrTmp1(i)) & ""
                            End If
                        Next i
                    .Update
            End If
NextStep:
        Loop
        
        Close #1
        .MoveFirst
        .Sort = "�q�渹�X,�a�}�O,��f���,�q�����O,�Ƶ�,�q�涵��"
End With
 
 '�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select updatesource from orders where storerkey = 'LTKK01' and rtrim(updatesource)='" & filLocalFile.FileName & "' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True: Exit Sub

'�ˬd�q���ƬO�_���T
rs_Src.MoveFirst
Do While Not rs_Src.EOF
    '��f����ˬd
    If Trim(rs_Src("��f���")) < Format(Now, "YYYYMMDD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True: Exit Sub
    
    '�f�D�f�� edit by eric 20140923@�ˬd���L�~���A���K��s����n�Ϊ��~���C
    str_SQL = "select sku from gv_skuxpack where storerkey = 'LTKK01' and (storersku = '" & Trim(rs_Src("�Ƹ�")) & "' or sku = '" & Trim(rs_Src("�Ƹ�")) & "') "
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
        MsgBox "�q��o�{�s�Ƹ� (" & Trim(rs_Src("�Ƹ�")) & " ) " & Trim(rs_Src("����W��")) & "�A�q����J�פ�!!"
        tmp_Rs.Close
        Exit Sub
    End If
    'if�Ƹ�����>20�h�ϥθ�Ʈw��SKU�A�_�h�����ϥέq��W���Ƹ�
    If Len(Trim(rs_Src("�Ƹ�"))) > 20 Then rs_Src("�Ƹ�") = tmp_Rs("sku")
    
'    '�ˬd�Ȥ�s��
'    str_SQL = "select consigneekey from trp01m where storerkey = 'LTKK01' and len(rtrim(consigneekey))>5 and substring(rtrim(consigneekey),5,20) = '" & Trim(rs_Src("�a�}�O")) & "'"
'
'    Call Confirm_Recordset_Closed(rsMainTK)
'    rsMainTK.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If rsMainTK.EOF Then
'        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
'        MsgBox "�q��o�{�s�Ȥ�渹:" & Trim(rs_Src("�a�}�O")) & Chr(13) & "�Х��s���Ȥ�A��t�Ϋإ߲Ĥ��X�}�l��:" & Trim(rs_Src("�a�}�O")) & "���Ȥ�渹" & Chr(13) & "EX: XXXX" & Trim(rs_Src("�a�}�O")) & Chr(13) & "�гs���Ȥ�A�s�W�Ȥ�D�ɸ��!!", vbCritical, "�q����J�פ�!!"
'        Exit Sub
'    End If
    
    '�ˬd��O
    If Trim(rs_Src("�q�����O")) = "C" Then
        MsgBox "�o�{�����q�����O�A�q��q�渹�X:" & Trim(rs_Src("�q�渹�X")) & " �q�����O:" & Trim(rs_Src("�q�����O")) & "�A�q����J�פ�!!�нT�{!!"
        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
        Exit Sub
    End If
    
    '�ˬdTK�ܧO Create by Gemini @20080521
    If (UCase(Trim(rs_Src("�ܧO"))) = "BL01" Or UCase(Trim(rs_Src("�ܧO"))) = "BLR68" Or UCase(Trim(rs_Src("�ܧO"))) = "BL02") = False Then
        MsgBox "�q���ɮסG" & filLocalFile.FileName & " (TK�渹�G" & Trim(rs_Src("��f�渹")) & ")�C" & vbCrLf & "�гq���Ȥ�ATK�ܮw�O���šA�нT�{�ӵ��q��O�_���~!?", vbOKOnly, "�o�{TK�ܧO�D�ըƹF�ܧO BL01, BLR68 ���q�����!"
        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
        Exit Sub
    End If
    
    '������� --�P�_�Ȥ�PO�q�渹�X�O�_����, ���ƮɤJ�t�Ψì��� ����>0�~�ˬd edit by eric20140923
    If Len(StrNoCH(Trim(rs_Src("�Ȥ�渹")))) > 0 Then
         Call Confirm_Recordset_Closed(tmp_Rs)
         str_SQL = "select o.orderkey from orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey " & _
             "where o.externorderkey = '" & Trim(rs_Src("�q�渹�X")) & "' " & _
             "and od.externlineno = '" & Trim(rs_Src("�q�涵��")) & Trim(rs_Src("��f�渹")) & "' " & _
             "and rtrim(substring(o.consigneekey,5,20)) = '" & Trim(rs_Src("�a�}�O")) & "' " & _
             "and rtrim(o.b_phone1)='" & StrNoCH(Trim(rs_Src("�Ȥ�渹"))) & "' " & _
             "and len(rtrim(isnull(o.b_phone1,''))) > 0 " & _
             "and isnull(o.type,'') <> '�R��' " & _
             "and o.priority = '" & Trim(rs_Src("�q�����O")) & "' and o.storerkey = 'LTKK01' "
    
         tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
         If tmp_Rs.EOF = False Then strRePoOrderkey = strRePoOrderkey & StrNoCH(Trim(rs_Src("�Ȥ�渹"))) & "','"
    End If
    rs_Src.MoveNext
Loop

Tran_Level = cn.BeginTrans
Dim int_OrderLine As Integer, int_Order As Integer, int_Repeat As Integer, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strTKLOC As String, strFacility As String, strSku As String, intOrderLinenumber As Integer
rs_Src.MoveFirst
Do While Not rs_Src.EOF
DoEvents: DoEvents
'
'    '�f�D�f��
'    str_SQL = "select sku from gv_skuxpack where storerkey = 'LTKK01' and (storersku = '" & Trim(rs_Src("�Ƹ�")) & "' or sku = '" & Trim(rs_Src("�Ƹ�")) & "') "
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.EOF Then
'        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
'        MsgBox "�q��o�{�s�Ƹ� (" & Trim(rs_Src("�Ƹ�")) & " ) " & Trim(rs_Src("����W��")) & "�A�q����J�פ�!!"
'        Exit Sub
'    End If
    
strSku = Trim(rs_Src("�Ƹ�"))

'If Len(Trim(rs_Src("�Ƹ�"))) > 20 Then strSku = tmp_Rs("sku")
        
            '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
            If strOrderNo <> Trim(rs_Src("�q�渹�X")) & Trim(rs_Src("�a�}�O")) & Trim(rs_Src("��f���")) & Trim(rs_Src("�q�����O")) & Trim(rs_Src("�Ƶ�")) Then
                    strOrderNo = Trim(rs_Src("�q�渹�X")) & Trim(rs_Src("�a�}�O")) & Trim(rs_Src("��f���")) & Trim(rs_Src("�q�����O")) & Trim(rs_Src("�Ƶ�"))
                    
                    '�q��D�ɷs�W�@��
                    str_SQL = "select isnull(max(orderkey),0) from orders"
                    Call Confirm_Recordset_Closed(tmp_Rs)
                    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                    str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
                    tmp_Rs.Close
                    
                    'edit by Eric 20150119�s�W�n��
                    If Right(UCase(Trim(rs_Src("�x��"))), 2) = "-S" Then
                        strFacility = "�ըƹF�n��"
                    Else
                        strFacility = "�ըƹF�_��"
                    End If
                    intOrderLinenumber = 0
                    
                    str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,ConsigneeKey,CustomerOrderkey,Notes,UpdateSource,Facility,b_phone1,type,addwho,editwho) " & _
                    "VALUES ('" & str_Orderkey & "','" & Trim(rs_Src("�q�渹�X")) & "','" & Trim(rs_Src("�q�����O")) & "','LTKK01','" & Trim(rs_Src("�q����")) & "','" & Trim(rs_Src("��f���")) & "', " & _
                    "'" & Trim(rs_Src("�a�}�O")) & "','" & Trim(rs_Src("�Ȥ�渹")) & "','" & Trim(rs_Src("�Ƶ�")) & "','" & filLocalFile.FileName & "','" & strFacility & "','" & StrNoCH(Trim(rs_Src("�Ȥ�渹"))) & "','','" & User_id & "','" & User_id & "')"
                    
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    int_Order = int_Order + 1
            End If
            
            '�������--�P�_�q����ӬO�_���ơA���Ƥ��W�[���ӡA���U�@����ơA�ҥH��b�̫e���]�L�ΡC
            Call Confirm_Recordset_Closed(tmp_Rs)
            str_SQL = "select o.orderkey from ORDERDETAIL od(nolock) join orders o(nolock) on o.orderkey = od.orderkey where o.storerkey = 'LTKK01' and o.ExternOrderKey='" & Trim(rs_Src("�q�渹�X")) & "' and rtrim(o.priority)= '" & Trim(rs_Src("�q�����O")) & "' and isnull(type,'') <> '�R��' and od.ExternlineNO= '" & Trim(rs_Src("�q�涵��")) & "_" & Trim(rs_Src("��f�渹")) & "' "
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF Then
                 
                 If Trim(rs_Src("�x��")) = "04A" Or Trim(rs_Src("�x��")) = "03A" Then
                    strTKLOC = "R" & Trim(rs_Src("�x��"))
                 Else
                    strTKLOC = Trim(rs_Src("�x��"))
                 End If
                 
                 intOrderLinenumber = intOrderLinenumber + 1
                                     
                '�q����Ӹ�Ʒs�W
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber, ExternlineNO ,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice ,CartonGroup,notes,updatesource)" & _
                "VALUES ('" & str_Orderkey & "','" & Format(intOrderLinenumber, "00000") & "','" & Trim(rs_Src("�q�涵��")) & "_" & Trim(rs_Src("��f�渹")) & "','" & Trim(rs_Src("�q�渹�X")) & "','" & strSku & "','LTKK01'," & _
                "'" & Trim(rs_Src("�̤p���ƶq")) & "','" & Trim(rs_Src("�̤p���ƶq")) & "','" & strTKLOC & "','" & Trim(rs_Src("�ܧO")) & "','" & Trim(rs_Src("�q����")) & "','" & Trim(rs_Src("���")) & "','" & Trim(rs_Src("�K��")) & "','" & Trim(rs_Src("�Ƶ�")) & "','" & IIf(UCase(Trim(rs_Src("�K��"))) = "Y", "�\��", "") & "') "
                
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                int_OrderLine = int_OrderLine + 1
                
            Else
                 int_Repeat = int_Repeat + 1
                Call FTPlog("�q����ӭ���" & str_SQL)
                
                '��������
                strReOrderkey = strReOrderkey & Trim(rs_Src("�q�渹�X")) & Trim(rs_Src("�q�涵��")) & Trim(rs_Src("��f�渹")) & "','"
'                GoTo Nextstep

            End If
    
'    Else
'        '�q�歫��
'        Call FTPlog("�q�歫��" & str_SQL)
'        '��������
'        strReOrderkey = strReOrderkey & RTrim(tmp_rs("externorderkey")) & "','"
'    End If
           
'nextloop:
rs_Src.MoveNext
Loop

'�ɫȤ���
cn.Execute "exec gs_ordersupdate 'LTKK01' ", RowsAffect, adExecuteNoRecords

'�s�Ȥ��ˬd
str_SQL = "select �f�D=storerkey , �f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �ˬd��� = getdate() " & _
        "from orders o (nolock) " & _
        "Where o.storerkey = 'LTKK01' and o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select t1m.consigneekey from trp01m t1m where t1m.storerkey = 'LTKK01') "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel���
    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
    If Dir("C:\LTKK01\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\LTKK01\�ʫȤ���"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\LTKK01\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True: cn.RollbackTrans: Tran_Level = 0
    MsgBox "�o�{�s�Ȥ�A�q����J����!", vbOKOnly, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'�����J���
Set dg_CustInv.DataSource = rs_Src

'�����e��
SetDataGridColWidth Me.Caption, dg_CustInv

With dg_CustInv
      .Columns(8).Alignment = dbgRight
      .Columns(9).Alignment = dbgRight
      .Columns(11).Alignment = dbgRight
 End With

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�ɮ� " & strTranFileName & " �ƥ��� C:\Orders\LTKK01\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�ɮ�:" & strTranFileName)
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
    
    If int_Repeat > 0 Then
        msg_text = "��" & int_Repeat & " ���q����ӭ�������!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If

'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then
'
'    str_SQL = "select �������O = 'TK�q����ӭ���-����J' , ��J�ɮצW�� = '" & filLocalFile.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
'        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(od.externorderkey)+rtrim(od.orderlinenumber) in ('" & strReOrderkey & "') " & _
'        "Union " & _
'        "select �������O = '�Ȥ�q�渹�X����-�w��J' , ��J�ɮצW�� = '" & filLocalFile.FileName & "' ,�q�渹�X = o.externorderkey ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource ,�ˬd�ɶ� = getdate() " & _
'        "From orders o join orderdetail od on o.orderkey = od.orderkey and isnull(o.type,'') <> '�R��' where rtrim(b_phone1) in ('" & strRePoOrderkey & "') and len(rtrim(isnull(b_phone1,''))) > 0 "

    str_SQL = "select �������O = 'TK�q����ӭ���-����J' , ��J�ɮצW�� = '" & filLocalFile.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = rtrim(replace(od.externlineno,'_','')) , �Ƹ� = isnull(s.storersku,s.sku) , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey left join gv_skuxpack s(nolock) on s.sku = od.sku and od.storerkey = s.storerkey where rtrim(od.externorderkey)+rtrim(od.orderlinenumber) in ('" & strReOrderkey & "') or rtrim(od.externorderkey)+rtrim(replace(od.externlineno,'_','')) in ('" & strReOrderkey & "') " & _
        "union " & _
        "select �������O = '�Ȥ�q�渹�X����-�w��J' , ��J�ɮצW�� = '" & filLocalFile.FileName & "' ,�q�渹�X = o.externorderkey ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = isnull(s.storersku,s.sku) , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource ,�ˬd�ɶ� = getdate() " & _
        "From orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey and isnull(o.type,'') <> '�R��' left join gv_skuxpack s(nolock) on s.sku = od.sku and od.storerkey = s.storerkey where rtrim(b_phone1) in ('" & strRePoOrderkey & "') and len(rtrim(isnull(b_phone1,''))) > 0 "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\LTKK01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\LTKK01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\LTKK01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'���ɸ�Ʈw�O��
str_SQL = "insert into gt_filelog(storerkey,filename,filedate,filelen,addwho) values('LTKK01','" & filLocalFile.FileName & "','" & Format(FileDateTime(strTranFileName), "YYYY/MM/DD hh:mm:ss") & "','" & FileLen(strTranFileName) & "','" & User_id & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String

'�����ɸ��
Call ReDim_Recordset(tmp_Rs)
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "select �ɮ׮ɶ� = filename , �ɮ׮ɶ� = filedate, ���ɮɶ� = gettime, �t���ɶ� = convert(char(20),gettime - filedate,20) from gv_FileTime where filename = '" & strFileName & "' "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then
    strTextbody = strFileName & "-�ɮ׮ɶ��G" & tmp_Rs("�ɮ׮ɶ�") & " �ɮפj�p�G" & FileLen(strTranFileName) & " �ɶ��t�G" & ((Mid(tmp_Rs("�t���ɶ�"), 9, 2) - 1) * 24) + Mid(tmp_Rs("�t���ɶ�"), 12, 2) & Mid(tmp_Rs("�t���ɶ�"), 14, 6) & "�G(�פJ)"
Else
    strTextbody = strFileName & "-�ɮ׮ɶ��G�L �ɮפj�p�G" & FileLen(strTranFileName) & " �ɶ��t�G�L" & "�G(�פJ)"
End If

''LTKK01�۰� Mail �q��
''�������w
''Exit Sub
'strFrom = "Tkedi@bestlog.com.tw"
'strTo = "jack@mail.kirin.com.tw,irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;simon@mail.kirin.com.tw"
'strCC = "tkedi@bestlog.com.tw"
'strBCC = strBCC
'strSubject = "���ɳq��(" & filLocalFile.FileName & ")"
'strTextbody = strTextbody
'strEmailID = "tkedi"
'strEmailPW = "tkedibl01"
'strAlways = "NO"
'
''�ǰe�l��
'Dim objEmail As Object
'Set objEmail = CreateObject("CDO.Message")
'
'objEmail.From = strFrom
'objEmail.To = strTo
'objEmail.CC = strCC   ' �ƥ�
'objEmail.BCC = strBCC ' �K��ƥ�
'objEmail.Subject = RTrim(strSubject)
'objEmail.TextBody = strTextbody
'objEmail.AddAttachment strAddAttachment
'
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
''SMTP ���A���ݭn���Ү�
'If Len(RTrim(strEmailID)) > 0 Then
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
'End If
'objEmail.Configuration.Fields.Update
'objEmail.Send
'
'Set objEmail = Nothing
'
'MsgBox "���ɳq��Email����", 64, strFileName
'
'tmp_Rs.Close

'�ƥ��ɮ�
If Dir("C:\LTKK01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\LTKK01\Orders\Backup"
If strTranFileName <> "C:\LTKK01\Orders\Backup\" & filLocalFile.FileName Then FileCopy strTranFileName, "C:\LTKK01\Orders\Backup\" & filLocalFile.FileName: Kill strTranFileName

'�ƥ���FTP
If Dir("O:\KIRIN\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\KIRIN\OrdersBackup"
FileCopy "C:\LTKK01\Orders\Backup\" & filLocalFile.FileName, "O:\KIRIN\OrdersBackup\" & filLocalFile.FileName

filLocalFile.Refresh
SSTab1.Tab = 1
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmd_Import_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
End Sub

'Private Sub cmdImport3_Click()
'On Error GoTo err_Handle
'
'If FileLen(filLocal3.Path & "\" & filLocal3.FileName) = 0 Then MsgBox "�ɮפj�p = 0,�ɦW:" & str_file, vbOKOnly + vbInformation, Me.Caption: Exit Sub
'
'If UCase(filLocal3.FileName) <> "XRSLDNL.TXT" Then
'    ConfirmYN = MsgBox("���פJ�ɮצW�٫D xrsldnl.txt ���ɮסA�T�w�פJ?", vbQuestion + vbYesNo, "Warning")
'    If ConfirmYN = vbNo Then Exit Sub
'End If
'
'cmdImport3.Enabled = False: Screen.MousePointer = 11: dg3.Enabled = False
'
''�}�l�פJ�ɮ�
''        If Len(Trim(filLocal3.FileName)) = 0 Then
''            Exit Sub
''        End If
'''Dim fl_file As Scripting.File
'''Set fso = New FileSystemObject
'''If fso.FileExists(strTranFileName) = False Then Exit Sub
'''
'''SSTab1.Tab = 3
''        If strTranFileName = "" Then
''            'Get_CustInv = False
''            Exit Sub
''        End If
'Dim rsReturnOrders As New ADODB.Recordset           '�h�f�q����
'Dim strRow As String    'Ū���C�@���r
'Dim strField As String  'Ū���C�ӰϹj�����
'Dim intPointer As Integer
''Set rsReturnOrders = New Recordset
'With rsReturnOrders
'    .Fields.Append "LoadNo", adChar, 12, adFldUpdatable
'    .Fields.Append "OrderNo", adChar, 12, adFldUpdatable
'    .Fields.Append "OrderDate", adChar, 8, adFldUpdatable
'    .Fields.Append "DeliveryDate", adChar, 8, adFldUpdatable
'    .Fields.Append "CustomerID", adChar, 10, adFldUpdatable
'    .Fields.Append "CustomerName1", adChar, 30, adFldUpdatable
'    .Fields.Append "CustomerName2", adChar, 30, adFldUpdatable
'    .Fields.Append "CustomerName3", adChar, 30, adFldUpdatable
'    .Fields.Append "Address1", adChar, 30, adFldUpdatable
'    .Fields.Append "Address2", adChar, 30, adFldUpdatable
'    .Fields.Append "Address3", adChar, 30, adFldUpdatable
'    .Fields.Append "DeliveryADRS1", adChar, 30, adFldUpdatable
'    .Fields.Append "DeliveryADRS2", adChar, 30, adFldUpdatable
'    .Fields.Append "DeliveryADRS3", adChar, 30, adFldUpdatable
'    .Fields.Append "ZIP", adChar, 3, adFldUpdatable
'    .Fields.Append "SKU", adChar, 14, adFldUpdatable
'    .Fields.Append "SKUDescription", adChar, 30, adFldUpdatable
'    .Fields.Append "OrderCS", adChar, 7, adFldUpdatable
'    .Fields.Append "OrderCT", adChar, 7, adFldUpdatable
'    .Fields.Append "OrderEA", adChar, 7, adFldUpdatable
'    .Fields.Append "OTotalEA", adChar, 7, adFldUpdatable
'    .Fields.Append "CS", adChar, 4, adFldUpdatable
'    .Fields.Append "Casecnt", adChar, 4, adFldUpdatable
'    .Fields.Append "EA", adChar, 4, adFldUpdatable
'    .Fields.Append "PickCS", adChar, 7, adFldUpdatable
'    .Fields.Append "PickCT", adChar, 7, adFldUpdatable
'    .Fields.Append "PickEA", adChar, 7, adFldUpdatable
'    .Fields.Append "PTotalEA", adChar, 7, adFldUpdatable
'    .Fields.Append "OrderComments", adChar, 60, adFldUpdatable
'    .Fields.Append "AssignedDate", adChar, 8, adFldUpdatable
'    .Fields.Append "OrderLine", adChar, 2, adFldUpdatable
'    .Fields.Append "Warehouse", adChar, 12, adFldUpdatable
'    .Fields.Append "DEPOTNO", adChar, 1, adFldUpdatable
'    .Fields.Append "REPACKAGE", adChar, 1, adFldUpdatable
'    .Fields.Append "WT", adChar, 11, adFldUpdatable
'    .Fields.Append "CBM", adChar, 11, adFldUpdatable
'    .Fields.Append "CARNO", adChar, 10, adFldUpdatable
'    .Fields.Append "AssignedFlag", adChar, 1, adFldUpdatable
'    .Fields.Append "CustomerPO", adChar, 12, adFldUpdatable
'    .Fields.Append "OrderType", adChar, 1, adFldUpdatable
'    .Fields.Append "RETURNCT", adChar, 7, adFldUpdatable
'    .Fields.Append "UTLOrderLine", adChar, 4, adFldUpdatable
'    .Fields.Append "ULTCustPO", adChar, 16, adFldUpdatable
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open
'End With
'
''txt to rs
'Open filLocal3.Path & "\" & filLocal3.FileName For Input As #1
'
'Do Until EOF(1)
'Line Input #1, strRow
'intPointer = 1
'With rsReturnOrders
'    .AddNew
'    .Fields(0) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(1) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(2) = Trim(GetWord(strRow, intPointer, 8))
'    .Fields(3) = Trim(GetWord(strRow, intPointer, 8))
'    .Fields(4) = Trim(GetWord(strRow, intPointer, 10))
'    .Fields(5) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(6) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(7) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(8) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(9) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(10) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(11) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(12) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(13) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(14) = Trim(GetWord(strRow, intPointer, 3))
'    .Fields(15) = Trim(GetWord(strRow, intPointer, 14))
'    .Fields(16) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(17) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(18) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(19) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(20) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(21) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(22) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(23) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(24) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(25) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(26) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(27) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(28) = Trim(GetWord(strRow, intPointer, 60))
'    .Fields(29) = Trim(GetWord(strRow, intPointer, 8))
'    .Fields(30) = Trim(GetWord(strRow, intPointer, 2))
'    .Fields(31) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(32) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(33) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(34) = Trim(GetWord(strRow, intPointer, 11))
'    .Fields(35) = Trim(GetWord(strRow, intPointer, 11))
'    .Fields(36) = Trim(GetWord(strRow, intPointer, 10))
'    .Fields(37) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(38) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(39) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(40) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(41) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(42) = Trim(GetWord(strRow, intPointer, 16))
'
'    .Update
'    .MoveFirst
'End With
'
'Loop
'Close
'
'rsReturnOrders.MoveFirst
'Set dg3.DataSource = rsReturnOrders
'With dg3
'    .Columns(17).Alignment = dbgRight
'    .Columns(18).Alignment = dbgRight
'    .Columns(19).Alignment = dbgRight
'    .Columns(20).Alignment = dbgRight
'    .Columns(21).Alignment = dbgRight
'    .Columns(22).Alignment = dbgRight
'    .Columns(23).Alignment = dbgRight
'    .Columns(24).Alignment = dbgRight
'    .Columns(25).Alignment = dbgRight
'    .Columns(26).Alignment = dbgRight
'    .Columns(27).Alignment = dbgRight
'
'End With
'
'rsReturnOrders.MoveFirst
'strOrderNo = ""
'int_repeat = 0
'int_order = 0
'int_orderline = 0
'
'cn.BeginTrans
'Do While Not rsReturnOrders.EOF
'    DoEvents: DoEvents
'
'    If strOrderNo <> rsReturnOrders("OrderNo") Then '�������--�P�_�q��s���w�q�O�_�n�b [�q��D��] ���s�W�@��
'        strOrderNo = rsReturnOrders("OrderNo")
'
'        '�������--�P�_�q��O�_����
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_ExternOrderKey = "R" & Trim(rsReturnOrders("OrderNO"))
'        str_SQL = "select ExternOrderKey from orders where ExternOrderKey='" & str_ExternOrderKey & "' "
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_rs.EOF Then
'        '�q��D�ɷs�W�@��
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_rs)
'            tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_rs.Fields(0))) + 1, 10, 0)
'            tmp_rs.Close
'            str_CustomerID = Trim(rsReturnOrders.Fields(4).Value)
'            '0 "LoadNo",1"OrderNo",2"OrderDate", 3"DeliveryDate",4"CustomerID",5"CustomerName1",6"CustomerName2",7"CustomerName3",8"Address1",9"Address2",10"Address3",
'            '11"DeliveryADRS1",12"DeliveryADRS2",13"DeliveryADRS3",14"ZIP",15"SKU",16"SKUDescription",17"OrderCS",18"OrderCT",19"OrderEA",20"OTotalEA",
'            '21"CS",22"CT",23"EA",24"PickCS",25"PickCT",26"PickEA",27"PTotalEA",28"OrderComments",29"AssignedDate",30"OrderLine",
'            '31"CPNO",32"DEPOTNO",33"REPACKAGE",34"RLNO",35"RDNO",36"CARNO",37"AssignedFlag",37"CustomerPO",39"OrderType",40"RETURNCT",
'            '41"UTLOrderLine",42, adChar, 16, adFldUpdatable
'
'            'OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_contact1,C_Contact2,C_Company,C_Address1,C_Address2,C_Address3,C_Address4,C_City,
'            'C_State,C_Zip,C_Country,C_ISOCntryCode,C_Phone1,C_Phone2,C_Fax1,C_Fax2,C_vat,BuyerPO, BillToKey, B_contact1, B_Contact2, B_Company, B_Address1, B_Address2, B_Address3, B_Address4, B_City, B_State, B_Zip, B_Country,
'            'B_ISOCntryCode, B_Phone1, B_Phone2, B_Fax1, B_Fax2, B_Vat, IncoTerm, PmtTerm, OpenQty, Status, DischargePlace, DeliveryPlace, IntermodalVehicle, CountryOfOrigin, CountryDestination, UpdateSource, Type, OrderGroup, Door,
'            'Route, Stop, Notes, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, ContainerType, ContainerQty, BilledContainerQty, DoRoute, CustomerOrderkey, PIVNO, Amount
'            str_Note = Trim(rsReturnOrders.Fields(28).Value)
'            str_Note = Replace(str_Note, "'", " ")
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,ConsigneeKey,C_Company,C_Zip,C_Address1,C_Address2,C_Address3,CustomerOrderkey,Notes,UpdateSource,DischargePlace,B_PHONE2,Type,OrderGroup,ContainerQty,BilledContainerQty,b_fax2) " & _
'                "VALUES ('" & str_Orderkey & "','" & str_ExternOrderKey & "','" & Trim(rsReturnOrders.Fields(39)) & "','UTL','" & Trim(rsReturnOrders.Fields(2)) & "','" & Trim(rsReturnOrders.Fields(3)) & "', " & _
'                "'" & str_CustomerID & "','" & Trim(rsReturnOrders.Fields(5)) & "','" & Trim(rsReturnOrders.Fields(14)) & "','" & Trim(rsReturnOrders.Fields(11)) & "','" & Trim(rsReturnOrders.Fields(12)) & "'," & _
'                "'" & Trim(rsReturnOrders.Fields(13)) & "','" & Trim(rsReturnOrders.Fields(38)) & "','" & str_Note & "','" & filLocal3.FileName & "','" & Trim(rsReturnOrders.Fields(31)) & "','01','" & Trim(rsReturnOrders.Fields(39)) & "','" & Trim(rsReturnOrders.Fields(33)) & "','" & Trim(rsReturnOrders.Fields(18)) & "','" & Trim(rsReturnOrders.Fields(18)) & "','0')"
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_order = int_order + 1
'        End If
'End If
'
'        '�������--�P�_�q����ӬO�_����
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_SQL = "select ExternOrderKey from ORDERDETAIL where ExternOrderKey='R" & Trim(rsReturnOrders.Fields(1)) & "' and OrderLineNumber= '" & Trim(rsReturnOrders.Fields(41)) & "'"
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If Not tmp_rs.EOF Then
'            int_repeat = int_repeat + 1
'            Call FTPlog("�q����ӭ���" & str_SQL)
'            GoTo nextloop
'        End If
'
'        '�q����Ӹ�Ʒs�W
'        'OrderKey, OrderLineNumber, OrderDetailSysId, ExternOrderKey, ExternLineNo, Sku, StorerKey, ManufacturerSku, RetailSku, AltSku, OriginalQty, OpenQty, ShippedQty, AdjustedQty, QtyPreAllocated, QtyAllocated, QtyPicked, UOM, PackKey, PickCode, CartonGroup, Lot, ID, Facility, Status, UnitPrice, Tax01, Tax02, ExtendedPrice, UpdateSource, Lottable01, Lottable02, Lottable03, Lottable04, Lottable05, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, TariffKey, Lottable06, Lottable07, Lottable08, Lottable09, Lottable10, Lottable11, Beginqty
'        str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber, ExternOrderKey,ExternLineNo,Sku,StorerKey,OriginalQty,openqty,qtypreallocated,shippedqty,qtyallocated,adjustedqty,qtypicked,Lottable01, Lottable02,PackKey,Lottable03,pickcode)" & _
'         "VALUES ('" & str_Orderkey & "','" & Trim(rsReturnOrders.Fields(41)) & "','" & str_ExternOrderKey & "','" & Trim(rsReturnOrders.Fields(40)) & "','" & Trim(rsReturnOrders.Fields(15)) & "','UTL', " & _
'         "'" & Round(CLng(rsReturnOrders.Fields(20)) / IIf(CLng(rsReturnOrders.Fields(22)) = 0, 1, CLng(rsReturnOrders.Fields(22))), 3) & "','" & Trim(rsReturnOrders.Fields(17)) & "','" & Trim(rsReturnOrders.Fields(17)) & "','" & Trim(rsReturnOrders.Fields(19)) & "','" & Trim(rsReturnOrders.Fields(19)) & "','" & Trim(rsReturnOrders.Fields(20)) & "','" & Trim(rsReturnOrders.Fields(20)) & "','" & Trim(rsReturnOrders.Fields(37)) & "','" & Trim(rsReturnOrders.Fields(29)) & "','" & Trim(rsReturnOrders.Fields(15)) & "','" & Trim(rsReturnOrders.Fields(27)) & "','" & Trim(rsReturnOrders.Fields(22)) & "')"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'        '�������--�P�_SKU�_�s�b
'        str_SQL = "select sku from sku where sku='" & Trim(rsReturnOrders("sku")) & "'"
'        Call Confirm_Recordset_Closed(tmp_rs)
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_rs.EOF Then
'            str_DESCR = Trim(rsReturnOrders("SKUDescription"))
'            str_DESCR = Replace(str_DESCR, "'", " ")
'            'StorerKey, Sku, DESCR, SUSR1, SUSR2, SUSR3, SUSR4, SUSR5, MANUFACTURERSKU, RETAILSKU, ALTSKU, PACKKey, STDGROSSWGT, STDNETWGT, STDCUBE, TARE, CLASS, ACTIVE, SKUGROUP, Tariffkey, BUSR1, BUSR2, BUSR3, BUSR4, BUSR5, LOTTABLE01LABEL, LOTTABLE02LABEL, LOTTABLE03LABEL, LOTTABLE04LABEL, LOTTABLE05LABEL, NOTES1, NOTES2, PickCode, StrategyKey, CartonGroup, PutCode, PutawayLoc, PutawayZone, InnerPack, Cube, GrossWgt, NetWgt, ABC, CycleCountFrequency, LastCycleCount, ReorderPoint, ReorderQty, StdOrderCost, CarryCost, Price, Cost, ReceiptHoldCode, ReceiptInspectionLoc, TrafficCop, ArchiveCop, IOFlag, TareWeight, LotxIdDetailOtherlabel1, LotxIdDetailOtherlabel2, LotxIdDetailOtherlabel3, AvgCaseWeight, TolerancePct, SkuRotat01, SkuRotat02, SkuRotat03, SkuRotat04, SkuRotat05, DefaultRotation, AllocParm, onreceiptcopypackkey, replenishmentQty
'            str_CSKU = StringCleaner(Trim(rsReturnOrders.Fields(15)), "'")
'            str_SQL = "INSERT sku (StorerKey, Sku, DESCR,PACKKey,STDGROSSWGT,BUSR4) " & _
'            "VALUES ('UTL','" & Trim(rsReturnOrders.Fields(15)) & "','" & str_DESCR & "','" & Trim(rsReturnOrders.Fields(15)) & "','" & Trim(rsReturnOrders.Fields(34)) / 1000 & "','" & Trim(rsReturnOrders.Fields(35)) / 1000 & "' )"
'            cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
'            'PackKey, PackDescr, PackUOM1, CaseCnt, ISWHQty1, ReplenishUOM1, ReplenishZone1, CartonizeUOM1, LengthUOM1, WidthUOM1, HeightUOM1, CubeUOM1, PackUOM2, InnerPack, ISWHQty2, ReplenishUOM2, ReplenishZone2, CartonizeUOM2, LengthUOM2, WidthUOM2, HeightUOM2, CubeUOM2, PackUOM3, Qty, ISWHQty3, ReplenishUOM3, ReplenishZone3, CartonizeUOM3, LengthUOM3, WidthUOM3, HeightUOM3, CubeUOM3, PackUOM4, Pallet, ISWHQty4, ReplenishUOM4, ReplenishZone4, CartonizeUOM4, LengthUOM4, WidthUOM4, HeightUOM4, CubeUOM4, PalletWoodLength, PalletWoodWidth, PalletWoodHeight, PalletTI, PalletHI, PackUOM5, Cube, ISWHQty5, PackUOM6, GrossWgt, ISWHQty6, PackUOM7, NetWgt, ISWHQty7, PackUOM8, OtherUnit1, ISWHQty8, ReplenishUOM8, ReplenishZone8, CartonizeUOM8, LengthUOM8, WidthUOM8, HeightUOM8, PackUOM9, OtherUnit2, ISWHQty9, ReplenishUOM9, ReplenishZone9, CartonizeUOM9, LengthUOM9, WidthUOM9, HeightUOM9, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, Timestamp
'            str_SQL = "INSERT pack (PACKKey,PackDescr,GrossWgt, PalletHI, PalletTI,CaseCnt) " & _
'            "VALUES ('" & Trim(rsReturnOrders.Fields(15)) & "','" & str_DESCR & "','" & Trim(rsReturnOrders.Fields(34)) / 1000 & "', " & _
'             "'0','0','1')"
'            cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
'        End If
'
'        tmp_rs.Close
'        int_orderline = int_orderline + 1
'nextloop:
'
'    rsReturnOrders.MoveNext
'Loop
'
'cn.CommitTrans
'
''�ƥ��ɮ�
'If Dir("C:\From_ids\Backup\UTLR", vbDirectory) = "" Then MkDir "C:\From_ids\Backup\UTLR"
'
'If filLocal3.Path = "C:\From_ids\Backup\UTLR" Then
'Else
'    FileCopy filLocal3.Path & "\" & filLocal3.FileName, "C:\From_ids\Backup\UTLR\xrsldnl" & Format(Now(), "yyyymmddhhmmss") & ".txt"
'    Kill filLocal3.Path & "\" & filLocal3.FileName
'End If
'
'If int_repeat > 0 Then
'    msg_text = "�� " & int_repeat & " ���q����ӭ�������"
'    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'End If
'
'msg_text = "�פJ " & int_order & " ���q��A " & int_orderline & " �����ӡA��r�� " & filLocal3.FileName & " �ƥ��� C:\From_ids\Backup\UTLR\"
'MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'Call FTPlog("�פJ " & int_order & " ���q��A " & int_orderline & " �����ӡA��r�� " & filLocal3.FileName)
'filLocal3.Refresh
'
'cmdImport3.Enabled = True: Screen.MousePointer = 0: dg3.Enabled = True
'Exit Sub
'
'err_Handle:
'Close #1
''cn.RollbackTrans
'Dim tmpString As String
'msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImport3_Click", tmpString
'MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'cmdImport3.Enabled = True: Screen.MousePointer = 0: dg3.Enabled = True
'End Sub

Private Sub cmdLogOn_Click()
    On Error GoTo LogOnError
    
    If txtServer = "" Or txtPassword = "" Then
        MsgBox "�A������Jftp Server�P�K�X", vbInformation + vbOKOnly, "LogOn Failure"
        Exit Sub
    End If
    
    'Set status label
    lblStatus = "Connecting"
    'Set protocol and server
    ITC.Protocol = icFTP
    ITC.url = txtServer
    'If no username is entered default to anonymous
    If txtUserName = "" Then
        ITC.UserName = "anonymous"
    Else
        ITC.UserName = txtUserName
    End If
    ITC.Cancel
    cmdLogOn.Enabled = False
    'Set the password and connect
    ITC.PassWord = txtPassword
    ITC.RequestTimeout = 20
    
    ITC.Execute , "CD " & Chr(34) & "/Bestg/Alc" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    
    'Set status label, disable the log on button, and enable the log off button
    lblStatus = "�w�s�u��alc��Ƨ�"
    cmdLogOn.Enabled = False
    cmdNewFolder.Enabled = True
    cmdDelete.Enabled = True
    cmdUpFolder.Enabled = True
    imgSendFile.Enabled = True
    imgReceiveFile.Enabled = True
    lstRemoteFile.Enabled = True
    cmdLogOff.Enabled = True
    
    Exit Sub
    
LogOnError:
    'If logon fails alert the user
    MsgBox "�n�J���~....", vbOKOnly + vbInformation, "�n�J���~"
    ITC.Cancel
    lblStatus = "Not Connected"
    cmdLogOn.Enabled = True
    cmdNewFolder.Enabled = False
    cmdDelete.Enabled = False
    cmdUpFolder.Enabled = False
    imgSendFile.Enabled = False
    imgReceiveFile.Enabled = False
    lstRemoteFile.Enabled = False
    cmdLogOff.Enabled = False
End Sub

Private Sub cmdLogOff_Click()
'Clear the list of remote files and log off
lstRemoteFile.Clear
ITC.Cancel

Do Until ITCReady(False)
    DoEvents: DoEvents: DoEvents: DoEvents
Loop

lblStatus = "Closing Connection"

If ITCReady(False) Then
    ITC.Execute , "CLOSE"
Else
    ITC.Cancel
End If
lblStatus = "Not Connected"
cmdLogOn.Enabled = False
cmdNewFolder.Enabled = False
cmdDelete.Enabled = False
cmdUpFolder.Enabled = False
imgSendFile.Enabled = False
imgReceiveFile.Enabled = False
lstRemoteFile.Enabled = False
cmdLogOff.Enabled = False
cmdLogOn.Enabled = True
End Sub
'Private Sub cmdLogOff3_Click()
'If int3.StillExecuting Then MsgBox "�еy��.  FTP���A�����椤", vbInformation + vbOKOnly, "���L��": Exit Sub
'    'Clear the list of remote files and log off
'    lstRemoteFile3.Clear
'
''    Do Until int3Ready(True)
''        DoEvents: DoEvents: DoEvents: DoEvents
''    Loop
''
'    lblStatus3 = "Closing Connection"
'
'    If int3Ready(False) Then
'        int3.Execute , "CLOSE"
'    Else
'        int3.Cancel
'    End If
'    lblStatus3 = "Not Connected"
'    cmdLogon3.Enabled = False
''    cmdNewFolder.Enabled = False
''    cmdDelete.Enabled = False
''    cmdUpFolder.Enabled = False
''    imgSendFile.Enabled = False
'    cmdImport3.Enabled = False
'    lstRemoteFile3.Enabled = False
'    cmdLogOff3.Enabled = False
'    cmdLogon3.Enabled = True
'    int3.Cancel
'End Sub

Private Sub cmdDelete_Click()
'�Ȥ��ϥ�
'If the itc is ready, ask user if they want to delete it, if so then delete
If ITCReady(True) Then
    If MsgBox("�T�w�R�� " & lstRemoteFile.Text & " ?", vbQuestion + vbOKCancel, "Delete") = vbOK Then
        If Right(lstRemoteFile.Text, 1) <> "/" Then
            ITC.Execute , "DELETE " & Chr(34) & lstRemoteFile.Text & Chr(34)
        Else
            ITC.Execute , "RMDIR " & Chr(34) & lstRemoteFile.Text & Chr(34)
        End If
        
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        lstRemoteFile.Clear
        ITC.Execute , "DIR"
        lblStatus = "�w�s�u"
    End If
End If
End Sub

Private Sub cmdNewFolder_Click()
    '�Ȥ��ϥ�
    'If the itc is ready then check to make sure the folder doesn't already exist
    Dim FolderName As String, i As Double
    If ITCReady(True) Then
        FolderName = InputBox("Enter new folder name", "New Folder")
        For i = 0 To lstRemoteFile.ListCount
            If FolderName & "/" = lstRemoteFile.List(i) Then
                MsgBox "Folder " & FolderName & " already exists.", vbInformation + vbOKOnly, "New Folder"
                Exit Sub
            End If
        Next i
        'Create the new folder then refresh the remote file list
        ITC.Execute , "MKDIR " & Chr(34) & FolderName & Chr(34)
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        
        lstRemoteFile.Clear
        ITC.Execute , "DIR"
        lblStatus = "�w�s�u"
    End If
End Sub

Private Sub cmdUpFolder_Click()
'�Ȥ��ϥ�
'If the itc is ready then move up one directory and refresh the remote files list
If ITCReady(True) Then
    ITC.Execute , "CDUP"
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    lstRemoteFile.Clear
    ITC.Execute , "DIR"
    lblStatus = "�w�s�u"
    
End If
End Sub

Private Sub Command2_Click()

strTranFileName = filLocalFileT7.Path & "\" & filLocalFileT7.FileName
If Len(RTrim(cboSheetT7)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT7.EOF Or rsMainT7 Is Nothing Then Exit Sub
On Error GoTo err_Handle

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT7.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'��f����ˬd
rsMainT7.MoveFirst
Do While Not rsMainT7.EOF

If Format(myExCharFilter(Trim(rsMainT7("Deliv.date"))), "YYYY/MM/DD") < Format(Now, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub

    rsMainT7.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT7.Enabled = False: dgMainT7.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngInnerpack As Long
Dim strOrderNo As String, strWHOrderNo As String, strWH As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strLot06 As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT7.MoveFirst

''CT or BEST�P�_
'strWHOrderNo = UCase(Trim(rsMainT7("Delivery")))
'
'If (Trim(rsMainT7("sloc")) = "0007" Or Trim(rsMainT7("sloc")) = "0008" Or Trim(rsMainT7("sloc")) = "0009" Or Trim(rsMainT7("sloc")) = "0010") Then
'    strWH = "BEST"
'Else
'    strWH = "CT"
'End If

Do While Not rsMainT7.EOF

    '�������--�P�_�q��ƬO�_��0
    If Trim(rsMainT7("Qty (stckpg unit)")) = 0 Then intNotBest = intNotBest + 1: GoTo next1

'    DoEvents: DoEvents
    
    'CT & BEST���ܥX�f�P�_
    If strWHOrderNo = UCase(Trim(rsMainT7("Delivery"))) Then
        If (strWH = "BEST" And (Trim(rsMainT7("sloc")) = "0001" Or Trim(rsMainT7("sloc")) = "0005")) Or (strWH = "CT" And (Trim(rsMainT7("sloc")) <> "0001" Or Trim(rsMainT7("sloc")) <> "0005") = False) Then
            cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True: cn.RollbackTrans: Tran_Level = 0
            MsgBox "�o�{ CT & BEST ���ܥX�f�q��A�лP�Ȥ�T�{�q�楿�T�L�~�I", 16, "�q����J�פ� "
            Exit Sub
        End If
    End If
    
    '�O�_���ըƹF�t�O�ܧO-->���U�@��
    If Trim(rsMainT7("plnt")) = "1119" And (Trim(rsMainT7("sloc")) = "0001" Or Trim(rsMainT7("sloc")) = "0005") Then
        
        strWH = "CT"
        '�ˬd�q��q�O�_�X�{�p���I
        If Val(rsMainT7("Qty (stckpg unit)")) <> Round(Val(rsMainT7("Qty (stckpg unit)")), 0) Then
            cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True: cn.RollbackTrans: Tran_Level = 0
            MsgBox "�o�{�q��q�X�{�p���I�A�лP�Ȥ�T�{���T�q��q�I", 16, "�q����J�פ� "
            Exit Sub
        End If
        
        '�������--�P�_SKU�O�_�s�b
        str_SQL = "select sku,innerpack from gv_skuxpack where sku='" & Trim(rsMainT7("Material")) & "' and Storerkey = 'LNSL01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            cn.RollbackTrans: Tran_Level = 0: cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT7("Material")) & ") �A�q����J�פ�!!"
            Exit Sub
        End If
        lngInnerpack = tmp_Rs("Innerpack")

    Else
        intNotBest = intNotBest + 1
        strWH = "BEST"
        GoTo next1

    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT7("Delivery"))) Then
        strOrderNo = UCase(Trim(rsMainT7("Delivery")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT7("Delivery"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,BillToKey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT7("Delivery"))) & "','I','LNSL01',getdate(),'" & myExCharFilter(Trim(rsMainT7("Deliv.date"))) & "','" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT7("Ship-To"))) & "','" & GetWord(myExCharFilter(Trim(rsMainT7("Name of the ship-to party"))), 1, 45) & "','','','','','','" & myExCharFilter(Trim(rsMainT7("Sold-to pt"))) & "','" & myExCharFilter(Trim(rsMainT7("po no"))) & "','" & myExCharFilter(Trim(rsMainT7("remarks"))) & "','" & filLocalFileT7.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT7("Delivery")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫�Ƥ��W�[����
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlineNumber = int_orderlineNumber + 1
            
            '�ƶq�ഫ
            intQTY = Trim(rsMainT7("Qty (stckpg unit)"))
'            If Trim(rsMainT7("Material")) = "12129314" Then intQTY = intQTY * IIf(lngInnerpack = 0, 1, lngInnerpack)
            'If lngInnerpack > 0 Then intQTY = intQTY * lngInnerpack
            
            '�ܧO�ഫ
            strLot06 = myExCharFilter(Trim(rsMainT7("sloc")))
            
            If strLot06 = "0001" Then
               strLot06 = "R01"
            ElseIf strLot06 = "0002" Then
               strLot06 = "R01"
            ElseIf strLot06 = "0005" Then
               strLot06 = "R08"
            End If
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable05,Lottable06, Facility,UOM,Unitprice,notes) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT7("Item"))) & "','" & myExCharFilter(Trim(rsMainT7("Delivery"))) & "','" & myExCharFilter(Trim(rsMainT7("material"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT7("Batch"))) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT7("BUn"))) & "','0','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:

        If (Trim(rsMainT7("sloc")) = "0001" Or Trim(rsMainT7("sloc")) = "0005") Then
            strWH = "CT"
        Else
            strWH = "BEST"
        End If

        strWHOrderNo = UCase(Trim(rsMainT7("Delivery")))
        rsMainT7.MoveNext
        
Loop

'�ɫȤ���
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'�s�Ȥ��ˬd
str_SQL = "select �f�D=storerkey , �f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �ˬd��� = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel���
    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�ʫȤ���"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "�o�{�s�Ȥ�A�q����J����!", 16, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT7.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT7.FileName & " �ƥ��� C:\BEST\LNSL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT7.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT7.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT7.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT7.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT7.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT7.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT7.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT7.Refresh: cboSheetT7.Clear
Screen.MousePointer = 0: cmdImportT7.Enabled = True: dgMainT7.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "��}��[������-�פJ", Me.Caption, "cmd2_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True
End Sub

Private Sub Command3_Click()

'��ƱƧ�
Recordset2Excel "TEST", rsMainT17_1

'..�b���s��EXCEL
With MyXlsApp
    
End With

Set MyXlsApp = Nothing
End Sub



Private Sub Command5_Click()

'��ƱƧ�
Recordset2Excel "�q��D��", rsMainT16
Recordset2Excel "�q�������", rsMainT16_1

'..�b���s��EXCEL
With MyXlsApp
    
End With

Set MyXlsApp = Nothing

End Sub

Private Sub dg_CustInv_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
'DataGrid ColResize�ƥ󤤥[�J�U�q�{���X�A�ΥH�O����e
If Len(dg_CustInv.Columns(ColIndex).DataField) = 0 Then Exit Sub
SaveSetting App.title, Me.Caption & "dg_CustInv", dg_CustInv.Columns(ColIndex).DataField, dg_CustInv.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16_1
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16_2
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_3_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16_3
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT17_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT17
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT18_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT18
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT15_Change()
    filLocalFileT15.Path = dirLocalDirT15.Path
End Sub

Private Sub dirLocalDirT16_Change()
    filLocalFileT16.Path = dirLocalDirT16.Path
End Sub


Private Sub dirLocalDirT17_Change()
    filLocalFileT17.Path = dirLocalDirT17.Path
End Sub

Private Sub dirLocalDirT18_Change()
    filLocalFileT18.Path = dirLocalDirT18.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveError
    dirLocalDirT20.Path = drvLocalDriveT20.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub dirLocalDirT19_Change()
    filLocalFileT19.Path = dirLocalDirT19.Path
End Sub

Private Sub dirLocalDirT20_Change()
       filLocalFileT20.Path = dirLocalDirT20.Path
End Sub


Private Sub dirLocalDirT21_Change()
    filLocalFileT21.Path = dirLocalDirT21.Path
End Sub

Private Sub dirLocalDirT22_Change()
    filLocalFileT22.Path = dirLocalDirT22.Path
End Sub


Private Sub drvLocalDrive_Change()
    On Error GoTo DriveError
    dirLocalDir.Path = drvLocalDrive.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub dirLocalDir_Change()
    filLocalFile.Path = dirLocalDir.Path
End Sub

Private Sub drvLocalDriveT15_Change()
On Error GoTo DriveError
dirLocalDirT15.Path = drvLocalDriveT15.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub drvLocalDriveT16_Change()
On Error GoTo DriveError
dirLocalDirT16.Path = drvLocalDriveT16.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub


Private Sub drvLocalDriveT17_Change()
On Error GoTo DriveError
dirLocalDirT17.Path = drvLocalDriveT17.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub drvLocalDriveT18_Change()
On Error GoTo DriveError
dirLocalDirT18.Path = drvLocalDriveT18.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub

Private Sub drvLocalDriveT19_Change()
On Error GoTo DriveError
dirLocalDirT19.Path = drvLocalDriveT19.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub



Private Sub drvLocalDriveT20_Click()
On Error GoTo DriveError
dirLocalDirT20.Path = drvLocalDriveT20.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub




Private Sub drvLocalDriveT20_Change()
    On Error GoTo DriveError
    dirLocalDirT20.Path = drvLocalDriveT20.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub drvLocalDriveT21_Change()
On Error GoTo DriveError
dirLocalDirT21.Path = drvLocalDriveT21.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub

Private Sub drvLocalDriveT22_Change()
On Error GoTo DriveError
dirLocalDirT22.Path = drvLocalDriveT22.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub


Private Sub filLocalFileT15_Click()

On Error GoTo err_Handle
Set rsMainT15 = Nothing: Set dgMainT15.DataSource = rsMainT15
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT15.Path, 1) = "\" Then
    strFilePath = filLocalFileT15.Path
Else
    strFilePath = filLocalFileT15.Path & "\"
End If

If Dir(strFilePath & filLocalFileT15.FileName) = "" Then: filLocalFileT15.Refresh: Exit Sub

cboSheetT15.Clear

If UCase(mySplit(filLocalFileT15.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT15.FileName)
    MyXlsApp.DisplayAlerts = False
  
    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT15.Clear
    
    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT15.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT15.ListIndex = -1
    
    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT15.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT16_Click()

On Error GoTo err_Handle
Set rsMainT16 = Nothing: Set dgMainT16.DataSource = rsMainT16
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

If Dir(strFilePath & filLocalFileT16.FileName) = "" Then: filLocalFileT16.Refresh: Exit Sub

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT17_Click()

On Error GoTo err_Handle
Set rsMainT17 = Nothing: Set dgMainT16.DataSource = rsMainT17
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT17.Path, 1) = "\" Then
    strFilePath = filLocalFileT17.Path
Else
    strFilePath = filLocalFileT17.Path & "\"
End If

If Dir(strFilePath & filLocalFileT17.FileName) = "" Then: filLocalFileT17.Refresh: Exit Sub

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT18_Click()

On Error GoTo err_Handle
Set rsMainT18 = Nothing: Set dgMainT18.DataSource = rsMainT18
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT18.Path, 1) = "\" Then
    strFilePath = filLocalFileT18.Path
Else
    strFilePath = filLocalFileT18.Path & "\"
End If

If Dir(strFilePath & filLocalFileT18.FileName) = "" Then: filLocalFileT18.Refresh: Exit Sub

cboSheetT18.Clear

If UCase(mySplit(filLocalFileT18.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT18.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT18.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT18.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT18.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT18.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub filLocalFileT19_Click()

On Error GoTo err_Handle
Set rsMainT19 = Nothing: Set dgMainT19.DataSource = rsMainT19
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT19.Path, 1) = "\" Then
    strFilePath = filLocalFileT19.Path
Else
    strFilePath = filLocalFileT19.Path & "\"
End If

If Dir(strFilePath & filLocalFileT19.FileName) = "" Then: filLocalFileT19.Refresh: Exit Sub

cboSheetT19.Clear

If UCase(mySplit(filLocalFileT19.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT19.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT19.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT19.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT19.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT19.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub filLocalFileT20_Click()
On Error GoTo err_Handle
Set rsMainT20 = Nothing: Set dgMainT20.DataSource = rsMainT20
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT20.Path, 1) = "\" Then
    strFilePath = filLocalFileT20.Path
Else
    strFilePath = filLocalFileT20.Path & "\"
End If

If Dir(strFilePath & filLocalFileT20.FileName) = "" Then: filLocalFileT20.Refresh: Exit Sub

cboSheetT20.Clear

If UCase(mySplit(filLocalFileT20.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT20.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT20.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        
        cboSheetT20.AddItem MyXlsApp.Sheets(i).Name
  
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT20.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT20.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT21_Click()

On Error GoTo err_Handle
Set rsMainT21 = Nothing: Set dgMainT21.DataSource = rsMainT21
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT21.Path, 1) = "\" Then
    strFilePath = filLocalFileT21.Path
Else
    strFilePath = filLocalFileT21.Path & "\"
End If

If Dir(strFilePath & filLocalFileT21.FileName) = "" Then: filLocalFileT21.Refresh: Exit Sub

cboSheetT21.Clear

If UCase(mySplit(filLocalFileT21.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT21.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT21.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        
        cboSheetT21.AddItem MyXlsApp.Sheets(i).Name
  
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT21.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT21.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT22_Click()

On Error GoTo err_Handle
Set rsMainT22 = Nothing: Set dgMainT22.DataSource = rsMainT22
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT22.Path, 1) = "\" Then
    strFilePath = filLocalFileT22.Path
Else
    strFilePath = filLocalFileT22.Path & "\"
End If

If Dir(strFilePath & filLocalFileT22.FileName) = "" Then: filLocalFileT22.Refresh: Exit Sub

cboSheetT22.Clear

If UCase(mySplit(filLocalFileT22.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT22.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT22.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT22.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT22.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT22.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub lstRemoteFile_DblClick()
'If the itc is ready, check that the selected is a folder and change the directory
If ITCReady(True) Then
    If Right(lstRemoteFile.Text, 1) = "/" Then
        ITC.Execute , "CD " & Chr(34) & lstRemoteFile.Text & Chr(34)
    Else
        Call imgReceiveFile_Click
        Exit Sub
    End If
    
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    
    lstRemoteFile.Clear
    ITC.Execute , "DIR"
    lblStatus = "�w�s�u"
End If
End Sub

Private Sub imgReceiveFile_Click()
Dim i As Double
    'If the ITC is not still executing then receive the file
    If ITCReady(True) = True Then
        'Check that they are not recieving a folder
        If Right(lstRemoteFile.Text, 1) = "/" Then
            MsgBox lstRemoteFile.Text & " is a folder and cannot be sent.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        'Check that the file does not already exist on the computer, if it does exit sub
        For i = 0 To filLocalFile.ListCount
            If lstRemoteFile.Text = filLocalFile.List(i) Then
                MsgBox "�ɮ� " & Right(lstRemoteFile.Text, 18) & " �w�s�b.", vbInformation + vbOKOnly, "Recieve"
                Exit Sub
            End If
        Next i
        str_file = Trim(Right(lstRemoteFile.Text, 18))
        ITC.Execute , "GET " & Chr(34) & str_file & Chr(34) & " " & Chr(34) & filLocalFile.Path & "\" & str_file & Chr(34)
        lblStatus = "�U�����еy��...."
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        filLocalFile.Refresh
        lblStatus = "�w�s�u"
        
        '�}�l�פJ�ɮ�
        strTranFileName = filLocalFile.Path & "\" & str_file
        If Len(Trim(strTranFileName)) = 0 Then
            Exit Sub
        End If
        If FileLen(strTranFileName) = 0 Then
            msg_text = "�ɮפj�p=0,�ɦW:" & str_file
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Exit Sub
        End If
        SSTab1.Tab = 1
        DoEvents: DoEvents
        Dim strRow As String    'Ū���C�@���r
        Dim strField As String  'Ū���C�ӰϹj�����
        Dim intPointer As Integer, int_OrderLine As Integer, int_Order As Integer, int_Repeat As Integer
        Set rs_Src = New Recordset
        With rs_Src
            .Fields.Append "OrderNo", adChar, 12, adFldUpdatable
            .Fields.Append "OrderType", adChar, 1, adFldUpdatable
            .Fields.Append "Division", adChar, 3, adFldUpdatable
            .Fields.Append "OrderDate", adChar, 8, adFldUpdatable
            .Fields.Append "DeliveryDate", adChar, 8, adFldUpdatable
            .Fields.Append "CustomerID", adChar, 10, adFldUpdatable
            .Fields.Append "CustomerName", adChar, 45, adFldUpdatable
            .Fields.Append "ZIP", adChar, 3, adFldUpdatable
            .Fields.Append "Address1", adChar, 45, adFldUpdatable
            .Fields.Append "Address2", adChar, 45, adFldUpdatable
            .Fields.Append "Address3", adChar, 45, adFldUpdatable
            .Fields.Append "CustomerPO", adChar, 12, adFldUpdatable
            .Fields.Append "OrderComments", adChar, 60, adFldUpdatable
            .Fields.Append "OrderLine", adChar, 4, adFldUpdatable
            .Fields.Append "SKU", adChar, 14, adFldUpdatable
            .Fields.Append "SKUDescription", adChar, 60, adFldUpdatable
            .Fields.Append "AllocateQTY", adChar, 11, adFldUpdatable
            .Fields.Append "Ship QTY", adChar, 11, adFldUpdatable
            .Fields.Append "Weight", adChar, 11, adFldUpdatable
            .Fields.Append "MSR", adChar, 11, adFldUpdatable
            .Fields.Append "AssignedFlag", adChar, 1, adFldUpdatable
            .Fields.Append "AssignedDate", adChar, 8, adFldUpdatable
            .Fields.Append "HI", adChar, 5, adFldUpdatable
            .Fields.Append "TI", adChar, 5, adFldUpdatable
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open
        End With
        ' �����_�l�ȡC
        intPointer = 1
        Open strTranFileName For Input As #1
        Do Until EOF(1)
            Line Input #1, strRow
            With rs_Src
                .AddNew
                .Fields(0) = Trim(GetWord(strRow, intPointer, 12))
                .Fields(1) = Trim(GetWord(strRow, intPointer, 1))
                .Fields(2) = Trim(GetWord(strRow, intPointer, 3))
                .Fields(3) = Trim(GetWord(strRow, intPointer, 8))
                .Fields(4) = Trim(GetWord(strRow, intPointer, 8))
                .Fields(5) = Trim(GetWord(strRow, intPointer, 10))
                .Fields(6) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(7) = Trim(GetWord(strRow, intPointer, 3))
                .Fields(8) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(9) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(10) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(11) = Trim(GetWord(strRow, intPointer, 12))
                .Fields(12) = Trim(GetWord(strRow, intPointer, 60))
                .Fields(13) = Trim(GetWord(strRow, intPointer, 4))
                .Fields(14) = Trim(GetWord(strRow, intPointer, 14))
                .Fields(15) = Trim(GetWord(strRow, intPointer, 60))
                .Fields(16) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(17) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(18) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(19) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(20) = Trim(GetWord(strRow, intPointer, 1))
                .Fields(21) = Trim(GetWord(strRow, intPointer, 8))
                .Fields(22) = Trim(GetWord(strRow, intPointer, 5))
                .Fields(23) = Trim(GetWord(strRow, intPointer, 5))
                .Update
                .MoveFirst
            End With
            intPointer = 1
        Loop
        Close
        rs_Src.MoveFirst
        Set dg_CustInv.DataSource = rs_Src
        With dg_CustInv
              .Columns(15).Alignment = dbgRight
              .Columns(16).Alignment = dbgRight
              .Columns(17).Alignment = dbgRight
              .Columns(18).Alignment = dbgRight
         End With
         
        rs_Src.MoveFirst
        strOrderNo = ""
        int_Repeat = 0
        int_Order = 0
        int_OrderLine = 0
        Do While Not rs_Src.EOF
           If strOrderNo <> rs_Src.Fields("OrderNo").Value Then '�������--�P�_�q��s���w�q�O�_�n�b [�q��D��] ���s�W�@��
                strOrderNo = rs_Src.Fields("OrderNo").Value
                '�������--�P�_�q��O�_����
                Call Confirm_Recordset_Closed(tmp_Rs)
                str_SQL = "select ExternOrderKey from Logictown.dbo.orders where ExternOrderKey='" & Trim(rs_Src.Fields(0)) & "' "
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If tmp_Rs.EOF Then
                    '�q��D�ɷs�W�@��
                    str_SQL = "select isnull(max(orderkey),0) from Logictown.dbo.orders"
                    Call Confirm_Recordset_Closed(tmp_Rs)
                    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                    str_Orderkey = StrPadLeft(Val(tmp_Rs.Fields(0)) + 1, 10, 0)
                    str_CustomerID = Trim(rs_Src.Fields(5).Value)
                    If Left(rs_Src.Fields(5).Value, 3) = "GO" Or Left(rs_Src.Fields(5).Value, 3) = "SO" Or Left(rs_Src.Fields(5).Value, 3) = "NO" Or Len(rs_Src.Fields(5).Value) = 0 Then
                        str_SQL = "select isnull(max(ConsigneeKey),0) from Logictown.dbo.orders where left(ConsigneeKey,3)='UTL'"
                        Call Confirm_Recordset_Closed(tmp_Rs)
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        str_CustomerID = "UTL" & StrPadLeft(Val(Right(tmp_Rs.Fields(0), 7)) + 1, 7, 0)
                        tmp_Rs.Close
                    End If
                    'tmp_rs.Close
                    cn.BeginTrans
                        'OrderNo,OrderType,Division,OrderDate,DeliveryDate,CustomerID,CustomerName,ZIP,Address1,Address2,Address3,CustomerPO,OrderComments
                        'OrderLine ,SKU,SKUDescription,AllocateQTY,Ship QTY,Weight,MSR,AssignedFlag,AssignedDate,HI,TI,
                        
                        'OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_contact1,C_Contact2,C_Company,C_Address1,C_Address2,C_Address3,C_Address4,C_City,
                        'C_State,C_Zip,C_Country,C_ISOCntryCode,C_Phone1,C_Phone2,C_Fax1,C_Fax2,C_vat,BuyerPO, BillToKey, B_contact1, B_Contact2, B_Company, B_Address1, B_Address2, B_Address3, B_Address4, B_City, B_State, B_Zip, B_Country,
                        'B_ISOCntryCode, B_Phone1, B_Phone2, B_Fax1, B_Fax2, B_Vat, IncoTerm, PmtTerm, OpenQty, Status, DischargePlace, DeliveryPlace, IntermodalVehicle, CountryOfOrigin, CountryDestination, UpdateSource, Type, OrderGroup, Door,
                        'Route, Stop, Notes, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, ContainerType, ContainerQty, BilledContainerQty, DoRoute, CustomerOrderkey, PIVNO, Amount
                        str_SQL = "INSERT Logictown.dbo.orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,ConsigneeKey,C_Company,C_Zip,C_Address1,C_Address2,C_Address3,CustomerOrderkey,Notes,UpdateSource) " & _
                            "VALUES ('" & str_Orderkey & "','" & Trim(rs_Src.Fields(0)) & "','" & Trim(rs_Src.Fields(1)) & "','" & Trim(rs_Src.Fields(2)) & "','" & Trim(rs_Src.Fields(3)) & "','" & Trim(rs_Src.Fields(4)) & "', " & _
                            "'" & str_CustomerID & "','" & Trim(rs_Src.Fields(6)) & "','" & Trim(rs_Src.Fields(7)) & "','" & Trim(rs_Src.Fields(8)) & "','" & Trim(rs_Src.Fields(9)) & "'," & _
                            "'" & Trim(rs_Src.Fields(10)) & "','" & Trim(rs_Src.Fields(11)) & "','" & Trim(rs_Src.Fields(12)) & "','" & Trim(filLocalFile.FileName) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        int_Order = int_Order + 1
                    cn.CommitTrans
                End If
                tmp_Rs.Close
           End If
           '�������--�P�_�q����ӬO�_����
           Call Confirm_Recordset_Closed(tmp_Rs)
           str_SQL = "select ExternOrderKey from Logictown.dbo.ORDERDETAIL where ExternOrderKey='" & Trim(rs_Src.Fields(0)) & "' and OrderLineNumber= '" & Trim(rs_Src.Fields(13)) & "'"
           tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
           If Not tmp_Rs.EOF Then
                int_Repeat = int_Repeat + 1
                GoTo nextloop
           End If
           '�q����Ӹ�Ʒs�W
           cn.BeginTrans
                'OrderNo,OrderType,Division,OrderDate,DeliveryDate,CustomerID,CustomerName,ZIP,Address1,Address2,Address3,CustomerPO,OrderComments
                'OrderLine ,SKU,SKUDescription,AllocateQTY,Ship QTY,Weight,MSR,AssignedFlag,AssignedDate,HI,TI,
                
                'OrderKey, OrderLineNumber, OrderDetailSysId, ExternOrderKey, ExternLineNo, Sku, StorerKey, ManufacturerSku, RetailSku, AltSku, OriginalQty, OpenQty, ShippedQty, AdjustedQty, QtyPreAllocated, QtyAllocated, QtyPicked, UOM, PackKey, PickCode, CartonGroup, Lot, ID, Facility, Status, UnitPrice, Tax01, Tax02, ExtendedPrice, UpdateSource, Lottable01, Lottable02, Lottable03, Lottable04, Lottable05, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, TariffKey, Lottable06, Lottable07, Lottable08, Lottable09, Lottable10, Lottable11, Beginqty
            str_SQL = "INSERT Logictown.dbo.ORDERDETAIL (OrderKey,OrderLineNumber, ExternOrderKey,Sku,StorerKey,OriginalQty,Lottable01, Lottable02,PackKey)" & _
                     "VALUES ('" & str_Orderkey & "','" & Trim(rs_Src.Fields(13)) & "','" & Trim(rs_Src.Fields(0)) & "','" & Trim(rs_Src.Fields(14)) & "','" & Trim(rs_Src.Fields(2)) & "', " & _
                     "'" & Trim(rs_Src.Fields(16)) / 1000 & "','" & Trim(rs_Src.Fields(20)) & "','" & Trim(rs_Src.Fields(21)) & "','" & Trim(rs_Src.Fields(14)) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            cn.CommitTrans
           '�������--�P�_SKU�_�s�b
            str_SQL = "select sku from Logictown.dbo.sku where sku='" & Trim(rs_Src.Fields(14)) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF Then
                cn.BeginTrans
                    'StorerKey, Sku, DESCR, SUSR1, SUSR2, SUSR3, SUSR4, SUSR5, MANUFACTURERSKU, RETAILSKU, ALTSKU, PACKKey, STDGROSSWGT, STDNETWGT, STDCUBE, TARE, CLASS, ACTIVE, SKUGROUP, Tariffkey, BUSR1, BUSR2, BUSR3, BUSR4, BUSR5, LOTTABLE01LABEL, LOTTABLE02LABEL, LOTTABLE03LABEL, LOTTABLE04LABEL, LOTTABLE05LABEL, NOTES1, NOTES2, PickCode, StrategyKey, CartonGroup, PutCode, PutawayLoc, PutawayZone, InnerPack, Cube, GrossWgt, NetWgt, ABC, CycleCountFrequency, LastCycleCount, ReorderPoint, ReorderQty, StdOrderCost, CarryCost, Price, Cost, ReceiptHoldCode, ReceiptInspectionLoc, TrafficCop, ArchiveCop, IOFlag, TareWeight, LotxIdDetailOtherlabel1, LotxIdDetailOtherlabel2, LotxIdDetailOtherlabel3, AvgCaseWeight, TolerancePct, SkuRotat01, SkuRotat02, SkuRotat03, SkuRotat04, SkuRotat05, DefaultRotation, AllocParm, onreceiptcopypackkey, replenishmentQty
                    str_CSKU = StringCleaner(Trim(rs_Src.Fields(15)), "'")
                    str_SQL = "INSERT Logictown.dbo.sku (StorerKey, Sku, DESCR,PACKKey,STDGROSSWGT,BUSR4) " & _
                        "VALUES ('" & Trim(rs_Src.Fields(2)) & "','" & Trim(rs_Src.Fields(14)) & "','" & str_CSKU & "','" & Trim(rs_Src.Fields(14)) & "','" & Trim(rs_Src.Fields(18)) / 10000 & "','" & Trim(rs_Src.Fields(19)) / 10000 & "' )"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    'PackKey, PackDescr, PackUOM1, CaseCnt, ISWHQty1, ReplenishUOM1, ReplenishZone1, CartonizeUOM1, LengthUOM1, WidthUOM1, HeightUOM1, CubeUOM1, PackUOM2, InnerPack, ISWHQty2, ReplenishUOM2, ReplenishZone2, CartonizeUOM2, LengthUOM2, WidthUOM2, HeightUOM2, CubeUOM2, PackUOM3, Qty, ISWHQty3, ReplenishUOM3, ReplenishZone3, CartonizeUOM3, LengthUOM3, WidthUOM3, HeightUOM3, CubeUOM3, PackUOM4, Pallet, ISWHQty4, ReplenishUOM4, ReplenishZone4, CartonizeUOM4, LengthUOM4, WidthUOM4, HeightUOM4, CubeUOM4, PalletWoodLength, PalletWoodWidth, PalletWoodHeight, PalletTI, PalletHI, PackUOM5, Cube, ISWHQty5, PackUOM6, GrossWgt, ISWHQty6, PackUOM7, NetWgt, ISWHQty7, PackUOM8, OtherUnit1, ISWHQty8, ReplenishUOM8, ReplenishZone8, CartonizeUOM8, LengthUOM8, WidthUOM8, HeightUOM8, PackUOM9, OtherUnit2, ISWHQty9, ReplenishUOM9, ReplenishZone9, CartonizeUOM9, LengthUOM9, WidthUOM9, HeightUOM9, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, Timestamp
                    str_SQL = "INSERT Logictown.dbo.pack (PACKKey,PackDescr,GrossWgt, PalletHI, PalletTI,CaseCnt) " & _
                        "VALUES ('" & Trim(rs_Src.Fields(14)) & "','" & str_CSKU & "','" & Trim(rs_Src.Fields(18)) / 10000 & "', " & _
                         "'" & Trim(rs_Src.Fields(22)) & "','" & Trim(rs_Src.Fields(23)) & "','1')"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    
                cn.CommitTrans
            Else
                cn.BeginTrans
                    str_CSKU = StringCleaner(Trim(rs_Src.Fields(15)), "'")
                    str_SQL = "UPDATE Logictown.dbo.sku set StorerKey='" & Trim(rs_Src.Fields(2)) & "',  DESCR ='" & str_CSKU & "',PACKKey='" & Trim(rs_Src.Fields(14)) & "' " & _
                        "where Sku='" & Trim(rs_Src.Fields(14)) & "'"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    
                    str_SQL = "UPDATE Logictown.dbo.pack set PackDescr='" & str_CSKU & "',GrossWgt='" & Trim(rs_Src.Fields(18)) / 10000 & "',CubeUOM1='" & Trim(rs_Src.Fields(19)) / 10000 & "' " & _
                        ", PalletHI='" & Trim(rs_Src.Fields(22)) & "', PalletTI='" & Trim(rs_Src.Fields(23)) & "',CaseCnt='1' " & _
                        "where PACKKey='" & Trim(rs_Src.Fields(14)) & "'"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    
                cn.CommitTrans
            End If
            tmp_Rs.Close
            int_OrderLine = int_OrderLine + 1
nextloop:
            rs_Src.MoveNext
        Loop
        '�ƥ��ɮ�
        Dim fl_file As Scripting.File

        Set fso = New FileSystemObject
        If fso.FileExists(strTranFileName) = True Then
            Set fl_file = fso.GetFile(strTranFileName)
            fl_file.copy ("C:\from_ids\backup\Alc\" & str_file)
            If fso.FileExists("C:\from_ids\backup\Alc\" & str_file) = True Then
                fl_file.Delete
            End If
        End If
        filLocalFile.Refresh
        '�R�Xftp�W���ɮ�
        ITC.Execute , "DELETE " & Chr(34) & str_file & Chr(34)
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        lstRemoteFile.Clear
        ITC.Execute , "DIR"
        lblStatus = "�w�s�u�malc��Ƨ�"
        
        If int_Repeat > 0 Then
            msg_text = "��" & int_Repeat & "���q����ӭ�������"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        End If
        msg_text = "�פJ" & int_Order & "���q�� " & int_OrderLine & "������,��r�ɳƥ���C:\from_ids\backup\Alc\"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
    End If
End Sub

'Private Sub imgReceiveFile3_Click()
'Call Command3_Click
''Call cmdImport3_Click
'End Sub

Private Sub imgSendFile_Click()
Dim i As Double
'If the ITC is not still executing then send the file
If ITCReady(True) = True Then
    'Check that a file has been selected
    If Trim(filLocalFile.FileName) = "" Then
        MsgBox "���I��n�W�Ǫ��ɮ�", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    'Check that the file does not already exist on the server
    For i = 0 To lstRemoteFile.ListCount
        If filLocalFile.FileName = lstRemoteFile.List(i) Then
            If MsgBox("�ɮ� " & filLocalFile.FileName & " �w�g�s�b" & vbCrLf & "�n�л\��?", vbQuestion + vbYesNo, "Overwrite") = vbNo Then
                Exit Sub
            End If
        End If
    Next i
           
    'Send the file and update the remote file list box
    ITC.Execute , "PUT " & Chr(34) & filLocalFile.Path & "\" & filLocalFile.FileName & Chr(34) & " " & Chr(34) & filLocalFile.FileName & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    lstRemoteFile.Clear
    ITC.Execute , "DIR"
    lblStatus = "�w�s�u"
End If
End Sub

Private Sub ITC_StateChanged(ByVal State As Integer)
    'Check the state of the itc, and change the status accordingly
    Dim Data1, RemoteFiles, i As Double
    Dim RemoteFileName As String
    Select Case State
        Case icResolvingHost
            lblStatus = "Finding Host IP Address"
        Case icHostResolved
            lblStatus = "IP Address Found"
        Case icConnecting
            lblStatus = "Connecting To Host"
        Case icConnected
            lblStatus = "�w�s�u"
        Case icRequesting
            lblStatus = "Sending Request"
        Case icRequestSent
            lblStatus = "Request Sent"
        Case icReceivingResponse
            lblStatus = "Receiving Response"
        Case icResponseReceived
            lblStatus = "Response Receiving"
        Case icDisconnecting
            lblStatus = "Disconnecting"
        Case icDisconnected
            lblStatus = "Not Connected"
        Case icError
            If ITC.ResponseCode = 12030 Then
                lblStatus = "Not Connected"
                cmdLogOn.Enabled = False
                cmdNewFolder.Enabled = False
                cmdDelete.Enabled = False
                'cmdRename.Enabled = False
                'cmdSize.Enabled = False
                cmdUpFolder.Enabled = False
                imgSendFile.Enabled = False
                imgReceiveFile.Enabled = False
                lstRemoteFile.Enabled = False
                cmdLogOff.Enabled = False
                cmdLogOn.Enabled = True
                ITC.Cancel
            End If
            If ITC.ResponseCode <> 87 Then
                MsgBox ITC.ResponseCode & " " & ITC.ResponseInfo, vbOKOnly + vbCritical, "Error"
            End If
        Case icResponseCompleted
            'loop until you get all data
            Do While True
                Data1 = ITC.GetChunk(4096, icString)
                If Len(Data1) = 0 Then Exit Do
                DoEvents
                RemoteFiles = RemoteFiles & Data1
            Loop
            'Beep
            'If it is recieving size data tell the user the size and then exit the sub
            If RecievingSize Then
                MsgBox "The size of file " & lstRemoteFile.Text & " is " & RemoteFiles & " bytes", vbInformation + vbOKOnly, "Size"
                Exit Sub
            End If
            'Loop through, check for carriage returns to get each file name and add to listbox
            For i = 1 To Len(RemoteFiles)
                If Mid(RemoteFiles, i, 1) = Chr(13) Then
                    If Trim(RemoteFileName) <> "" Then
                        lstRemoteFile.AddItem RemoteFileName
                        RemoteFileName = ""
                    End If
                Else
                    If Mid(RemoteFiles, i, 1) <> Chr(10) Then
                        RemoteFileName = RemoteFileName & Mid(RemoteFiles, i, 1)
                    End If
                End If
            Next i
    End Select
End Sub

'Private Sub int3_StateChanged(ByVal State As Integer)
'    'Check the state of the itc, and change the status accordingly
'    Dim Data1, RemoteFiles
'    Dim RemoteFileName As String
'    Select Case State
'        Case icResolvingHost
'            lblStatus3 = "Finding Host IP Address"
'        Case icHostResolved
'            lblStatus3 = "IP Address Found"
'        Case icConnecting
'            lblStatus3 = "Connecting To Host"
'        Case icConnected
'            lblStatus3 = "�w�s�u"
'        Case icRequesting
'            lblStatus3 = "Sending Request"
'        Case icRequestSent
'            lblStatus3 = "Request Sent"
'        Case icReceivingResponse
'            lblStatus3 = "Receiving Response"
'        Case icResponseReceived
'            lblStatus3 = "Response Receiving"
'        Case icDisconnecting
'            lblStatus3 = "Disconnecting"
'        Case icDisconnected
'            lblStatus3 = "Not Connected"
'        Case icError
'            If int3.ResponseCode = 12030 Then
'                lblStatus3 = "Not Connected"
'                cmdLogon3.Enabled = False
''                cmdNewFolder.Enabled = False
''                cmdDelete.Enabled = False
'                'cmdRename.Enabled = False
'                'cmdSize.Enabled = False
''                cmdUpFolder.Enabled = False
''                imgSendFile.Enabled = False
'                imgReceiveFile3.Enabled = False
'                lstRemoteFile3.Enabled = False
'                cmdLogOff3.Enabled = False
'                cmdLogon3.Enabled = True
'                int3.Cancel
'            End If
'            If int3.ResponseCode <> 87 Then
'                MsgBox int3.ResponseCode & " " & int3.ResponseInfo, vbOKOnly + vbCritical, "Error"
'            End If
'        Case icResponseCompleted
'            'loop until you get all data
'            Do While True
'                Data1 = int3.GetChunk(4096, icString)
'                If Len(Data1) = 0 Then Exit Do
'                DoEvents
'                RemoteFiles = RemoteFiles & Data1
'            Loop
'            'Beep
'            'If it is recieving size data tell the user the size and then exit the sub
'            If RecievingSize Then
'                MsgBox "The size of file " & lstRemoteFile3.Text & " is " & RemoteFiles & " bytes", vbInformation + vbOKOnly, "Size"
'                Exit Sub
'            End If
'            'Loop through, check for carriage returns to get each file name and add to listbox
'            For i = 1 To Len(RemoteFiles)
'                If Mid(RemoteFiles, i, 1) = Chr(13) Then
'                    If Trim(RemoteFileName) <> "" Then
'                        lstRemoteFile3.AddItem RemoteFileName
'                        RemoteFileName = ""
'                    End If
'                Else
'                    If Mid(RemoteFiles, i, 1) <> Chr(10) Then
'                        RemoteFileName = RemoteFileName & Mid(RemoteFiles, i, 1)
'                    End If
'                End If
'            Next i
'    End Select
'End Sub

Public Function GetWord(ByVal strData As String, ByRef intStart As Integer, ByVal intLen As Integer) As String
    Dim intloop, z As Integer
    Dim strTemp As String
    '�r�ꤣ���ɸɪť�
    If Len(strData) < intStart + intLen Then
        strData = Left(strData & String(intStart + intLen, " "), intStart + intLen)
    End If
    
    intloop = 0
    z = Len(strData)
    Do While intloop <= intLen - 1
        strTemp = Mid(strData, intStart + intloop, 1)
        If intloop = intLen - 1 Then        '�P�_�̫�@�X�O�_������,�]����r��e�����^��ɤ��Ϋ�i��|�h�@��
            If Asc(strTemp) < 0 Then        '�p�G�r���O����
                intLen = intLen - 1
                GetWord = GetWord & " "     '�r�ꪽ���[�@��ť�,���A�[����
            Else
                GetWord = GetWord + strTemp
            End If
            intloop = intloop + 1
        Else
            If Asc(strTemp) < 0 Then
                intLen = intLen - 1                         '�p�G�r���O����
            End If
            GetWord = GetWord + strTemp
            intloop = intloop + 1
        End If
            
    Loop
    intStart = intStart + intloop
End Function

Public Function StringCleaner(S As String, Search As String) As String
    Dim i As Integer, res As String
    res = S
    Do While InStr(res, Search)
        i = InStr(res, Search)
        res = Left(res, i - 1) & Mid(res, i + 1)
    Loop
    StringCleaner = res
End Function

Private Sub FTPlog(ByVal strActionName As String)
    '���͵{������O��
    Dim fso As Scripting.FileSystemObject
    Dim ts_LogFile As Scripting.TextStream
    Dim strTmp As String
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then MkDirs App.Path & "\Log"
    
    '���o�{������O���ɡA�Y���s�b�h�۰ʷs�W
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(App.Path & "\Log\Import.log") Then
       Set ts_LogFile = fso.OpenTextFile(App.Path & "\log\Import.log", ForAppending)  'open TextStream Object
    Else
       Set ts_LogFile = fso.CreateTextFile(App.Path & "\log\Import.log", True)       'create TextStream Object
    End If
    '�g�J���A��
    strTmp = Format(Now, "yyyy-mm-dd ttttt") & "�A" & strActionName & " �פJ�� : " & User_id
    
    ts_LogFile.WriteLine (strTmp)
    ts_LogFile.WriteBlankLines (1)
    ts_LogFile.Close
    Set ts_LogFile = Nothing
    Set fso = Nothing
End Sub

Private Sub cmd_Alc_Click()
    lstRemoteFile.Clear
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    ITC.Execute , "CD " & Chr(34) & "/Bestg/Alc" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    'ITC.Execute , "LS"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    lblStatus = "�w�s�u��alc��Ƨ�"
End Sub

Private Sub cmd_shp_Click()
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    lstRemoteFile.Clear
    ITC.Execute , "CD " & Chr(34) & "/Bestg/shp" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    'ITC.Execute , "LS"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    lblStatus = "�w�s�u��shp��Ƨ�"
End Sub

Private Sub cmd_CFM_Click()
    lstRemoteFile.Clear
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    ITC.Execute , "CD " & Chr(34) & "/Bestg/CFM" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    'ITC.Execute , "LS"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    lblStatus = "�w�s�u��cfm��Ƨ�"
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "FTP�W�U��"
End Sub

Private Sub Form_Load()
    '�]�w Form �j�p�B��m
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 11475
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    Me.txtServer = "10.190.1.2"
    'Me.txtServer = "202.145.173.225"
    Me.txtUserName = "NLJL"
    Me.txtPassword = "NLJL"
    
'    txtServer3 = "chutney.unilever.com"
'    chutney.unilever.com
'    txtUserName3 = "jean.chang"
'    txtPassWord3 = "re7ated"
    
    '�p�G�O�t�κ޲z���A�h�}�ҥߨ�
    If UCase(User_id) = "ADMINISTRATOR" Then
        SSTab1.Tab = 5: SSTab1.Caption = "�ߨ��h�f�q��"
        SSTab1.Tab = 11: SSTab1.Caption = "�ߨ��q��פJ"
    End If
    SSTab1.Tab = 18
    
    If Dir("C:\LTKK01\Orders", vbDirectory) = "" Then MkDirs "C:\LTKK01\Orders"
    If Dir("C:\BEST\LVTL01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LVTL01\Orders"
    If Dir("C:\BEST\LFYY01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LFYY01\Orders"
    If Dir("C:\BEST\LKAO01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LKAO01\Orders"
    If Dir("C:\BEST\LSJR01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LSJR01\Orders"
    If Dir("C:\BEST\LNSL01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders"
    If Dir("C:\BEST\LNIP01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\Orders"
    If Dir("C:\BEST\Other\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\Other\Orders"
       
    drvLocalDrive.Drive = "C:": dirLocalDir.Path = "C:\LTKK01\Orders"
    drvLocalDriveT2.Drive = "C:": dirLocalDirT2.Path = "C:\BEST\LVTL01\Orders"
    drvLocalDriveT3.Drive = "C:": dirLocalDirT3.Path = "C:\BEST\LFYY01\Orders"
    drvLocalDriveT4.Drive = "C:": dirLocalDirT4.Path = "C:\BEST\LKAO01\Orders"
    drvLocalDriveT5.Drive = "C:": dirLocalDirT5.Path = "C:\BEST\LNIP01\Orders"
    drvLocalDriveT6.Drive = "C:": dirLocalDirT6.Path = "C:\BEST\LSJR01\Orders"
    drvLocalDriveT7.Drive = "C:": dirLocalDirT7.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT8.Drive = "C:": dirLocalDirT8.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT9.Drive = "C:": dirLocalDirT9.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT10.Drive = "C:": dirLocalDirT10.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT11.Drive = "C:": dirLocalDirT11.Path = "C:\BEST\LNIP01\Orders"
    drvLocalDriveT12.Drive = "C:": dirLocalDirT12.Path = "C:\BEST\LKAO01\Orders"
    drvLocalDriveT13.Drive = "C:": dirLocalDirT13.Path = "C:\BEST\Other\Orders"
    drvLocalDriveT14.Drive = "C:": dirLocalDirT14.Path = "C:\BEST\LNSL01\Orders"
'    drvLocalDriveT15.Drive = "C:": dirLocalDirT15.Path = "C:\BEST\LMYS01\Orders"
'    drvLocalDriveT16.Drive = "C:": dirLocalDirT16.Path = "C:\BEST\LMYS01\Orders"
'    drvLocalDriveT17.Drive = "C:": dirLocalDirT17.Path = "C:\BEST\LAPP01\Orders"
'    drvLocalDriveT18.Drive = "C:": dirLocalDirT18.Path = "C:\BEST\LAPP01\Orders"
    '���X�Ҧ��f�D���--TRP16M
    Dim tmp_cnt As Integer
    cboStorerkeyT13.Clear
    str_SQL = "Select Rtrim(StorerKey) as 'StorerKey',Isnull(Rtrim(Short_Name),'') as 'StorerName' From TRP16M where storer_status > 0 Order by StorerKey"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
    Do While Not tmp_Rs.EOF
       cboStorerkeyT13.AddItem tmp_Rs.Fields("StorerKey")
    tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    
    cboStorerkeyT13 = "LHPT01"

    RecievingSize = False
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > SSTab1.Top + Frame1.Top + Frame1.Height + 500 Then
    SSTab1.Height = (Me.ScaleHeight)
    dg_CustInv.Height = SSTab1.Height - SSTab1.Top - 1200
    dgMainT2.Height = SSTab1.Height - Frame1.Top - Frame1.Height - 240
    dgMainT3.Height = SSTab1.Height - Frame2.Top - Frame2.Height - 240
    dgMainT4.Height = SSTab1.Height - Frame7.Top - Frame7.Height - 240
    dgMainT5.Height = SSTab1.Height - Frame8.Top - Frame8.Height - 240
    dgMainT6.Height = SSTab1.Height - Frame9.Top - Frame9.Height - 240
    dgMainT7.Height = SSTab1.Height - Frame10.Top - Frame10.Height - 240
    dgMainT8.Height = SSTab1.Height - Frame11.Top - Frame11.Height - 240
    dgMainT9.Height = SSTab1.Height - Frame12.Top - Frame12.Height - 240
    dgMainT10.Height = SSTab1.Height - Frame13.Top - Frame13.Height - 240
    dgMainT11.Height = SSTab1.Height - Frame3.Top - Frame3.Height - 240
    dgMainT12.Height = SSTab1.Height - Frame4.Top - Frame4.Height - 240
    dgMainT13.Height = SSTab1.Height - Frame5.Top - Frame5.Height - 240
    dgMainT14.Height = SSTab1.Height - Frame6.Top - Frame6.Height - 240
    dgMainT15.Height = SSTab1.Height - Frame14.Top - Frame14.Height - 240
    dgMainT16.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT16_1.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT16_2.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT16_3.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT17.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 240
    dgMainT17_1.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 240
    dgMainT18.Height = SSTab1.Height - Frame17.Top - Frame17.Height - 240
    dgMainT19.Height = SSTab1.Height - Frame18.Top - Frame18.Height - 240
    dgMainT20.Height = SSTab1.Height - Frame19.Top - Frame19.Height - 240
    dgMainT21.Height = SSTab1.Height - Frame20.Top - Frame18.Height - 240
    dgMainT22.Height = SSTab1.Height - Frame21.Top - Frame19.Height - 240
End If

If Me.ScaleHeight > SSTab2.Top + Frame1.Top + Frame1.Height + 500 Then
    SSTab2.Height = (Me.ScaleHeight)
    dgMainT15.Height = SSTab1.Height - Frame14.Top - Frame14.Height - 800
    dgMainT16.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_1.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_2.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_3.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT17.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT17_1.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT18.Height = SSTab1.Height - Frame17.Top - Frame17.Height - 800
    dgMainT19.Height = SSTab1.Height - Frame18.Top - Frame18.Height - 800
    dgMainT20.Height = SSTab1.Height - Frame19.Top - Frame19.Height - 800
    dgMainT21.Height = SSTab1.Height - Frame20.Top - Frame18.Height - 800
    dgMainT22.Height = SSTab1.Height - Frame21.Top - Frame19.Height - 800
End If

If Me.ScaleHeight > SSTab3.Top + Frame1.Top + Frame1.Height + 500 Then
    SSTab3.Height = (Me.ScaleHeight)
    dgMainT15.Height = SSTab1.Height - Frame14.Top - Frame14.Height - 800
    dgMainT16.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_1.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_2.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_3.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT17.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT17_1.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT18.Height = SSTab1.Height - Frame17.Top - Frame17.Height - 800
    dgMainT19.Height = SSTab1.Height - Frame18.Top - Frame18.Height - 800
    dgMainT20.Height = SSTab1.Height - Frame19.Top - Frame19.Height - 800
    dgMainT21.Height = SSTab1.Height - Frame20.Top - Frame20.Height - 800
    dgMainT22.Height = SSTab1.Height - Frame21.Top - Frame21.Height - 800
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab1.Width = (Me.ScaleWidth - 240)
    dg_CustInv.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT2.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT3.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT4.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT5.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT6.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT7.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT8.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT9.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT10.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT11.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT12.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT13.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT14.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT15.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16_1.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16_2.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16_3.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT17.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT17_1.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT18.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT19.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT20.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT21.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT22.Width = SSTab1.Width - SSTab1.Left - 600
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab2.Width = (Me.ScaleWidth - 850)
    dgMainT15.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_2.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_3.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT18.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT19.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT20.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT21.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT22.Width = SSTab1.Width - SSTab1.Left - 750
    
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab3.Width = (Me.ScaleWidth - 850)
    dgMainT15.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_2.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_3.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT18.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT19.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT20.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT21.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT22.Width = SSTab1.Width - SSTab1.Left - 750
End If
End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
ITC.Cancel
'int3.Cancel
'Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel any tasks that the itc is doing
ITC.Cancel
'int3.Cancel
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_FTP = Nothing
Set rsMainT2 = Nothing
Set rsMainT3 = Nothing
Set rsMainT4 = Nothing
Set rsMainT5 = Nothing
Set rsMainT6 = Nothing
Set rsMainT7 = Nothing
Set rsMainT8 = Nothing
Set rsMainT9 = Nothing
Set rsMainT10 = Nothing
Set rsMainT11 = Nothing
Set rsMainT12 = Nothing
Set rsMainT13 = Nothing
Set rsMainT14 = Nothing
Set rsMainT15 = Nothing
Set rsMainT16 = Nothing
Set rsMainT16_1 = Nothing
Set rsMainT16_2 = Nothing
Set rsMainT16_3 = Nothing
Set rsMainT17 = Nothing
Set rsMainT18 = Nothing
Set rsMainT19 = Nothing
Set rsMainT20 = Nothing
Set rsMainT21 = Nothing
Set rsMainT22 = Nothing
End Sub

Private Function ITCReady(ShowMessage As Boolean)
'Check the state of itc, if it is not executing return true
If ITC.StillExecuting Then
    ITCReady = False
    If ShowMessage Then
        MsgBox "�еy��.  FTP���A�����椤", vbInformation + vbOKOnly, "���L��"
    End If
Else
    ITCReady = True
End If
End Function

'Private Function int3Ready(ShowMessage As Boolean)
''Check the state of itc, if it is not executing return true
'If int3.StillExecuting Then
'    int3Ready = False
'    If ShowMessage Then
'        MsgBox "�еy��.  FTP���A�����椤", vbInformation + vbOKOnly, "���L��"
'    End If
'Else
'    int3Ready = True
'End If
'End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
If Left(Trim(SSTab1.Caption), 2) = "--" Then SSTab1.Tab = PreviousTab
filLocalFile.Refresh
dirLocalDirT3.Refresh
If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab

End Sub
Private Sub cmdImportT2_Click()
'�|������q����ӡA�]��VTL���X�f���|��ܩ��ӡA��L�f�D���|

strTranFileName = filLocalFileT2.Path & "\" & filLocalFileT2.FileName
If Len(Trim(filLocalFileT2.FileName)) = 0 Then Exit Sub
On Error GoTo err_Handle

Dim strLineTmp As String, i As Integer, j As Integer, k As Integer, arrLen, blDuplicationOrder As Boolean, Intcheck As Integer, Str_check As String
cmdImportT2.Enabled = False: dgMainT2.Enabled = False: Screen.MousePointer = 11
Intcheck = 0
Str_check = ""

Set rsMainT2 = Nothing

    arrLen = Array(11, 9, 9, 9, 9, 16, 37, 6, 9, 9, 9, 2, 11, 3, 3, 12, 31, 3, 9, 17, 255)

    Open filLocalFileT2.Path & "\" & filLocalFileT2.FileName For Input As #1
    
    Set rsMainT2 = New ADODB.Recordset
    With rsMainT2
        .Fields.Append "�X�f�渹", adChar, arrLen(0), adFldUpdatable
        .Fields.Append "�b�ګȤ�", adChar, arrLen(1), adFldUpdatable
        .Fields.Append "�b�ګȤ�W��", adChar, arrLen(2), adFldUpdatable
        .Fields.Append "�e�f�Ȥ�", adChar, arrLen(3), adFldUpdatable
        .Fields.Append "�e�f�Ȥ�W��", adChar, arrLen(4), adFldUpdatable
        .Fields.Append "�Ȥ�q��", adChar, arrLen(5), adFldUpdatable
        .Fields.Append "�Ȥ�a�}", adChar, arrLen(6), adFldUpdatable
        .Fields.Append "���ӥN��", adChar, arrLen(7), adFldUpdatable
        .Fields.Append "���ӦW��", adChar, arrLen(8), adFldUpdatable
        .Fields.Append "�w�X���", adChar, arrLen(9), adFldUpdatable
        .Fields.Append "�ƥX���", adChar, arrLen(10), adFldUpdatable
        .Fields.Append "�ϥδ̪O", adChar, arrLen(11), adFldUpdatable
        .Fields.Append "����", adDouble, arrLen(12), adFldUpdatable
        .Fields.Append "����", adUnsignedSmallInt, arrLen(13), adFldUpdatable
        .Fields.Append "�X�f��]", adChar, arrLen(14), adFldUpdatable
        .Fields.Append "���~�s��", adChar, arrLen(15), adFldUpdatable
        .Fields.Append "���~�W��", adChar, arrLen(16), adFldUpdatable
        .Fields.Append "���", adChar, arrLen(17), adFldUpdatable
        .Fields.Append "�ƶq", adDouble, arrLen(18), adFldUpdatable
        .Fields.Append "�Ȥ�渹", adChar, arrLen(19), adFldUpdatable
        .Fields.Append "�Ƶ�", adChar, arrLen(20), adFldUpdatable
        .Fields.Append "���O", adChar, 10, adFldUpdatable
        .Fields.Append "�ܧO", adChar, 18, adFldUpdatable
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
  
        '�}���ɮ�
        k = 0
        Do While Not EOF(1)
                k = k + 1
                Line Input #1, strLineTmp
                If Len(RTrim(strLineTmp)) > 135 Then
                    .AddNew
                    j = 1
                    
                    For i = 0 To rsMainT2.Fields.Count - 3
                        .Fields(i) = RTrim(GetWord(strLineTmp, j, arrLen(i)))
                    Next
                    
                    '�q�����O-�O�_���J�w
                    rsMainT2("���O") = "I": rsMainT2("�ܧO") = "R01"
                    If UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW327" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW328" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DP1324" Then rsMainT2("���O") = "RC": rsMainT2("�ܧO") = "R01"
                    If UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW427" Then rsMainT2("���O") = "RC": rsMainT2("�ܧO") = "R01-C"
                    If UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW815" Then rsMainT2("���O") = "RC": rsMainT2("�ܧO") = "R01-S"
                    
                    '�q��X�w�ܧO
                    If RTrim(rsMainT2("���O")) = "I" Then
                        If UCase(Trim(rsMainT2("���ӥN��"))) = "W3270" Or UCase(Trim(rsMainT2("���ӥN��"))) = "W3280" Or UCase(Trim(rsMainT2("���ӥN��"))) = "P1324" Or UCase(Trim(rsMainT2("���ӥN��"))) = "P2324" Then rsMainT2("�ܧO") = "R01" '�_�ܥX�f
                        If UCase(Trim(rsMainT2("���ӥN��"))) = "W4270" Then rsMainT2("�ܧO") = "R01-C" '���ܥX�f
                        If UCase(Trim(rsMainT2("���ӥN��"))) = "W8150" Then rsMainT2("�ܧO") = "R01-S" '�n�ܥX�f
                    End If
                End If
                
        Loop
            Close #1
        
           .MoveFirst
    
    End With
    rsMainT2.Sort = "�X�f�渹,���ӥN��,����"
    Set dgMainT2.DataSource = rsMainT2
    
    With dgMainT2
    
    For i = 0 To rsMainT2.Fields.Count - 1
    .Columns(i).Caption = rsMainT2.Fields(i).Name
    Next
    
        .ColumnHeaders = True        '���D�����
        .RowHeight = 300

    End With
    
    SetDataGridColWidth Me.Caption, dgMainT2
'
'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders(nolock) where storerkey = 'LVTL01' and rtrim(updatesource)='" & filLocalFileT2.FileName & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: Exit Sub

'�q�����ˬd add by Eric
rsMainT2.MoveFirst

Do While Not rsMainT2.EOF
    '����ˬdDZ,EA
    If UCase(Trim(rsMainT2("���"))) <> "DZ" And UCase(Trim(rsMainT2("���"))) <> "EA" Then
        MsgBox "�q�榳EA,DZ�H�~�����A�нT�{�ɮ׮榡�O�_���T�C", vbOKOnly, Me.Caption: cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: Exit Sub
    End If
    
    '��f����ˬd
    If Trim(rsMainT2("�w�X���")) < Format(Now, "YYYYMMDD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: Exit Sub

    '�������--�P�_SKU�O�_�s�b
    str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where Storerkey = 'LVTL01' and sku = '" & Trim(rsMainT2("���~�s��")) & "'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT2("���~�s��")) & " ) " & Trim(rsMainT2("���~�W��")) & "�A�q����J�פ�!!": cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: Exit Sub
    End If

    '�������--�P�_�O�_�ݨըƹF�q��
    If UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW327" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW328" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW427" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW815" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DP1324" Or UCase(Trim(rsMainT2("���ӥN��"))) = "W3270" Or UCase(Trim(rsMainT2("���ӥN��"))) = "W3280" Or UCase(Trim(rsMainT2("���ӥN��"))) = "W4270" Or UCase(Trim(rsMainT2("���ӥN��"))) = "W8150" Or UCase(Trim(rsMainT2("���ӥN��"))) = "WA500" Or UCase(Trim(rsMainT2("���ӥN��"))) = "WB500" Or UCase(Trim(rsMainT2("���ӥN��"))) = "WD500" Or UCase(Trim(rsMainT2("���ӥN��"))) = "P1324" Or UCase(Trim(rsMainT2("���ӥN��"))) = "P2324" Then
    Else
        MsgBox "�Ȥ�渹�G" & Trim(rsMainT2("�X�f�渹")) & vbCrLf & "�гq���Ȥ�A�T�{�ӵ��q��O�_���~!?", vbOKOnly, "�o�{�D�ըƹF���q����"
        cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: Exit Sub
    End If

    '�s�Ȥ��ˬd1
    str_SQL = "select storerkey from trp01m(nolock) where storerkey = 'LVTL01' and consigneekey = '" & Trim(rsMainT2("�b�ګȤ�")) & "'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close
        MsgBox "�o�{�s�Ȥ�A�q����J����!", vbOKOnly, Me.Caption
        Exit Sub
    End If

    '�s�Ȥ��ˬd2
    str_SQL = "select storerkey from trp01m(nolock) where storerkey = 'LVTL01' and consigneekey = '" & Trim(rsMainT2("�e�f�Ȥ�")) & "'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
            cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close
            MsgBox "�o�{�s�Ȥ�:" & Trim(rsMainT2("�e�f�Ȥ�")) & "�A�q����J����!", vbOKOnly, Me.Caption
            Exit Sub
    End If
    
    '���q�涵��
    If Trim(rsMainT2("�X�f�渹")) <> Str_check Then
        Str_check = Trim(rsMainT2("�X�f�渹"))
        Intcheck = 1
        If Val(Trim(rsMainT2("����"))) <> Intcheck Then
                cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0
                MsgBox "�o�{�q�涵�����~�A�q����J����!", vbOKOnly, Me.Caption
                Exit Sub
        End If
        Intcheck = Intcheck + 1
    Else
        If Val(Trim(rsMainT2("����"))) <> Intcheck Then
                cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0
                MsgBox "�o�{�q�涵�����~�A�q����J����!", vbOKOnly, Me.Caption
                Exit Sub
        End If
        Intcheck = Intcheck + 1
    End If
    rsMainT2.MoveNext
Loop

'�}�l�פJ
Tran_Level = cn.BeginTrans
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, intQTY As Long, strFacility As String

rsMainT2.MoveFirst
Do While Not rsMainT2.EOF
DoEvents: DoEvents

'�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
If strOrderNo <> UCase(Trim(rsMainT2("�X�f�渹"))) Then
    strOrderNo = UCase(Trim(rsMainT2("�X�f�渹")))
    blDuplicationOrder = False

    '�������--�P�_�q��O�_���ơA���Ƥ��W�[
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select orderkey from orders where rtrim(ExternOrderKey) ='" & Trim(rsMainT2("�X�f�渹")) & "' and storerkey = 'LVTL01' and isnull(type,'') <> '�R��' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then

        If RTrim(rsMainT2("���O")) = "RC" Then int_Asn = int_Asn + 1
        '���q�渹�X
        str_SQL = "select isnull(max(orderkey),0) from orders"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
        tmp_Rs.Close

        strFacility = UCase(Trim(rsMainT2("���ӦW��")))
        If UCase(Trim(rsMainT2("���ӥN��"))) = "W3270" Or UCase(Trim(rsMainT2("���ӥN��"))) = "W3280" Or UCase(Trim(rsMainT2("���ӥN��"))) = "P1324" Or UCase(Trim(rsMainT2("���ӥN��"))) = "P2324" Then strFacility = "�ըƹF�_��"
        If UCase(Trim(rsMainT2("���ӥN��"))) = "W4270" Then strFacility = "�ըƹF����"
        If UCase(Trim(rsMainT2("���ӥN��"))) = "W8150" Then strFacility = "�ըƹF�n��"
        If UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW327" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW328" Or UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DP1324" Then strFacility = "�ըƹF�_��"
        If UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW427" Then strFacility = "�ըƹF����"
        If UCase(Trim(rsMainT2("�e�f�Ȥ�"))) = "DW815" Then strFacility = "�ըƹF�n��"
        
        If Len(Trim(rsMainT2("�w�X���"))) = 0 Then rsMainT2("�w�X���") = Format(Now() + 1, "YYYYMMDD")

        str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,Stop,Door,Facility,ConsigneeKey,billtokey,c_company,b_contact1,c_phone1,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,amount,addwho,editwho) " & _
        "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT2("�X�f�渹")) & "','" & rsMainT2("���O") & "','LVTL01','" & Trim(rsMainT2("�ƥX���")) & "','" & Trim(rsMainT2("�w�X���")) & "','" & Trim(rsMainT2("���ӥN��")) & "','" & Trim(rsMainT2("���ӦW��")) & "','" & strFacility & "', " & _
        "'" & Trim(rsMainT2("�e�f�Ȥ�")) & "','" & Trim(rsMainT2("�b�ګȤ�")) & "','" & Trim(rsMainT2("�e�f�Ȥ�W��")) & "','" & Trim(rsMainT2("�b�ګȤ�W��")) & "','" & Trim(rsMainT2("�Ȥ�q��")) & "',substring('" & Trim(rsMainT2("�Ȥ�a�}")) & "', 1, 60),substring('" & Trim(rsMainT2("�Ȥ�a�}")) & "', 61, 45),'" & Trim(rsMainT2("�Ȥ�渹")) & "','" & Trim(rsMainT2("�Ƶ�")) & "','" & filLocalFileT2.FileName & "','','" & Trim(rsMainT2("����")) & "','" & User_id & "','" & User_id & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        int_Order = int_Order + 1
    Else
        '�q�歫��
        Call FTPlog("�q�歫��" & str_SQL)
        '��������
        strReOrderkey = strReOrderkey & Trim(rsMainT2("�X�f�渹")) & Trim(rsMainT2("����")) & "','"
        blDuplicationOrder = True

    End If
End If

'    '�������--�P�_�q����ӬO�_���ơA���Ƥ��W�[���ӡA���U�@�����
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    str_SQL = "select o.orderkey from ORDERDETAIL od (nolock) join orders o (nolock) on o.orderkey = od.orderkey where od.editdate >getdate()-1 and rtrim(o.ExternOrderKey) + rtrim(od.OrderLineNumber) ='" & Trim(rsMainT2("�X�f�渹")) & Trim(rsMainT2("����")) & "' and o.storerkey = 'LVTL01' and isnull(o.type,'') <> '�R��' "
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.EOF And blDuplicationOrder = False Then
    If blDuplicationOrder = False Then
         intQTY = Trim(rsMainT2("�ƶq"))
         If UCase(Trim(rsMainT2("���"))) = "DZ" Then intQTY = Trim(rsMainT2("�ƶq")) * 12

        '�q����Ӹ�Ʒs�W
        str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber, ExternlineNO ,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice ,CartonGroup,notes)" & _
        "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT2("����")) & "','" & Trim(rsMainT2("�X�f��]")) & "','" & Trim(rsMainT2("�X�f�渹")) & "','" & Trim(rsMainT2("���~�s��")) & "','LVTL01'," & _
        "'" & intQTY & "','" & intQTY & "','" & rsMainT2("�ܧO") & "','','" & Trim(rsMainT2("���")) & "','0','" & Trim(rsMainT2("�ϥδ̪O")) & "','" & Trim(rsMainT2("�Ƶ�")) & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        int_OrderLine = int_OrderLine + 1
'
'    Else
'        '�q����ӭ���
'        Call FTPlog("�q����ӭ���" & str_SQL)
'        '��������
'        strReOrderkey = strReOrderkey & Trim(rsMainT2("�X�f�渹")) & Trim(rsMainT2("����")) & "','"
'    End If
    End If
    
    rsMainT2.MoveNext
Loop

'�ɫȤ���
cn.Execute "exec gs_ordersupdate 'LVTL01' ", RowsAffect, adExecuteNoRecords


'��f�q���� 1.���J�ƨ��t�θ�� 2.���͹w�����ʳ�
Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = 3

'orderlinenumber �g�Jpodetail.externpokey�A���w�ťո`�٤l�d�߳t�� edit by Eric 20140311
str_SQL = "select od.storerkey " & _
            ",o.orderkey " & _
            ", od.externorderkey " & _
            ", o.priority " & _
            ", orderlinenumber = ' '" & _
            ", od.sku " & _
            ", s.descr " & _
            ", openqty = sum(od.openqty) " & _
            ", notes = cast(o.notes as varchar(300)) " & _
            ", o.consigneekey " & _
            ", o.c_company " & _
            ", ContainerKey = case when len(rtrim(o.customerorderkey)) = 0 then od.externorderkey else o.customerorderkey end " & _
            "from orders o (nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey " & _
            "join " & strWMSDB & "..sku s(nolock) on s.sku = od.sku and s.storerkey = od.storerkey " & _
            "where o.storerkey = 'LVTL01' and o.priority = 'RC' and o.B_PHONE2 is null " & _
            "group by od.storerkey ,o.orderkey , od.externorderkey , o.priority ,od.storerkey ,o.orderkey , od.externorderkey , o.priority , o.consigneekey , o.c_company ,od.sku, s.descr,cast(o.notes as varchar(300)) , o.customerorderkey " & _
            "order by od.storerkey , o.orderkey "
            
rsTmp.Open str_SQL, cn
If Not rsTmp.EOF Then

    Dim strKeycount As String, strOrderkey As String, intLineNumber As Integer
        
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
    
        '�g�JWMS
        If Trim(rsTmp("orderkey")) <> strOrderkey Then
            intLineNumber = 1
            strOrderkey = Trim(rsTmp("orderkey"))
    
            '���t��PO�渹
            Dim rsKeycount As New ADODB.Recordset
            rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
            '�渹+1
            cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
            strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
            rsKeycount.Close: Set rsKeycount = Nothing
    
            '�g�J���Y
            str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,BuyersReference ,  BuyerVAT , sellername,selleraddress1,externpokey,potype,notes) " & _
                        "values( '" & strKeycount & "','" & rsTmp("StorerKey") & "','" & rsTmp("ExternOrderKey") & "','" & RTrim(rsTmp("ContainerKey")) & "','" & rsTmp("consigneekey") & "','" & rsTmp("C_company") & "','" & rsTmp("OrderKey") & "','" & rsTmp("priority") & "','" & rsTmp("notes") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '���O�w��ƨ��t��
            cn.Execute "update orders set B_PHONE2='00',trafficCop=null where orderkey = '" & rsTmp("OrderKey") & "' ", RowsAffect, adExecuteNoRecords
            
        End If
    
            '�g�J��
            str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered) " & _
                    "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & rsTmp("OrderLineNumber") & "','" & rsTmp("SKU") & "','" & rsTmp("descr") & "','" & rsTmp("StorerKey") & "','" & rsTmp("openqty") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
            intLineNumber = intLineNumber + 1
    
        rsTmp.MoveNext

    Loop
    
End If
rsTmp.Close: Set rsTmp = Nothing

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�ɮ� " & filLocalFileT2.FileName & " �ƥ��� C:\BEST\LVTL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�ɮ�:" & strTranFileName)
    
    If int_Asn > 0 Then MsgBox "�� " & int_Asn & " ����f�q������J!", vbOKOnly + vbInformation, Me.Caption: Call FTPlog("�פJ " & int_Asn & " ����f�q���q��A�ɮ�:" & strTranFileName)
    If int_Repeat > 0 Then MsgBox "�� " & int_Repeat & " ���q����ӭ�������!", vbOKOnly + vbInformation, Me.Caption: Call FTPlog("�פJ " & int_Repeat & " �����ƭq����ӡA�ɮ�:" & strTranFileName)

'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT2.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) + rtrim(od.OrderLineNumber) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LVTL01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LVTL01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LVTL01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT2.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LVTL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LVTL01\Orders\Backup"
If Dir("C:\BEST\LVTL01\Orders\Backup\" & filLocalFileT2.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LVTL01\Orders\Backup\" & filLocalFileT2.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LVTL01\Orders\Backup\" & mySplit(filLocalFileT2.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & ".txt"
End If

Kill strTranFileName
    
filLocalFileT2.Refresh
Screen.MousePointer = 0: cmdImportT2.Enabled = True: dgMainT2.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmd_Import_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT2.Enabled = True: Screen.MousePointer = 0: dgMainT2.Enabled = True

End Sub

Private Sub dgMainT2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT2
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, SSTab1.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT2_Change()
    filLocalFileT2.Path = dirLocalDirT2.Path
End Sub

Private Sub drvLocalDriveT2_Change()

On Error GoTo DriveError
dirLocalDirT2.Path = drvLocalDriveT2.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub
Private Sub cmdImportT4_Click()

strTranFileName = filLocalFileT4.Path & "\" & filLocalFileT4.FileName
If Len(RTrim(cboSheetT4)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT4.EOF Or rsMainT4 Is Nothing Then Exit Sub
Dim strStorerkey As String
On Error GoTo err_Handle

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT4.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'�f�D�s��
strStorerkey = "LKAO01"
cmdImportT4.Enabled = False: dgMainT4.Enabled = False

'��f����ˬd,�P�_SKU�O�_�s�b
rsMainT4.MoveFirst
Do While Not rsMainT4.EOF
'
'    If Replace(myExCharFilter(Trim(rsMainT4("��f��"))), ".", "/") < Format(Now - 1, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub

    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & myExCharFilter(Trim(rsMainT4("�ӫ~�N��"))) & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "�q��o�{�s�~�� (" & myExCharFilter(Trim(rsMainT4("�ӫ~�N��"))) & " ) " & Trim(rsMainT4("�ӫ~�W��")) & "�A�q����J�פ�!!": cmdImportT4.Enabled = True: dgMainT4.Enabled = True
        tmp_Rs.Close
        Exit Sub

'        '�s�WSKU
'        '�ˬdPackkey�j��10�X���sPackkey
'        If Len(strSku) > 10 Then
'
'            '��Packkey�y����
'            Call Confirm_Recordset_Closed(tmp_rs)
'            'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strstorerkey & "' and left(consigneekey,4) = '" & strstorerkey & "' order by consigneekey desc "
'            'tmp_rs.Open str_SQL, cn
'            '
'            'If Not tmp_rs.EOF Then intTmp = Val(tmp_rs("consigneekey"))
'            '
'            'tmp_rs.Close
'
'        Else
'            strPackkey = strSku
'        End If
'
'        lngCasecnt = Val(rsMainT4("�e�f PCs")) / Val(rsMainT4("�q��/�X�f�ƶq"))
'        dblStdCube = Round(Val(rsMainT4("��n")) / Val(rsMainT4("�e�f PCs")) / 28316, 10)
'        dblStdGrossWGT = Val(rsMainT4("�b��")) / Val(rsMainT4("�e�f PCs"))
'
'        str_SQL = "insert into sku(Storerkey,SKU,SKUGROUP,DESCR,STDCUBE,STDGROSSWGT,SUSR1,SUSR2,SUSR3,SUSR4,SUSR5,BUSR1,BUSR2,BUSR3,BUSR4,BUSR5,Packkey,AllocParm,DefaultRotation,IOFlag,PickCode,PutAwayLoc,PutCode,PutAwayZone,ReceiptInspectionLoc,SKURotat01,StrategyKey,LOTTABLE01LABEL,LOTTABLE02LABEL,LOTTABLE03LABEL,LOTTABLE04LABEL,LOTTABLE05LABEL,LOTTABLE06LABEL,LOTTABLE07LABEL,LOTTABLE08LABEL,LOTTABLE09LABEL,LOTTABLE10LABEL,LOTTABLE11LABEL) Values " & _
'                  "('LKAO01','" & strSku & "','STD000N',convert(char(60),'" & myExCharFilter(Trim(rsMainT4("�ӫ~�W��"))) & "')," & dblStdCube & "," & dblStdGrossWGT & ",'',0,'',5000,1,'" & myExCharFilter(Trim(rsMainT4("BUn"))) & "','','" & myExCharFilter(Trim(rsMainT4("SU"))) & "','','','" & strPackkey & "','FLOAT PICK','FIFO','N','NSPFIFO','UNKNOWN','NSPPASTD','RACK','QC','LOTTABLE05','ZONEA','Pack Key','��B�渹','�Ͳ��帹','�s�y��','�����','�ܧO','�̪OID','�̪O���O','','','') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'        '�s�WPACK
'        str_SQL = "insert into pack(Packkey,Packdescr,PackUOM1,Casecnt,LengthUOM1,WidthUOM1,HeightUOM1,CubeUOM1,PackUOM2,Innerpack,PackUOM3,Qty,LengthUOM3,WidthUOM3,HeightUOM3,CubeUOM3,PackUOM4,Pallet,PalletTI,PalletHI,ADDDate,ADDWho,EditDate,EditWho,replenishzone1,replenishzone2,replenishzone3,replenishzone4,replenishzone8,replenishzone9,CartonizeUOM3) Values " & _
'                  "('" & strSku & "','" & strPackkey & "','CS','" & lngCasecnt & "',0,0,0,0,'IP',0,'EA',1,0,0,0,0,'PL',0,0,0,getdate(),'SA',getdate(),'SA','CASE','PICK','PICK','PICK','PICK','PICK','Y') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    End If
    tmp_Rs.Close

rsMainT4.MoveNext
Loop

Tran_Level = cn.BeginTrans
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, dblStdCube As Double, dblStdGrossWGT As Double, lngCasecnt As Long, intPointer As Integer
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strOrderKeyS As String, strPackkey As String, strSku As String, strAddress As String

Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
Screen.MousePointer = 11
            
'���̫�Ȥ�s��
'Call Confirm_Recordset_Closed(tmp_rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strstorerkey & "' and left(consigneekey,4) = '" & strstorerkey & "' order by consigneekey desc "
'tmp_rs.Open str_SQL, cn
'
'If Not tmp_rs.EOF Then intTmp = Val(tmp_rs("consigneekey"))
'
'tmp_rs.Close

rsMainT4.MoveFirst
Do While Not rsMainT4.EOF
'    DoEvents: DoEvents
    
    intPointer = 1
    strSku = myExCharFilter(Trim(rsMainT4("�ӫ~�N��")))
    strAddress = myExCharFilter(Trim(rsMainT4("����"))) & myExCharFilter(Trim(rsMainT4("��}"))) & myExCharFilter(Trim(rsMainT4("���P")))
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT4("�q�ʳ渹"))) Then
        strOrderNo = UCase(Trim(rsMainT4("�q�ʳ渹")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
'        '�L���Ȥ�s���s�W
'        str_SQL = "select * from trp01m where storerkey = '" & strStorerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT4("���f�Ȥ�"))) & "' "
'        Call Confirm_Recordset_Closed(tmp_rs)
'        tmp_rs.CursorLocation = 3
'        tmp_rs.Open str_SQL, cn
'
'        If tmp_rs.EOF Then
'
'            '�s�W�Ȥ�D��
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
'            " values('" & strStorerkey & "','','" & myExCharFilter(Trim(rsMainT4("���f�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT4("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT4("�Ȥ�W��"))) & "','','','" & myExCharFilter(Trim(rsMainT4("����"))) & myExCharFilter(Trim(rsMainT4("���P"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
'        End If
    
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT4("�q�ʳ渹"))) & "' and storerkey = '" & strStorerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
                     
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,B_company,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT4("�q�ʳ渹"))) & "','C','" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT4("�B�e���"))) & "','" & myExCharFilter(Trim(rsMainT4("��f��"))) & "','','" & _
            myExCharFilter(Trim(rsMainT4("���f�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT4("���f�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT4("�Ȥ�W��"))) & "','','','','" & GetWord(strAddress, intPointer, 58) & "','" & GetWord(strAddress, intPointer, 45) & "','','','" & filLocalFileT4.FileName & "','','" & User_id & "','" & User_id & "','' )"

'            '����q��s�Wcustomerorderkey,notes
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,B_company,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT4("�q�ʳ渹"))) & "','A2B','" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT4("�B�e���"))) & "','" & myExCharFilter(Trim(rsMainT4("��f��"))) & "','','" & _
'            myExCharFilter(Trim(rsMainT4("���f�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT4("���f�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT4("�Ȥ�W��"))) & "','','','','" & GetWord(strAddress, intPointer, 58) & "','" & GetWord(strAddress, intPointer, 45) & "','" & myExCharFilter(Trim(rsMainT4("���ʳ渹"))) & "','" & myExCharFilter(Trim(rsMainT4("�Ƶ�"))) & "','" & filLocalFileT4.FileName & "','','" & User_id & "','" & User_id & "','' )"
'
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & myExCharFilter(Trim(rsMainT4("�q�ʳ渹"))) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Val(myExCharFilter(Trim(rsMainT4("�e�f PCs"))))
            
            strLot06 = "R01"
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT4("�q�ʳ渹"))) & "','" & myExCharFilter(Trim(rsMainT4("�ӫ~�N��"))) & "','" & strStorerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT4("BUn"))) & "','0','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If

        rsMainT4.MoveNext
Loop

'�s�Ȥ��ˬd
str_SQL = "select �f�D=storerkey , �f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �Ȥ�W��=c_company , �ˬd��� = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel���
    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
    If Dir("C:\BEST\" & strStorerkey & "\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\�ʫȤ���"
    MyXlsApp.Range("h:h").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT4.Enabled = True: dgMainT4.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "�o�{�s�Ȥ�A�q����J����!", vbOKOnly, Me.Caption
    Exit Sub
End If

'�ɫȤ���
cn.Execute "exec gs_ordersupdate '" & strStorerkey & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA(�@ " & rsMainT4.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT4.FileName & " �ƥ��� C:\BEST\" & strStorerkey & "\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�ɮ� " & filLocalFileT4.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , �����ɮצW�� = '" & filLocalFileT4.FileName & "' , �W���ɮצW�� = o.updatesource ,���ƭq�渹�X = rtrim(o.externorderkey) ,�W���Ȥ�渹 = rtrim(o.customerorderkey) ,  �W���q���� = convert(varchar,o.orderdate,111) , �W����f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �W���Ƹ� = od.sku , �W���ƶq = od.openqty ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & strStorerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\Orders\Backup"
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT4.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT4.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & mySplit(filLocalFileT4.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT4.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT4.Refresh: cboSheetT4.Clear
Screen.MousePointer = 0: cmdImportT4.Enabled = True: dgMainT4.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportT4_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT4.Enabled = True: Screen.MousePointer = 0: dgMainT4.Enabled = True

End Sub

Private Sub dgMainT4_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT4
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT4_Change()
    filLocalFileT4.Path = dirLocalDirT4.Path
End Sub
Private Sub drvLocalDriveT4_Change()

On Error GoTo DriveError
dirLocalDirT4.Path = drvLocalDriveT4.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT4_Click()

On Error GoTo err_Handle
Set rsMainT4 = Nothing: Set dgMainT4.DataSource = rsMainT4
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT4.Path, 1) = "\" Then
    strFilePath = filLocalFileT4.Path
Else
    strFilePath = filLocalFileT4.Path & "\"
End If

If Dir(strFilePath & filLocalFileT4.FileName) = "" Then: filLocalFileT4.Refresh: Exit Sub

cboSheetT4.Clear

If UCase(mySplit(filLocalFileT4.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT4.FileName)
  
    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT4.Clear
    
    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT4.AddItem MyXlsApp.Sheets(i).Name
        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT4.ListIndex = -1
    
    MyXlsApp.DisplayAlerts = False:  MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT4.FileName
    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT4.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT4_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String, IntI As Integer
IntI = 1
'�T�{���|�O�_�a"\"
If Right(filLocalFileT4.Path, 1) = "\" Then
    strFilePath = filLocalFileT4.Path
Else
    strFilePath = filLocalFileT4.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = ""

If Right(filLocalFileT4.Path, 1) <> "\" Then
    strFilePath = filLocalFileT4.Path & "\"
Else
    strFilePath = filLocalFileT4.Path
End If

''�إ����W�ٰ}�C
'strFieldName = "�q�ʳ渹" & Chr(9) & "�X�f�渹" & Chr(9) & "�B�e���" & Chr(9) & "�B�e���" & Chr(9) & "��f��" & Chr(9) & "���f�Ȥ�" & Chr(9) & "�Ȥ�W��" & Chr(9) & "����" & Chr(9) & "��}" & Chr(9) & "���P" & Chr(9) & "�ӫ~�N��" & Chr(9) & "�ӫ~�W��" & Chr(9) & "�q��/�X�f�ƶq" & Chr(9) & "SU" & Chr(9) & "�e�fPCs" & Chr(9) & "BUn" & Chr(9) & "�b��" & Chr(9) & "WUn" & Chr(9) & "��n" & Chr(9) & "VUn" & Chr(9) & "�w�O" & Chr(9)
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
Set rsMainT4 = New ADODB.Recordset
'Call Excel2RecordsetT4(strFilePath & filLocalFileT4.FileName, cboSheetT4, strFieldName, rsMainT4)
Call Excel2Recordset(strFilePath & filLocalFileT4.FileName, cboSheetT4, strFieldName, tmp_Rs)
tmp_Rs.MoveFirst

Call Replication_Recordset(tmp_Rs, rsMainT4)
tmp_Rs.Close


Set dgMainT4.DataSource = rsMainT4

If rsMainT4 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT4
    MsgBox "���u�@��@ " & rsMainT4.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

rsMainT4.Sort = "�q�ʳ渹,�s��"  'add by Eric ���ӹq�l�ɪ����ǥh�Ƨ�

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Sub Excel2RecordsetT4(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'Create by Gemini @20090312 4 Excel�פJRecordset
'�ϥλ���
'1.�p�G�ӷ�Excel�u�@���a���W�١A�Щ�strFieldName���w�A�åHchar(9)�@�����j�Ÿ�
'strFieldName = "�Ȥ�" & Chr(9) & "�νs" & Chr(9) & "�P�f���" & Chr(9) & "�P�f�渹" & Chr(9) & "�p���H" & Chr(9) & "�q��" & Chr(9) & "�e�f�a�}" & Chr(9) & "�o�����X" & Chr(9) & "�~�ȭ�" & Chr(9) & "�Ƶ�" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�ܮw" & Chr(9) & "�ƶq" & Chr(9) & "���" & Chr(9) & "�e�m���/�Ƶ�/�Ȥ�渹" & Chr(9)

'�Ѽƻ���
'strFileName:�ӷ��ɮצW�ٸ��|
'strSheetName:�ӷ��u�@��
'strFieldName:���W��
'rs:�^�Ǫ�Recordset
'�d��
'call Excel2Recordset ("C:\book1.xls","Sheet1", "�Ȥ�" & Chr(9) & "�νs" & Chr(9) & "�P�f���" & Chr(9),rsMain)
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '�䤣���ɮ�

On Error GoTo err_Handle
Screen.MousePointer = 11

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '���Ĥ@�Ӥu�@��W��
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '�䤣����w�u�@��A��βĤ@��
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(1, i) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        k = 2 '�ѲĤG�C�}�l�פJ
    End If
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp)
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    k = 7 '��7�C�}�l
    
    '�g�JRecordset
    Do While Len(RTrim(.Cells(k, 2))) > 0
    rsTmp.AddNew
        For j = 2 To UBound(arrTmp) + 2 ''��B7�x�s��}�l
            rsTmp(j - 2) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & Chr(64 + j) & k & ")�A��ƬO�_���~�I"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub
Private Sub cmdImportT5_Click()

On Error GoTo err_Handle
strTranFileName = filLocalFileT5.Path & "\" & filLocalFileT5.FileName
If Len(RTrim(cboSheetT5)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT5.EOF Or rsMainT5 Is Nothing Then Exit Sub

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT5.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT5.Enabled = True: dgMainT5.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT5.MoveFirst
Do While Not rsMainT5.EOF

    '��f����ˬd
    arrTmp = Split(Trim(rsMainT5("�P�f���")), "/")
    If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub
    
    '�ƶq�ˬd
    If Trim(rsMainT5("�ƶq")) > 0 Then
        MsgBox "�o�{�q��ƶq�j��0�A" & Trim(rsMainT5("�~��")) & "-" & Trim(rsMainT5("�~�W")) & "(" & Trim(rsMainT5("�ƶq")) & Trim(rsMainT5("���")) & ")�A�q����J�פ�!!", , "�h�f��פJ": Exit Sub
        Exit Sub
    End If
    
    rsMainT5.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT5.Enabled = False: dgMainT5.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'���̫�Ȥ�s��
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LNIP01' and left(consigneekey,4) = 'LNIP' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT5.MoveFirst
Do While Not rsMainT5.EOF
    DoEvents: DoEvents
    
'    '�������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If UCase(Trim(rsMaint5("�ܮw"))) = "" Then
''        MsgBox "�Ȥ�渹�G" & Trim(rsMainT4("�P�f�渹")) & "( " & Trim(rsMainT4("�ܮw")) & " )" & vbCrLf & "�гq���Ȥ�A�T�{�ӵ��q��O�_���~!?", vbOKOnly, "�D�ըƹF���q�椣��J"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '�������--�P�_SKU�O�_�s�b
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT5("�~��")) & "' and Storerkey = 'LNIP01' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cn.RollbackTrans: Tran_Level = 0
        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT5("�~��")) & " ) " & Trim(rsMainT5("�~�W")) & "�A�q����J�פ�!!": cmdImportT5.Enabled = True: dgMainT5.Enabled = True: Screen.MousePointer = 0
        Exit Sub
    End If

'    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT5("�P�f�渹"))) Then
        strOrderNo = UCase(Trim(rsMainT5("�P�f�渹")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�ˬd�O�_�����Ȥ�W��
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LNIP01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT5("�Ȥ�"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '�L���Ȥ�W�٫h�s�W
            intTmp = intTmp + 1
            strConsigneeKey = "LNIP" & Format(intTmp, "000000")
            
            '�s�W�Ȥ�D��
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT5("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT5("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT5("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT5("�q��"))) & "','" & myExCharFilter(Trim(rsMainT5("�e�f�a�}"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '�����s�W���Ȥ�s��
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = 'LNIP01' and full_name = '" & myExCharFilter(Trim(rsMainT5("�Ȥ�"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT5("�p���H"))) & "' " & _
                        "and rtrim(phone) = '" & myExCharFilter(Trim(rsMainT5("�q��"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT5("�e�f�a�}"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '�p���H�B�q�ܻP��f�a�}����
                intTmp = intTmp + 1
                strConsigneeKey = "LNIP" & Format(intTmp, "000000")
                
                '�s�W�Ȥ�D��
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes) " & _
                " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT5("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT5("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT5("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT5("�q��"))) & "','" & myExCharFilter(Trim(rsMainT5("�e�f�a�}"))) & "','" & myExCharFilter(Trim(tmp_Rs("notes"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '�����s�W���Ȥ�s��
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '�۲Ūu���«Ƚs
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
            rsTmp.Close
        End If
        tmp_Rs.Close
    
'        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMaint5("�P�f�渹"))) & "' and storerkey = 'LNIP01' and isnull(type,'') <> '�R��' "
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = Trim(rsMainT5("�ܮw"))
            strFacility = "�ըƹF�_��"
            arrTmp = Split(Trim(rsMainT5("�P�f���")), "/")
            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT5("�P�f�渹"))) & "','R','LNIP01','" & strOrderDate & "','" & CDate(strOrderDate) + 1 & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT5("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT5("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT5("�~�ȭ�"))) & "','" & myExCharFilter(Trim(rsMainT5("�νs"))) & "','" & myExCharFilter(Trim(rsMainT5("�q��"))) & "','','" & myExCharFilter(Trim(GetWord(rsMainT5("�e�f�a�}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT5("�e�f�a�}"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT5("�Ƶ�"))) & "','" & filLocalFileT5.FileName & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT5("�o�����X"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LNIP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
'        Else
'            '�q�歫��
'            Call FTPlog("�q�歫��" & str_SQL)
'            '��������
'            strReOrderkey = strReOrderkey & Trim(rsMaint5("�P�f�渹")) & "','"
'            blDuplicationOrder = True
'
'        End If
    End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Abs(Trim(rsMainT5("�ƶq")))
            strLot06 = IIf(UCase(Trim(rsMainT5("�ܮw"))) = "A06", "A06-S", Trim(rsMainT5("�ܮw")))
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT5("�P�f�渹"))) & "','" & myExCharFilter(Trim(rsMainT5("�~��"))) & "','LNIP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','" & myExCharFilter(rsMainT5("�ܮw")) & "','" & myExCharFilter(Trim(rsMainT5("���"))) & "','0','" & myExCharFilter(Trim(rsMainT5("�e�m���/�Ƶ�/�Ȥ�渹"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT5.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT5.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT5.FileName & " �ƥ��� C:\BEST\LNIP01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT5.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT5.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\LNIP01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNIP01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LNIP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\Orders\Backup"
If Dir("C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT5.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT5.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & mySplit(filLocalFileT5.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT5.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT5.Refresh: cboSheetT5.Clear
Screen.MousePointer = 0: cmdImportT5.Enabled = True: dgMainT5.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportt5_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT5.Enabled = True: Screen.MousePointer = 0: dgMainT5.Enabled = True

End Sub
'Sub ExcelSheet2Recordset()
'On Error GoTo err_Handle
'Dim strExcel As String, arrTmp, strFilePath As String
'
''�T�{���|�O�_�a"\"
'If Right(filLocalFileT4.Path, 1) = "\" Then
'    strFilePath = filLocalFileT4.Path
'Else
'    strFilePath = filLocalFileT4.Path & "\"
'End If
'
''�إ����W�ٰ}�C
'arrTmp = Array("�Ȥ�", "�νs", "�P�f���", "�P�f�渹", "�p���H", "�q��", "�e�f�a�}", "�o�����X", "�~�ȭ�", "�Ƶ�", "�~��", "�~�W", "�ܮw", "�ƶq", "���", "�e�m���/�Ƶ�/�Ȥ�渹")
'
''�إ� Excel �����Ʈw�s��
'Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
'cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
'cnExcel.ConnectionString = "Data Source=" & filLocalFileT4.Path & "\" & filLocalFileT4.FileName & ";Extended Properties=""Excel 8.0; HDR=No;ReadOnly=True;"""
'cnExcel.Open
'
'Call ReDim_Recordset(tmp_rs)
'
'tmp_rs.CursorLocation = 3
'tmp_rs.Open "select * from [" & cboSheetT4 & "$] ", cnExcel ', adOpenStatic, adLockOptimistic
'tmp_rs.Sort = "F4,F16"
'
'Set rsMainT4 = New ADODB.Recordset
'
''�N�ǤJ tmp_rs ����ƻs�� rsMainT4
'Dim fldcnt As Integer, reccnt As Double
'
''�إ� Recordset �� Table �[�c (�b�O���餤�� ADO Recordset)
'rsMainT4.Fields.Append "�s��", adDouble
'For fldcnt = 0 To tmp_rs.Fields.Count - 1
'    rsMainT4.Fields.Append arrTmp(fldcnt), tmp_rs.Fields(fldcnt).Type, tmp_rs.Fields(fldcnt).DefinedSize
'Next fldcnt
'
'With rsMainT4
'     .CursorType = adOpenStatic
'     .LockType = adLockOptimistic
'     .Open    '���ݳs������
'End With
'
'reccnt = 0
'Do While Not tmp_rs.EOF
'   reccnt = reccnt + 1
'   rsMainT4.AddNew
'   rsMainT4.Fields(0).Value = reccnt
'   For fldcnt = 0 To tmp_rs.Fields.Count - 1
'    rsMainT4.Fields(fldcnt + 1).Value = tmp_rs.Fields(fldcnt).Value & ""
'   Next fldcnt
'   rsMainT4.Update
'   tmp_rs.MoveNext
'Loop
'
'tmp_rs.Close: Set tmp_rs = Nothing
'
'Set dgMainT4.DataSource = rsMainT4: dgMainT4.Visible = False
'
'rsMainT4.MoveFirst
'
'SetDataGridColWidth Me.Caption, dgMainT4
'dgMainT4.RowHeight = 300
'Screen.MousePointer = 0: dgMainT4.Visible = True
'MsgBox "���u�@��@ " & rsMainT4.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "�u�@��}��"
'
'cnExcel.Close: Set cnExcel = Nothing
'
'Exit Sub
'err_Handle:
'Set cnExcel = Nothing
'Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

'End Sub

Private Sub dgMainT5_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT5
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT5_Change()
    filLocalFileT5.Path = dirLocalDirT5.Path
End Sub
Private Sub drvLocalDriveT5_Change()

On Error GoTo DriveError
dirLocalDirT5.Path = drvLocalDriveT5.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT5_Click()

On Error GoTo err_Handle
Set rsMainT5 = Nothing: Set dgMainT5.DataSource = rsMainT5
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If

If Dir(strFilePath & filLocalFileT5.FileName) = "" Then: filLocalFileT5.Refresh: Exit Sub

cboSheetT5.Clear

If UCase(mySplit(filLocalFileT5.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT5.FileName)
    MyXlsApp.DisplayAlerts = False
  
    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT5.Clear
    
    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT5.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT5.ListIndex = -1
    
    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT5.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT5_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "�Ȥ�" & Chr(9) & "�νs" & Chr(9) & "�P�f���" & Chr(9) & "�P�f�渹" & Chr(9) & "�p���H" & Chr(9) & "�q��" & Chr(9) & "�e�f�a�}" & Chr(9) & "�o�����X" & Chr(9) & "�~�ȭ�" & Chr(9) & "�Ƶ�" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�ܮw" & Chr(9) & "�ƶq" & Chr(9) & "���" & Chr(9) & "�e�m���/�Ƶ�/�Ȥ�渹" & Chr(9)

If Right(filLocalFileT5.Path, 1) <> "\" Then
    strFilePath = filLocalFileT5.Path & "\"
Else
    strFilePath = filLocalFileT5.Path
End If

Set rsMainT5 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT5.FileName, cboSheetT5, strFieldName, rsMainT5)

Set dgMainT5.DataSource = rsMainT5

If rsMainT5 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT5
    MsgBox "���u�@��@ " & rsMainT5.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Sub ExcelSheet2RecordsetT5()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If

'�إ� Excel �����Ʈw�s��
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT5.Path & "\" & filLocalFileT5.FileName & ";Extended Properties=""Excel 8.0;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT5 & "$] where len(rtrim(�Ȥ�)) > 0 ", cnExcel ', adOpenStatic, adLockOptimistic
tmp_Rs.Sort = "�P�f�渹,�~��"

Set rsMainT5 = New ADODB.Recordset

'�N�ǤJ tmp_rs ����ƻs�� rsMainT5
Dim fldcnt As Integer, reccnt As Double

'�إ� Recordset �� Table �[�c (�b�O���餤�� ADO Recordset)
rsMainT5.Fields.Append "�s��", adDouble
For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT5.Fields.Append tmp_Rs.Fields(fldcnt).Name, tmp_Rs.Fields(fldcnt).Type, tmp_Rs.Fields(fldcnt).DefinedSize
Next fldcnt

With rsMainT5
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With

reccnt = 0
Do While Not tmp_Rs.EOF
   reccnt = reccnt + 1
   rsMainT5.AddNew
   rsMainT5.Fields(0).Value = reccnt
   For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT5.Fields(fldcnt + 1).Value = myExCharFilter(tmp_Rs.Fields(fldcnt).Value & "")
   Next fldcnt
   rsMainT5.Update
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMainT5.DataSource = rsMainT5: dgMainT5.Visible = False

rsMainT5.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT5

Screen.MousePointer = 0: dgMainT5.Visible = True
MsgBox "���u�@��@ " & rsMainT5.RecordCount & " ������", 64, "�u�@��}��"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT6_Click()

strTranFileName = filLocalFileT6.Path & "\" & filLocalFileT6.FileName
If Len(RTrim(cboSheetT6)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT6.EOF Or rsMainT6 Is Nothing Then Exit Sub
Dim strStorerkey As String, strSku As String, strPackkey As String, lngCasecnt As Long, lngPallet As Long, dblStdCube As Double, dblStdGrossWGT As Double
Dim rsTmp As New ADODB.Recordset
On Error GoTo err_Handle

'�f�D�s��
strStorerkey = "LSJR01"

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select updatesource from orders where rtrim(updatesource)='" & filLocalFileT6.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'�������
rsMainT6.MoveFirst
Do While Not rsMainT6.EOF

    '��f����ˬd
'    If Replace(myExCharFilter(Trim(rsMainT6("�Ȥ���w�e�f��"))), ".", "/") < Format(Now - 1, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub

    '�P�_�X�f�ܬO�_�s�b
    strSku = RTrim(myExCharFilter(Trim(rsMainT6("�X�f��"))))
    str_SQL = "select * from trp01m where consigneekey='" & RTrim(myExCharFilter(Trim(rsMainT6("�X�f��")))) & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "�L�X�f�ܸ�� (" & myExCharFilter(Trim(rsMainT6("�X�f��"))) & ")�A�зs�W��A��J�C ", 16, "�q����J�פ�!!"
        tmp_Rs.Close: Exit Sub
    End If
    tmp_Rs.Close

    '�P�_SKU�O�_�s�b
    strSku = RTrim(myExCharFilter(Trim(rsMainT6("�Ƹ�"))))
    str_SQL = "select * from " & strWMSDB & "..sku where sku='" & strSku & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
'        MsgBox "�q��o�{�s�~�� (" & myExCharFilter(Trim(rsMainT6("�ӫ~�N��"))) & " ) " & Trim(rsMainT6("�ӫ~�W��")) & "�A�q����J�פ�!!"
'        tmp_Rs.Close: Exit Sub

       '�s�WSKU
       '���sPackkey
        Call Confirm_Recordset_Closed(rsTmp)
        str_SQL = "select top 1 substring(packkey,7,20) as packkey from " & strWMSDB & "..sku where storerkey = '" & strStorerkey & "' and left(packkey,6) = '" & strStorerkey & "' order by substring(packkey,7,20) desc "
        rsTmp.Open str_SQL, cn
        
        If rsTmp.EOF Then
            strPackkey = strStorerkey & "0001"
        Else
            strPackkey = strStorerkey & Format(Val(rsTmp("packkey")) + 1, "0000")
        End If
        
        rsTmp.Close

        lngPallet = Val(rsMainT6("�j���Ӽ�")) * Val(rsMainT6("�C�O�c��"))
        lngCasecnt = Val(rsMainT6("�j���Ӽ�"))
        dblStdCube = Round(Val(rsMainT6("�p�����n") / rsMainT6("�j���Ӽ�") / 28316), 10)
        dblStdGrossWGT = Val(rsMainT6("�p��쭫�q") / rsMainT6("�j���Ӽ�"))
        
        '�s�WSKU
        str_SQL = "insert into sku(Storerkey,SKU,SKUGROUP,DESCR,STDCUBE,STDGROSSWGT,SUSR1,SUSR2,SUSR3,SUSR4,SUSR5,BUSR1,BUSR2,BUSR3,BUSR4,BUSR5,Packkey,AllocParm,DefaultRotation,IOFlag,PickCode,PutAwayLoc,PutCode,PutAwayZone,ReceiptInspectionLoc,SKURotat01,StrategyKey,LOTTABLE01LABEL,LOTTABLE02LABEL,LOTTABLE03LABEL,LOTTABLE04LABEL,LOTTABLE05LABEL,LOTTABLE06LABEL,LOTTABLE07LABEL,LOTTABLE08LABEL,LOTTABLE09LABEL,LOTTABLE10LABEL,LOTTABLE11LABEL) Values " & _
                  "('" & strStorerkey & "','" & strSku & "','STD000N',convert(char(60),'" & myExCharFilter(Trim(rsMainT6("���~�W��"))) & "')," & dblStdCube & "," & dblStdGrossWGT & ",'',0,'',5000,1,'EA','','CS','','','" & strPackkey & "','FLOAT PICK','FIFO','N','NSPFIFO','UNKNOWN','NSPPASTD','RACK','QC','LOTTABLE05','ZONEA','Pack Key','��B�渹','�Ͳ��帹','�s�y��','�����','�ܧO','�̪OID','�̪O���O','','','') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

        '�s�WPACK
        str_SQL = "insert into pack(Packkey,Packdescr,PackUOM1,Casecnt,LengthUOM1,WidthUOM1,HeightUOM1,CubeUOM1,PackUOM2,Innerpack,PackUOM3,Qty,LengthUOM3,WidthUOM3,HeightUOM3,CubeUOM3,PackUOM4,Pallet,PalletTI,PalletHI,ADDDate,ADDWho,EditDate,EditWho,replenishzone1,replenishzone2,replenishzone3,replenishzone4,replenishzone8,replenishzone9,CartonizeUOM3) Values " & _
                  "('" & strPackkey & "','" & strPackkey & "_" & strSku & "','CS','" & lngCasecnt & "',0,0,0,0,'IP',0,'EA',1,0,0,0,0,'PL'," & lngPallet & ",0,0,getdate(),'SA',getdate(),'SA','CASE','PICK','PICK','PICK','PICK','PICK','Y') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    Else '��s�ӫ~�D��
    
        lngPallet = Val(rsMainT6("�j���Ӽ�")) * Val(rsMainT6("�C�O�c��"))
        lngCasecnt = Val(rsMainT6("�j���Ӽ�"))
        dblStdCube = Round(Val(rsMainT6("�p�����n") / rsMainT6("�j���Ӽ�") / 28316), 10)
        dblStdGrossWGT = Val(rsMainT6("�p��쭫�q") / rsMainT6("�j���Ӽ�"))
        
        '��ssku
        str_SQL = "update " & strWMSDB & "..sku " & _
                  "set DESCR = '" & myExCharFilter(Trim(rsMainT6("���~�W��"))) & "' " & _
                  ",STDCUBE = " & dblStdCube & " " & _
                  ",STDGROSSWGT = " & dblStdGrossWGT & " " & _
                  "where storerkey = '" & strStorerkey & "' and sku = '" & strSku & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

        '��sPACK
        str_SQL = "update " & strWMSDB & "..pack " & _
        "set Casecnt = " & lngCasecnt & " " & _
        ",Pallet = " & lngPallet & " " & _
        "where packkey = '" & tmp_Rs("packkey") & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    End If

rsMainT6.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT6.Enabled = False: dgMainT6.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, intPointer As Integer
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strOrderKeyS As String, strAddress As String, strNotes As String, strDeleteOrder As String

Dim rsTmp1 As New ADODB.Recordset
            
Screen.MousePointer = 11
            
'���̫�Ȥ�s��
'Call Confirm_Recordset_Closed(tmp_rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strstorerkey & "' and left(consigneekey,4) = '" & strstorerkey & "' order by consigneekey desc "
'tmp_rs.Open str_SQL, cn
'
'If Not tmp_rs.EOF Then intTmp = Val(tmp_rs("consigneekey"))
'
'tmp_rs.Close

rsMainT6.MoveFirst
Do While Not rsMainT6.EOF
    
    '�R���ˬd
    If UCase(myExCharFilter(Trim(rsMainT6("�q�檬�A")))) = "DELETE" Then

    strDeleteOrder = strDeleteOrder + UCase(Trim(rsMainT6("�f�D�q�渹�X"))) & "','"

    GoTo nextLine
    End If
    
    intPointer = 1
    strSku = myExCharFilter(Trim(rsMainT6("�Ƹ�")))
    strAddress = myExCharFilter(Trim(rsMainT6("��f�a�}")))
                     
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT6("�f�D�q�渹�X"))) Then
        strOrderNo = UCase(Trim(rsMainT6("�f�D�q�渹�X")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
           
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT6("�f�D�q�渹�X"))) & "' and storerkey = '" & strStorerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '��TMS�渹
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�L���Ȥ�s���s�W
            str_SQL = "select consigneekey from trp01m where storerkey = '" & strStorerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT6("�Ȥ�s��"))) & "' "
            Call Confirm_Recordset_Closed(rsTmp)
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
    
            If rsTmp.EOF Then
                '�s�W�Ȥ�D��
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Fax,channel_type,Address,updatesource) " & _
                " values('" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT6("�l���ϸ�"))) & "','" & myExCharFilter(Trim(rsMainT6("�Ȥ�s��"))) & "','" & myExCharFilter(Trim(rsMainT6("�Ȥ�W��"))) & "','" & myExCharFilter(Trim(rsMainT6("�Ȥ�²��"))) & "','" & myExCharFilter(Trim(rsMainT6("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT6("�q��"))) & "','" & myExCharFilter(Trim(rsMainT6("�ǯu"))) & "','" & myExCharFilter(Trim(rsMainT6("�q���O"))) & "','" & strAddress & "','" & strOrderKeyS & "') ", RowsAffect, adExecuteNoRecords
            End If
            rsTmp.Close
            
            If Len(myExCharFilter(Trim(rsMainT6("���w��f�ɶ�")))) > 0 Then
                strNotes = "���w��f�ɶ�:" & myExCharFilter(Trim(rsMainT6("���w��f�ɶ�"))) & ";"
            End If
            
            strNotes = strNotes & myExCharFilter(Trim(rsMainT6("�q��Ƶ�")))
            
            '�s�W�q����Y
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,b_company) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT6("�f�D�q�渹�X"))) & "','A2B','" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT6("�q����"))) & "','" & myExCharFilter(Trim(rsMainT6("��f��"))) & "','','" & _
            myExCharFilter(Trim(rsMainT6("�X�f��"))) & "','','','','','','','" & myExCharFilter(Trim(rsMainT6("�Ȥ�渹"))) & "','" & strNotes & "','" & filLocalFileT6.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT6("�Ȥ�s��"))) & "')"
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & myExCharFilter(Trim(rsMainT6("�f�D�q�渹�X"))) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Val(myExCharFilter(Trim(rsMainT6("�q��ƶq")))) * Val(myExCharFilter(Trim(rsMainT6("�j���Ӽ�"))))
            
            strLot06 = "R01"
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternLineNo,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable04,Lottable05,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT6("����"))) & "','" & myExCharFilter(Trim(rsMainT6("�f�D�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT6("�Ƹ�"))) & "','" & strStorerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT6("���w�s�y��"))) & "','" & myExCharFilter(Trim(rsMainT6("���w�����"))) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT6("�p���W��"))) & "','" & myExCharFilter(Trim(rsMainT6("���"))) & "','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If

nextLine:
        rsMainT6.MoveNext
Loop

'�R��q��
If Len(RTrim(strDeleteOrder)) > 0 Then

MsgBox "�Ъ`�N���R��q���I", 64, "�q����J"

str_SQL = "select �f�D=storerkey ,TMS�渹 = rtrim(o.orderkey),�f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �Ȥ�W��=c_company,�q�檬�A = rtrim(o.type) ,�q���ɮ� = '" & filLocalFileT6.FileName & "', �ˬd��� = getdate() " & _
        "from orders o " & _
        "Where o.externorderkey in ('" & strDeleteOrder & "') "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF = False Then
    
        'Excel���
        Call Recordset2Excel("�R��q��", tmp_Rs)
        If Dir("C:\BEST\" & strStorerkey & "\�R��q��", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\�R��q��"
        MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\�R��q��\�R��q��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
        Set MyXlsApp = Nothing: tmp_Rs.Close
    
    End If

End If

''�s�Ȥ��ˬd
'str_SQL = "select �f�D=storerkey , �q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �Ȥ�W��=c_company , �ˬd��� = getdate() " & _
'        "from orders o " & _
'        "Where o.b_phone2 Is Null " & _
'        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
'
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF = False Then
'
'    'Excel���
'    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
'    If Dir("C:\BEST\" & strStorerkey & "\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\�ʫȤ���"
'    MyXlsApp.Range("h:h").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
'    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'    Set MyXlsApp = Nothing: tmp_Rs.Close
'
'    cmdImportT6.Enabled = True: dgMainT6.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
'    MsgBox "�o�{�s�Ȥ�A�q����J����!", vbOKOnly, Me.Caption
'    Exit Sub
'End If

'�ɫȤ���
cn.Execute "exec gs_ordersupdate '" & strStorerkey & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA(�@ " & rsMainT6.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT6.FileName & " �ƥ��� C:\BEST\" & strStorerkey & "\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�ɮ� " & filLocalFileT6.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , �����ɮצW�� = '" & filLocalFileT6.FileName & "' , �W���ɮצW�� = o.updatesource ,���ƭq�渹�X = rtrim(o.externorderkey) ,�W���Ȥ�渹 = rtrim(o.customerorderkey) ,  �W���q���� = convert(varchar,o.orderdate,111) , �W����f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �W���Ƹ� = od.sku , �W���ƶq = od.openqty ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & strStorerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFilet6.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\Orders\Backup"
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT6.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT6.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & mySplit(filLocalFileT6.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT6.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT6.Refresh: cboSheetT6.Clear
Screen.MousePointer = 0: cmdImportT6.Enabled = True: dgMainT6.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportt6_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT6.Enabled = True: Screen.MousePointer = 0: dgMainT6.Enabled = True

End Sub

Private Sub dgMainT6_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT6
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT6_Change()
    filLocalFileT6.Path = dirLocalDirT6.Path
End Sub
Private Sub drvLocalDriveT6_Change()

On Error GoTo DriveError
dirLocalDirT6.Path = drvLocalDriveT6.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT6_Click()

On Error GoTo err_Handle
Set rsMainT6 = Nothing: Set dgMainT6.DataSource = rsMainT6
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT6.Path, 1) = "\" Then
    strFilePath = filLocalFileT6.Path
Else
    strFilePath = filLocalFileT6.Path & "\"
End If

If Dir(strFilePath & filLocalFileT6.FileName) = "" Then: filLocalFileT6.Refresh: Exit Sub

cboSheetT6.Clear

If UCase(mySplit(filLocalFileT6.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT6.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT6.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT6.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT6.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT6.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT6_Click()

On Error GoTo err_Handle

Dim strFilePath As String
If Right(filLocalFileT6.Path, 1) <> "\" Then
    strFilePath = filLocalFileT6.Path & "\"
Else
    strFilePath = filLocalFileT6.Path
End If

Set rsMainT6 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT6.FileName, cboSheetT6, "", rsMainT6)

Set dgMainT6.DataSource = rsMainT6

If rsMainT6 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT6
    MsgBox "���u�@��@ " & rsMainT6.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
rsMainT6.Sort = "�f�D�q�渹�X,����"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
'Private Sub cboSheetT7_Click()
'If blDo = True Then Call ExcelSheet2RecordsetT7
'End Sub


Sub Excel2Recordset2(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset, intRow As Integer)
'**************************************************
'�]�����W�٭��s�w�q�A�ҥH�W�ߦ��Ƶ{���A���F���L�Ĥ@�����W��
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '�䤣���ɮ�

On Error GoTo err_Handle
Screen.MousePointer = 11

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '���Ĥ@�Ӥu�@��W��
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '�䤣����w�u�@��A��βĤ@��
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = intRow '�ѫ��w�C�}�l�פJ
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '�g�JRecordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & Chr(64 + j) & k & ")�A��ƬO�_���~�I"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)
End Sub
Private Sub cboSheetT7_Click()

On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT7.Path, 1) <> "\" Then
    strFilePath = filLocalFileT7.Path & "\"
Else
    strFilePath = filLocalFileT7.Path
End If

'strFieldName = "�q��s��" & Chr(9) & "���f�Ȥ�N��" & Chr(9) & "�f��" & Chr(9) & "�X�f�c��" & Chr(9) & "�X�f�]��" & Chr(9) & "���" & Chr(9) & "�X�f��" & Chr(9) & "�妸" & Chr(9) & "PO" & Chr(9)

Set rsMainT7 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT7.FileName, cboSheetT7, strFieldName, rsMainT7)

Set dgMainT7.DataSource = rsMainT7

If rsMainT7 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"

Else

'rsMainT7.Sort = "��ڸ��X,���~�~��"

    SetDataGridColWidth Me.Caption, dgMainT7
    MsgBox "���u�@��@ " & rsMainT7.RecordCount & "������", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT7_Click()

strTranFileName = filLocalFileT7.Path & "\" & filLocalFileT7.FileName
If Len(RTrim(cboSheetT7)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT7.EOF Or rsMainT7 Is Nothing Then Exit Sub
On Error GoTo err_Handle

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT7.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'�������
rsMainT7.MoveFirst
Do While Not rsMainT7.EOF

    '��f����ˬd
    If Format(myExCharFilter(Trim(rsMainT7("��f���"))), "YYYY/MM/DD") < Format(Now, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub
    
    '�P�_�q��ƬO�_��0
    If Val(myExCharFilter(Trim(rsMainT7("�X�f�ƶq")))) = 0 Then MsgBox "�X�f�ƶq�� 0�A�q����J�פ�!!": Exit Sub
    
    '�P�_SKU�O�_�s�b
    str_SQL = "select sku,casecnt from gv_skuxpack where sku='" & Trim(rsMainT7("���~�~��")) & "' and Storerkey = 'LPSI01' "
'
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0
        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT7("���~�~��")) & ") �A�q����J�פ�!!"
        Exit Sub
    End If

    '�ഫ�v�ˬd
    If Val(rsMainT7("�X�f�ƶq")) * Val(tmp_Rs("casecnt")) <> Val(rsMainT7("�X�f�]��")) Then MsgBox "�q��X�f�c�ƻP�X�f�]�Ƥ���(�c�]�ഫ�v�P�Ȥᤣ�P)�A�q����J�פ�!!", 16, Me.Caption: Exit Sub

    rsMainT7.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT7.Enabled = False: dgMainT7.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngInnerpack As Long
Dim strOrderNo As String, strWHOrderNo As String, strWH As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strLot06 As String, strPickMark As String, strPono As String, strTmp As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT7.MoveFirst

Do While Not rsMainT7.EOF
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT7("��ڸ��X"))) Then
        strOrderNo = UCase(Trim(rsMainT7("��ڸ��X")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select OrderKey from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT7("��ڸ��X"))) & "' and storerkey = 'LPSI01' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close

            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders(OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,BillToKey,CustomerOrderkey,BuyerPO,Notes,UpdateSource,type,addwho,editwho,b_phone1,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT7("��ڸ��X"))) & "','I','LPSI01',getdate(),'" & myExCharFilter(Trim(rsMainT7("��f���"))) & "','" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT7("�Ȥ�N��"))) & "','','','','','','','','" & myExCharFilter(Trim(rsMainT7("���ʭq�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT7("�q�渹�X"))) & "','','" & filLocalFileT7.FileName & "','','" & User_id & "','" & User_id & "','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT7("��ڸ��X")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫�Ƥ��W�[����
        If blDuplicationOrder = False Then
            
            '�W�[����
            int_orderlineNumber = int_orderlineNumber + 1
            
            intQTY = Val(myExCharFilter(Trim(rsMainT7("�X�f�]��")))) 'Val(myExCharFilter(Trim(rsMainT7("�X�f�ƶq"))))
            strLot06 = "FG01"
'            str_Orderkey = StrPadLeft(int_orderlineNumber, 10, 0)
                        
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes,updatesource) " & _
            "VALUES ('" & str_Orderkey & "','" & StrPadLeft(Val(int_orderlineNumber), 5, 0) & "','" & StrPadLeft(Val(int_orderlineNumber), 5, 0) & "','" & myExCharFilter(Trim(rsMainT7("��ڸ��X"))) & "','" & myExCharFilter(Trim(rsMainT7("���~�~��"))) & "','LPSI01'," & _
            "'" & intQTY & "','" & intQTY & "','" & Left(myExCharFilter(Trim(rsMainT7("�妸"))), 8) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT7("���s�X"))) & "','0','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' and OrderLineNumber = '" & StrPadLeft(Val(int_orderlineNumber), 5, 0) & "' "
                       
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If

        strWHOrderNo = UCase(Trim(rsMainT7("��ڸ��X")))
        rsMainT7.MoveNext
        
Loop

'�ɫȤ���
cn.Execute "exec gs_Ordersupdate 'LPSI01' ", RowsAffect, adExecuteNoRecords

'�s�Ȥ��ˬd
str_SQL = "select �f�D=storerkey , �f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �ˬd��� = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel���
    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
    If Dir("C:\BEST\LPSI01\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\BEST\LPSI01\�ʫȤ���"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LPSI01\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "�o�{�s�Ȥ�A�q����J����!", 16, Me.Caption
    rsMainT7.MoveFirst
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT7.FileName & " �ƥ��� C:\BEST\LPSI01\Orders\Backup " & strTmp
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�ɮ� " & filLocalFileT7.FileName & strTmp)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT7.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LPSI01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LPSI01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LPSI01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LPSI01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LPSI01\OrdersBackup"
'FileCopy strTranFileName, "O:\LPSI01\OrdersBackup\" & filLocalFileT7.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LPSI01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LPSI01\Orders\Backup"
If Dir("C:\BEST\LPSI01\Orders\Backup\" & filLocalFileT7.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LPSI01\Orders\Backup\" & filLocalFileT7.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LPSI01\Orders\Backup\" & mySplit(filLocalFileT7.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT7.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT7.Refresh: cboSheetT7.Clear
Screen.MousePointer = 0: cmdImportT7.Enabled = True: dgMainT7.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportT7_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True

End Sub

Private Sub dgMainT7_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT7
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT7_Change()
    filLocalFileT7.Path = dirLocalDirT7.Path
End Sub

Private Sub drvLocalDriveT7_Change()

On Error GoTo DriveError
dirLocalDirT7.Path = drvLocalDriveT7.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT7_Click()

On Error GoTo err_Handle
Set rsMainT7 = Nothing: Set dgMainT7.DataSource = rsMainT7
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT7.Path, 1) = "\" Then
    strFilePath = filLocalFileT7.Path
Else
    strFilePath = filLocalFileT7.Path & "\"
End If

If Dir(strFilePath & filLocalFileT7.FileName) = "" Then: filLocalFileT7.Refresh: Exit Sub

cboSheetT7.Clear

If UCase(mySplit(filLocalFileT7.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT7.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT7.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT7.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT7.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT7.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT8_Click()
If blDo = True Then Call ExcelSheet2RecordsetT8
End Sub
Private Sub ExcelSheet2RecordsetT8()

Dim strFilePath As String
If Right(filLocalFileT8.Path, 1) <> "\" Then
    strFilePath = filLocalFileT8.Path & "\"
Else
    strFilePath = filLocalFileT8.Path
End If

Set rsMainT8 = New ADODB.Recordset
Call Excel2RecordsetT8(strFilePath & filLocalFileT8.FileName, cboSheetT8, "Document.date" & Chr(9) & "Requested.delivery.d" & Chr(9) & "Sold-to.party" & Chr(9) & "Sold.to.name" & Chr(9) & "Ship.to" & Chr(9) & "Ship.to.name" & Chr(9) & "PO.number" & Chr(9) & "Sales.document" & Chr(9) & "Delivery.number" & Chr(9) & "Billing.Doc.no" & Chr(9) & "Order.Type" & Chr(9) & "Material" & Chr(9) & "Material.Description" & Chr(9) & "Order.Quantity" & Chr(9) & "Order.Confirmed.Quan" & Chr(9) & "Sales.unit" & Chr(9) & "Order.Reason" & Chr(9) & "Description" & Chr(9) & "Reason.for.Rejection" & Chr(9) & "Created.By" & Chr(9) & "Remarks" & Chr(9), rsMainT8)
rsMainT8.Sort = "Sales.document"

Set dgMainT8.DataSource = rsMainT8

If rsMainT8 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"

Else
    SetDataGridColWidth Me.Caption, dgMainT8
    MsgBox "���u�@��@ " & rsMainT8.RecordCount & "������", 64, "Excel2Recordset"

End If

End Sub

Sub ExcelSheet2RecordsetT8_old()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT8.Path, 1) = "\" Then
    strFilePath = filLocalFileT8.Path
Else
    strFilePath = filLocalFileT8.Path & "\"
End If

'�إ����W�ٰ}�C
arrTmp = Array("Document.date", "Requested.delivery.d", "Sold-to.party", "Sold.to.name", "Ship.to", "Ship.to.name", "PO.number", "Sales.document", "Delivery.number", "Billing.Doc.no", "Order.Type", "Material", "Material.Description", "Order.Quantity", "Order.Confirmed.Quan", "Sales.unit", "Order.Reason", "Description", "Reason.for.Rejection", "Created.By", "Remarks")

'�إ� Excel �����Ʈw�s��
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT8.Path & "\" & filLocalFileT8.FileName & ";Extended Properties=""Excel 8.0; HDR=YES;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT8 & "$] ", cnExcel ', adOpenStatic, adLockOptimistic
'tmp_rs.Sort = "[Sales document],Material"

Set rsMainT8 = New ADODB.Recordset

'�N�ǤJ tmp_rs ����ƻs�� rsMainT8
Dim fldcnt As Integer, reccnt As Double

'�إ� Recordset �� Table �[�c (�b�O���餤�� ADO Recordset)
rsMainT8.Fields.Append "�s��", adDouble
For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT8.Fields.Append arrTmp(fldcnt), tmp_Rs.Fields(fldcnt).Type, tmp_Rs.Fields(fldcnt).DefinedSize
Next fldcnt

With rsMainT8
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With

reccnt = 0
Do While Not tmp_Rs.EOF
If IsNull(tmp_Rs.Fields(0).Value) = True Then GoTo NextRow 'Add by Gemini @20110721
   reccnt = reccnt + 1
   rsMainT8.AddNew
   rsMainT8.Fields(0).Value = reccnt
   For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT8.Fields(fldcnt + 1).Value = tmp_Rs.Fields(fldcnt).Value & ""
   Next fldcnt
   rsMainT8.Update
NextRow:
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMainT8.DataSource = rsMainT8: dgMainT8.Visible = False

rsMainT8.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT8
dgMainT8.RowHeight = 300
Screen.MousePointer = 0: dgMainT8.Visible = True
MsgBox "���u�@��@ " & rsMainT8.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "�u�@��}��"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT8_Click()

strTranFileName = filLocalFileT8.Path & "\" & filLocalFileT8.FileName
If Len(RTrim(cboSheetT8)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT8.EOF Or rsMainT8 Is Nothing Then Exit Sub

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT8.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT8.Enabled = True: dgMainT8.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

On Error GoTo err_Handle

'��f����ˬd
rsMainT8.MoveFirst
Do While Not rsMainT8.EOF

'    If Format(myExCharFilter(Trim(rsMainT8("Document.date"))), "YYYY/MM/DD") < Format(Now, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub

    '�������--�P�_SKU�O�_�s�b
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT8("Material")) & "' and Storerkey = 'LNSL01' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT8("Material")) & " )" & Trim(rsMainT8("Material.Description")) & "�A�q����J�פ�!!": cmdImportT8.Enabled = True: dgMainT8.Enabled = True: Screen.MousePointer = 0
        Exit Sub
    End If

    rsMainT8.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT8.Enabled = False: dgMainT8.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngCase As Long, lngInnerpack As Long
Dim strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT8.MoveFirst
Do While Not rsMainT8.EOF
    DoEvents: DoEvents
    
'    '�������--�P�_�O�_�q��ƬO�_��0-->���U�@��
'    If Trim(rsMainT8("Qty(stckpg.unit)")) = 0 Then
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
'    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT8("Sales.document"))) Then
        strOrderNo = UCase(Trim(rsMainT8("Sales.document")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT8("Sales.document"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,billtokey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT8("Sales.document"))) & "','R','LNSL01','" & myExCharFilter(Trim(rsMainT8("Document.date"))) & "','" & myExCharFilter(Trim(rsMainT8("Requested.delivery.d"))) & "','" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT8("Ship.To"))) & "','" & myExCharFilter(Trim(rsMainT8("ship.to.name"))) & "','','','','','','" & myExCharFilter(Trim(rsMainT8("Sold-to.party"))) & "','" & myExCharFilter(Trim(rsMainT8("po.number"))) & "','" & myExCharFilter(Trim(rsMainT8("Remarks"))) & "','" & filLocalFileT8.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT8("Sales.document")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫�Ƥ��W�[����
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlineNumber = int_orderlineNumber + 1
            
            '���ӫ~�D�ɸ��
            str_SQL = "select casecnt = isnull(s.casecnt,0) ,innerpack from gv_skuxpack s where s.storerkey = 'LNSL01' and s.sku = '" & myExCharFilter(Trim(rsMainT8("material"))) & "' "
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            lngCase = tmp_Rs("casecnt")
            lngInnerpack = tmp_Rs("innerpack")
            tmp_Rs.Close
                        
            If Trim(rsMainT8("Sales.unit")) = "CS" Then
                intQTY = Trim(rsMainT8("order.quantity")) * IIf(lngCase = 0, 1, lngCase)
            Else
                intQTY = Trim(rsMainT8("order.quantity"))
'                If myExCharFilter(Trim(rsMainT8("material"))) = "12129314" Then intQTY = intQTY * IIf(lngInnerpack = 0, 1, lngInnerpack)
                If lngInnerpack > 0 Then intQTY = intQTY * lngInnerpack
            End If
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT8("Sales.document"))) & "','" & myExCharFilter(Trim(rsMainT8("material"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "',null,null,'','" & myExCharFilter(Trim(rsMainT8("Sales.unit"))) & "','0','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT8.MoveNext

Loop

'�ɫȤ���
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'�s�Ȥ��ˬd
str_SQL = "select �f�D=storerkey , �f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �ˬd��� = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel���
    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�ʫȤ���"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT8.Enabled = True: dgMainT8.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "�o�{�s�Ȥ�A�q����J����!", vbOKOnly, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT8.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT8.FileName & " �ƥ��� C:\BEST\LNSL01\Orders\LNSL01\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT8.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT8.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT8.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT8.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT8.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT8.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT8.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT8.Refresh: cboSheetT8.Clear
Screen.MousePointer = 0: cmdImportT8.Enabled = True: dgMainT8.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportT8_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT8.Enabled = True: Screen.MousePointer = 0: dgMainT8.Enabled = True

End Sub

Private Sub dgMainT8_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT8
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT8_Change()
    filLocalFileT8.Path = dirLocalDirT8.Path
End Sub
Private Sub drvLocalDriveT8_Change()

On Error GoTo DriveError
dirLocalDirT8.Path = drvLocalDriveT8.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT8_Click()

On Error GoTo err_Handle
Set rsMainT8 = Nothing: Set dgMainT8.DataSource = rsMainT8
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT8.Path, 1) = "\" Then
    strFilePath = filLocalFileT8.Path
Else
    strFilePath = filLocalFileT8.Path & "\"
End If

If Dir(strFilePath & filLocalFileT8.FileName) = "" Then: filLocalFileT8.Refresh: Exit Sub

cboSheetT8.Clear

If UCase(mySplit(filLocalFileT8.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT8.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT8.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT8.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT8.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT8.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT9_Click()
If blDo = True Then Call ExcelSheet2RecordsetT9
End Sub
Private Sub ExcelSheet2RecordsetT9()

Dim strFilePath As String
If Right(filLocalFileT9.Path, 1) <> "\" Then
    strFilePath = filLocalFileT9.Path & "\"
Else
    strFilePath = filLocalFileT9.Path
End If

Set rsMainT9 = New ADODB.Recordset
Call Excel2RecordsetT9(strFilePath & filLocalFileT9.FileName, cboSheetT9, "Delivery" & Chr(9) & "Expr1" & Chr(9) & "TO.Number" & Chr(9) & "DlvTy" & Chr(9) & "Sold-to.pt" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Deliv.date" & Chr(9) & "Expr2" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Act.qty(dest)" & Chr(9) & "BUn" & Chr(9) & "Actual.qty" & Chr(9) & "AUn" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "SLED/BBD" & Chr(9) & "Route" & Chr(9) & "Notes" & Chr(9), rsMainT9)
rsMainT9.Sort = "[Delivery]"

Set dgMainT9.DataSource = rsMainT9

If rsMainT9 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"

Else
    SetDataGridColWidth Me.Caption, dgMainT9
    MsgBox "���u�@��@ " & rsMainT9.RecordCount & "������", 64, "Excel2Recordset"

End If

End Sub
Sub ExcelSheet2RecordsetT9_old()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT9.Path, 1) = "\" Then
    strFilePath = filLocalFileT9.Path
Else
    strFilePath = filLocalFileT9.Path & "\"
End If

'�إ����W�ٰ}�C
arrTmp = Array("Delivery", "Expr1", "TO.Number", "DlvTy", "Sold-to.pt", "Ship-To.Pt", "Deliv.date", "Expr2", "Item", "Material", "Act.qty(dest)", "BUn", "Actual.qty", "AUn", "Plnt", "SLoc", "Batch", "SLED/BBD", "Route")

'�إ� Excel �����Ʈw�s��
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT9.Path & "\" & filLocalFileT9.FileName & ";Extended Properties=""Excel 8.0; HDR=YES;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT9 & "$] where len(rtrim(Delivery)) > 0 ", cnExcel ', adOpenStatic, adLockOptimistic
tmp_Rs.Sort = "[Delivery]"

Set rsMainT9 = New ADODB.Recordset

'�N�ǤJ tmp_rs ����ƻs�� rsMainT9
Dim fldcnt As Integer, reccnt As Double

'�إ� Recordset �� Table �[�c (�b�O���餤�� ADO Recordset)
rsMainT9.Fields.Append "�s��", adDouble
For fldcnt = 0 To 18
    rsMainT9.Fields.Append arrTmp(fldcnt), tmp_Rs.Fields(fldcnt).Type, tmp_Rs.Fields(fldcnt).DefinedSize
Next fldcnt

With rsMainT9
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With

reccnt = 0
Do While Not tmp_Rs.EOF
   reccnt = reccnt + 1
   rsMainT9.AddNew
   rsMainT9.Fields(0).Value = reccnt
   For fldcnt = 0 To 18
    rsMainT9.Fields(fldcnt + 1).Value = tmp_Rs.Fields(fldcnt).Value & ""
   Next fldcnt
   rsMainT9.Update
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMainT9.DataSource = rsMainT9: dgMainT9.Visible = False

rsMainT9.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT9
dgMainT9.RowHeight = 300
Screen.MousePointer = 0: dgMainT9.Visible = True
MsgBox "���u�@��@ " & rsMainT9.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "�u�@��}��"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT9_Click()

strTranFileName = filLocalFileT9.Path & "\" & filLocalFileT9.FileName
If Len(RTrim(filLocalFileT9.FileName)) = 0 Then MsgBox "�п���ɮ�", 64, Me.Caption: Exit Sub
If Len(RTrim(cboSheetT9)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT9.EOF Or rsMainT9 Is Nothing Then Exit Sub

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT9.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT9.Enabled = True: dgMainT9.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

On Error GoTo err_Handle

'��f����ˬd
rsMainT9.MoveFirst
Do While Not rsMainT9.EOF

If Format(myExCharFilter(Trim(rsMainT9("Deliv.date"))), "YYYY/MM/DD") < Format(Now - 1, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub

    rsMainT9.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT9.Enabled = False: dgMainT9.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngInnerpack As Long
Dim strOrderNo As String, strWHOrderNo As String, strWH As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strLot06 As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT9.MoveFirst

Do While Not rsMainT9.EOF

    '�������--�P�_�q��ƬO�_��0
    If Trim(rsMainT9("Act.qty(dest)")) = 0 Then intNotBest = intNotBest + 1: GoTo next1

    DoEvents: DoEvents
    
'    'CT & BEST���ܥX�f�P�_
'    If strWHOrderNo = UCase(Trim(rsMainT9("Delivery"))) Then
'        If (strWH = "BEST" And (Trim(rsMainT9("sloc")) = "0001" Or Trim(rsMainT9("sloc")) = "0005")) Or (strWH = "CT" And (Trim(rsMainT9("sloc")) <> "0001" Or Trim(rsMainT9("sloc")) <> "0005") = False) Then
'            cmdImportT9.Enabled = True: Screen.MousePointer = 0: dgMainT9.Enabled = True: cn.RollbackTrans: Tran_Level = 0
'            MsgBox "�o�{ CT & BEST ���ܥX�f�q��A�лP�Ȥ�T�{�q�楿�T�L�~�I", 16, "�q����J�פ� "
'            Exit Sub
'        End If
'    End If
    
    '�O�_���ըƹF�t�O�ܧO-->���U�@��
    If Trim(rsMainT9("plnt")) = "1119" And (Trim(rsMainT9("sloc")) = "0001" Or Trim(rsMainT9("sloc")) = "0002" Or Trim(rsMainT9("sloc")) = "0005" Or Trim(rsMainT9("sloc")) = "0007" Or Trim(rsMainT9("sloc")) = "0010") Then
        
        strWH = "CT"
        '�ˬd�q��q�O�_�X�{�p���I
        If Val(rsMainT9("Act.qty(dest)")) <> Round(Val(rsMainT9("Act.qty(dest)")), 0) Then
            cmdImportT9.Enabled = True: Screen.MousePointer = 0: dgMainT9.Enabled = True: cn.RollbackTrans: Tran_Level = 0
            MsgBox "�o�{�q��q�X�{�p���I�A�лP�Ȥ�T�{���T�q��q�I", 16, "�q����J�פ� "
            Exit Sub
        End If
        
        '�������--�P�_SKU�O�_�s�b
        str_SQL = "select sku,innerpack from gv_skuxpack where sku='" & Trim(rsMainT9("Material")) & "' and Storerkey = 'LNSL01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            cn.RollbackTrans: Tran_Level = 0: cmdImportT9.Enabled = True: dgMainT9.Enabled = True: Screen.MousePointer = 0
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT9("Material")) & ") �A�q����J�פ�!!"
            Exit Sub
        End If
        lngInnerpack = tmp_Rs("Innerpack")

    Else
        intNotBest = intNotBest + 1
        strWH = "BEST"
        GoTo next1

    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT9("Delivery"))) Then
        strOrderNo = UCase(Trim(rsMainT9("Delivery")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT9("Delivery"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,BillToKey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT9("Delivery"))) & "','I','LNSL01', '" & myExCharFilter(Trim(rsMainT9("ExPr1"))) & "',cast('" & myExCharFilter(Trim(rsMainT9("Deliv.date"))) & "' as datetime)+1,'" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT9("Route"))) & "','','','','','','','" & myExCharFilter(Trim(rsMainT9("Sold-to.pt"))) & "','" & myExCharFilter(Trim(rsMainT9("to.number"))) & "','" & myExCharFilter(Trim(rsMainT9("Notes"))) & "','" & filLocalFileT9.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT9("Delivery")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫�Ƥ��W�[����
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlineNumber = int_orderlineNumber + 1
            
            '�ƶq�ഫ
            intQTY = Trim(rsMainT9("Act.qty(dest)"))
'            If Trim(rsMainT9("Material")) = "12129314" Then intQTY = intQTY * IIf(lngInnerpack = 0, 1, lngInnerpack)
            If lngInnerpack > 0 Then intQTY = intQTY * lngInnerpack
            
            '�ܧO�ഫ
            strLot06 = myExCharFilter(Trim(rsMainT9("sloc")))
            
            If strLot06 = "0001" Then
               strLot06 = "R01"
            ElseIf strLot06 = "0002" Then
               strLot06 = "R02"
            ElseIf strLot06 = "0005" Then
               strLot06 = "R05"
            ElseIf strLot06 = "0007" Then
               strLot06 = "R07"
            ElseIf strLot06 = "0010" Then
               strLot06 = "R10"
            End If
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes,updatesource) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT9("Item"))) & "','" & myExCharFilter(Trim(rsMainT9("Delivery"))) & "','" & myExCharFilter(Trim(rsMainT9("material"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT9("Batch"))) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT9("BUn"))) & "','0','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:

        If (Trim(rsMainT9("sloc")) = "0001" Or Trim(rsMainT9("sloc")) = "0005") Then
            strWH = "CT"
        Else
            strWH = "BEST"
        End If

        strWHOrderNo = UCase(Trim(rsMainT9("Delivery")))
        rsMainT9.MoveNext
        
Loop

'�ɫȤ���
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'�s�Ȥ��ˬd
str_SQL = "select �f�D=storerkey , �f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �ˬd��� = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel���
    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�ʫȤ���"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT9.Enabled = True: dgMainT9.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "�o�{�s�Ȥ�A�q����J����!", 16, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT9.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT9.FileName & " �ƥ��� C:\BEST\LNSL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT9.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT9.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFilet9.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT9.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT9.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT9.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT9.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT9.Refresh: cboSheetT9.Clear
Screen.MousePointer = 0: cmdImportT9.Enabled = True: dgMainT9.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportt9_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT9.Enabled = True: Screen.MousePointer = 0: dgMainT9.Enabled = True

End Sub

Private Sub dgMainT9_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT9
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT9_Change()
    filLocalFileT9.Path = dirLocalDirT9.Path
End Sub
Private Sub drvLocalDriveT9_Change()

On Error GoTo DriveError
dirLocalDirT9.Path = drvLocalDriveT9.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT9_Click()

On Error GoTo err_Handle
Set rsMainT9 = Nothing: Set dgMainT9.DataSource = rsMainT9
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT9.Path, 1) = "\" Then
    strFilePath = filLocalFileT9.Path
Else
    strFilePath = filLocalFileT9.Path & "\"
End If

If Dir(strFilePath & filLocalFileT9.FileName) = "" Then: filLocalFileT9.Refresh: Exit Sub

cboSheetT9.Clear

If UCase(mySplit(filLocalFileT9.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT9.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT9.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT9.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT9.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT9.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT10_Click()
If blDo = True Then Call ExcelSheet2Recordsett10
End Sub

Sub ExcelSheet2Recordsett10()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String, i As Integer, strQty As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT10.Path, 1) = "\" Then
    strFilePath = filLocalFileT10.Path
Else
    strFilePath = filLocalFileT10.Path & "\"
End If

'�إ����W�ٰ}�C
arrTmp = Array("Ship.to(Return.Code)", "Ship.to.name", "PO.number", "ZRR�����", "�~��", "���", "Sales����ñ��", "E-mail.to����", "�Ƶ�")

'�إ� Excel �����Ʈw�s��
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT10.Path & "\" & filLocalFileT10.FileName & ";Extended Properties=""Excel 8.0; HDR=YES;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT10 & "$] ", cnExcel ', adOpenStatic, adLockOptimistic
'tmp_rs.Sort = "[Ship to(Return Code)]"

Set rsMainT10 = New ADODB.Recordset

'�N�ǤJ tmp_rs ����ƻs�� rsMaint10
Dim fldcnt As Integer, reccnt As Double

'�إ� Recordset �� Table �[�c (�b�O���餤�� ADO Recordset)
rsMainT10.Fields.Append "�s��", adDouble
For fldcnt = 0 To 8
'    rsMainT10.Fields.Append arrTmp(fldcnt), tmp_rs.Fields(fldcnt).Type, tmp_rs.Fields(fldcnt).DefinedSize
rsMainT10.Fields.Append arrTmp(fldcnt), adVarChar, 255
Next fldcnt

With rsMainT10
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With

reccnt = 0
tmp_Rs.MoveFirst
Do While Not tmp_Rs.EOF
    If IsNull(tmp_Rs("Ship to(Return Code)")) = False Or Len(Trim(tmp_Rs("Ship to(Return Code)"))) > 0 Then
        reccnt = reccnt + 1
        rsMainT10.AddNew
        rsMainT10.Fields(0).Value = reccnt
        For fldcnt = 0 To 8
         rsMainT10.Fields(fldcnt + 1).Value = Trim(tmp_Rs.Fields(fldcnt)) & ""
        Next fldcnt
        rsMainT10.Update
    End If
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

rsMainT10.MoveFirst

'���O�X���x�s��
i = 1
strQty = rsMainT10("���")
Do While Not rsMainT10.EOF
    
    If rsMainT10("���") = "" Then
        i = i + 1
        rsMainT10("���") = "**" & i & "/" & strQty
        
    Else
        strQty = rsMainT10("���")
        i = 1
    End If

rsMainT10.MoveNext
Loop

i = 0
'�έp�����ƻP�`���
rsMainT10.MoveLast
Do While Not rsMainT10.BOF

        If Left(rsMainT10("���"), 2) = "**" Then
        rsMainT10("�Ƶ�") = "(" & Val(Replace(mySplit(rsMainT10("���"), "/", 0), "**", "")) & "/" & Val(Replace(mySplit(rsMainT10("���"), "/", 0), "**", "")) + i & ")�@" & mySplit(rsMainT10("���"), "/", -1) & "��A" & rsMainT10("�Ƶ�")
        i = i + 1
        rsMainT10("���") = 1
    Else
        If i <> 0 Then rsMainT10("�Ƶ�") = "(1/" & i + 1 & ")�@" & rsMainT10("���") & "��A" & rsMainT10("�Ƶ�")
        strQty = rsMainT10("���")
        i = 0
    End If

rsMainT10.MovePrevious
Loop

Set dgMainT10.DataSource = rsMainT10: dgMainT10.Visible = False


SetDataGridColWidth Me.Caption, dgMainT10
dgMainT10.RowHeight = 300
Screen.MousePointer = 0: dgMainT10.Visible = True
MsgBox "���u�@��@ " & rsMainT10.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "�u�@��}��"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT10_Click()

On Error GoTo err_Handle
strTranFileName = filLocalFileT10.Path & "\" & filLocalFileT10.FileName
If Len(RTrim(cboSheetT10)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT10.EOF Or rsMainT10 Is Nothing Then Exit Sub

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT10.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT10.Enabled = True: dgMainT10.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'��f����ˬd
rsMainT10.MoveFirst
'Do While Not rsMainT10.EOF
'
'    If myExCharFilter(Trim(rsMainT10("ZRR�����"))) < Format(Now - 1, "YYYYMMDD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub
'    rsMainT10.MoveNext
'Loop

Tran_Level = cn.BeginTrans
cmdImportT10.Enabled = False: dgMainT10.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer
Dim strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT10.MoveFirst
Do While Not rsMainT10.EOF
    DoEvents: DoEvents
    
    '�������--�P�_�O�_�~���O�_��D-->���U�@��
'    If UCase(Trim(rsMainT10("�~��"))) <> "D" Then
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '�������--�P�_SKU�O�_�s�b
        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT10("�~��")) & "' and Storerkey = 'LNSL01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT10("�~��")) & ")" & "�A�q����J�פ�!!": cmdImportT10.Enabled = True: dgMainT10.Enabled = True: Screen.MousePointer = 0
            cn.RollbackTrans: Tran_Level = 0
            Exit Sub
        End If

'    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT10("PO.number"))) Then
        strOrderNo = UCase(Trim(rsMainT10("PO.number")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT10("PO.number"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,billtokey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT10("PO.number"))) & "','R','LNSL01','" & myExCharFilter(Trim(rsMainT10("ZRR�����"))) & "',cast(" & "'" & myExCharFilter(Trim(rsMainT10("ZRR�����"))) & "' as datetime)+1,'" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT10("Ship.to(Return.Code)"))) & "','" & myExCharFilter(Trim(rsMainT10("ship.to.name"))) & "','','','','','','','','" & myExCharFilter(Trim(rsMainT10("�Ƶ�"))) & "','" & filLocalFileT10.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT10("PO.number")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫�Ƥ��W�[����
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlineNumber = int_orderlineNumber + 1

            intQTY = Trim(rsMainT10("���"))
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT10("PO.number"))) & "','" & myExCharFilter(Trim(rsMainT10("�~��"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "',null,null,'','EA','0','" & myExCharFilter(Trim(rsMainT10("�Ƶ�"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT10.MoveNext

Loop

'�ɫȤ���
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'�s�Ȥ��ˬd
str_SQL = "select �f�D=storerkey , �f�D�q��渹=externorderkey , �q�����O = priority , �q����=orderdate , ��f���=deliverydate , �Ȥ�s��=consigneekey , �ˬd��� = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel���
    Call Recordset2Excel("�ʫȤ���", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�ʫȤ���", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�ʫȤ���"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�ʫȤ���\�ʫȤ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT10.Enabled = True: dgMainT10.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "�o�{�s�Ȥ�A�q����J����!", vbOKOnly, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT10.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT10.FileName & " �ƥ��� C:\BEST\LNSL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT10.FileName)
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dgMainT10.Enabled = True
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT10.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LNSL01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFilet10.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT10.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT10.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT10.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT10.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT10.Refresh: cboSheetT10.Clear
Screen.MousePointer = 0: cmdImportT10.Enabled = True: dgMainT10.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportt10_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT10.Enabled = True: Screen.MousePointer = 0: dgMainT10.Enabled = True

End Sub

Private Sub dgMainT10_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT10
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT10_Change()
    filLocalFileT10.Path = dirLocalDirT10.Path
End Sub
Private Sub drvLocalDriveT10_Change()

On Error GoTo DriveError
dirLocalDirT10.Path = drvLocalDriveT10.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT10_Click()

On Error GoTo err_Handle
Set rsMainT10 = Nothing: Set dgMainT10.DataSource = rsMainT10
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT10.Path, 1) = "\" Then
    strFilePath = filLocalFileT10.Path
Else
    strFilePath = filLocalFileT10.Path & "\"
End If

If Dir(strFilePath & filLocalFileT10.FileName) = "" Then: filLocalFileT10.Refresh: Exit Sub

cboSheetT10.Clear

If UCase(mySplit(filLocalFileT10.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT10.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT10.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT10.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT10.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT10.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT3_Click()
If blDo = True Then Call ExcelSheet2RecordsetT3
End Sub

Sub ExcelSheet2RecordsetT3()
On Error GoTo err_Handle
Screen.MousePointer = 11

Dim strFilePath As String
If Right(filLocalFileT3.Path, 1) <> "\" Then
    strFilePath = filLocalFileT3.Path & "\"
Else
    strFilePath = filLocalFileT3.Path
End If

Set rsMainT3 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT3.FileName, cboSheetT3, "", rsMainT3)

rsMainT3.Sort = "�q��s��"

Set dgMainT3.DataSource = rsMainT3

If rsMainT3 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT3
    MsgBox "���u�@��@ " & rsMainT3.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

Screen.MousePointer = 0
Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT3_Click()

strTranFileName = filLocalFileT3.Path & "\" & filLocalFileT3.FileName
If Len(RTrim(cboSheetT3)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT3.EOF Or rsMainT3 Is Nothing Then Exit Sub

On Error GoTo err_Handle
'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT3.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT3.Enabled = True: dgMainT3.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

''��f����ˬd
'rsMainT3.MoveFirst
'Do While Not rsMainT3.EOF
'
'    If myExCharFilter(Trim(rsMainT3("���e���")))< Format(Now, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub
'
'    rsMainT3.MoveNext
'Loop

Tran_Level = cn.BeginTrans: cmdImportT3.Enabled = False: dgMainT3.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'���̫�Ȥ�s��
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,7,20) as consigneekey from trp01m where storerkey = 'LFYY01' and left(consigneekey,6) = 'LFYY00' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT3.MoveFirst
Do While Not rsMainT3.EOF
    DoEvents: DoEvents
    
    '�������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If UCase(Trim(rsMainT3("�ܮw"))) <> "A02" And UCase(Trim(rsMainT3("�ܮw"))) <> "A02A" And UCase(Trim(rsMainT3("�ܮw"))) <> "A02C" Then
''        MsgBox "�Ȥ�渹�G" & Trim(rsMainT4("�P�f�渹")) & "( " & Trim(rsMainT4("�ܮw")) & " )" & vbCrLf & "�гq���Ȥ�A�T�{�ӵ��q��O�_���~!?", vbOKOnly, "�D�ըƹF���q�椣��J"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '�������--�P�_SKU�O�_�s�b
        str_SQL = "select casecnt = case when isnull(casecnt,0) = 0 then 1 else casecnt end from gv_skuxpack where sku='" & Trim(rsMainT3("�����~��")) & "' and Storerkey = 'LFYY01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT3("�����~��")) & " ) " & Trim(rsMainT3("�ӫ~�W��")) & "�A�q����J�פ�!!"
             cmdImportT3.Enabled = True: dgMainT3.Enabled = True: Screen.MousePointer = 0
            cn.RollbackTrans: Tran_Level = 0
            tmp_Rs.Close
            Exit Sub
        End If
        tmp_Rs.Close
        
'    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT3("�q��s��"))) Then
        strOrderNo = UCase(Trim(rsMainT3("�q��s��")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�ˬd�O�_�����Ȥ�W��
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LFYY01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT3("DC�W��"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '�L���Ȥ�W�٫h�s�W
            intTmp = intTmp + 1
            strConsigneeKey = "LFYY" & Format(intTmp, "000000")
            
            '�s�W�Ȥ�D��
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('LFYY01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT3("DC�W��"))) & "','" & myExCharFilter(Trim(rsMainT3("DC�W��"))) & "','" & myExCharFilter(Trim(rsMainT3("�q�f�H��"))) & "','','" & myExCharFilter(Trim(rsMainT3("DC��}"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '�����s�W���Ȥ�s��
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = 'LFYY01' and full_name = '" & myExCharFilter(Trim(rsMainT3("DC�W��"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT3("�q�f�H��"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT3("DC��}"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '�p���H�P��f�a�}����
                intTmp = intTmp + 1
                strConsigneeKey = "LFYY" & Format(intTmp, "000000")
                
                '�s�W�Ȥ�D��
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
                " values('LFYY01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT3("DC�W��"))) & "','" & myExCharFilter(Trim(rsMainT3("DC�W��"))) & "','" & myExCharFilter(Trim(rsMainT3("�q�f�H��"))) & "','','" & myExCharFilter(Trim(rsMainT3("DC��}"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '�����s�W���Ȥ�s��
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '�۲Ūu���«Ƚs
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
        End If
        tmp_Rs.Close
    
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT3("�q��s��"))) & "' and storerkey = 'LFYY01' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT3("�q��s��"))) & "','I','LFYY01','" & myExCharFilter(Trim(rsMainT3("�q�f���"))) & "','" & myExCharFilter(Trim(rsMainT3("��f���"))) & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT3("DC�W��"))) & "','" & myExCharFilter(Trim(rsMainT3("�q�f�H��"))) & "','','','" & myExCharFilter(Trim(GetWord(rsMainT3("DC��}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT3("DC��}"), intPointer, 45))) & "','','','" & filLocalFileT3.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT3("DC�s��"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LFYY01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT3("�q��s��")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Val(rsMainT3("�q�f��"))
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes,addwho,editwho)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT3("�q��s��")) & "','" & Trim(rsMainT3("�����~��")) & "','LFYY01'," & _
            "'" & intQTY & "','" & intQTY & "','R01','','EA','0','','" & User_id & "','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT3.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT3.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT3.FileName & " �ƥ��� C:\BEST\LFYY01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT3.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT3.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\LFYY01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LFYY01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LFYY01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LFYY01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LFYY01\Orders\Backup"
If Dir("C:\BEST\LFYY01\Orders\Backup\" & filLocalFileT3.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LFYY01\Orders\Backup\" & filLocalFileT3.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LFYY01\Orders\Backup\" & mySplit(filLocalFileT3.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT3.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT3.Refresh: cboSheetT3.Clear
Screen.MousePointer = 0: cmdImportT3.Enabled = True: dgMainT3.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportT3_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT3.Enabled = True: Screen.MousePointer = 0: dgMainT3.Enabled = True

End Sub

Private Sub dgMainT3_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT3
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT3_Change()
    filLocalFileT3.Path = dirLocalDirT3.Path
End Sub
Private Sub drvLocalDriveT3_Change()

On Error GoTo DriveError
dirLocalDirT3.Path = drvLocalDriveT3.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT3_Click()

On Error GoTo err_Handle
Set rsMainT3 = Nothing: Set dgMainT3.DataSource = rsMainT10
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT3.Path, 1) = "\" Then
    strFilePath = filLocalFileT3.Path
Else
    strFilePath = filLocalFileT3.Path & "\"
End If

If Dir(strFilePath & filLocalFileT3.FileName) = "" Then: filLocalFileT3.Refresh: Exit Sub

cboSheetT3.Clear

If UCase(mySplit(filLocalFileT3.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT3.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT3.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT3.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT3.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT3.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT11_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT11.Path, 1) = "\" Then
    strFilePath = filLocalFileT11.Path
Else
    strFilePath = filLocalFileT11.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = "�Ȥ�" & Chr(9) & "�νs" & Chr(9) & "�P�f���" & Chr(9) & "�P�f�渹" & Chr(9) & "�p���H" & Chr(9) & "�q��" & Chr(9) & "�e�f�a�}" & Chr(9) & "�o�����X" & Chr(9) & "�~�ȭ�" & Chr(9) & "�Ƶ�" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�ܮw" & Chr(9) & "�ƶq" & Chr(9) & "���" & Chr(9) & "�e�m���/�Ƶ�/�Ȥ�渹" & Chr(9)

If Right(filLocalFileT11.Path, 1) <> "\" Then
    strFilePath = filLocalFileT11.Path & "\"
Else
    strFilePath = filLocalFileT11.Path
End If

Set rsMainT11 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT11.FileName, cboSheetT11, strFieldName, rsMainT11)

Set dgMainT11.DataSource = rsMainT11

If rsMainT11 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT11
    MsgBox "���u�@��@ " & rsMainT11.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT11_Click()

On Error GoTo err_Handle
strTranFileName = filLocalFileT11.Path & "\" & filLocalFileT11.FileName
If Len(RTrim(cboSheetT11)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT11.EOF Or rsMainT11 Is Nothing Then Exit Sub

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT11.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT11.Enabled = True: dgMainT11.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT11.MoveFirst
Do While Not rsMainT11.EOF

    '��f����ˬd
    arrTmp = Split(Trim(rsMainT11("�P�f���")), "/")
    If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub
            
    '�ƶq�ˬd
    If Val(rsMainT11("�ƶq")) < 0 Then
        MsgBox "�q��ƶq�p��1�A" & Trim(rsMainT11("�~��")) & "-" & Trim(rsMainT11("�~�W")) & "(" & Trim(rsMainT11("�ƶq")) & Trim(rsMainT11("���")) & ")�A�q����J�פ�!!", , "�q���ɶפJ": Exit Sub
        Exit Sub
    End If
                 
    rsMainT11.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT11.Enabled = False: dgMainT11.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'���̫�Ȥ�s��
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LNIP01' and left(consigneekey,4) = 'LNIP' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT11.MoveFirst
Do While Not rsMainT11.EOF
    DoEvents: DoEvents
    
'    '�������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If UCase(Trim(rsMainT11("�ܮw"))) = "" Then
''        MsgBox "�Ȥ�渹�G" & Trim(rsMainT4("�P�f�渹")) & "( " & Trim(rsMainT4("�ܮw")) & " )" & vbCrLf & "�гq���Ȥ�A�T�{�ӵ��q��O�_���~!?", vbOKOnly, "�D�ըƹF���q�椣��J"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '�������--�P�_SKU�O�_�s�b
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT11("�~��")) & "' and Storerkey = 'LNIP01' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cn.RollbackTrans: Tran_Level = 0
        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT11("�~��")) & " ) " & Trim(rsMainT11("�~�W")) & "�A�q����J�פ�!!": cmdImportT11.Enabled = True: dgMainT11.Enabled = True: Screen.MousePointer = 0
        Exit Sub
    End If

'    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT11("�P�f�渹"))) Then
        strOrderNo = UCase(Trim(rsMainT11("�P�f�渹")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�ˬd�O�_�����Ȥ�W��
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LNIP01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT11("�Ȥ�"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '�L���Ȥ�W�٫h�s�W
            intTmp = intTmp + 1
            strConsigneeKey = "LNIP" & Format(intTmp, "000000")
            
            '�s�W�Ȥ�D��
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT11("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT11("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT11("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT11("�q��"))) & "','" & myExCharFilter(Trim(rsMainT11("�e�f�a�}"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '�����s�W���Ȥ�s��
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = 'LNIP01' and full_name = '" & myExCharFilter(Trim(rsMainT11("�Ȥ�"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT11("�p���H"))) & "' " & _
                        "and rtrim(phone) = '" & myExCharFilter(Trim(rsMainT11("�q��"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT11("�e�f�a�}"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '�p���H�B�q�ܻP��f�a�}����
                intTmp = intTmp + 1
                strConsigneeKey = "LNIP" & Format(intTmp, "000000")
                
                '�s�W�Ȥ�D��
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes) " & _
                " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT11("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT11("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT11("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT11("�q��"))) & "','" & myExCharFilter(Trim(rsMainT11("�e�f�a�}"))) & "','" & myExCharFilter(Trim(tmp_Rs("notes"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '�����s�W���Ȥ�s��
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '�۲Ūu���«Ƚs
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
            rsTmp.Close
        End If
        tmp_Rs.Close
    
'        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT11("�P�f�渹"))) & "' and storerkey = 'LNIP01' and isnull(type,'') <> '�R��' "
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = Trim(rsMainT11("�ܮw"))
            strFacility = "�ըƹF�_��"
            arrTmp = Split(Trim(rsMainT11("�P�f���")), "/")
            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT11("�P�f�渹"))) & "','I','LNIP01','" & strOrderDate & "','" & CDate(strOrderDate) + 1 & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT11("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT11("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT11("�~�ȭ�"))) & "','" & myExCharFilter(Trim(rsMainT11("�νs"))) & "','" & myExCharFilter(Trim(rsMainT11("�q��"))) & "','','" & myExCharFilter(Trim(GetWord(rsMainT11("�e�f�a�}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT11("�e�f�a�}"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT11("�Ƶ�"))) & "','" & filLocalFileT11.FileName & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT11("�o�����X"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LNIP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
'        Else
'            '�q�歫��
'            Call FTPlog("�q�歫��" & str_SQL)
'            '��������
'            strReOrderkey = strReOrderkey & Trim(rsMainT11("�P�f�渹")) & "','"
'            blDuplicationOrder = True
'
'        End If
    End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Trim(rsMainT11("�ƶq"))
            strLot06 = IIf(UCase(Trim(rsMainT11("�ܮw"))) = "A06", "A06-S", Trim(rsMainT11("�ܮw")))
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT11("�P�f�渹"))) & "','" & myExCharFilter(Trim(rsMainT11("�~��"))) & "','LNIP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','" & myExCharFilter(rsMainT11("�ܮw")) & "','" & myExCharFilter(Trim(rsMainT11("���"))) & "','0','" & myExCharFilter(Trim(rsMainT11("�e�m���/�Ƶ�/�Ȥ�渹"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT11.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT11.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT11.FileName & " �ƥ��� C:\BEST\LNIP01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT11.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT11.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �Ƹ� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\LNIP01\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNIP01\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\LNIP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\Orders\Backup"
If Dir("C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT11.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT11.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & mySplit(filLocalFileT11.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT11.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT11.Refresh: cboSheetT11.Clear
Screen.MousePointer = 0: cmdImportT11.Enabled = True: dgMainT11.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportT11_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT11.Enabled = True: Screen.MousePointer = 0: dgMainT11.Enabled = True

End Sub

Private Sub dgMainT11_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT11
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT11_Change()
    filLocalFileT11.Path = dirLocalDirT11.Path
End Sub
Private Sub drvLocalDriveT11_Change()

On Error GoTo DriveError
dirLocalDirT11.Path = drvLocalDriveT11.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT11_Click()

On Error GoTo err_Handle
Set rsMainT11 = Nothing: Set dgMainT11.DataSource = rsMainT11
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT11.Path, 1) = "\" Then
    strFilePath = filLocalFileT11.Path
Else
    strFilePath = filLocalFileT11.Path & "\"
End If

If Dir(strFilePath & filLocalFileT11.FileName) = "" Then: filLocalFileT11.Refresh: Exit Sub

cboSheetT11.Clear

If UCase(mySplit(filLocalFileT11.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT11.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT11.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT11.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT11.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT11.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT12_Click()
Dim bl_Error As Boolean '�O�����p���I���X��
Dim str_Error As String '�O�����p���I���~�����
Dim x As Long

bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT12 Is Nothing Then Exit Sub
If rsMainT12.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT12.Enabled = False: cmdImportT12.Enabled = False
strTranFileName = filLocalFileT12.Path & "\" & filLocalFileT12.FileName

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT12.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT12.RecordCount = 0 Or rsMainT12 Is Nothing Then
Else
rsMainT12.MoveFirst
str_Storerkey = "LPSI01"

Do While Not rsMainT12.EOF
    '��f����ˬd
    If Len(Trim(rsMainT12("�ѳf���"))) = 0 Then
         MsgBox "����渹:" & Trim(rsMainT12("��f")) & "���ѳf������ťաA�q����J�פ�!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT12("�ѳf���"))) > 0 And Len(Trim(rsMainT12("�ѳf���"))) < 8 Then
         MsgBox "����渹:" & Trim(rsMainT12("��f")) & "���ѳf���:" & Trim(rsMainT12("�ѳf���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    Else
        '�ˬd��f�餣�i�p�󤵤�
        If Trim(rsMainT12("�ѳf���")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '�̰��v�����ˬd��f��
                 x = MsgBox("�ѳf����p�󤵤�A�A�T�w�n�~���?", vbQuestion + vbYesNo, "�̰��v����f���ˬd") '�������U���O�T�w�άO����
                    If x = 6 Then
                        '�~��
                    Else
                        '���}
                         dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
    '�q����ˬd
    If Len(Trim(rsMainT12("�ѳf���"))) = 0 Then
         MsgBox "�ѳf������ťաA�q����J�פ�!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT12("�ѳf���"))) > 0 And Len(Trim(rsMainT12("�ѳf���"))) < 8 Then
         MsgBox "�ѳf���:" & Trim(rsMainT12("�ѳf���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
'    Else
'        If Trim(rsMainT12("�ѳf���")) > Trim(rsMainT12("�ѳf���")) Then MsgBox "�q�渹�X:" & Trim(rsMainT12("�q�渹�X")) & "���q���:" & Trim(rsMainT12("�ѳf���")) & "�A�j���f��A�q����J�פ�!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    End If
    
    '�ƶq�ˬd
    If Val(rsMainT12("��f�ƶq")) < 1 Then
        MsgBox "�ƶq�p��1�A" & Trim(rsMainT12("����渹")) & "-�~���G" & Trim(rsMainT12("����")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '������� --�P�_SKU�O�_�s�b
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT12("����")) & "' and Storerkey = '" & str_Storerkey & "'"
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT12("����")) & ")�A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
            Exit Sub
        End If
        
'        If Trim(rsMainT12("�q�����O")) = "A2B" Then
'        '�ˬdA2B�q��Ȥ�s���O�_�s�b
'                str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT12("���f�Ȥ�s��")) & "' and Storerkey = '" & Str_storerkey & "'"
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'                If tmp_Rs.EOF Then  '���s���ǭn��
'                    MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT12("���f�Ȥ�s��")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                    dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'                    Exit Sub
'                End If
'        End If
'
'        '�ˬdA2B�q��H�~���Ȥ�s���O�_�s�b
'        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT12("��f�Ȥ�s��")) & "' and Storerkey = '" & Str_storerkey & "'"
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then  '���s���ǭn��
'            MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT12("��f�Ȥ�s��")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
        '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT12("��f�ƶq")), ".") <> 0 Then
            str_Error = "����渹:" & Trim(rsMainT12("��f")) & "�A�~��:" & Trim(rsMainT12("����")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
'        '�ˬd�f�D--mark by Gemini @20160602
'        If UCase(Trim(rsMainT12("�f�D"))) <> "LABT01" And UCase(Trim(rsMainT12("�f�D"))) <> "LLFA01" Then
'            MsgBox "�q��o�{�D�Ȱ����f�D: " & Trim(rsMainT12("�f�D")) & " )�A���פJ�{���ȨѶפJ�Ȱ��ΧQ�׭q��A�нT�{��A�פJ�A�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If

        '�P�_��O--add by  Gemini @20160602
'        If Trim(rsMainT12("�q�����O")) <> "C" Or Trim(rsMainT12("�q�����O")) <> "R" Or Trim(rsMainT12("�q�����O")) <> "RC" Or Trim(rsMainT12("�q�����O")) <> "I" Or Trim(rsMainT12("�q�����O")) <> "A2B" Then
'        Else
'            MsgBox "�t�εL����O:" & Trim(rsMainT12("�q�����O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
        '�ˬd�f�D�O�_�s�b
        str_SQL = "select storerkey from trp16m where storerkey = '" & str_Storerkey & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�f�D ( " & str_Storerkey & " )�A�Х���f�D�D�ɷs�سf�D��ơA�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
            Exit Sub
        End If
        
        '�P�_C��
'        If Trim(rsMainT12("�q�����O")) = "C" And (Trim(rsMainT12("�f�D")) = "LKAO01" Or Trim(rsMainT12("�f�D")) = "LABT01") Then
'        Else
'            MsgBox "���f�D:" & Trim(rsMainT12("�f�D")) & "���q�����O���i��:" & Trim(rsMainT12("�q�����O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
'        '�P�_�O�_��A2B��R��
'        If Trim(rsMainT12("�q�����O")) <> "A2B" And Trim(rsMainT12("�q�����O")) <> "R" And UCase(Trim(rsMainT12("�f�D"))) = "LABT01" Then
'            MsgBox "���f�D:" & Trim(rsMainT12("�f�D")) & "���q�����O���i��:" & Trim(rsMainT12("�q�����O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
    rsMainT12.MoveNext
Loop
rsMainT12.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT12.Enabled = False: dgMainT12.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long, IntCasecnt As Integer
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'�}�l�פJ
Do While Not rsMainT12.EOF
    DoEvents: DoEvents
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT12("��f"))) Then
        strOrderNo = UCase(Trim(rsMainT12("��f")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�w�g�ˬd�L�A�h�����A������X�Ȥ�D�ɤ��� �Ȥ�s���A�l���ϸ��A�s���H�A�q�ܡA�a�}
        '��O��A2B�h�A�촣�f�Ƚs�A�DA2B�h���f�Ƚs
'        If myExCharFilter(Trim(rsMainT12("�q�����O"))) = "A2B" Then
'            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & myExCharFilter(Trim(rsMainT12("�f�D"))) & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT12("���f�Ȥ�s��"))) & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.CursorLocation = 3
'            tmp_Rs.Open str_SQL, cn
'        Else
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT12("�u�t"))) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
'        End If
        '�۲Ūu���«Ƚs
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close

        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT12("��f"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
'            If UCase(Right(Trim(rsMainT12("�ܧO")), 2)) = "-C" Then
'                strFacility = "�ըƹF����"
'            ElseIf UCase(Right(Trim(rsMainT12("�ܧO")), 2)) = "-S" Then
'                strFacility = "�ըƹF�n��"
'            Else
'                strFacility = "�ըƹF�_��"
'            End If
            
'            If Trim(rsMainT12("�ܧO")) = "" Then strFacility = ""
            
            strOrderDate = Trim(rsMainT12("�ѳf���"))
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
'            If myExCharFilter(Trim(rsMainT12("�q�����O"))) = "A2B" Then
            'A2B�A�h�����@��B�I���ȽsB_company
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT12("�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT12("�q�����O"))) & "','" & myExCharFilter(Trim(rsMainT12("�f�D"))) & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT12("��f��"))) & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT12("���f�Ȥ�s��"))) & "','" & myExCharFilter(Trim(rsMainT12("��f�Ȥ�s��"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT12("�q��Ƶ�"))) & "','" & filLocalFileT12.FileName & "','','" & User_id & "','" & User_id & "','','','" & Trim(rsMainT12("�ճ����O")) & "','" & Val(Trim(rsMainT12("���"))) & "') "
'            Else
            'not A2B
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT12("��f"))) & "','" & "RC" & "','" & str_Storerkey & "','" & myExCharFilter(Trim(rsMainT12("�ѳf���"))) & "','" & myExCharFilter(Trim(rsMainT12("�ѳf���"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT12("�u�t"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & "" & "','" & filLocalFileT12.FileName & "','','" & User_id & "','" & User_id & "','','','" & "RC" & "','" & "" & "') "
'            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT12("��f")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            
            str_SQL = "select CaseCnt as CN, sku from gv_skuxpack where sku='" & Trim(rsMainT12("����")) & "' and Storerkey = '" & str_Storerkey & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            
            IntCasecnt = Trim(tmp_Rs("CN"))
            intQTY = Val(rsMainT12("��f�ƶq") * IntCasecnt)
            
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,otherUOM)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT12("��f"))) & "','" & myExCharFilter(Trim(rsMainT12("����"))) & "','" & str_Storerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT12("�妸"))) & "','','" & strFacility & "','','" & myExCharFilter(Trim(rsMainT12("BUn"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT12.MoveNext
Loop

'����gs_ordersupdate   �ΫȤ�D�ɧ�s�q����
'
'cn.Execute "exec gs_ordersupdate 'LYFY09'", RowsAffect, adExecuteNoRecords

'����gs_ordersupdate
cn.Execute "exec gs_ordersupdate 'LPSI01'", RowsAffect, adExecuteNoRecords

'�ˬd���L���`�q��  �ܧO�w�O���~ es_Checklot06_by_storer '�f�D','�q���ɦW'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT12.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '���^�ǿ��~���q���ơA����excel
'    Recordset2Excel "�t�e�ܧO�P���ӭܧO���Ū��q����", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT12.Enabled = True


'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�ɮ� " & filLocalFileT12.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT12.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ��ɮ�
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT12.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT12.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT12.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT12.FileName, ".", -1)
End If

''�ƥ���FTP
'If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
'FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT12.FileName


Kill strTranFileName
    
filLocalFileT12.Refresh:
Screen.MousePointer = 0: cmdImportT12.Enabled = True: dgMainT12.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT12.Enabled = True: Screen.MousePointer = 0: dgMainT12.Enabled = True

End Sub
Private Sub dgMainT12_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT12
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT12_Change()
    filLocalFileT12.Path = dirLocalDirT12.Path
End Sub
Private Sub drvLocalDriveT12_Change()

On Error GoTo DriveError
dirLocalDirT12.Path = drvLocalDriveT12.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT12_Click()

On Error GoTo err_Handle
Set rsMainT12 = Nothing: Set dgMainT12.DataSource = rsMainT12
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT12.Path, 1) = "\" Then
    strFilePath = filLocalFileT12.Path
Else
    strFilePath = filLocalFileT12.Path & "\"
End If

If Dir(strFilePath & filLocalFileT12.FileName) = "" Then: filLocalFileT12.Refresh: Exit Sub

cboSheetT12.Clear

If UCase(mySplit(filLocalFileT12.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT12.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT12.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT12.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT12.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT12.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT12_Click()

On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT12.Path, 1) <> "\" Then
    strFilePath = filLocalFileT12.Path & "\"
Else
    strFilePath = filLocalFileT12.Path
End If

'strFieldName = "Delivery" & Chr(9) & "Sold-to.Pt" & Chr(9) & "Name.of.sold-to" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Name.of.the.ship-to.Party" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "Route" & Chr(9) & "Deliv.date" & Chr(9) & "Qty(stckpg.unit)" & Chr(9) & "BUn" & Chr(9) & "Delivery.qty" & Chr(9) & "SU" & Chr(9) & "order.no" & Chr(9) & "po.no" & Chr(9) & "remarks" & Chr(9)

Set rsMainT12 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT12.FileName, cboSheetT12, strFieldName, rsMainT12)

Set dgMainT12.DataSource = rsMainT12

If rsMainT12 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
'
Else

rsMainT12.Sort = "��f"

    SetDataGridColWidth Me.Caption, dgMainT12
    MsgBox "���u�@��@ " & rsMainT12.RecordCount & "������", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Sub Excel2RecordsetT12(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'�]�����W�٭��ơA�ҥH�W�ߦ��Ƶ{���A���F���L�Ĥ@�����W��
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '�䤣���ɮ�

On Error GoTo err_Handle
Screen.MousePointer = 11

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '���Ĥ@�Ӥu�@��W��
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '�䤣����w�u�@��A��βĤ@��
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 2 '�ѲĤG�C�}�l�פJ
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '�g�JRecordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & Chr(64 + j) & k & ")�A��ƬO�_���~�I"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub
Private Sub cboSheetT13_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT13.Path, 1) = "\" Then
    strFilePath = filLocalFileT13.Path
Else
    strFilePath = filLocalFileT13.Path & "\"
End If

'�إ����W�ٰ}�C
'strFieldName = "�Ȥ�" & Chr(9) & "�νs" & Chr(9) & "�P�f���" & Chr(9) & "�P�f�渹" & Chr(9) & "�p���H" & Chr(9) & "�q��" & Chr(9) & "�e�f�a�}" & Chr(9) & "�o�����X" & Chr(9) & "�~�ȭ�" & Chr(9) & "�Ƶ�" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�ܮw" & Chr(9) & "�ƶq" & Chr(9) & "���" & Chr(9) & "�e�m���/�Ƶ�/�Ȥ�渹" & Chr(9)

If Right(filLocalFileT13.Path, 1) <> "\" Then
    strFilePath = filLocalFileT13.Path & "\"
Else
    strFilePath = filLocalFileT13.Path
End If

Set rsMainT13 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT13.FileName, cboSheetT13, strFieldName, rsMainT13)
rsMainT13.Sort = "�P�f�渹,�~��"

Set dgMainT13.DataSource = rsMainT13

If rsMainT13 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT13
    MsgBox "���u�@��@ " & rsMainT13.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT13_Click()

If Len(RTrim(cboStorerkeyT13)) = 0 Then MsgBox "�п�f�D�s���I", 64, "�q����J": Exit Sub

On Error GoTo err_Handle

strTranFileName = filLocalFileT13.Path & "\" & filLocalFileT13.FileName
If Len(RTrim(cboSheetT13)) = 0 Then MsgBox "�п�ܤu�@��", 64, Me.Caption: Exit Sub
If rsMainT13.EOF Or rsMainT13 Is Nothing Then Exit Sub
Dim strStorerkey As String

strStorerkey = cboStorerkeyT13

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource) = '" & filLocalFileT13.FileName & "' and storerkey = '" & strStorerkey & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT13.Enabled = True: dgMainT13.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT13.MoveFirst
Do While Not rsMainT13.EOF

    '��f����ˬd
    arrTmp = Split(Trim(rsMainT13("�P�f���")), "/")
    If UBound(arrTmp) < 2 Then MsgBox "�P�f����榡���~(YYYY/MM/DD)�A�q����J�פ�!", 16, Me.Caption: Exit Sub
    If IsDate(Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)) = False Then MsgBox "�P�f������~(" & Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) & ")�A�q����J�פ�!", 16, Me.Caption: Exit Sub
    If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then MsgBox "��f��p�󤵤�A�q����J�פ�!", 16, Me.Caption: Exit Sub
            
    '�ƶq�ˬd
    If Val(rsMainT13("�ƶq")) < 0 Then
        MsgBox "�q��ƶq�p��1�A" & Trim(rsMainT13("�~��")) & "-" & Trim(rsMainT13("�~�W")) & "(" & Trim(rsMainT13("�ƶq")) & Trim(rsMainT13("���")) & ")�A�q����J�פ�!!", , "�q���ɶפJ": Exit Sub
        Exit Sub
    End If
                 
    rsMainT13.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT13.Enabled = False: dgMainT13.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'���̫�Ȥ�s��
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strStorerkey & "' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT13.MoveFirst
Do While Not rsMainT13.EOF
    DoEvents: DoEvents
    
'    '�������--�P�_�O�_�ݨըƹF�q��-->���U�@��
'    If UCase(Trim(rsMainT11("�ܮw"))) = "" Then
''        MsgBox "�Ȥ�渹�G" & Trim(rsMainT4("�P�f�渹")) & "( " & Trim(rsMainT4("�ܮw")) & " )" & vbCrLf & "�гq���Ȥ�A�T�{�ӵ��q��O�_���~!?", vbOKOnly, "�D�ըƹF���q�椣��J"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '�������--�P�_SKU�O�_�s�b
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT13("�~��")) & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cn.RollbackTrans: Tran_Level = 0
        MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT13("�~��")) & " ) " & Trim(rsMainT13("�~�W")) & "�A�q����J�פ�!!": cmdImportT13.Enabled = True: dgMainT13.Enabled = True: Screen.MousePointer = 0
        tmp_Rs.Close
        Exit Sub
    End If
    tmp_Rs.Close
    
'    End If
                  
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT13("�P�f�渹"))) Then
        strOrderNo = UCase(Trim(rsMainT13("�P�f�渹")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�ˬd�O�_�����Ȥ�W��
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = '" & strStorerkey & "' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT13("�Ȥ�"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '�L���Ȥ�W�٫h�s�W
            intTmp = intTmp + 1
            strConsigneeKey = "BEST" & Format(intTmp, "000000")
            
            '�s�W�Ȥ�D��
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('" & strStorerkey & "','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT13("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT13("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT13("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT13("�q��"))) & "','" & myExCharFilter(Trim(rsMainT13("�e�f�a�}"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '�����s�W���Ȥ�s��
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '����p���H�B�q�ܻP��f�a�}�O�_�۲�
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = '" & strStorerkey & "' and full_name = '" & myExCharFilter(Trim(rsMainT13("�Ȥ�"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT13("�p���H"))) & "' " & _
                        "and rtrim(phone) = '" & myExCharFilter(Trim(rsMainT13("�q��"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT13("�e�f�a�}"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '�p���H�B�q�ܻP��f�a�}����
                intTmp = intTmp + 1
                strConsigneeKey = Left(strStorerkey, 4) & Format(intTmp, "000000")
                
                '�s�W�Ȥ�D��
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes) " & _
                " values('" & strStorerkey & "','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT13("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT13("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT13("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT13("�q��"))) & "','" & myExCharFilter(Trim(rsMainT13("�e�f�a�}"))) & "','" & myExCharFilter(Trim(tmp_Rs("notes"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '�����s�W���Ȥ�s��
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '�۲Ūu���«Ƚs
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
            rsTmp.Close
        End If
        tmp_Rs.Close
    
        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT13("�P�f�渹"))) & "' and storerkey = '" & strStorerkey & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            strFacility = "�ըƹF�_��"
            arrTmp = Split(Trim(rsMainT13("�P�f���")), "/")
            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT13("�P�f�渹"))) & "','I','" & strStorerkey & "','" & strOrderDate & "','" & CDate(strOrderDate) + 1 & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT13("�Ȥ�"))) & "','" & myExCharFilter(Trim(rsMainT13("�p���H"))) & "','" & myExCharFilter(Trim(rsMainT13("�~�ȭ�"))) & "','" & myExCharFilter(Trim(rsMainT13("�νs"))) & "','" & myExCharFilter(Trim(rsMainT13("�q��"))) & "','','" & myExCharFilter(Trim(GetWord(rsMainT13("�e�f�a�}"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT13("�e�f�a�}"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT13("�Ƶ�"))) & "','" & filLocalFileT13.FileName & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT13("�o�����X"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & strStorerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT13("�P�f�渹")) & "','"
            blDuplicationOrder = True

        End If
    End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Trim(rsMainT13("�ƶq"))
            strLot06 = Trim(rsMainT13("�ܮw"))
            
            '�q����Ӹ�Ʒs�W
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT13("�P�f�渹"))) & "','" & myExCharFilter(Trim(rsMainT13("�~��"))) & "','" & strStorerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','F1','" & myExCharFilter(Trim(rsMainT13("���"))) & "','0','" & myExCharFilter(Trim(rsMainT13("�e�m���/�Ƶ�/�Ȥ�渹"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT13.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'�T�����
    msg_text = "�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " ������ (�@ " & rsMainT13.RecordCount & " ������)�ATMS�渹�G " & strOrderKeyS & "~" & str_Orderkey & "�A�ɮ� " & filLocalFileT13.FileName & " �ƥ��� C:\BEST\Other\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��A " & int_OrderLine & " �����ӡA�D�ըƹF�q�� " & intNotBest & " �����ӡA�ɮ� " & filLocalFileT13.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , �����ɮצW�� = '" & filLocalFileT13.FileName & "' , �W���ɮצW�� = o.updatesource ,���ƭq�渹�X = rtrim(o.externorderkey) ,�W���Ȥ�渹 = rtrim(o.customerorderkey) ,  �W���q���� = convert(varchar,o.orderdate,111) , �W����f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �W���~�� = od.sku , �W���ƶq = od.openqty ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\Best\Other\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\Other\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\Other\�q�歫��\" & strStorerkey & "�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''�ƥ���FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'�ƥ��ɮ�
If Dir("C:\BEST\Other\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\Other\Orders\Backup"
If Dir("C:\BEST\Other\Orders\Backup\" & filLocalFileT11.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\Other\Orders\Backup\" & filLocalFileT13.FileName
Else
    FileCopy strTranFileName, "C:\BEST\Other\Orders\Backup\" & mySplit(filLocalFileT13.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT13.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT13.Refresh: cboSheetT13.Clear
Screen.MousePointer = 0: cmdImportT13.Enabled = True: dgMainT13.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�פJ", Me.Caption, "cmdImportT13_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT13.Enabled = True: Screen.MousePointer = 0: dgMainT13.Enabled = True

End Sub

Private Sub dgMainT13_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT13
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT13_Change()
    filLocalFileT13.Path = dirLocalDirT13.Path
End Sub
Private Sub drvLocalDriveT13_Change()

On Error GoTo DriveError
dirLocalDirT13.Path = drvLocalDriveT13.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT13_Click()

On Error GoTo err_Handle
Set rsMainT13 = Nothing: Set dgMainT13.DataSource = rsMainT13
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT13.Path, 1) = "\" Then
    strFilePath = filLocalFileT13.Path
Else
    strFilePath = filLocalFileT13.Path & "\"
End If

If Dir(strFilePath & filLocalFileT13.FileName) = "" Then: filLocalFileT13.Refresh: Exit Sub

cboSheetT13.Clear

If UCase(mySplit(filLocalFileT13.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT13.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT13.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT13.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT13.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT13.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel "TEST", rsMainT4

'..�b���s��EXCEL
With MyXlsApp
    
End With

Set MyXlsApp = Nothing

End Sub
Private Sub cboSheetT14_Click()

On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT14.Path, 1) <> "\" Then
    strFilePath = filLocalFileT14.Path & "\"
Else
    strFilePath = filLocalFileT14.Path
End If

'strFieldName = "Delivery" & Chr(9) & "Sold-to.Pt" & Chr(9) & "Name.of.sold-to" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Name.of.the.ship-to.Party" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "Route" & Chr(9) & "Deliv.date" & Chr(9) & "Qty(stckpg.unit)" & Chr(9) & "BUn" & Chr(9) & "Delivery.qty" & Chr(9) & "SU" & Chr(9) & "order.no" & Chr(9) & "po.no" & Chr(9) & "remarks" & Chr(9)

Set rsMainT14 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT14.FileName, cboSheetT14, strFieldName, rsMainT14)

Set dgMainT14.DataSource = rsMainT14

If rsMainT14 Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
'
Else

rsMainT14.Sort = "�f�D,�q�渹�X"

    SetDataGridColWidth Me.Caption, dgMainT14
    MsgBox "���u�@��@ " & rsMainT14.RecordCount & "������", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT14_Click()
Dim bl_Error As Boolean '�O�����p���I���X��
Dim str_Error As String '�O�����p���I���~�����
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT14 Is Nothing Then Exit Sub
If rsMainT14.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT14.Enabled = False: cmdImportT14.Enabled = False
strTranFileName = filLocalFileT14.Path & "\" & filLocalFileT14.FileName

'�������--�P�_�ɮ׬O�_�w��J
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT14.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "�ɮצW�٬ۦP�A�нT�{�O�_������J!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT14.RecordCount = 0 Or rsMainT14 Is Nothing Then
Else
rsMainT14.MoveFirst
str_Storerkey = myExCharFilter(Trim(rsMainT14("�f�D")))

Do While Not rsMainT14.EOF
    '��f����ˬd
    If Len(Trim(rsMainT14("��f��"))) = 0 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "����f�鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT14("��f��"))) > 0 And Len(Trim(rsMainT14("��f��"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "����f��:" & Trim(rsMainT14("��f��")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT14("��f��")), 4) + "/" + Mid(Trim(rsMainT14("��f��")), 5, 2) + "/" + Right(Trim(rsMainT14("��f��")), 2)) = False Then
         MsgBox "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "����f��:" & Trim(rsMainT14("��f��")) & "�A���O�@�ӥ��`����A�q����J�פ�!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    Else
        '�ˬd��f�餣�i�p�󤵤�
        If Trim(rsMainT14("��f��")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '�̰��v�����ˬd��f��
                 x = MsgBox("��f��p�󤵤�A�A�T�w�n�~���?", vbQuestion + vbYesNo, "�̰��v����f���ˬd") '�������U���O�T�w�άO����
                    If x = 6 Then
                        '�~��
                    Else
                        '���}
                         dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
    '�q����ˬd
    If Len(Trim(rsMainT14("�q���"))) = 0 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "���q��鬰�ťաA�q����J�פ�!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT14("�q���"))) > 0 And Len(Trim(rsMainT14("�q���"))) < 8 Then
         MsgBox "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "���q���:" & Trim(rsMainT14("�q���")) & "�A�榡����A�иɻ�8�X�A�q����J�פ�!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT14("�q���")), 4) + "/" + Mid(Trim(rsMainT14("�q���")), 5, 2) + "/" + Right(Trim(rsMainT14("�q���")), 2)) = False Then
         MsgBox "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "���q���:" & Trim(rsMainT14("��f��")) & "�A���O�@�ӥ��`����A�q����J�פ�!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub

    Else
        If Trim(rsMainT14("�q���")) > Trim(rsMainT14("��f��")) Then MsgBox "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "���q���:" & Trim(rsMainT14("�q���")) & "�A�j���f��A�q����J�פ�!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    End If
    
    '�ƶq�ˬd
    If Val(rsMainT14("�ƶq")) < 1 Then
        MsgBox "�ƶq�p��1�A" & Trim(rsMainT14("�q�渹�X")) & "-�~���G" & Trim(rsMainT14("�~��")) & "�A�q����J�פ�!!�нT�{!!", , "�q���ɶפJ": dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '������� --�P�_SKU�O�_�s�b
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT14("�~��")) & "' and Storerkey = '" & Trim(rsMainT14("�f�D")) & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�~�� (" & Trim(rsMainT14("�~��")) & ")�A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        If Trim(rsMainT14("�q�����O")) = "A2B" Then
        '�ˬdA2B�q��Ȥ�s���O�_�s�b
                str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT14("���f�Ȥ�s��")) & "' and Storerkey = '" & Trim(rsMainT14("�f�D")) & "' "
            
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            
                If tmp_Rs.EOF Then  '���s���ǭn��
                    MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT14("���f�Ȥ�s��")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
                    dgMainT14.Enabled = True: cmdImportT14.Enabled = True
                    Exit Sub
                End If
        End If
        
        '�ˬdA2B�q��H�~���Ȥ�s���O�_�s�b
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT14("��f�Ȥ�s��")) & "' and Storerkey = '" & Trim(rsMainT14("�f�D")) & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�Ȥ�s�� ( " & Trim(rsMainT14("��f�Ȥ�s��")) & " )�A�Х���Ȥ�D�ɷs�ثȤ��ơA�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        '�ˬd�ƶq���L�p���I
        If InStr(Trim(rsMainT14("�ƶq")), ".") <> 0 Then
            str_Error = "�q�渹�X:" & Trim(rsMainT14("�q�渹�X")) & "�A�~��:" & Trim(rsMainT14("�~��")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
'        '�ˬd�f�D--mark by Gemini @20160602
'        If UCase(Trim(rsMainT14("�f�D"))) <> "LABT01" And UCase(Trim(rsMainT14("�f�D"))) <> "LLFA01" Then
'            MsgBox "�q��o�{�D�Ȱ����f�D: " & Trim(rsMainT14("�f�D")) & " )�A���פJ�{���ȨѶפJ�Ȱ��ΧQ�׭q��A�нT�{��A�פJ�A�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
'            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
'            Exit Sub
'        End If

        '�P�_��O--add by  Gemini @20160602
        If Trim(rsMainT14("�q�����O")) <> "C" Or Trim(rsMainT14("�q�����O")) <> "R" Or Trim(rsMainT14("�q�����O")) <> "RC" Or Trim(rsMainT14("�q�����O")) <> "I" Or Trim(rsMainT14("�q�����O")) <> "A2B" Then
        Else
            MsgBox "�t�εL����O:" & Trim(rsMainT14("�q�����O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        '�ˬd�f�D�O�_�s�b
        str_SQL = "select storerkey from trp16m where storerkey = '" & Trim(rsMainT14("�f�D")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then  '���s���ǭn��
            MsgBox "�q��o�{�s�f�D ( " & Trim(rsMainT14("�f�D")) & " )�A�Х���f�D�D�ɷs�سf�D��ơA�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        '�P�_C��
'        If Trim(rsMainT14("�q�����O")) = "C" And (Trim(rsMainT14("�f�D")) = "LKAO01" Or Trim(rsMainT14("�f�D")) = "LABT01") Then
'        Else
'            MsgBox "���f�D:" & Trim(rsMainT14("�f�D")) & "���q�����O���i��:" & Trim(rsMainT14("�q�����O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
'                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
'            Exit Sub
'        End If
        
'        '�P�_�O�_��A2B��R��
'        If Trim(rsMainT14("�q�����O")) <> "A2B" And Trim(rsMainT14("�q�����O")) <> "R" And UCase(Trim(rsMainT14("�f�D"))) = "LABT01" Then
'            MsgBox "���f�D:" & Trim(rsMainT14("�f�D")) & "���q�����O���i��:" & Trim(rsMainT14("�q�����O")) & "�A�нT�{���q�����O�O�_���T�A�q����J�פ�!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
'                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
'            Exit Sub
'        End If
        
    rsMainT14.MoveNext
Loop
rsMainT14.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "��Ʀ��p���I�I�Э��s�פJ"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT14.Enabled = False: dgMainT14.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'�}�l�פJ
Do While Not rsMainT14.EOF
    DoEvents: DoEvents
    '�������--�ӷ��q��ۦP�渹�P�_�A���P�W�[HEAD
    If strOrderNo <> UCase(Trim(rsMainT14("�q�渹�X"))) Then
        strOrderNo = UCase(Trim(rsMainT14("�q�渹�X")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '�w�g�ˬd�L�A�h�����A������X�Ȥ�D�ɤ��� �Ȥ�s���A�l���ϸ��A�s���H�A�q�ܡA�a�}
        '��O��A2B�h�A�촣�f�Ƚs�A�DA2B�h���f�Ƚs
        If myExCharFilter(Trim(rsMainT14("�q�����O"))) = "A2B" Then
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT14("���f�Ȥ�s��"))) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
        Else
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT14("��f�Ȥ�s��"))) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
        End If
        '�۲Ūu���«Ƚs
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close

        '�������--�P�_�q��O�_���ơA���Ƥ��W�[
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT14("�q�渹�X"))) & "' and storerkey = '" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "' and isnull(type,'') <> '�R��' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '���q�渹�X
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '�t�e�ܧO�P�_
            If UCase(Right(Trim(rsMainT14("�ܧO")), 2)) = "-C" Then
                strFacility = "�ըƹF����"
            ElseIf UCase(Right(Trim(rsMainT14("�ܧO")), 2)) = "-S" Then
                strFacility = "�ըƹF�n��"
            Else
                strFacility = "�ըƹF�_��"
            End If
            
            If Trim(rsMainT14("�ܧO")) = "" Then strFacility = ""
            
            strOrderDate = Trim(rsMainT14("�q���"))
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            If myExCharFilter(Trim(rsMainT14("�q�����O"))) = "A2B" Then
                
            'A2B�A�h�����@��B�I���ȽsB_company
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT14("�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT14("�q�����O"))) & "','" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT14("��f��"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT14("���f�Ȥ�s��"))) & "','" & myExCharFilter(Trim(rsMainT14("��f�Ȥ�s��"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT14("���ʳ渹")) & "','" & myExCharFilter(Trim(rsMainT14("�q��Ƶ�"))) & "','" & filLocalFileT14.FileName & "','','" & User_id & "','" & User_id & "','','','" & Trim(rsMainT14("�ճ����O")) & "','" & Val(Trim(rsMainT14("���"))) & "') "
            Else
            'not A2B
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT14("�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT14("�q�����O"))) & "','" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT14("��f��"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT14("��f�Ȥ�s��"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT14("���ʳ渹")) & "','" & myExCharFilter(Trim(rsMainT14("�q��Ƶ�"))) & "','" & filLocalFileT14.FileName & "','','" & User_id & "','" & User_id & "','','','" & Trim(rsMainT14("�ճ����O")) & "','" & Val(Trim(rsMainT14("���"))) & "') "
            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
'            '�p�G�Ȥ�D�ɬ۲šA��s�q��l���ϸ��A�H�K�ݫȤ�T�{
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '�q�歫��
            Call FTPlog("�q�歫��" & str_SQL)
            '��������
            strReOrderkey = strReOrderkey & Trim(rsMainT14("�q�渹�X")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '�q�歫���ˬd
        If blDuplicationOrder = False Then
        
            '�W�[����
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Val(rsMainT14("�ƶq"))
            
            
            '�q����Ӹ�Ʒs�W
            If Trim(rsMainT14("���W��")) = "�c" Or Trim(rsMainT14("���W��")) = "CS" Or Trim(rsMainT14("���W��")) = "CASE" Then
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
                " select '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT14("�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT14("�~��"))) & "','" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "'," & _
                "'" & intQTY & "' * p.casecnt ,'" & intQTY & "' * p.casecnt,'" & myExCharFilter(Trim(rsMainT14("�ܧO"))) & "','',''" & _
                "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT14("�~��")) & "' and s.storerkey = '" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "' "
            Else
                 str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
                "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT14("�q�渹�X"))) & "','" & myExCharFilter(Trim(rsMainT14("�~��"))) & "','" & myExCharFilter(Trim(rsMainT14("�f�D"))) & "'," & _
                "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT14("�ܧO"))) & "','','')"
           End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��spackkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT14.MoveNext
Loop

'����gs_ordersupdate   �ΫȤ�D�ɧ�s�q����
'
'cn.Execute "exec gs_ordersupdate 'LYFY09'", RowsAffect, adExecuteNoRecords

'�ˬd���L���`�q��  �ܧO�w�O���~ es_Checklot06_by_storer '�f�D','�q���ɦW'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT14.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '���^�ǿ��~���q���ơA����excel
'    Recordset2Excel "�t�e�ܧO�P���ӭܧO���Ū��q����", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT14.Enabled = True


'�T�����
    msg_text = "�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "TMS�渹�G " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("�פJ " & int_Order & " ���q��" & Chr(13) & "�פJ" & int_OrderLine & " ������" & Chr(13) & "�ɮ� " & filLocalFileT14.FileName)
    
'�q�歫�����
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select �������O = '�q�歫��-����J' , ��J�ɮצW�� = '" & filLocalFileT14.FileName & "' ,�q�渹�X = rtrim(o.externorderkey) ,�Ȥ�渹 = rtrim(o.customerorderkey) ,  �q���� = convert(varchar,o.orderdate,111) , ��f�� = convert(varchar,o.deliverydate,111) , ���� = od.orderlinenumber , �~�� = od.sku , �ƶq = od.openqty , �W���ɮצW�� = o.updatesource  ,�ˬd�ɶ� = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "�o�{�q�歫��!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("�q�歫��", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\�q�歫��", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\�q�歫��"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\�q�歫��\�q�歫��_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'�ƥ��ɮ�
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT14.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT14.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT14.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT14.FileName, ".", -1)
End If

'�ƥ���FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT14.FileName



Kill strTranFileName
    
filLocalFileT14.Refresh:
Screen.MousePointer = 0: cmdImportT14.Enabled = True: dgMainT14.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "�ɮצW�١G " & strTranFileName)
    cmdImportT14.Enabled = True: Screen.MousePointer = 0: dgMainT14.Enabled = True

End Sub

Private Sub dgMainT14_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT14
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT14_Change()
    filLocalFileT14.Path = dirLocalDirT14.Path
End Sub

Private Sub drvLocalDriveT14_Change()

On Error GoTo DriveError
dirLocalDirT14.Path = drvLocalDriveT14.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT14_Click()

On Error GoTo err_Handle
Set rsMainT14 = Nothing: Set dgMainT14.DataSource = rsMainT14
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFileT14.Path, 1) = "\" Then
    strFilePath = filLocalFileT14.Path
Else
    strFilePath = filLocalFileT14.Path & "\"
End If

If Dir(strFilePath & filLocalFileT14.FileName) = "" Then: filLocalFileT14.Refresh: Exit Sub

cboSheetT14.Clear

If UCase(mySplit(filLocalFileT14.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT14.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheetT14.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT14.AddItem MyXlsApp.Sheets(i).Name
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT14.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT14.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Sub Excel2RecordsetT8(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'�]�����W�٭��ơA�ҥH�W�ߦ��Ƶ{���A���F���L�Ĥ@�����W��
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '�䤣���ɮ�

On Error GoTo err_Handle
Screen.MousePointer = 11

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '���Ĥ@�Ӥu�@��W��
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '�䤣����w�u�@��A��βĤ@��
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 2 '�ѲĤG�C�}�l�פJ
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '�g�JRecordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & Chr(64 + j) & k & ")�A��ƬO�_���~�I"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub
Sub Excel2RecordsetT9(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'�]�����W�٭��ơA�ҥH�W�ߦ��Ƶ{���A���F���L�Ĥ@�����W��
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '�䤣���ɮ�

On Error GoTo err_Handle
Screen.MousePointer = 11

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '���Ĥ@�Ӥu�@��W��
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '�䤣����w�u�@��A��βĤ@��
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 2 '�ѲĤG�C�}�l�פJ
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '�g�JRecordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & Chr(64 + j) & k & ")�A��ƬO�_���~�I"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub

Private Sub DB_Connect_Self(connection_string As String)
'ADO [Connection] Object connect
On Error GoTo err_Handle
Set cn_Self = New ADODB.Connection
cn_Self.CommandTimeout = 300
cn_Self.ConnectionTimeout = 20
cn_Self.ConnectionString = connection_string
cn_Self.Open Options:=adAsyncConnect
Do While cn_Self.State = adStateConnecting
   DoEvents: DoEvents
Loop
Exit Sub

err_Handle:
   msg_text = "�s�u���~�G�L�k�P��Ʈw�إ߳s�u�A�гq�� ��T�� "
   MsgBox msg_text, vbOKOnly + vbInformation, ""
   End
End Sub
Sub XML2Recordset(FileName As String)
Dim i As Integer, arrLen
arrLen = Array(12, 6, 12, 2, 50, 15, 19, 12, 6, 19, 60, 15, 15, 15, 10, 100, 20, 60, 8, 6, 3, 3, 10, 19, 20, 76, 76, 3, 10, 10, 60, 255, 19, 20)

'�ɮת��׬�0
If FileLen(FileName) = 0 Then Call ErrorMsgbox(Me.Caption, err.Number, err.Description, FileName & "�ɮת��׬� 0 "): Exit Sub

Set rs_Src = Nothing

Dim objXMLDOM As New MSXML2.DOMDocument40
Dim objNodes As IXMLDOMNodeList
Dim objBookNode As IXMLDOMNode

'�}�lŪ��xml
objXMLDOM.async = False

'�}��xml�ɿ��~
If Not objXMLDOM.Load(FileName) Then Call ErrorMsgbox(Me.Caption, err.Number, objXMLDOM.parseError.reason, FileName): Exit Sub

Set objNodes = objXMLDOM.selectNodes("/ROWSET/ROW")
If objNodes.Length <= 0 Then Call ErrorMsgbox(Me.Caption, err.Number, "objNodes.length = " & objNodes.Length, FileName): Exit Sub

Set rs_Src = New ADODB.Recordset
With rs_Src
    .Fields.Append "DELIVERY_NO", adChar, 12, adFldUpdatable
    .Fields.Append "GOODS_OWNER", adChar, 6, adFldUpdatable
    .Fields.Append "DELIVERY_DETAIL_ID", adChar, 12, adFldUpdatable
    .Fields.Append "ORDER_TYPE", adChar, 2, adFldUpdatable
    .Fields.Append "TRANSACTION_TYPE_DESC", adChar, 50, adFldUpdatable
    .Fields.Append "SALESREP", adChar, 15, adFldUpdatable
    .Fields.Append "ORDERED_DATE", adChar, 19, adFldUpdatable
    .Fields.Append "SALES_ORDER_NUMBER", adChar, 12, adFldUpdatable
    .Fields.Append "SEQUENCE_NO", adChar, 6, adFldUpdatable
    .Fields.Append "NOTIFY_DATE", adChar, 19, adFldUpdatable
    .Fields.Append "CUSTOMER_NAME", adChar, 60, adFldUpdatable
    .Fields.Append "CONTACT", adChar, 15, adFldUpdatable
    .Fields.Append "TELEPHONE", adChar, 15, adFldUpdatable
    .Fields.Append "SINGLE_SHOP_CODE", adChar, 15, adFldUpdatable
    .Fields.Append "SUPPLIER_CODE", adChar, 10, adFldUpdatable
    .Fields.Append "REQUEST_ADDRESS", adChar, 100, adFldUpdatable
    .Fields.Append "ITEM_NAME", adChar, 20, adFldUpdatable
    .Fields.Append "ITEM_DESCRIPTION", adChar, 60, adFldUpdatable
    .Fields.Append "QTY1", adChar, 8, adFldUpdatable
    .Fields.Append "QTY2", adChar, 6, adFldUpdatable
    .Fields.Append "UOM1", adChar, 3, adFldUpdatable
    .Fields.Append "UOM2", adChar, 3, adFldUpdatable
    .Fields.Append "SELLING_PRICE", adChar, 10, adFldUpdatable
    .Fields.Append "SCHEDULE_SHIP_DATE", adChar, 19, adFldUpdatable
    .Fields.Append "CUST_PO_NUMBER", adChar, 20, adFldUpdatable
    .Fields.Append "DELIVERY_INFORMATIONS", adChar, 76, adFldUpdatable
    .Fields.Append "SHIPPING_INSTRUCTIONS", adChar, 76, adFldUpdatable
    .Fields.Append "ORGANIZATION_CODE", adChar, 3, adFldUpdatable
    .Fields.Append "SUBINVENTORY", adChar, 10, adFldUpdatable
    .Fields.Append "POSTAL_CODE", adChar, 10, adFldUpdatable
    .Fields.Append "SHIP_TO_CUSTOMER_NAME", adChar, 60, adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open

'�g�Jrs

For Each objBookNode In objNodes

       .AddNew
       For i = 0 To rs_Src.Fields.Count - 1
         .Fields(i) = Replace(GetWord((objBookNode.selectSingleNode(rs_Src.Fields(i).Name).nodeTypedValue), 1, arrLen(i)), "'", """")
       Next i
       .Update

Next objBookNode

.Sort = "SALES_ORDER_NUMBER,SEQUENCE_NO"
.MoveFirst

Set dgMainT20.DataSource = rs_Src

End With

End Sub
Sub ExceltoRecordset(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'Create by Gemini @20090312 4 Excel�פJRecordset
'�ϥλ���
'1.�p�G�ӷ�Excel�u�@���a���W�١A�Щ�strFieldName���w�A�åHchar(9)�@�����j�Ÿ�
'strFieldName = "�Ȥ�" & Chr(9) & "�νs" & Chr(9) & "�P�f���" & Chr(9) & "�P�f�渹" & Chr(9) & "�p���H" & Chr(9) & "�q��" & Chr(9) & "�e�f�a�}" & Chr(9) & "�o�����X" & Chr(9) & "�~�ȭ�" & Chr(9) & "�Ƶ�" & Chr(9) & "�~��" & Chr(9) & "�~�W" & Chr(9) & "�ܮw" & Chr(9) & "�ƶq" & Chr(9) & "���" & Chr(9) & "�e�m���/�Ƶ�/�Ȥ�渹" & Chr(9)

'�Ѽƻ���
'strFileName:�ӷ��ɮצW�ٸ��|
'strSheetName:�ӷ��u�@��
'strFieldName:���W��
'rs:�^�Ǫ�Recordset
'�d��
'call Excel2Recordset ("C:\book1.xls","Sheet1", "�Ȥ�" & Chr(9) & "�νs" & Chr(9) & "�P�f���" & Chr(9),rsMain)
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '�䤣���ɮ�

On Error GoTo err_Handle
Screen.MousePointer = 11

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '�M����w�u�@��
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (strSheetName) Then .Sheets(strSheetName).Select: Exit For '��w�u�@��
    Next
    
    If (.ActiveSheet.Name) <> (strSheetName) Then
        MsgBox "�䤣�� " & strSheetName & " �u�@��I", vbOKOnly + vbInformation, "Excel2Recordset"
        .Quit: Set MyXlsApp = Nothing
        Exit Sub
    End If

    k = 1 '�w�]�ѲĤ@�C�}�l�פJ
    
    '�Y�L�ӷ����W��
    If strFieldName = "" Then
        '�����W��
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        k = 2 '�ѲĤG�C�}�l�פJ
    End If
    
    '�������W��
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '�g�JRecordset  '�q�o��}�l���U�g
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & k & Chr(64 + j) & ")�A��ƬO�_���~�I"
    
ElseIf err.Number = 13 Then
    str = "�нT�{�x�s��(" & k & Chr(64 + j) & ")�A��ƬO�_���~�I"

Else
     str = "�нT�{�x�s��(" & k & Chr(64 + j) & ")�A��ƬO�_���~�I"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
If Trim(SSTab2.Caption) = "" Then SSTab2.Tab = PreviousTab
End Sub


Public Sub LCDisConnect(IpAddress)
 
Dim str1 As String, strRun As String
 'yes���߰�
str1 = "NET use \\" & IpAddress & " /Delete /yes" & vbCrLf
strRun = str1
Shell "cmd.exe /c " & strRun, vbHide
 
End Sub

Public Sub LCConnect(IpAddress As String, ACC As String, PassWord As String)
 '���s�s�u
 'LCConnect "192.168.2.202", "share", "share"
Dim str1 As String, strRun As String
str1 = "NET use \\" & IpAddress & " " & PassWord & " /user:" & IpAddress & "\" & ACC & " /PERSISTENT:NO" & vbCrLf
strRun = str1
Shell "cmd.exe /c " & strRun, vbHide
 
End Sub
