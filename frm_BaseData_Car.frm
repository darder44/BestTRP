VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_BaseData_Car 
   Caption         =   "����/�f�B���q �򥻸�ƺ��@�@�~"
   ClientHeight    =   8130
   ClientLeft      =   285
   ClientTop       =   750
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   11475
   Begin TabDlg.SSTab SSTab1 
      Height          =   8040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   14182
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14215660
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm_BaseData_Car.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape1(1)"
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(2)=   "dg_Tab0_ConsigneeList"
      Tab(0).Control(3)=   "cmd_Tab0_ConsigneeQuery"
      Tab(0).Control(4)=   "cmd_Tab0_2Excel"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(6)=   "fam_Tab0_Consignee"
      Tab(0).Control(7)=   "cmd_Tab0_ConsigneeShow"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "�������"
      TabPicture(1)   =   "frm_BaseData_Car.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dg_Tab1_CarList"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmd_Tab1_CarQuery"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmd_Tab1_2Excel"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fam_Tab1_Car"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmd_Tab1_CarShow"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "�f�B���q"
      TabPicture(2)   =   "frm_BaseData_Car.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fam_Tab2_FunctionArea"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "fam_Tab2_Company"
      Tab(2).Control(3)=   "dg_Tab2_TRPCompanyList"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frm_BaseData_Car.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dg_Tab3_SkuList"
      Tab(3).Control(1)=   "fam_Tab3_Sku"
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(3)=   "Frame5"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frm_BaseData_Car.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dg_Tab4_AcceptableList"
      Tab(4).Control(1)=   "fam_Tab4_Acceptable"
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(3)=   "Frame8"
      Tab(4).Control(4)=   "CmnDialog"
      Tab(4).Control(5)=   "Frame7"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frm_BaseData_Car.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame"
      Tab(5).Control(1)=   "gd_Tab5_AaccessTable"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame7 
         Appearance      =   0  '����
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   550
         Left            =   -66720
         TabIndex        =   1
         Top             =   2880
         Width           =   2025
         Begin VB.CommandButton cmd_Tab4_SaveToExcel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�d�ߵ��G��Excel"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   50
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   2
            Top             =   120
            Width           =   1950
         End
      End
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   -67320
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_Tab0_ConsigneeShow 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��ܩҦ��Ȥ�"
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
         Left            =   -73005
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   223
         Top             =   390
         Width           =   1830
      End
      Begin VB.CommandButton cmd_Tab1_CarShow 
         BackColor       =   &H00FFC0FF&
         Caption         =   "��ܩҦ�����"
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
         Left            =   1980
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   222
         Top             =   405
         Width           =   1830
      End
      Begin VB.Frame fam_Tab0_Consignee 
         Appearance      =   0  '����
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   6360
         Left            =   -71160
         TabIndex        =   161
         Top             =   1545
         Width           =   7440
         Begin VB.ComboBox cmb_Tab0_ExtraDemand2 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   192
            Top             =   3615
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_ExtraDemand1 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   191
            Top             =   3210
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_VehicleType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   190
            Top             =   2805
            Width           =   6000
         End
         Begin VB.TextBox txt_Tab0_Phone 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3180
            MaxLength       =   30
            TabIndex        =   189
            Top             =   2430
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab0_Contact 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            TabIndex        =   188
            Top             =   2430
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab0_Fax 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   5355
            MaxLength       =   18
            TabIndex        =   187
            Top             =   2430
            Width           =   1695
         End
         Begin VB.TextBox txt_Tab0_CodeDate1 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   186
            ToolTipText     =   "�Q���s������;P&G������"
            Top             =   5160
            Width           =   840
         End
         Begin VB.TextBox txt_Tab0_CodeDate2 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2925
            MaxLength       =   10
            TabIndex        =   185
            ToolTipText     =   "�Q��M�s������"
            Top             =   5160
            Width           =   840
         End
         Begin VB.TextBox txt_Tab0_Notes 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            TabIndex        =   184
            Top             =   5880
            Width           =   5880
         End
         Begin VB.TextBox txt_Tab0_PalletSpec 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2925
            MaxLength       =   20
            TabIndex        =   183
            Top             =   5520
            Width           =   3960
         End
         Begin VB.TextBox txt_Tab0_PalletType 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   20
            TabIndex        =   182
            Top             =   5520
            Width           =   840
         End
         Begin VB.TextBox txt_Tab0_Stamp 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1005
            MaxLength       =   1
            TabIndex        =   181
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox txt_Tab0_Penalties 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2445
            MaxLength       =   1
            TabIndex        =   180
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox txt_Tab0_Channel 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3060
            MaxLength       =   20
            TabIndex        =   179
            Top             =   4440
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab0_ChannelType 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   178
            Top             =   4440
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab0_ConsigneeKey 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1020
            TabIndex        =   177
            Top             =   240
            Width           =   1980
         End
         Begin VB.ComboBox cmb_Tab0_Storer 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   300
            Left            =   3525
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   176
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txt_Tab0_FullName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   60
            TabIndex        =   175
            Top             =   600
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_Zip 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   174
            Top             =   960
            Width           =   1995
         End
         Begin VB.TextBox txt_Tab0_Class 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3885
            MaxLength       =   10
            TabIndex        =   173
            ToolTipText     =   "�Ӽh�ɶK���п�J�Ʀr�A�t�ιw�]=0"
            Top             =   960
            Width           =   705
         End
         Begin VB.TextBox txt_Tab0_GridCode 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   172
            Top             =   960
            Width           =   1395
         End
         Begin VB.TextBox txt_Tab0_ShortName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   60
            TabIndex        =   171
            Top             =   1680
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_AreaCode 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   170
            Top             =   1320
            Width           =   6000
         End
         Begin VB.TextBox txt_Tab0_Address 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            TabIndex        =   169
            Top             =   2040
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_PickTool 
            BackColor       =   &H00C0FFC0&
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
            Left            =   1020
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   168
            Top             =   4080
            Width           =   1890
         End
         Begin VB.TextBox txt_Tab0_UnLoad 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4080
            TabIndex        =   167
            Top             =   4080
            Width           =   420
         End
         Begin VB.CheckBox chk_Tab0_MultiCustomer 
            BackColor       =   &H8000000C&
            Caption         =   "���e�Ȥ�"
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
            Height          =   315
            Left            =   4440
            TabIndex        =   166
            Top             =   4800
            Width           =   1140
         End
         Begin VB.CheckBox chkDC 
            BackColor       =   &H8000000C&
            Caption         =   "�έܫȤ�"
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
            Height          =   315
            Left            =   3240
            TabIndex        =   165
            ToolTipText     =   "�Q�藍������ɶK�з�"
            Top             =   4800
            Width           =   1260
         End
         Begin VB.ComboBox cmdCodeDateRate 
            BackColor       =   &H00C0FFC0&
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
            ItemData        =   "frm_BaseData_Car.frx":00A8
            Left            =   6480
            List            =   "frm_BaseData_Car.frx":00B2
            TabIndex        =   164
            ToolTipText     =   "�Y�����w�A���_eOrder��J�ɹw�]��1/2�Ĵ�"
            Top             =   5160
            Width           =   930
         End
         Begin VB.ComboBox cmb_Tab0_Group 
            BackColor       =   &H00C0FFC0&
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
            Left            =   5100
            TabIndex        =   163
            Text            =   "cmb_Tab0_Group"
            Top             =   4420
            Width           =   1875
         End
         Begin VB.TextBox txt_Tab0_CodeDate3 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4845
            MaxLength       =   10
            TabIndex        =   162
            ToolTipText     =   "�Q�ﶼ�Ƥ�����"
            Top             =   5160
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�S��ݨD 2"
            Height          =   180
            Index           =   12
            Left            =   150
            TabIndex        =   221
            Top             =   3690
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�S��ݨD 1"
            Height          =   180
            Index           =   11
            Left            =   150
            TabIndex        =   220
            Top             =   3285
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���إN�X"
            Height          =   180
            Index           =   10
            Left            =   285
            TabIndex        =   219
            Top             =   2880
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            Height          =   180
            Index           =   9
            Left            =   2790
            TabIndex        =   218
            Top             =   2475
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p���H"
            Height          =   180
            Index           =   8
            Left            =   465
            TabIndex        =   217
            Top             =   2475
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ǯu"
            Height          =   180
            Index           =   55
            Left            =   4980
            TabIndex        =   216
            Top             =   2475
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "������1"
            Height          =   180
            Index           =   57
            Left            =   120
            TabIndex        =   215
            Top             =   5205
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "������2"
            Height          =   180
            Index           =   58
            Left            =   1920
            TabIndex        =   214
            Top             =   5205
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ƶ�"
            Height          =   180
            Index           =   59
            Left            =   240
            TabIndex        =   213
            Top             =   5925
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�̪O�W��"
            Height          =   180
            Index           =   60
            Left            =   2160
            TabIndex        =   212
            Top             =   5565
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�̪O����"
            Height          =   180
            Index           =   61
            Left            =   240
            TabIndex        =   211
            Top             =   5565
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�K��"
            Height          =   180
            Index           =   62
            Left            =   600
            TabIndex        =   210
            Top             =   4845
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�@�ګȤ�"
            Height          =   180
            Index           =   63
            Left            =   1680
            TabIndex        =   209
            Top             =   4845
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q���O"
            Height          =   180
            Index           =   45
            Left            =   2400
            TabIndex        =   208
            Top             =   4485
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q�����A"
            Height          =   180
            Index           =   13
            Left            =   240
            TabIndex        =   207
            Top             =   4485
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�s��"
            Height          =   180
            Index           =   16
            Left            =   240
            TabIndex        =   206
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D"
            Height          =   180
            Index           =   17
            Left            =   3120
            TabIndex        =   205
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�W��"
            Height          =   180
            Index           =   4
            Left            =   240
            TabIndex        =   204
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�l���ϸ�"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   203
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ӽh�ɶK"
            Height          =   180
            Index           =   7
            Left            =   3120
            TabIndex        =   202
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�x�}�ϽX"
            Height          =   180
            Index           =   18
            Left            =   4800
            TabIndex        =   201
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�²��"
            Height          =   180
            Index           =   6
            Left            =   240
            TabIndex        =   200
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϽX"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   199
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�a�}"
            Height          =   180
            Index           =   5
            Left            =   240
            TabIndex        =   198
            Top             =   2085
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�h�B�u��"
            Height          =   180
            Index           =   43
            Left            =   240
            TabIndex        =   197
            Top             =   4170
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���f������"
            Height          =   180
            Index           =   15
            Left            =   3120
            TabIndex        =   196
            Top             =   4140
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��������"
            Height          =   180
            Index           =   64
            Left            =   5760
            TabIndex        =   195
            Top             =   5205
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�s��"
            Height          =   180
            Index           =   69
            Left            =   4320
            TabIndex        =   194
            Top             =   4485
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "������3"
            Height          =   180
            Index           =   70
            Left            =   3840
            TabIndex        =   193
            Top             =   5205
            Width           =   630
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  '����
         BackColor       =   &H00004000&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71145
         TabIndex        =   154
         Top             =   345
         Width           =   7440
         Begin VB.CommandButton cmd_Tab0_AddNew 
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
            Left            =   1290
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   160
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Modify 
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
            Left            =   75
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   159
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Save 
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
            Left            =   2505
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   158
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Delete 
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
            Left            =   4935
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   157
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
            Index           =   0
            Left            =   6150
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   156
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Cancel 
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
            Left            =   3720
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   155
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame fam_Tab1_Car 
         Appearance      =   0  '����
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   5400
         Left            =   3855
         TabIndex        =   105
         Top             =   1545
         Width           =   7440
         Begin VB.CheckBox chkActive 
            BackColor       =   &H8000000C&
            Caption         =   "�ҥ�"
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
            Left            =   4680
            TabIndex        =   230
            ToolTipText     =   "�O�_�ϥ�PND��f�l�ܨt��"
            Top             =   300
            Width           =   780
         End
         Begin VB.ComboBox cmb_Tab1_AreaCode 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   127
            Top             =   1500
            Width           =   6375
         End
         Begin VB.ComboBox cmb_Tab1_ZIP 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   126
            Top             =   1110
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Tab1_Company 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   125
            Top             =   2280
            Width           =   6375
         End
         Begin VB.TextBox txt_Tab1_WeightCapacity 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   945
            MaxLength       =   7
            TabIndex        =   124
            Top             =   4260
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_VolumnCapacity 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3360
            MaxLength       =   7
            TabIndex        =   123
            Top             =   4260
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_PalletCapacity 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6060
            MaxLength       =   7
            TabIndex        =   122
            Top             =   4260
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_CarWeight 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   945
            TabIndex        =   121
            Top             =   2670
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_CarHeight 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3375
            TabIndex        =   120
            Top             =   2670
            Width           =   885
         End
         Begin VB.ComboBox cmb_Tab1_CarBox 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   119
            Top             =   3045
            Width           =   2715
         End
         Begin VB.ComboBox cmb_Tab1_EmployType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   118
            Top             =   3435
            Width           =   2715
         End
         Begin VB.ComboBox cmb_Tab1_UnloadType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   117
            Top             =   3840
            Width           =   2715
         End
         Begin VB.ComboBox cmb_Tab1_VehicleType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   116
            Top             =   1890
            Width           =   6375
         End
         Begin VB.TextBox txt_Tab1_Description 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   930
            TabIndex        =   115
            Top             =   4635
            Width           =   6255
         End
         Begin VB.ComboBox cmb_Tab1_CarType 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "frm_BaseData_Car.frx":00C0
            Left            =   5685
            List            =   "frm_BaseData_Car.frx":00D3
            TabIndex        =   114
            Top             =   1125
            Width           =   1635
         End
         Begin VB.TextBox txt_Tab1_CarID 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   945
            TabIndex        =   113
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab1_Driver 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   975
            TabIndex        =   112
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt_Tab1_Phone 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2940
            TabIndex        =   111
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab1_Receiver 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   5700
            TabIndex        =   110
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkPND 
            BackColor       =   &H8000000C&
            Caption         =   "PND��f�l��"
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
            Left            =   3000
            TabIndex        =   109
            ToolTipText     =   "�O�_�ϥ�PND��f�l�ܨt��"
            Top             =   300
            Width           =   1620
         End
         Begin VB.TextBox txtAPFix 
            Alignment       =   1  '�a�k���
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3780
            TabIndex        =   108
            Top             =   1140
            Width           =   735
         End
         Begin VB.TextBox txtAdd 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   915
            TabIndex        =   107
            ToolTipText     =   "�ϥΪ� / �ɶ�"
            Top             =   5040
            Width           =   2805
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4395
            TabIndex        =   106
            ToolTipText     =   "�ϥΪ� / �ɶ�"
            Top             =   5040
            Width           =   2805
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�e�ϽX"
            Height          =   180
            Index           =   23
            Left            =   150
            TabIndex        =   153
            Top             =   1575
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�B���q"
            Height          =   180
            Index           =   24
            Left            =   150
            TabIndex        =   152
            Top             =   2370
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�˸����q"
            Height          =   180
            Index           =   25
            Left            =   150
            TabIndex        =   151
            Top             =   4320
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�˸����n"
            Height          =   180
            Index           =   26
            Left            =   2610
            TabIndex        =   150
            Top             =   4335
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�˸��O��"
            Height          =   180
            Index           =   27
            Left            =   5340
            TabIndex        =   149
            Top             =   4320
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�`��"
            Height          =   180
            Index           =   28
            Left            =   510
            TabIndex        =   148
            Top             =   2730
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���ɰ���"
            Height          =   180
            Index           =   29
            Left            =   2595
            TabIndex        =   147
            Top             =   2730
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���[�Φ�"
            Height          =   180
            Index           =   30
            Left            =   150
            TabIndex        =   146
            Top             =   3135
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���Τ覡"
            Height          =   180
            Index           =   31
            Left            =   150
            TabIndex        =   145
            Top             =   3525
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�˨��覡"
            Height          =   180
            Index           =   32
            Left            =   150
            TabIndex        =   144
            Top             =   3930
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��    ��"
            Height          =   180
            Index           =   33
            Left            =   330
            TabIndex        =   143
            Top             =   1980
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   180
            Index           =   34
            Left            =   495
            TabIndex        =   142
            Top             =   4695
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�l���ϸ�"
            Height          =   180
            Index           =   22
            Left            =   150
            TabIndex        =   141
            Top             =   1185
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p�O���O"
            Height          =   180
            Index           =   44
            Left            =   4920
            TabIndex        =   140
            Top             =   1185
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���P���X"
            Height          =   180
            Index           =   19
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�r�p�H"
            Height          =   180
            Index           =   20
            Left            =   360
            TabIndex        =   138
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q��"
            Height          =   180
            Index           =   21
            Left            =   2520
            TabIndex        =   137
            Top             =   780
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�дڤH"
            Height          =   180
            Index           =   56
            Left            =   5040
            TabIndex        =   136
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "KG"
            Height          =   180
            Index           =   14
            Left            =   1920
            TabIndex        =   135
            Top             =   2730
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "KG"
            Height          =   180
            Index           =   65
            Left            =   1920
            TabIndex        =   134
            Top             =   4320
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            Height          =   180
            Index           =   66
            Left            =   4320
            TabIndex        =   133
            Top             =   4320
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�O"
            Height          =   180
            Index           =   67
            Left            =   7080
            TabIndex        =   132
            Top             =   4320
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "CM"
            Height          =   180
            Index           =   68
            Left            =   4320
            TabIndex        =   131
            Top             =   2730
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�O�վ�"
            Height          =   180
            Index           =   71
            Left            =   3000
            TabIndex        =   130
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�s�W"
            Height          =   180
            Index           =   72
            Left            =   480
            TabIndex        =   129
            Top             =   5100
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   180
            Index           =   73
            Left            =   3960
            TabIndex        =   128
            Top             =   5100
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '����
         BackColor       =   &H00404080&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   3870
         TabIndex        =   98
         Top             =   360
         Width           =   7485
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
            Left            =   6195
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   104
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Delete 
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
            Left            =   4980
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   103
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Save 
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
            Left            =   2535
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   102
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Modify 
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
            Left            =   90
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   101
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_AddNew 
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
            Left            =   1320
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   100
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Cancel 
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
            Left            =   3765
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   99
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmd_Tab0_2Excel 
         Appearance      =   0  '����
         BackColor       =   &H00808080&
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73650
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   97
         Top             =   420
         Width           =   585
      End
      Begin VB.CommandButton cmd_Tab0_ConsigneeQuery 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�Ȥ�j�M"
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
         Left            =   -74775
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   96
         Top             =   420
         Width           =   1110
      End
      Begin VB.CommandButton cmd_Tab1_2Excel 
         Appearance      =   0  '����
         BackColor       =   &H00808080&
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1350
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   95
         Top             =   420
         Width           =   585
      End
      Begin VB.CommandButton cmd_Tab1_CarQuery 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�����j�M"
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
         Left            =   225
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   94
         Top             =   420
         Width           =   1110
      End
      Begin VB.Frame fam_Tab2_FunctionArea 
         Appearance      =   0  '����
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -74910
         TabIndex        =   90
         Top             =   435
         Width           =   3795
         Begin VB.CommandButton cmd_Tab2_CompanyShow 
            BackColor       =   &H00FFC0FF&
            Caption         =   "��ܩҦ����q"
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
            Left            =   75
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   93
            Top             =   180
            Width           =   1830
         End
         Begin VB.CommandButton cmd_Tab2_CarQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "��Ʒj�M"
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
            Left            =   1950
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   92
            Top             =   180
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab2_TRPCompanyReset 
            Appearance      =   0  '����
            BackColor       =   &H00808080&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3105
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   91
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '����
         BackColor       =   &H00004040&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71100
         TabIndex        =   83
         Top             =   435
         Width           =   7440
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
            Left            =   3735
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   89
            Top             =   195
            Width           =   1200
         End
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
            Left            =   1305
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   88
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
            Left            =   75
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   87
            Top             =   195
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
            Left            =   2520
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   86
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
            Left            =   4935
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   85
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
            Index           =   2
            Left            =   6165
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   84
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame fam_Tab2_Company 
         BackColor       =   &H8000000C&
         Height          =   2355
         Left            =   -74730
         TabIndex        =   66
         Top             =   1215
         Width           =   10890
         Begin VB.TextBox txt_Tab2_CompanyCode 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1050
            TabIndex        =   74
            Top             =   285
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab2_CName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   73
            Top             =   780
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab2_EName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   72
            Top             =   1155
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab2_Address 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   71
            Top             =   1530
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab2_ShortName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   70
            Top             =   1530
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab2_Contact 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   69
            Top             =   780
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab2_Phone 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   68
            Top             =   1155
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab2_Descr 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   67
            Top             =   1920
            Width           =   4620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���q�N�X"
            Height          =   180
            Index           =   35
            Left            =   225
            TabIndex        =   82
            Top             =   405
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����W��"
            Height          =   180
            Index           =   36
            Left            =   225
            TabIndex        =   81
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�^��W��"
            Height          =   180
            Index           =   37
            Left            =   225
            TabIndex        =   80
            Top             =   1215
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�a  �}"
            Height          =   180
            Index           =   38
            Left            =   495
            TabIndex        =   79
            Top             =   1590
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "²  ��"
            Height          =   180
            Index           =   39
            Left            =   6240
            TabIndex        =   78
            Top             =   1590
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�p���H"
            Height          =   180
            Index           =   40
            Left            =   6150
            TabIndex        =   77
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q  ��"
            Height          =   180
            Index           =   41
            Left            =   6240
            TabIndex        =   76
            Top             =   1215
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��  ��"
            Height          =   180
            Index           =   42
            Left            =   495
            TabIndex        =   75
            Top             =   1980
            Width           =   450
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '����
         BackColor       =   &H00004040&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71070
         TabIndex        =   59
         Top             =   480
         Width           =   7440
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
            Index           =   3
            Left            =   6165
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   65
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Delete 
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
            Left            =   4935
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   64
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Save 
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
            Left            =   2520
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   63
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Modify 
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
            Left            =   75
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   62
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_AddNew 
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
            Left            =   1305
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   61
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Cancel 
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
            Left            =   3735
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   60
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '����
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -74880
         TabIndex        =   55
         Top             =   480
         Width           =   3795
         Begin VB.CommandButton cmd_Tab2_SkuReset 
            Appearance      =   0  '����
            BackColor       =   &H00808080&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3105
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   58
            Top             =   180
            Width           =   585
         End
         Begin VB.CommandButton cmd_Tab2_SkuQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "��Ʒj�M"
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
            Left            =   1950
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   57
            Top             =   180
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab2_SkuShow 
            BackColor       =   &H00FFC0FF&
            Caption         =   "��ܥ���"
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
            Left            =   75
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   56
            Top             =   180
            Width           =   1830
         End
      End
      Begin VB.Frame fam_Tab3_Sku 
         BackColor       =   &H8000000C&
         Height          =   2355
         Left            =   -74820
         TabIndex        =   36
         Top             =   1260
         Width           =   10890
         Begin VB.TextBox txt_Tab3_Sku 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1050
            TabIndex        =   45
            Top             =   285
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab3_DESCR 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   44
            Top             =   780
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab3_NOTES1 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   43
            Top             =   1155
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab3_NOTES2 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   42
            Top             =   1530
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab3_BUSR1 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   41
            Top             =   1530
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_STDGROSSWGT 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   40
            Top             =   780
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_BUSR4 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   39
            Top             =   1155
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_SKUGROUP 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   38
            Top             =   1920
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_StorerKey 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   37
            Top             =   1920
            Width           =   1965
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f���N�X"
            Height          =   180
            Index           =   53
            Left            =   225
            TabIndex        =   54
            Top             =   405
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����W��"
            Height          =   180
            Index           =   52
            Left            =   225
            TabIndex        =   53
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ƶ��@"
            Height          =   180
            Index           =   51
            Left            =   225
            TabIndex        =   52
            Top             =   1215
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ƶ��G"
            Height          =   180
            Index           =   50
            Left            =   225
            TabIndex        =   51
            Top             =   1590
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���~�O"
            Height          =   180
            Index           =   49
            Left            =   6150
            TabIndex        =   50
            Top             =   1590
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�C�c��"
            Height          =   180
            Index           =   48
            Left            =   6150
            TabIndex        =   49
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�C�c��"
            Height          =   180
            Index           =   47
            Left            =   6150
            TabIndex        =   48
            Top             =   1215
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���O"
            Height          =   180
            Index           =   46
            Left            =   6150
            TabIndex        =   47
            Top             =   1980
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D"
            Height          =   180
            Index           =   54
            Left            =   225
            TabIndex        =   46
            Top             =   1980
            Width           =   360
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  '����
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -75000
         TabIndex        =   32
         Top             =   360
         Width           =   3795
         Begin VB.CommandButton cmd_Tab4_AcceptableShow 
            BackColor       =   &H00FFC0FF&
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
            Height          =   495
            Left            =   75
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   35
            Top             =   180
            Width           =   1830
         End
         Begin VB.CommandButton cmd_Tab4_AcceptableQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "��Ʒj�M"
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
            Left            =   1920
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   34
            Top             =   180
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab4_AcceptableReset 
            Appearance      =   0  '����
            BackColor       =   &H00808080&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3105
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   33
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '����
         BackColor       =   &H00004000&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71280
         TabIndex        =   25
         Top             =   360
         Width           =   7440
         Begin VB.CommandButton cmd_Tab4_Cancel 
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
            Left            =   3720
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   31
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
            Index           =   4
            Left            =   6150
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   30
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_Delete 
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
            Left            =   4935
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   29
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_Save 
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
            Left            =   2505
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   28
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_Modify 
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
            Left            =   75
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   27
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_AddNew 
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
            Left            =   1290
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   26
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame fam_Tab4_Acceptable 
         Appearance      =   0  '����
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2385
         Left            =   -75000
         TabIndex        =   11
         Top             =   1200
         Width           =   11100
         Begin VB.TextBox txt_Tab4_ConsigneeKey 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1080
            TabIndex        =   17
            Top             =   915
            Width           =   1980
         End
         Begin VB.ComboBox cmb_Tab4_Storer 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            ItemData        =   "frm_BaseData_Car.frx":00F9
            Left            =   1080
            List            =   "frm_BaseData_Car.frx":00FB
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   16
            Top             =   330
            Width           =   1935
         End
         Begin VB.TextBox txt_Tab4_FullName 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  '�S���ؽu
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   340
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   925
            Width           =   6000
         End
         Begin VB.TextBox txt_Tab4_Sku 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1080
            TabIndex        =   14
            Top             =   1320
            Width           =   1980
         End
         Begin VB.TextBox txt_Tab4_AllowDays 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   13
            Top             =   1800
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab4_DESCR 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  '�S���ؽu
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   340
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1330
            Width           =   6000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f        �D"
            Height          =   180
            Index           =   79
            Left            =   285
            TabIndex        =   24
            Top             =   390
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�s��"
            Height          =   180
            Index           =   78
            Left            =   285
            TabIndex        =   23
            Top             =   1035
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�Ȥ�W��"
            Height          =   180
            Index           =   77
            Left            =   3525
            TabIndex        =   22
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���~�W��"
            Height          =   180
            Index           =   76
            Left            =   3525
            TabIndex        =   21
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���~�s��"
            Height          =   180
            Index           =   75
            Left            =   285
            TabIndex        =   20
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�����Ѽ�"
            Height          =   180
            Index           =   74
            Left            =   285
            TabIndex        =   19
            Top             =   1940
            Width           =   720
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '�z��
            Caption         =   "�� �� �� �� �� �� �� �� �� �@ �@ �~"
            BeginProperty Font 
               Name            =   "�з���"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   375
            Left            =   4230
            TabIndex        =   18
            Top             =   320
            Width           =   6135
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "�����ѼƶפJ"
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
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   11040
         Begin VB.CommandButton cmdImportT5 
            BackColor       =   &H0080FFFF&
            Caption         =   "�פJ"
            Height          =   375
            Left            =   3720
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT5 
            Height          =   300
            Left            =   135
            TabIndex        =   8
            ToolTipText     =   "Local Drive List"
            Top             =   240
            Width           =   2640
         End
         Begin VB.DirListBox dirLocalDirT5 
            Height          =   1560
            Left            =   135
            TabIndex        =   7
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   5640
         End
         Begin VB.FileListBox filLocalFileT5 
            Height          =   1890
            Left            =   5880
            Pattern         =   "*.xls"
            TabIndex        =   6
            ToolTipText     =   "����� ""*.xls"" �ɮ�"
            Top             =   240
            Width           =   4950
         End
         Begin VB.CommandButton cmdOpenFilesT5 
            BackColor       =   &H0080FFFF&
            Caption         =   "�}��"
            Height          =   375
            Left            =   4800
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid gd_Tab5_AaccessTable 
         Height          =   3945
         Left            =   -74880
         TabIndex        =   3
         Top             =   2760
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   6959
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
      Begin MSDataGridLib.DataGrid dg_Tab4_AcceptableList 
         Height          =   3360
         Left            =   -75000
         TabIndex        =   10
         Top             =   3600
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   5927
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
      Begin MSDataGridLib.DataGrid dg_Tab0_ConsigneeList 
         Height          =   7080
         Left            =   -74820
         TabIndex        =   224
         Top             =   810
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   12488
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
      Begin MSDataGridLib.DataGrid dg_Tab1_CarList 
         Height          =   6120
         Left            =   180
         TabIndex        =   225
         Top             =   810
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   10795
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
      Begin MSDataGridLib.DataGrid dg_Tab2_TRPCompanyList 
         Height          =   3345
         Left            =   -74715
         TabIndex        =   226
         Top             =   3585
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   5900
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
      Begin MSDataGridLib.DataGrid dg_Tab3_SkuList 
         Height          =   3345
         Left            =   -74805
         TabIndex        =   227
         Top             =   3630
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   5900
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�� �� �� �� �� �� �� �@ �@ �~"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   -69915
         TabIndex        =   229
         Top             =   1185
         Width           =   5070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�B �� �� �� �� �� �� �� �� �@ �@ �~"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   1
         Left            =   4605
         TabIndex        =   228
         Top             =   1185
         Width           =   6120
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00400040&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   435
         Index           =   0
         Left            =   180
         Top             =   405
         Width           =   1800
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00400040&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   465
         Index           =   1
         Left            =   -74805
         Top             =   390
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frm_BaseData_Car"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private blTab0ConsignEventEnable As Boolean     '�Ȥ��� List ���ƥ� Enable ����
Private blTab1CarEventEnable As Boolean         '������� List ���ƥ� Enable ����
Private blTab2CompanyEventEnable As Boolean     '�f�B���q List ���ƥ� Enable ����
Private blTab3skuEventEnable As Boolean         '�f�B���q List ���ƥ� Enable ����
Private blTab4AcceptableEventEnable As Boolean  '�Ȥ᤹���Ѽ� List ���ƥ� Enable ����

Private arStorer() As String            '�f�D
Private arZip() As String               '�l���ϸ�
Private arZIPArea() As String           '�l���ϸ��ɳ]�w�� AreaCode
Private arAreaCode() As String          '�ϰ�N�X
Private arVehicleType() As String       '��������
Private arExtraDemand() As String       '�S��ݨD
Private arPickTool() As String          '�h�B�u��
Private arCompany() As String           '�����G�f�B���q
Private arCarBox() As String            '�����G���[�Φ�
Private arEmployType() As String        '�����G���Τ覡
Private arUnloadType() As String        '�����G�˨��覡

Private rs_Tab0_ConsigneeList As ADODB.Recordset       '��ܩҦ��Ȥ���
Private rs_Tab1_CarList As ADODB.Recordset             '��ܨ����򥻸��
Private rs_Tab2_TRPCompanyList As ADODB.Recordset      '��ܹB�餽�q�򥻸��
Private rs_Tab3_SkuList As ADODB.Recordset             '��ܳf���򥻸��
Private rs_Tab4_AcceptableList As ADODB.Recordset      '��ܩҦ��Ȥ᤹���ѼƸ��

Private MyXlsAppV2 As Excel.Application     '�����ѼƸ����Excel
Private rsMain As ADODB.Recordset
Private fso As Scripting.FileSystemObject

Private Sub cmb_Tab0_Zip_Click()
'�Ȥ��� >> �l���ϸ�
If fam_Tab0_Consignee.Enabled = True Then
    If cmb_Tab0_Zip.ListIndex <> -1 Then
        Dim i As Double
        For i = 0 To cmb_Tab0_AreaCode.ListCount - 1
            If arAreaCode(i) = UCase(arZIPArea(cmb_Tab0_Zip.ListIndex)) Then
                cmb_Tab0_AreaCode.ListIndex = i
                Exit Sub
            End If
        Next i
    End If
End If
End Sub

Private Sub cmb_Tab1_VehicleType_Click()
cmb_Tab1_CarType = mySplit(cmb_Tab1_VehicleType, "/", -1)
End Sub

Private Sub cmb_Tab1_ZIP_Click()
'������� >> �l���ϸ�
If fam_Tab1_Car.Enabled = True Then
    If cmb_Tab1_ZIP.ListIndex <> -1 Then
        Dim i As Double
        For i = 0 To cmb_Tab1_AreaCode.ListCount - 1
            If arAreaCode(i) = arZIPArea(cmb_Tab1_ZIP.ListIndex) Then
                cmb_Tab1_AreaCode.ListIndex = i
                Exit Sub
            End If
        Next i
    End If
End If
End Sub

Private Sub cmd_Tab0_2Excel_Click()

Dim rsTmp As New ADODB.Recordset
Screen.MousePointer = 11
rsTmp.Open "select * from gv_customer ", cn
Recordset2Excel "�Ȥ�D��", rsTmp
Set MyXlsApp = Nothing
rsTmp.Close: Set rsTmp = Nothing
Screen.MousePointer = 0

End Sub
Private Sub cmd_Tab1_2Excel_Click()

Dim rsTmp As New ADODB.Recordset
Screen.MousePointer = 11

Recordset2Excel "�����D��", rs_Tab1_CarList
Set MyXlsApp = Nothing

Screen.MousePointer = 0

End Sub

Private Sub cmd_Tab0_AddNew_Click()
'�Ȥ��� >> �ഫ�ܷs�W�Ҧ�
If Not rs_Tab0_ConsigneeList Is Nothing Then
    If dg_Tab0_ConsigneeList.SelBookmarks.Count > 0 Then dg_Tab0_ConsigneeList.SelBookmarks.Remove 0
End If
fam_Tab0_Consignee.BackColor = &HC0FFC0
fam_Tab0_Consignee.Enabled = True
txt_Tab0_ConsigneeKey.Enabled = True
cmb_Tab0_Storer.Enabled = True
Call Clear_ConsigneeData
cmd_Tab0_Save.Enabled = True
cmd_Tab0_Cancel.Enabled = True
cmd_Tab0_AddNew.Enabled = False
cmd_Tab0_Modify.Enabled = False
cmd_Tab0_Delete.Enabled = False
End Sub

Private Sub cmd_Tab0_Cancel_Click()
'�Ȥ��� >> �����ק�
Call Clear_ConsigneeData
If txt_Tab0_ConsigneeKey.Enabled = False Then
    If Not rs_Tab0_ConsigneeList Is Nothing Then
        dg_Tab0_ConsigneeList.SelBookmarks.Add rs_Tab0_ConsigneeList.Bookmark
        Call Display_SelectedConsignData(rs_Tab0_ConsigneeList.Fields("�f�D").Value, rs_Tab0_ConsigneeList.Fields("�Ȥ�s��").Value)
    End If
End If
fam_Tab0_Consignee.BackColor = &H8000000C
fam_Tab0_Consignee.Enabled = False
cmd_Tab0_Cancel.Enabled = False
cmd_Tab0_Save.Enabled = False
cmd_Tab0_AddNew.Enabled = True
cmd_Tab0_Modify.Enabled = True
cmd_Tab0_Delete.Enabled = True
End Sub

Private Sub cmd_Tab0_ConsigneeQuery_Click()
'�Ȥ��� >> �Ȥ�j�M
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
If rs_Tab0_ConsigneeList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab0_ConsigneeList"

If ShowForm_RS_FilterAndSort(rs_Tab0_ConsigneeList, "�Ȥ���", Me.Tag) = False Then
    MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
Me.WindowState = vbNormal

End Sub

Private Sub cmd_Tab0_ConsigneeReset_Click()
'�Ȥ�򥻸�� >> �����z��Ƨ�
'�����z�����A���]�ƧǨ̾�
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
 blTab0ConsignEventEnable = False
 rs_Tab0_ConsigneeList.Filter = adFilterNone
 rs_Tab0_ConsigneeList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
 blTab0ConsignEventEnable = True

End Sub

Private Sub cmd_Tab0_ConsigneeShow_Click()
'�Ȥ��� >> ��ܩҦ��Ȥ�
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0_ConsigneeList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0_ConsigneeList)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT Rtrim(t1.StorerKey) as �f�D , Rtrim(t1.ConsigneeKey) as �Ȥ�s�� , Rtrim(Isnull(t1.Full_Name,'')) as �Ȥ�W��  " & _
          "From TRP01M t1 join trp16m t16 on t1.storerkey = t16.storerkey and t16.storer_status <> '0' Order by t1.StorerKey,ConsigneeKey"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C�Ȥ���"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0_ConsigneeList)
tmp_Rs.Close

blTab0ConsignEventEnable = False
With dg_Tab0_ConsigneeList
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab0_ConsigneeList.MoveFirst
Set dg_Tab0_ConsigneeList.DataSource = rs_Tab0_ConsigneeList
With dg_Tab0_ConsigneeList
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 600        '�f�D
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '�Ȥ�s��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '�Ȥ�W��
    .Columns(3).Alignment = dbgLeft
End With
blTab0ConsignEventEnable = True
Call Clear_ConsigneeData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�Ȥ���-��ܩҦ����", Me.Caption, "cmd_Tab0-ConsignShow_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Delete_Click()
'�Ȥ��� >> �R��
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

'1.�ˮ� TRP02W �O�_�����Ȥ�q����
str_SQL = "Select Count(*) as RecCnt From TRP02W Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
    blDelete = False
    If msg_text = "" Then
        msg_text = "   �ݱƨ��q�� [TRP02W] �����Ȥ�q����"
    Else
        msg_text = msg_text & vbCrLf & "   �ݱƨ��q�� [TRP02W] �����Ȥ�q����"
    End If
End If
tmp_Rs.Close
'2.�ˮ� TRP02T �O�_�����Ȥ�q����
str_SQL = "Select Count(*) as RecCnt From TRP02T Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
    blDelete = False
    If msg_text = "" Then
        msg_text = "   �w�ƨ��q�� [TRP02T] �����Ȥ�q����"
    Else
        msg_text = msg_text & vbCrLf & "   �w�ƨ��q�� [TRP02T] �����Ȥ�q����"
    End If
End If
tmp_Rs.Close
'3.�ˮ� Orders �O�_�����Ȥ�q����
str_SQL = "Select Count(*) as RecCnt From Orders Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   �q��D�� [Orders] �����Ȥ�q����"
   Else
      msg_text = msg_text & vbCrLf & "   �q��D�� [Orders] �����Ȥ�q����"
   End If
End If
tmp_Rs.Close

'�ˮ֬O�_���\�i��R���X�Э�
If blDelete = False Then
   msg_text = "�Ȥ��ƵL�k�R���G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'���\�R��
str_SQL = "Delete From TRP01M Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

fam_Tab0_Consignee.BackColor = &H8000000C
fam_Tab0_Consignee.Enabled = False
cmd_Tab0_Cancel.Enabled = False
cmd_Tab0_Save.Enabled = False
cmd_Tab0_AddNew.Enabled = True
cmd_Tab0_Modify.Enabled = False
cmd_Tab0_Delete.Enabled = False
'���s��ܩҦ��Ȥ���
Call cmd_Tab0_ConsigneeShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�Ȥ���-�R��", Me.Caption, "cmd_Tab0_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Modify_Click()
'�Ȥ��� >> �ഫ�ק�Ҧ�
'�T�{����Ȥ��Ƥ褹�\ [�ק�] �\��
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
If dg_Tab0_ConsigneeList.SelBookmarks.Count <> 0 Then
   fam_Tab0_Consignee.BackColor = &HC0E0FF
   fam_Tab0_Consignee.Enabled = True
   txt_Tab0_ConsigneeKey.Enabled = False
   cmd_Tab0_Save.Enabled = True
   cmd_Tab0_Cancel.Enabled = True
   cmd_Tab0_AddNew.Enabled = False
   cmd_Tab0_Modify.Enabled = False
   cmd_Tab0_Delete.Enabled = False
End If
End Sub

Private Sub cmd_Tab0_Save_Click()
'�Ȥ��� >> �Ȥ��Ʀs��

If Len(RTrim((cmb_Tab0_Storer.Text))) = 0 Or Len(RTrim(txt_Tab0_ConsigneeKey)) = 0 Then MsgBox "�п�J�f�D�P�Ȥ�s��", 16, "�`�N": Exit Sub

'�M���S��r��
Call myFormExCharFilter(Me)

On Error GoTo err_Handle

'�Ӽh�ɶK���ˬd�A�@�w�n�Ʀr�C����J�H0�p��
If Len(RTrim(txt_Tab0_Class.Text)) = 0 Then
            MsgBox "�Ӽh�ɶK�S��J�A�t�ιw�]�a0", 64, "�`�N"
            txt_Tab0_Class.Text = 0
End If
If Not IsNumeric(txt_Tab0_Class.Text) Then
            MsgBox "�Ӽh�ɶK���п�J�Ʀr", 64, "�`�N"
            Exit Sub
End If

'�P�_���i���t��
If Left(RTrim(txt_Tab0_Class.Text), 1) = "-" Then MsgBox "�Ӽh�ɶK�Фſ�J�t��", 64, "�`�N": Exit Sub
'���p���I2��L����i��A*1���F��0
If Val(txt_Tab0_Class.Text) < 1 Then
    txt_Tab0_Class.Text = ("0" & Format(txt_Tab0_Class.Text, ".##")) * 1
Else
    txt_Tab0_Class.Text = Format(txt_Tab0_Class.Text, ".##") * 1
End If

'�Ȥ�s�������ˬd
If txt_Tab0_ConsigneeKey.Enabled = True Then
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Open "select consigneekey from trp01m where rtrim(consigneekey) = '" & RTrim(txt_Tab0_ConsigneeKey) & "' and rtrim(storerkey) = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = False Then
        MsgBox "�P�@�f�D�A�s�W�Ȥ�s������!!", 64, "�`�N"
           txt_Tab0_ConsigneeKey.SelStart = 0: txt_Tab0_ConsigneeKey.SelLength = Len(txt_Tab0_ConsigneeKey.Text)
       txt_Tab0_ConsigneeKey.SetFocus
       Exit Sub
    End If
End If

'�s�ɸ���ˮ�
If Check_ComsigneeData = False Then Exit Sub

'LTKK01�a�}�O���ƧP�w
If Left(cmb_Tab0_Storer, 6) = "LTKK01" Then
    str_SQL = "select * from trp01m " & _
                "where consigneekey <> '" & RTrim(txt_Tab0_ConsigneeKey) & "' " & _
                "and substring(consigneekey , 5,20) = '" & RTrim(Mid(txt_Tab0_ConsigneeKey, 5, 20)) & "' and storerkey = 'LTKK01'"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn
    If Not tmp_Rs.EOF Then MsgBox "�f�D LTKK01 �a�}�O�s�X�ӵu�νs�����ơA�w�s�b�Ȥ�s��(" & RTrim(tmp_Rs("consigneekey")) & ") �A�Ȥ�W��(" & RTrim(tmp_Rs("short_name")) & ")�C", vbOKOnly, "�Ȥ�D�ɷs�W": Exit Sub
End If

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_ConsigneeData_UPDATE"

'�f�D
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("StorerKey").Value = arStorer(cmb_Tab0_Storer.ListIndex)

'�Ȥ�s��
Set tmp_para = tmp_Cmd.CreateParameter("ConsigneeKey", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_Tab0_ConsigneeKey.Text)

'�l���ϸ�
Set tmp_para = tmp_Cmd.CreateParameter("ZIP", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_Zip.ListIndex <> -1 Then
   tmp_Cmd.Parameters("ZIP").Value = arZip(cmb_Tab0_Zip.ListIndex)
Else
   tmp_Cmd.Parameters("ZIP").Value = ""
End If


'�B�e�ϽX
Set tmp_para = tmp_Cmd.CreateParameter("Area_Code", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_AreaCode.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Area_Code").Value = arAreaCode(cmb_Tab0_AreaCode.ListIndex)
Else
   tmp_Cmd.Parameters("Area_Code").Value = Null
End If

'�B�e�a�}
Set tmp_para = tmp_Cmd.CreateParameter("Address", adVarChar, adParamInput, 200)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Address.Text) = "" Then
   tmp_Cmd.Parameters("Address").Value = ""
Else
   tmp_Cmd.Parameters("Address").Value = Trim(txt_Tab0_Address.Text)
End If

'�p���H 'Terry 20180123 contact ���ץ�30�אּ80
Set tmp_para = tmp_Cmd.CreateParameter("Contact", adVarChar, adParamInput, 80)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Contact.Text) = "" Then
   tmp_Cmd.Parameters("Contact").Value = ""
Else
   tmp_Cmd.Parameters("Contact").Value = Trim(txt_Tab0_Contact.Text)
End If

'�q��
Set tmp_para = tmp_Cmd.CreateParameter("Phone", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Phone.Text) = "" Then
   tmp_Cmd.Parameters("Phone").Value = ""
Else
   tmp_Cmd.Parameters("Phone").Value = Trim(txt_Tab0_Phone.Text)
End If

'�Ȥᵥ��
Set tmp_para = tmp_Cmd.CreateParameter("Class", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Class.Text) = "" Then
   tmp_Cmd.Parameters("Class").Value = Null
Else
   tmp_Cmd.Parameters("Class").Value = Trim(txt_Tab0_Class.Text)
End If

'�S��ݨD 1
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_ExtraDemand1.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = arExtraDemand(cmb_Tab0_ExtraDemand1.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = Null
End If

'�S��ݨD 2
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code2", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_ExtraDemand2.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = arExtraDemand(cmb_Tab0_ExtraDemand2.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = Null
End If

'�Ȥ�W��
Set tmp_para = tmp_Cmd.CreateParameter("Full_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_FullName.Text) = "" Then
   tmp_Cmd.Parameters("Full_Name").Value = ""
Else
   tmp_Cmd.Parameters("Full_Name").Value = Trim(txt_Tab0_FullName.Text)
End If

'�Ȥ�²��
Set tmp_para = tmp_Cmd.CreateParameter("Short_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_ShortName.Text) = "" Then
   tmp_Cmd.Parameters("Short_Name").Value = ""
Else
   tmp_Cmd.Parameters("Short_Name").Value = Trim(txt_Tab0_ShortName.Text)
End If

'�q�����A
Set tmp_para = tmp_Cmd.CreateParameter("Channel_Type", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_ChannelType.Text) > 0 Then
   tmp_Cmd.Parameters("Channel_Type").Value = Trim(txt_Tab0_ChannelType.Text)
Else
   tmp_Cmd.Parameters("Channel_Type").Value = Null
End If

'���d������
Set tmp_para = tmp_Cmd.CreateParameter("Unload_Type", adVarChar, adParamInput, 3)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_UnLoad.Text) = "" Then
   tmp_Cmd.Parameters("Unload_Type").Value = Null
Else
   tmp_Cmd.Parameters("Unload_Type").Value = Trim(txt_Tab0_UnLoad.Text)
End If

'Billing_Type
Set tmp_para = tmp_Cmd.CreateParameter("BILLING_TYPE", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("BILLING_TYPE").Value = Null

'Payment_Type
Set tmp_para = tmp_Cmd.CreateParameter("Payment_Type", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Payment_Type").Value = Null

'Special_Charge
Set tmp_para = tmp_Cmd.CreateParameter("Special_Charge", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Special_Charge").Value = Null

'���e�Ȥ�
Set tmp_para = tmp_Cmd.CreateParameter("Multi_Customer", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
If chk_Tab0_MultiCustomer.Value = vbChecked Then
   tmp_Cmd.Parameters("Multi_Customer").Value = "Y"
Else
   tmp_Cmd.Parameters("Multi_Customer").Value = "N"
End If

'�έܫȤ�
Set tmp_para = tmp_Cmd.CreateParameter("dc", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
If chkDC.Value = vbChecked Then
   tmp_Cmd.Parameters("dc").Value = "Y"
Else
   tmp_Cmd.Parameters("dc").Value = "N"
End If

'Grid_Code �x�}�ϽX
Set tmp_para = tmp_Cmd.CreateParameter("Grid_Code", adVarChar, adParamInput, 5)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_GridCode.Text) = "" Then
   tmp_Cmd.Parameters("Grid_Code").Value = Null
Else
   tmp_Cmd.Parameters("Grid_Code").Value = Trim(txt_Tab0_GridCode.Text)
End If

'���إN�X
Set tmp_para = tmp_Cmd.CreateParameter("Vehicle_Type", adVarChar, adParamInput, 2)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_VehicleType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Vehicle_Type").Value = arVehicleType(cmb_Tab0_VehicleType.ListIndex)
Else
   tmp_Cmd.Parameters("Vehicle_Type").Value = Null
End If

'�h�B�u��
Set tmp_para = tmp_Cmd.CreateParameter("PICK_TOOL", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_PickTool.ListIndex <> -1 Then
   tmp_Cmd.Parameters("PICK_TOOL").Value = arPickTool(cmb_Tab0_PickTool.ListIndex)
Else
   tmp_Cmd.Parameters("PICK_TOOL").Value = Null
End If

'�q���O
Set tmp_para = tmp_Cmd.CreateParameter("Channel", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Channel.Text) > 0 Then
   tmp_Cmd.Parameters("Channel").Value = Trim(txt_Tab0_Channel.Text)
Else
   tmp_Cmd.Parameters("Channel").Value = Null
End If

'�ǯu
Set tmp_para = tmp_Cmd.CreateParameter("fax", adChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("fax").Value = Trim(txt_Tab0_Fax.Text)

'Codedate1
Set tmp_para = tmp_Cmd.CreateParameter("codedate1", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("codedate1").Value = Trim(txt_Tab0_CodeDate1.Text)

'Codedate2
Set tmp_para = tmp_Cmd.CreateParameter("codedate2", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("codedate2").Value = Trim(txt_Tab0_CodeDate2.Text)

'�K��
Set tmp_para = tmp_Cmd.CreateParameter("stamp", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("stamp").Value = Trim(txt_Tab0_Stamp.Text)

'�@�ګȤ�
Set tmp_para = tmp_Cmd.CreateParameter("Penalties", adChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Penalties").Value = Trim(txt_Tab0_Penalties.Text)

'�̪O����
Set tmp_para = tmp_Cmd.CreateParameter("PalletType", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("PalletType").Value = Trim(txt_Tab0_PalletType.Text)

'�̪O�W��
Set tmp_para = tmp_Cmd.CreateParameter("Palletspec", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Palletspec").Value = Trim(txt_Tab0_PalletSpec.Text)

'�Ƶ�
Set tmp_para = tmp_Cmd.CreateParameter("Notes", adChar, adParamInput, 255)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Notes").Value = Trim(txt_Tab0_Notes.Text)

'�Ȥ�s��
Set tmp_para = tmp_Cmd.CreateParameter("CustGroup", adChar, adParamInput, 255)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("CustGroup").Value = Trim(cmb_Tab0_Group.Text)

'������
Set tmp_para = tmp_Cmd.CreateParameter("CodeDateRate", adChar, adParamInput, 255)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("CodeDateRate").Value = Trim(cmdCodeDateRate)


'Codedate3
Set tmp_para = tmp_Cmd.CreateParameter("codedate3", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("codedate3").Value = Trim(txt_Tab0_CodeDate3.Text)


Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'�D�P�B����
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
Loop
Set tmp_Cmd = Nothing

fam_Tab0_Consignee.BackColor = &H8000000C
fam_Tab0_Consignee.Enabled = False
cmd_Tab0_Cancel.Enabled = False
cmd_Tab0_Save.Enabled = False
cmd_Tab0_AddNew.Enabled = True
cmd_Tab0_Modify.Enabled = True
cmd_Tab0_Delete.Enabled = False

'����EditWho
str_SQL = "update trp01m set editwho = '" & User_id & "' , editdate = getdate() where storerkey = '" & mySplit(cmb_Tab0_Storer, " ", 0) & "' and consigneekey = '" & txt_Tab0_ConsigneeKey & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'����ADDWho
str_SQL = "update trp01m set addwho = '" & User_id & "' where addwho is null and storerkey = '" & mySplit(cmb_Tab0_Storer, " ", 0) & "' and consigneekey = '" & txt_Tab0_ConsigneeKey & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'LTKK01�Ȥ�D�ɲ��ʦ۰� Mail �q��
If mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then Call SendMail(txt_Tab0_ConsigneeKey)

If rs_Tab0_ConsigneeList Is Nothing = False Then rs_Tab0_ConsigneeList("�Ȥ�W��") = Trim(txt_Tab0_FullName.Text)

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�Ȥ���-�s��", Me.Caption, "cmd_Tab0_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_AddNew_Click()
'������� >> �s�W�Ҧ��ഫ
If Not rs_Tab1_CarList Is Nothing Then
   If dg_Tab1_CarList.SelBookmarks.Count > 0 Then dg_Tab1_CarList.SelBookmarks.Remove 0
End If
fam_Tab1_Car.BackColor = &HC0FFC0
fam_Tab1_Car.Enabled = True
txt_Tab1_CarID.Enabled = True
Call Clear_CarData
cmd_Tab1_Save.Enabled = True
cmd_Tab1_Cancel.Enabled = True
cmd_Tab1_AddNew.Enabled = False
cmd_Tab1_Modify.Enabled = False
cmd_Tab1_Delete.Enabled = False
End Sub

Private Sub cmd_Tab1_Cancel_Click()
'������� >> ����
Call Clear_CarData
If txt_Tab1_CarID.Enabled = False Then
   If Not rs_Tab1_CarList Is Nothing Then
      dg_Tab1_CarList.SelBookmarks.Add rs_Tab1_CarList.Bookmark
      Call Display_SelectedCarData(rs_Tab1_CarList.Fields("���P���X").Value)
   End If
End If
fam_Tab1_Car.BackColor = &H8000000C
chkPND.BackColor = &H8000000C
fam_Tab1_Car.Enabled = False
cmd_Tab1_Cancel.Enabled = False
cmd_Tab1_Save.Enabled = False
cmd_Tab1_AddNew.Enabled = True
cmd_Tab1_Modify.Enabled = True
cmd_Tab1_Delete.Enabled = True
End Sub

Private Sub cmd_Tab1_CarQuery_Click()
'������� >> �����j�M
If rs_Tab1_CarList Is Nothing Then Exit Sub
If rs_Tab1_CarList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab1_CarList"

If ShowForm_RS_FilterAndSort(rs_Tab1_CarList, "�������", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = vbNormal

End Sub

Private Sub cmd_Tab1_CarReset_Click()
'�����򥻸�� >> �����z��Ƨ�
'�����z�����A���]�ƧǨ̾�
If rs_Tab1_CarList Is Nothing Then Exit Sub
 blTab1CarEventEnable = False
 rs_Tab1_CarList.Filter = adFilterNone
 rs_Tab1_CarList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
 blTab1CarEventEnable = True
End Sub

Private Sub cmd_Tab1_CarShow_Click()

'������� >> ��ܩҦ�����
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab1_CarList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab1_CarList)
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "Select Rtrim(a1.Vehicle_ID_No) as ���P���X , Rtrim(Isnull(a1.Driver,'')) as �r�p�H , Rtrim(Isnull(b1.Description,'')) as ���� , Rtrim(Isnull(a1.receiver,'')) as �дڤH   " & _
          "From TRP09M a1 Left outer join TRP15M b1 on b1.Vehicle_Type = a1.Vehicle_Type Order by A1.Vehicle_ID_No"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C�������"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab1_CarList)
tmp_Rs.Close

blTab1CarEventEnable = False
With dg_Tab1_CarList
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With

rs_Tab1_CarList.MoveFirst
Set dg_Tab1_CarList.DataSource = rs_Tab1_CarList
With dg_Tab1_CarList
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '���P���X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '�q��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '����
    .Columns(3).Alignment = dbgLeft
End With

blTab1CarEventEnable = True
Call Clear_CarData
Screen.MousePointer = vbDefault
Call Display_SelectedCarData(rs_Tab1_CarList.Fields("���P���X").Value)
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�������-��ܩҦ����", Me.Caption, "cmd_Tab1-CarShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_AddNew_Click()
'�f�B���q��� >> �s�W
If Not rs_Tab2_TRPCompanyList Is Nothing Then
   If dg_Tab2_TRPCompanyList.SelBookmarks.Count > 0 Then dg_Tab2_TRPCompanyList.SelBookmarks.Remove 0
End If
fam_Tab2_Company.BackColor = &HC0FFC0
fam_Tab2_Company.Enabled = True
txt_Tab2_CompanyCode.Enabled = True
Call Clear_CompanyData
cmd_Tab2_Save.Enabled = True
cmd_Tab2_Cancel.Enabled = True
cmd_Tab2_AddNew.Enabled = False
cmd_Tab2_Modify.Enabled = False
cmd_Tab2_Delete.Enabled = False
End Sub

Private Sub cmd_Tab2_Cancel_Click()
'�f�B���q��� >> ����
Call Clear_CompanyData
If txt_Tab2_CompanyCode.Enabled = False Then
   If Not rs_Tab2_TRPCompanyList Is Nothing Then
      dg_Tab2_TRPCompanyList.SelBookmarks.Add rs_Tab2_TRPCompanyList.Bookmark
      Call Display_SelectedCompanyData(rs_Tab2_TRPCompanyList.Fields("���q�N�X").Value)
   End If
End If
fam_Tab2_Company.BackColor = &H8000000C
fam_Tab2_Company.Enabled = False
cmd_Tab2_Cancel.Enabled = False
cmd_Tab2_Save.Enabled = False
cmd_Tab2_AddNew.Enabled = True
cmd_Tab2_Modify.Enabled = True
cmd_Tab2_Delete.Enabled = True
End Sub

Private Sub cmd_Tab2_CarQuery_Click()
'�f�B���q��� >> �����j�M
If rs_Tab2_TRPCompanyList Is Nothing Then Exit Sub
If rs_Tab2_TRPCompanyList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab2_TRPCompanyList"

If ShowForm_RS_FilterAndSort(rs_Tab2_TRPCompanyList, "�f�B���q���", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = vbNormal
End Sub

Private Sub cmd_Tab2_CompanyShow_Click()
'�f�B���q��� >> ��ܩҦ��f�B���q
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab2_TRPCompanyList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2_TRPCompanyList)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select Rtrim(Company_Code) as ���q�N�X,Rtrim(Isnull(C_Name,'')) as ����W��,Rtrim(Isnull(E_Name,'')) as �^��W��,Rtrim(Isnull(Short_Name,'')) as ²��," & _
          "   Rtrim(isnull(Phone,'')) as �q�� , Rtrim(Isnull(Contact,'')) as �p���H ,Rtrim(Isnull(Address,'')) as �a�} , Rtrim(Isnull(Description,'')) as Descr  " & _
          "From TRP08M Order by Company_Code"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C�f�B���q���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab2_TRPCompanyList)
tmp_Rs.Close

blTab2CompanyEventEnable = False
With dg_Tab2_TRPCompanyList
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab2_TRPCompanyList.MoveFirst
Set dg_Tab2_TRPCompanyList.DataSource = rs_Tab2_TRPCompanyList
With dg_Tab2_TRPCompanyList
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '�f�B���q�N�X
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2500       '����W��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '�^��W��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1500       '²��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1100       '�q��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '�p���H
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 2500       '�a�}
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 2000       '����
    .Columns(8).Alignment = dbgLeft
End With
blTab2CompanyEventEnable = True
Call Clear_CompanyData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�f�B���q���-��ܩҦ����", Me.Caption, "cmd_Tab2-CompanyShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Delete_Click()
'������� >> �R��
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

'1.�ˮ� TRP05T �O�_���������˸��X�����
str_SQL = "Select Count(*) as RecCnt From TRP05T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   ���u�s���B�e��� [TRP05T] ���������˸��X�����"
   Else
      msg_text = msg_text & vbCrLf & "   ���u�s���B�e��� [TRP05T] ���������˸��X�����"
   End If
End If
tmp_Rs.Close

'2.�ˮ� TRP02T �O�_���������˸��X�����
str_SQL = "Select Count(*) as RecCnt From TRP02T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   �w�ƨ��q�� [TRP02T] ���������˸��X�����"
   Else
      msg_text = msg_text & vbCrLf & "   �w�ƨ��q�� [TRP02T] ���������˸��X�����"
   End If
End If
tmp_Rs.Close

'3.�ˮ� SDN02T �O�_���������˸��X�����
str_SQL = "Select Count(*) as RecCnt From SDN02T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   �w�X���q�� [SDN02T] ���������˸��X�����"
   Else
      msg_text = msg_text & vbCrLf & "   �w�ƨ��q�� [SDN02T] ���������˸��X�����"
   End If
End If
tmp_Rs.Close

'4.�ˮ� SDN01T �O�_���������˸��X�����
str_SQL = "Select Count(*) as RecCnt From SDN01T Where C_Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   �w�X���q�� [SDN01T] ���������˸��X�����"
   Else
      msg_text = msg_text & vbCrLf & "   �w�ƨ��q�� [SDN01T] ���������˸��X�����"
   End If
End If
tmp_Rs.Close

'5.�ˮ� ORT02T �O�_���������˸��X�����
str_SQL = "Select Count(*) as RecCnt From ORT02T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   �w�X���q�� [ORT02T] ���������˸��X�����"
   Else
      msg_text = msg_text & vbCrLf & "   �w�ƨ��q�� [ORT02T] ���������˸��X�����"
   End If
End If
tmp_Rs.Close

'�ˮ֬O�_���\�i��R���X�Э�
If blDelete = False Then
   msg_text = "������ƵL�k�R���G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'���\�R��
str_SQL = "Delete From TRP09M Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

fam_Tab1_Car.BackColor = &H8000000C
fam_Tab1_Car.Enabled = False
cmd_Tab1_Cancel.Enabled = False
cmd_Tab1_Save.Enabled = False
cmd_Tab1_AddNew.Enabled = True
cmd_Tab1_Modify.Enabled = False
cmd_Tab1_Delete.Enabled = False
'���s��ܩҦ��Ȥ���
Call cmd_Tab1_CarShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�������-�R��", Me.Caption, "cmd_Tab1_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Modify_Click()
'������� >> �ק�
'�T�{���������Ƥ褹�\ [�ק�] �\��
If rs_Tab1_CarList Is Nothing Then Exit Sub
'If txt_Tab1_Driver = "" Then Exit Sub mark by Gemini @20151019
If dg_Tab1_CarList.SelBookmarks.Count <> 0 Then
   fam_Tab1_Car.BackColor = &HC0E0FF
   chkPND.BackColor = &HC0E0FF
   fam_Tab1_Car.Enabled = True
   txt_Tab1_CarID.Enabled = False
   cmd_Tab1_Save.Enabled = True
   cmd_Tab1_Cancel.Enabled = True
   cmd_Tab1_AddNew.Enabled = False
   cmd_Tab1_Modify.Enabled = False
   cmd_Tab1_Delete.Enabled = False
End If

End Sub

Private Sub cmd_Tab1_Save_Click()
'������� >> �s��

'�M���S��r��
Call myFormExCharFilter(Me)

On Error GoTo err_Handle

'�s�ɸ���ˮ�
If Check_CarData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_CarDara_Update"

'���P���X
Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = Trim(txt_Tab1_CarID.Text)

'�l���ϸ�
Set tmp_para = tmp_Cmd.CreateParameter("ZIP", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_ZIP.ListIndex <> -1 Then
   tmp_Cmd.Parameters("ZIP").Value = arZip(cmb_Tab1_ZIP.ListIndex)
Else
   tmp_Cmd.Parameters("ZIP").Value = Null
End If

'�B�e�ϽX
Set tmp_para = tmp_Cmd.CreateParameter("Area_Code", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_AreaCode.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Area_Code").Value = arAreaCode(cmb_Tab1_AreaCode.ListIndex)
Else
   tmp_Cmd.Parameters("Area_Code").Value = Null
End If

'�f�B���q
Set tmp_para = tmp_Cmd.CreateParameter("TRP_COMPANY_CODE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_Company.ListIndex <> -1 Then
   tmp_Cmd.Parameters("TRP_COMPANY_CODE").Value = arCompany(cmb_Tab1_Company.ListIndex)
Else
   tmp_Cmd.Parameters("TRP_COMPANY_CODE").Value = Null
End If

'����
Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_TYPE", adVarChar, adParamInput, 2)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_VehicleType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("VEHICLE_TYPE").Value = arVehicleType(cmb_Tab1_VehicleType.ListIndex)
Else
   tmp_Cmd.Parameters("VEHICLE_TYPE").Value = Null
End If

'�i�Ӹ����q
Set tmp_para = tmp_Cmd.CreateParameter("LOADING_SIZE", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_WeightCapacity.Text) = "" Then
   tmp_Cmd.Parameters("LOADING_SIZE").Value = Null
Else
   tmp_Cmd.Parameters("LOADING_SIZE").Value = Trim(txt_Tab1_WeightCapacity.Text)
End If

'�i�Ӹ����n
Set tmp_para = tmp_Cmd.CreateParameter("MAX_CUBIC_CAPACITY", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_VolumnCapacity.Text) = "" Then
   tmp_Cmd.Parameters("MAX_CUBIC_CAPACITY").Value = Null
Else
   tmp_Cmd.Parameters("MAX_CUBIC_CAPACITY").Value = Trim(txt_Tab1_VolumnCapacity.Text)
End If

'�q��
Set tmp_para = tmp_Cmd.CreateParameter("DRIVER", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_Driver.Text) = "" Then
   tmp_Cmd.Parameters("DRIVER").Value = Null
Else
   tmp_Cmd.Parameters("DRIVER").Value = Trim(txt_Tab1_Driver.Text)
End If

'�q��
Set tmp_para = tmp_Cmd.CreateParameter("DRIVER_PHONE", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_Phone.Text) = "" Then
   tmp_Cmd.Parameters("DRIVER_PHONE").Value = Null
Else
   tmp_Cmd.Parameters("DRIVER_PHONE").Value = Trim(txt_Tab1_Phone.Text)
End If

'����
Set tmp_para = tmp_Cmd.CreateParameter("Description", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_Description.Text) = "" Then
   tmp_Cmd.Parameters("Description").Value = Null
Else
   tmp_Cmd.Parameters("Description").Value = Trim(txt_Tab1_Description.Text)
End If

'�i�˸��O��
Set tmp_para = tmp_Cmd.CreateParameter("PALLET_CAPACITY", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_PalletCapacity.Text) = "" Then
   tmp_Cmd.Parameters("PALLET_CAPACITY").Value = Null
Else
   tmp_Cmd.Parameters("PALLET_CAPACITY").Value = Trim(txt_Tab1_PalletCapacity.Text)
End If

'����
Set tmp_para = tmp_Cmd.CreateParameter("CAR_WIEGHT", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_CarWeight.Text) = "" Then
   tmp_Cmd.Parameters("CAR_WIEGHT").Value = "0"
Else
   tmp_Cmd.Parameters("CAR_WIEGHT").Value = Trim(txt_Tab1_CarWeight.Text)
End If

'���[�Φ�
Set tmp_para = tmp_Cmd.CreateParameter("CARBOX_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_CarBox.ListIndex <> -1 Then
   tmp_Cmd.Parameters("CARBOX_TYPE").Value = arCarBox(cmb_Tab1_CarBox.ListIndex)
Else
   tmp_Cmd.Parameters("CARBOX_TYPE").Value = Null
End If

'���ɰ���
Set tmp_para = tmp_Cmd.CreateParameter("CAR_HEIGHT", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_CarHeight.Text) = "" Then
   tmp_Cmd.Parameters("CAR_HEIGHT").Value = Null   'Trim(txt_Tab1_CarHeight.Text)
Else
   tmp_Cmd.Parameters("CAR_HEIGHT").Value = Trim(txt_Tab1_CarHeight.Text)
End If

'�˨��覡
Set tmp_para = tmp_Cmd.CreateParameter("UNLAODING_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_UnloadType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("UNLAODING_TYPE").Value = arUnloadType(cmb_Tab1_UnloadType.ListIndex)
Else
   tmp_Cmd.Parameters("UNLAODING_TYPE").Value = Null
End If

'���Τ覡
Set tmp_para = tmp_Cmd.CreateParameter("EMPLOY_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_EmployType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("EMPLOY_TYPE").Value = arEmployType(cmb_Tab1_EmployType.ListIndex)
Else
   tmp_Cmd.Parameters("EMPLOY_TYPE").Value = Null
End If

'�p�O���O
Set tmp_para = tmp_Cmd.CreateParameter("CAR_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(cmb_Tab1_CarType.Text)) > 0 Then
   tmp_Cmd.Parameters("CAR_TYPE").Value = cmb_Tab1_CarType.Text
Else
   tmp_Cmd.Parameters("CAR_TYPE").Value = Null
End If

'�дڤH
Set tmp_para = tmp_Cmd.CreateParameter("Receiver", adVarChar, adParamInput, 50)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab1_Receiver)) > 0 Then
   tmp_Cmd.Parameters("Receiver").Value = txt_Tab1_Receiver
Else
   tmp_Cmd.Parameters("Receiver").Value = Null
End If

'PND
Set tmp_para = tmp_Cmd.CreateParameter("PND", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
If chkPND.Value = vbChecked Then
   tmp_Cmd.Parameters("PND").Value = "Y"
Else
   tmp_Cmd.Parameters("PND").Value = "N"
End If

'APFix
Set tmp_para = tmp_Cmd.CreateParameter("APFix", adDouble, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("APFix").Value = txtAPFix

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'�D�P�B����
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
Loop

'����EditWho
str_SQL = "update trp09m set editwho = '" & User_id & "' , editdate = getdate() where VEHICLE_ID_NO = '" & txt_Tab1_CarID & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'����ADDWho
str_SQL = "update trp09m set addwho = '" & User_id & "' where addwho is null and VEHICLE_ID_NO = '" & txt_Tab1_CarID & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'���ʬ���
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select * From TRP09M Where VEHICLE_ID_NO = '" & txt_Tab1_CarID & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then
    Dim str As String, i As Integer
    For i = 0 To tmp_Rs.Fields.Count - 1
        str = str & RTrim(tmp_Rs.Fields(i)) & ","
    Next i
    
    '�g�J��Ʈw����
    str_SQL = "Insert into gt_Logs(APName,APVer,APCaption,Code,Description,Notes,ComputerName,AddWho) Values ('" & _
                    App.EXEName & "','" & App.Major & "." & App.Minor & "." & App.Revision & "','" & Me.Caption & "','0','�����D�ɲ��ʬ���','" & str & "','" & strComputerName & "','" & User_id & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End If

fam_Tab1_Car.BackColor = &H8000000C
chkPND.BackColor = &H8000000C
fam_Tab1_Car.Enabled = False
cmd_Tab1_Cancel.Enabled = False
cmd_Tab1_Save.Enabled = False
cmd_Tab1_AddNew.Enabled = True
cmd_Tab1_Modify.Enabled = False
cmd_Tab1_Delete.Enabled = False

'���s��ܩҦ��Ȥ���
'Call cmd_Tab1_CarShow_Click'�אּ�����s�d�ߡA�קKUSER�q�Y�d��
If Not rs_Tab1_CarList Is Nothing Then
    If rs_Tab1_CarList("���P���X") = txt_Tab1_CarID Then '�D�s�W�ɧ�s�M����
        rs_Tab1_CarList("�r�p�H") = txt_Tab1_Driver
        rs_Tab1_CarList("����") = cmb_Tab1_VehicleType
        Call Display_SelectedCarData(rs_Tab1_CarList.Fields("���P���X").Value)
    End If
End If
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�������-�s��", Me.Caption, "cmd_Tab1_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Delete_Click()
'�f�B���q >> �R��
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

'1.�ˮ� TRP05T �O�_�����f�B���q�˸��X�����
str_SQL = "Select Count(*) as RecCnt From TRP05T Where TRP_Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   ���u�s���B�e��� [TRP05T] �����f�B���q�˸��X�����"
   Else
      msg_text = msg_text & vbCrLf & "   ���u�s���B�e��� [TRP05T] �����f�B���q�˸��X�����"
   End If
End If
tmp_Rs.Close

'2.�ˮ� ORT05T �O�_�����f�B���q�˸��X�����
str_SQL = "Select Count(*) as RecCnt From ORT05T Where TRP_Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   ���u�s���B�e��� [ORT05T] �����f�B���q�˸��X�����"
   Else
      msg_text = msg_text & vbCrLf & "   ���u�s���B�e��� [TRP05T] �����f�B���q�˸��X�����"
   End If
End If
tmp_Rs.Close

'3.�ˮ� TRP09M �O�_�����f�B���q�˸��X�����
str_SQL = "Select Count(*) as RecCnt From TRP09M Where TRP_Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   �����򥻸���� [TRP09M] �����f�B���q�������"
   Else
      msg_text = msg_text & vbCrLf & "   �����򥻸���� [TRP09M] �����f�B���q�������"
   End If
End If
tmp_Rs.Close

'�ˮ֬O�_���\�i��R���X�Э�
If blDelete = False Then
   msg_text = "�f�B���q��ƵL�k�R���G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'���\�R��
str_SQL = "Delete From TRP08M Where Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

fam_Tab2_Company.BackColor = &H8000000C
fam_Tab2_Company.Enabled = False
cmd_Tab2_Cancel.Enabled = False
cmd_Tab2_Save.Enabled = False
cmd_Tab2_AddNew.Enabled = False
cmd_Tab2_Modify.Enabled = False
cmd_Tab2_Delete.Enabled = False
'���s��ܩҦ��Ȥ���
Call cmd_Tab1_CarShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�f�B���q���-�R��", Me.Caption, "cmd_Tab2_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Modify_Click()
'�f�B���q��� >> �ק�
'�T�{���������Ƥ褹�\ [�ק�] �\��
If rs_Tab2_TRPCompanyList Is Nothing Then Exit Sub
If dg_Tab2_TRPCompanyList.SelBookmarks.Count <> 0 Then
   fam_Tab2_Company.BackColor = &HC0E0FF
   fam_Tab2_Company.Enabled = True
   txt_Tab2_CompanyCode.Enabled = False
   cmd_Tab2_Save.Enabled = True
   cmd_Tab2_Cancel.Enabled = True
   cmd_Tab2_AddNew.Enabled = False
   cmd_Tab2_Modify.Enabled = False
   cmd_Tab2_Delete.Enabled = False
End If

End Sub

Private Sub cmd_Tab2_Save_Click()

'�M���S��r��
Call myFormExCharFilter(Me)

'�f�B���q��� >> �s��
On Error GoTo err_Handle

'�s�ɸ���ˮ�
If Check_CompanyData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_CompanyData_Update"
'���q�N�X
Set tmp_para = tmp_Cmd.CreateParameter("Company_Code", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Company_Code").Value = Trim(txt_Tab2_CompanyCode.Text)
'����W��
Set tmp_para = tmp_Cmd.CreateParameter("C_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("C_Name").Value = Trim(txt_Tab2_CName.Text)
'�^��W��
Set tmp_para = tmp_Cmd.CreateParameter("E_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab2_EName.Text)) > 0 Then
   tmp_Cmd.Parameters("E_Name").Value = Trim(txt_Tab2_EName.Text)
Else
   tmp_Cmd.Parameters("E_Name").Value = Null
End If
'�a�}
Set tmp_para = tmp_Cmd.CreateParameter("Address", adVarChar, adParamInput, 45)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab2_Address.Text)) <> 0 Then
   tmp_Cmd.Parameters("Address").Value = Trim(txt_Tab2_Address.Text)
Else
   tmp_Cmd.Parameters("Address").Value = Null
End If
'²��
Set tmp_para = tmp_Cmd.CreateParameter("Short_Name", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab2_ShortName.Text)) <> 0 Then
   tmp_Cmd.Parameters("Short_Name").Value = Trim(txt_Tab2_ShortName.Text)
Else
   tmp_Cmd.Parameters("Short_Name").Value = Null
End If
'�p���H 'Terry 20180123 contact ���ץ�30�אּ80
Set tmp_para = tmp_Cmd.CreateParameter("Contact", adVarChar, adParamInput, 80)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab2_Contact.Text) = "" Then
   tmp_Cmd.Parameters("Contact").Value = Null
Else
   tmp_Cmd.Parameters("Contact").Value = Trim(txt_Tab2_Contact.Text)
End If
'�q��
Set tmp_para = tmp_Cmd.CreateParameter("Phone", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Phone").Value = Null
If Trim(txt_Tab2_Contact.Text) = "" Then
   tmp_Cmd.Parameters("Phone").Value = Null
Else
   tmp_Cmd.Parameters("Phone").Value = Trim(txt_Tab2_Phone.Text)
End If
'����
Set tmp_para = tmp_Cmd.CreateParameter("Description", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab2_Descr.Text) = "" Then
   tmp_Cmd.Parameters("Description").Value = Null
Else
   tmp_Cmd.Parameters("Description").Value = Trim(txt_Tab2_Descr.Text)
End If

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'�D�P�B����
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
Loop

fam_Tab2_Company.BackColor = &H8000000C
fam_Tab2_Company.Enabled = False
cmd_Tab2_Cancel.Enabled = False
cmd_Tab2_Save.Enabled = False
cmd_Tab2_AddNew.Enabled = True
cmd_Tab2_Modify.Enabled = False
cmd_Tab2_Delete.Enabled = False
'���s��ܩҦ��Ȥ���
Call cmd_Tab1_CarShow_Click
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�B�餽�q���-�s��", Me.Caption, "cmd_Tab2_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_SkuQuery_Click()
'�f����� >> �f����Ʒj�M
If rs_Tab3_SkuList Is Nothing Then Exit Sub
If rs_Tab3_SkuList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab3_SkuList"

If ShowForm_RS_FilterAndSort(rs_Tab3_SkuList, "�f�����", Me.Tag) = False Then
    MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
Me.WindowState = vbNormal
End Sub

Private Sub cmd_Tab2_SkuShow_Click()
'�f����� >> ��ܩҦ��f��
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab2_TRPCompanyList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2_TRPCompanyList)
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "select StorerKey as �f�D, Sku as �f��, DESCR as ����W��,  STDGROSSWGT as �C�c��, busr4  as �C�c��, " & _
        "isnull(SKUGROUP,'') as ���O,rtrim(SUSR1) as ���~�O,NOTES1 as �Ƶ��@,NOTES2 as �Ƶ��G  from gv_SKUxpack"

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C�f�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab3_SkuList)
tmp_Rs.Close

blTab3skuEventEnable = False
With dg_Tab3_SkuList
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab3_SkuList.MoveFirst
Set dg_Tab3_SkuList.DataSource = rs_Tab3_SkuList
With dg_Tab3_SkuList
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '�f�D
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 900       '�f��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 2500       '����W��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 900       '�C�c��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 900       '�C�c��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 900        '���O
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 900       '���~�O
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 2000       '�^��W��
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 2000       '����
    .Columns(9).Alignment = dbgLeft
End With
blTab3skuEventEnable = True
'Call Clear_CompanyData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�f�����-��ܩҦ����", Me.Caption, "cmd_Tab3_SKUShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_TRPCompanyReset_Click()
'�f�B���q�򥻸�� >> �����z��Ƨ�
'�����z�����A���]�ƧǨ̾�
If rs_Tab2_TRPCompanyList Is Nothing Then Exit Sub
 blTab2CompanyEventEnable = False
 rs_Tab2_TRPCompanyList.Filter = adFilterNone
 rs_Tab2_TRPCompanyList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
 blTab2CompanyEventEnable = True

End Sub

Private Sub cmd_Tab3_AddNew_Click()
'�f�B���q��� >> �s�W
If Not rs_Tab3_SkuList Is Nothing Then
   If dg_Tab3_SkuList.SelBookmarks.Count > 0 Then dg_Tab3_SkuList.SelBookmarks.Remove 0
End If
fam_Tab3_Sku.BackColor = &HC0FFC0
fam_Tab3_Sku.Enabled = True
txt_Tab3_Sku.Enabled = True
Call Clear_SkuData
cmd_Tab3_Save.Enabled = True
cmd_Tab3_Cancel.Enabled = True
cmd_Tab3_AddNew.Enabled = False
cmd_Tab3_Modify.Enabled = False
cmd_Tab3_Delete.Enabled = False
End Sub

Private Sub cmd_Tab3_Cancel_Click()
'�f����� >> ����
Call Clear_SkuData
If txt_Tab3_Sku.Enabled = False Then
   If Not rs_Tab3_SkuList Is Nothing Then
      dg_Tab3_SkuList.SelBookmarks.Add rs_Tab3_SkuList.Bookmark
      Call Display_SelectedSkuData(rs_Tab3_SkuList.Fields("�f��").Value)
   End If
End If
fam_Tab3_Sku.BackColor = &H8000000C
fam_Tab3_Sku.Enabled = False
cmd_Tab3_Cancel.Enabled = False
cmd_Tab3_Save.Enabled = False
cmd_Tab3_AddNew.Enabled = True
cmd_Tab3_Modify.Enabled = True
cmd_Tab3_Delete.Enabled = True
End Sub

Private Sub cmd_Tab3_Modify_Click()
'�f�����q��� >> �ק�
'�T�{���������Ƥ褹�\ [�ק�] �\��
If rs_Tab3_SkuList Is Nothing Then Exit Sub
If dg_Tab3_SkuList.SelBookmarks.Count <> 0 Then
   fam_Tab3_Sku.BackColor = &HC0E0FF
   fam_Tab3_Sku.Enabled = True
   txt_Tab3_Sku.Enabled = False
   cmd_Tab3_Save.Enabled = True
   cmd_Tab3_Cancel.Enabled = True
   cmd_Tab3_AddNew.Enabled = False
   cmd_Tab3_Modify.Enabled = False
   cmd_Tab3_Delete.Enabled = False
End If
End Sub

Private Sub cmd_Tab3_Save_Click()
'�f�B���q��� >> �s��

'�M���S��r��
Call myFormExCharFilter(Me)

On Error GoTo err_Handle
'select StorerKey as �f�D, Sku as �f��, DESCR as ����W��, STDGROSSWGT as �C�c��,rtrim(BUSR4) as �C�c��, " & _
        "SKUGROUP as ���O,rtrim(BUSR1) as ���~�O,NOTES1 as �^��W��,NOTES2 as ����  from dbo.SKU"
'�s�ɸ���ˮ�
If Check_SkuData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_SkuData_Update"
'�f�D
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("StorerKey").Value = Trim(txt_Tab3_StorerKey.Text)
'�f��
Set tmp_para = tmp_Cmd.CreateParameter("Sku", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Sku").Value = Trim(txt_Tab3_Sku.Text)
'����W��
Set tmp_para = tmp_Cmd.CreateParameter("DESCR", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
'If Len(Trim(txt_Tab3_DESCR.Text)) > 0 Then
   tmp_Cmd.Parameters("DESCR").Value = Trim(txt_Tab3_DESCR.Text)
'Else
'   tmp_cmd.Parameters("DESCR").Value = Null
'End If
'�C�c��
Set tmp_para = tmp_Cmd.CreateParameter("STDGROSSWGT", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab3_STDGROSSWGT.Text)) <> 0 Then
   tmp_Cmd.Parameters("STDGROSSWGT").Value = Trim(txt_Tab3_STDGROSSWGT.Text)
Else
   tmp_Cmd.Parameters("STDGROSSWGT").Value = 0
End If
'�C�c��
Set tmp_para = tmp_Cmd.CreateParameter("BUSR4", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab3_BUSR4.Text)) <> 0 Then
   tmp_Cmd.Parameters("BUSR4").Value = Trim(txt_Tab3_BUSR4.Text)
Else
   tmp_Cmd.Parameters("BUSR4").Value = 0
End If
'���O
Set tmp_para = tmp_Cmd.CreateParameter("SKUGROUP", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_SKUGROUP.Text) = "" Then
'   tmp_cmd.Parameters("SKUGROUP").Value = Null
'Else
   tmp_Cmd.Parameters("SKUGROUP").Value = Trim(txt_Tab3_SKUGROUP.Text)
'End If
'���~�O
Set tmp_para = tmp_Cmd.CreateParameter("BUSR1", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_BUSR1.Text) = "" Then
'   tmp_cmd.Parameters("BUSR1").Value = Null
'Else
   tmp_Cmd.Parameters("BUSR1").Value = Trim(txt_Tab3_BUSR1.Text)
'End If
'�^��W��
Set tmp_para = tmp_Cmd.CreateParameter("NOTES1", adVarChar, adParamInput, 40)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_NOTES1.Text) = "" Then
'   tmp_cmd.Parameters("NOTES1").Value = Null
'Else
   tmp_Cmd.Parameters("NOTES1").Value = Trim(txt_Tab3_NOTES1.Text)
'End If
'����
Set tmp_para = tmp_Cmd.CreateParameter("NOTES2", adVarChar, adParamInput, 40)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_NOTES2.Text) = "" Then
'   tmp_cmd.Parameters("NOTES2").Value = Null
'Else
   tmp_Cmd.Parameters("NOTES2").Value = Trim(txt_Tab3_NOTES2.Text)
'End If
Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'�D�P�B����
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
Loop

fam_Tab3_Sku.BackColor = &H8000000C
fam_Tab3_Sku.Enabled = False
cmd_Tab3_Cancel.Enabled = False
cmd_Tab3_Save.Enabled = False
cmd_Tab3_AddNew.Enabled = True
cmd_Tab3_Modify.Enabled = False
cmd_Tab3_Delete.Enabled = False
'���s��ܩҦ��Ȥ���
Call cmd_Tab2_SkuShow_Click
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�f�����-�s��", Me.Caption, "cmd_Tab3_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Public Sub dg_Tab0_ConsigneeList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�Ȥ��ƦC��G�����
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
If blTab0ConsignEventEnable Then
   If Not rs_Tab0_ConsigneeList.EOF Then
      dg_Tab0_ConsigneeList.SelBookmarks.Add rs_Tab0_ConsigneeList.Bookmark
      Call Display_SelectedConsignData(rs_Tab0_ConsigneeList.Fields("�f�D").Value, rs_Tab0_ConsigneeList.Fields("�Ȥ�s��").Value)
      fam_Tab0_Consignee.BackColor = &H8000000C
      fam_Tab0_Consignee.Enabled = False
      cmd_Tab0_Cancel.Enabled = False
      cmd_Tab0_Save.Enabled = False
      cmd_Tab0_AddNew.Enabled = True
      cmd_Tab0_Modify.Enabled = True
      cmd_Tab0_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab1_CarList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'������ƦC��G�����
If blTab1CarEventEnable Then
   If Not rs_Tab1_CarList.EOF Then
      dg_Tab1_CarList.SelBookmarks.Add rs_Tab1_CarList.Bookmark
      Call Display_SelectedCarData(rs_Tab1_CarList.Fields("���P���X").Value)
      fam_Tab1_Car.BackColor = &H8000000C
      chkPND.BackColor = &H8000000C
      fam_Tab1_Car.Enabled = False
      cmd_Tab1_Cancel.Enabled = False
      cmd_Tab1_Save.Enabled = False
      cmd_Tab1_AddNew.Enabled = True
      cmd_Tab1_Modify.Enabled = True
      cmd_Tab1_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab2_TRPCompanyList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'������ƦC��G�����
If blTab2CompanyEventEnable Then
   If Not rs_Tab2_TRPCompanyList.EOF Then
      dg_Tab2_TRPCompanyList.SelBookmarks.Add rs_Tab2_TRPCompanyList.Bookmark
      Call Display_SelectedCompanyData(rs_Tab2_TRPCompanyList.Fields("���q�N�X").Value)
      fam_Tab2_Company.BackColor = &H8000000C
      fam_Tab2_Company.Enabled = False
      cmd_Tab2_Cancel.Enabled = False
      cmd_Tab2_Save.Enabled = False
      cmd_Tab2_AddNew.Enabled = True
      cmd_Tab2_Modify.Enabled = True
      cmd_Tab2_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab3_SkuList_Click()
'������ƦC��G�����
If blTab3skuEventEnable Then
   If Not rs_Tab3_SkuList.EOF Then
      dg_Tab3_SkuList.SelBookmarks.Add rs_Tab3_SkuList.Bookmark
'      Call Display_SelectedSkuData(rs_Tab3_SkuList.Fields("�f��").Value)
      fam_Tab3_Sku.BackColor = &H8000000C
      fam_Tab3_Sku.Enabled = False
      cmd_Tab3_Cancel.Enabled = False
      cmd_Tab3_Save.Enabled = False
      cmd_Tab3_AddNew.Enabled = True
      cmd_Tab3_Modify.Enabled = True
      cmd_Tab3_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab3_SkuList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'�P�@����
If LastRow = Empty Then Exit Sub

'�O�_�����
If rs_Tab3_SkuList Is Nothing Then Exit Sub
If rs_Tab3_SkuList.RecordCount = 0 Then Exit Sub

Call Display_SelectedSkuData(rs_Tab3_SkuList.Fields("�f��").Value)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub drvLocalDriveT5_Change()
    On Error GoTo DriveError
    dirLocalDirT5.Path = drvLocalDriveT5.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�Ȥ�/�����򥻸�ƺ��@�@�~"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475

Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

Dim tmp_cnt As Integer

''���X�Ҧ��f�D���--TRP16M

'cmb_Tab0_Storer.Clear
'cmb_Tab4_Storer.Clear   '�Ȥ᤹���ѼƳf�D
'str_SQL = "Select Rtrim(StorerKey) as 'StorerKey',Isnull(Rtrim(Short_Name),'') as 'StorerName' From TRP16M Order by StorerKey"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
'ReDim arStorer(1) As String
'If Not tmp_Rs.EOF Then
'   tmp_cnt = 0
'   Do While Not tmp_Rs.EOF
'      arStorer(tmp_cnt) = tmp_Rs.Fields("StorerKey").Value
'      cmb_Tab0_Storer.AddItem tmp_Rs.Fields("StorerKey").Value & Space(7 - Len(Trim(tmp_Rs.Fields("StorerKey").Value))) & tmp_Rs.Fields("StorerName").Value
'      cmb_Tab4_Storer.AddItem tmp_Rs.Fields("StorerKey").Value & Space(7 - Len(Trim(tmp_Rs.Fields("StorerKey").Value))) & tmp_Rs.Fields("StorerName").Value
'      tmp_Rs.MoveNext
'      tmp_cnt = tmp_cnt + 1
'      If tmp_cnt = UBound(arStorer) Then
'         ReDim Preserve arStorer(UBound(arStorer) + 10) As String
'      End If
'   Loop
'End If
'tmp_Rs.Close

'���X�Ҧ��l���ϸ� TRP02M
cmb_Tab0_Zip.Clear: cmb_Tab1_ZIP.Clear
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
      cmb_Tab0_Zip.AddItem tmp_Rs.Fields("ZIP").Value & Space(5 - Len(Trim(tmp_Rs.Fields("ZIP").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab1_ZIP.AddItem tmp_Rs.Fields("ZIP").Value & Space(5 - Len(Trim(tmp_Rs.Fields("ZIP").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arZip) Then
         ReDim Preserve arZip(UBound(arZip) + 10) As String
         ReDim Preserve arZIPArea(UBound(arZIPArea) + 10) As String
      End If
   Loop
End If

'���X�Ҧ��B�e�ϰ�N�X TRP03M
cmb_Tab0_AreaCode.Clear: cmb_Tab1_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_Tab0_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab1_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'���X�Ҧ����ظ��--TRP15M
cmb_Tab0_VehicleType.Clear: cmb_Tab1_VehicleType.Clear
str_SQL = "Select Rtrim(Vehicle_Type) as 'VType',Isnull(Rtrim(Description),'') as 'VTypeDescr',car_type  From TRP15M Order by Vehicle_Type"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arVehicleType(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arVehicleType(tmp_cnt) = tmp_Rs.Fields("VType").Value
      cmb_Tab0_VehicleType.AddItem tmp_Rs.Fields("VType").Value & Space(4 - Len(Trim(tmp_Rs.Fields("VType").Value))) & tmp_Rs.Fields("VTypeDescr").Value
      cmb_Tab1_VehicleType.AddItem tmp_Rs("VType") & Space(4 - Len(Trim(tmp_Rs.Fields("VType").Value))) & tmp_Rs.Fields("VTypeDescr") & "/" & tmp_Rs("Car_type")
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arVehicleType) Then
         ReDim Preserve arVehicleType(UBound(arVehicleType) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'���X�Ҧ��S��ݨD--TRP04M
cmb_Tab0_ExtraDemand1.Clear: cmb_Tab0_ExtraDemand2.Clear
str_SQL = "Select Rtrim(Extra_Demand_Code) as 'ECode',Isnull(Rtrim(Description),'') as 'ECodeDescr' From TRP04M Order by Extra_Demand_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arExtraDemand(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arExtraDemand(tmp_cnt) = tmp_Rs.Fields("ECode").Value
      cmb_Tab0_ExtraDemand1.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      cmb_Tab0_ExtraDemand2.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arExtraDemand) Then
         ReDim Preserve arExtraDemand(UBound(arExtraDemand) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'���X�Ҧ��f�B���q--TRP09M
cmb_Tab1_Company.Clear
str_SQL = "Select Rtrim(Company_Code) as 'CCode',Isnull(Rtrim(C_Name),'') as 'CName' From TRP08M Order by Company_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arCompany(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arCompany(tmp_cnt) = tmp_Rs.Fields("CCode").Value
      cmb_Tab1_Company.AddItem tmp_Rs.Fields("CCode").Value & Space(5 - Len(Trim(tmp_Rs.Fields("CCode").Value))) & tmp_Rs.Fields("CName").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arCompany) Then
         ReDim Preserve arCompany(UBound(arCompany) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'���X�Ҧ����[�Φ�--CODELKUP.ListName = [CARBOXTYPE]
cmb_Tab1_CarBox.Clear
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS ���[�Φ� " & _
          "From CodeLKUP Where ListName = 'CARBOXTYPE'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arCarBox(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arCarBox(tmp_cnt) = tmp_Rs.Fields("�N�X").Value
      cmb_Tab1_CarBox.AddItem tmp_Rs.Fields("�N�X").Value & Space(5 - Len(Trim(tmp_Rs.Fields("�N�X").Value))) & tmp_Rs.Fields("���[�Φ�").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arCarBox) Then
         ReDim Preserve arCarBox(UBound(arCarBox) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'���X�Ҧ����Τ覡--CODELKUP.ListName = [EMPLOYTYPE]
cmb_Tab1_EmployType.Clear
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS ���Τ覡 " & _
          "From CodeLKUP Where ListName = 'EMPLOYTYPE'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arEmployType(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arEmployType(tmp_cnt) = tmp_Rs.Fields("�N�X").Value
      cmb_Tab1_EmployType.AddItem tmp_Rs.Fields("�N�X").Value & Space(5 - Len(Trim(tmp_Rs.Fields("�N�X").Value))) & tmp_Rs.Fields("���Τ覡").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arEmployType) Then
         ReDim Preserve arEmployType(UBound(arEmployType) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'���X�Ҧ��˨��覡--CODELKUP.ListName = [LOADUNLOADTYPE]
cmb_Tab1_UnloadType.Clear
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS �˨��覡 " & _
          "From CodeLKUP Where ListName = 'LOADUNLOADTYPE'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arUnloadType(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arUnloadType(tmp_cnt) = tmp_Rs.Fields("�N�X").Value
      cmb_Tab1_UnloadType.AddItem tmp_Rs.Fields("�N�X").Value & Space(5 - Len(Trim(tmp_Rs.Fields("�N�X").Value))) & tmp_Rs.Fields("�˨��覡").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arUnloadType) Then
         ReDim Preserve arUnloadType(UBound(arUnloadType) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'���o �h�B�u��
cmb_Tab0_PickTool.Clear: tmp_cnt = 0
ReDim arPickTool(1) As String
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS �h�B�u�� " & _
          "From CodeLKUP Where ListName = 'MOVETOOL'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_Tab0_PickTool.AddItem tmp_Rs.Fields("�N�X").Value & "  " & tmp_Rs.Fields("�h�B�u��").Value
   tmp_cnt = tmp_cnt + 1
   If UBound(arPickTool) < tmp_cnt Then
      ReDim Preserve arPickTool(tmp_cnt) As String
   End If
   arPickTool(tmp_cnt - 1) = tmp_Rs.Fields("�N�X").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_Tab0_PickTool.ListIndex = -1

'���X�Ҧ��q����t--TRP18M.consigneekey
cmb_Tab0_Group.Clear
str_SQL = "Select distinct custgroup as 'custGroup' from TRP01M Order by custgroup"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If Not tmp_Rs.EOF Then
   Do While Not tmp_Rs.EOF
      cmb_Tab0_Group.AddItem RTrim(tmp_Rs.Fields("custGroup").Value)
      tmp_Rs.MoveNext
   Loop
End If
cmb_Tab0_Group = ""
tmp_Rs.Close

blTab0ConsignEventEnable = True

SSTab1.Tab = 1

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
Set frm_BaseData_Car = Nothing
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub Display_SelectedConsignData(ByVal strStorerkey As String, ByVal strConsigneeKey As String)
'��ܶǤJ���Ȥ���
Call Clear_ConsigneeData

str_SQL = "Select * From TRP01M Where ConsigneeKey = '" & strConsigneeKey & "' and Storerkey = '" & strStorerkey & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫Ȥ�򥻸��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab0_ConsigneeKey.Text = Trim(tmp_Rs.Fields("ConsigneeKey").Value)
DoEvents: DoEvents
For i = 0 To cmb_Tab0_Storer.ListCount - 1
    If arStorer(i) = Trim(tmp_Rs.Fields("StorerKey").Value) Then
       cmb_Tab0_Storer.ListIndex = i
       Exit For
    End If
Next i
If IsNull(tmp_Rs.Fields("ZIP").Value) Then
   cmb_Tab0_Zip.ListIndex = -1
Else
DoEvents: DoEvents
   For i = 0 To cmb_Tab0_Zip.ListCount - 1
       If arZip(i) = Trim(tmp_Rs.Fields("ZIP").Value) Then
          cmb_Tab0_Zip.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab0_Class.Text = IIf(IsNull(tmp_Rs.Fields("Class").Value), "", Trim(tmp_Rs.Fields("Class").Value))
If IsNull(tmp_Rs.Fields("Area_Code").Value) Then
   cmb_Tab0_AreaCode.ListIndex = -1
Else
DoEvents: DoEvents
   For i = 0 To cmb_Tab0_AreaCode.ListCount - 1
       If arAreaCode(i) = Trim(tmp_Rs.Fields("Area_Code").Value) Then
          cmb_Tab0_AreaCode.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab0_FullName.Text = IIf(IsNull(tmp_Rs.Fields("Full_Name").Value), "", Trim(tmp_Rs.Fields("Full_Name").Value))
txt_Tab0_Address.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
txt_Tab0_Contact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_Tab0_Phone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_Tab0_ShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))
txt_Tab0_GridCode.Text = IIf(IsNull(tmp_Rs.Fields("Grid_Code").Value), "", Trim(tmp_Rs.Fields("Grid_Code").Value))
If IsNull(tmp_Rs.Fields("Vehicle_Type").Value) Then
   cmb_Tab0_VehicleType.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_VehicleType.ListCount - 1
       If arVehicleType(i) = Trim(tmp_Rs.Fields("Vehicle_Type").Value) Then
          cmb_Tab0_VehicleType.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Extra_Demand_Code").Value) Then
   cmb_Tab0_ExtraDemand1.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_ExtraDemand1.ListCount - 1
       If arExtraDemand(i) = Trim(tmp_Rs.Fields("Extra_Demand_Code").Value) Then
          cmb_Tab0_ExtraDemand1.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Extra_Demand_Code2").Value) Then
   cmb_Tab0_ExtraDemand2.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_ExtraDemand2.ListCount - 1
       If arExtraDemand(i) = Trim(tmp_Rs.Fields("Extra_Demand_Code2").Value) Then
          cmb_Tab0_ExtraDemand2.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab0_ChannelType.Text = IIf(IsNull(tmp_Rs.Fields("Channel_Type").Value), "", Trim(tmp_Rs.Fields("Channel_Type").Value))
txt_Tab0_Channel.Text = IIf(IsNull(tmp_Rs.Fields("Channel").Value), "", Trim(tmp_Rs.Fields("Channel").Value))
txt_Tab0_UnLoad.Text = IIf(IsNull(tmp_Rs.Fields("Unload_Type").Value), "", Trim(tmp_Rs.Fields("Unload_Type").Value))

If IsNull(tmp_Rs.Fields("Multi_Customer").Value) Then
   chk_Tab0_MultiCustomer.Value = vbUnchecked
Else
   If Trim(tmp_Rs.Fields("Multi_Customer").Value) = "N" Then
      chk_Tab0_MultiCustomer.Value = vbUnchecked
   Else
      chk_Tab0_MultiCustomer.Value = vbChecked
   End If
End If

If IsNull(tmp_Rs.Fields("DC").Value) Then
   chk_Tab0_MultiCustomer.Value = vbUnchecked
Else
   If Trim(tmp_Rs.Fields("DC").Value) = "N" Then
      chkDC.Value = vbUnchecked
   Else
      chkDC.Value = vbChecked
   End If
End If


If IsNull(tmp_Rs.Fields("PICK_TOOL").Value) Then
   cmb_Tab0_PickTool.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_PickTool.ListCount - 1
       If arPickTool(i) = Trim(tmp_Rs.Fields("pick_tool").Value) Then
          cmb_Tab0_PickTool.ListIndex = i
          Exit For
       End If
   Next i
End If

'�q����t
cmb_Tab0_Group = RTrim(tmp_Rs.Fields("CustGroup"))

txt_Tab0_CodeDate1.Text = IIf(IsNull(tmp_Rs.Fields("CodeDate1").Value), "", Trim(tmp_Rs.Fields("CodeDate1").Value))
txt_Tab0_CodeDate2.Text = IIf(IsNull(tmp_Rs.Fields("CodeDate2").Value), "", Trim(tmp_Rs.Fields("CodeDate2").Value))
txt_Tab0_CodeDate3.Text = IIf(IsNull(tmp_Rs.Fields("CodeDate3").Value), "", Trim(tmp_Rs.Fields("CodeDate3").Value))
txt_Tab0_Fax.Text = IIf(IsNull(tmp_Rs.Fields("fax").Value), "", Trim(tmp_Rs.Fields("fax").Value))
txt_Tab0_Stamp.Text = IIf(IsNull(tmp_Rs.Fields("stamp").Value), "", Trim(tmp_Rs.Fields("stamp").Value))
txt_Tab0_Penalties.Text = IIf(IsNull(tmp_Rs.Fields("Penalties").Value), "", Trim(tmp_Rs.Fields("Penalties").Value))
txt_Tab0_PalletType.Text = IIf(IsNull(tmp_Rs.Fields("PalletType").Value), "", Trim(tmp_Rs.Fields("PalletType").Value))
txt_Tab0_PalletSpec.Text = IIf(IsNull(tmp_Rs.Fields("PalletSpec").Value), "", Trim(tmp_Rs.Fields("PalletSpec").Value))
txt_Tab0_Notes.Text = IIf(IsNull(tmp_Rs.Fields("Notes").Value), "", Trim(tmp_Rs.Fields("Notes").Value))

'������
cmdCodeDateRate = tmp_Rs("CodeDateRate")

tmp_Rs.Close

End Sub

Private Sub Clear_ConsigneeData()
'�M�� �Ȥ��� �e��������
txt_Tab0_ConsigneeKey.Text = ""
cmb_Tab0_Storer.ListIndex = -1
cmb_Tab0_Zip.ListIndex = -1
txt_Tab0_Class.Text = ""
cmb_Tab0_AreaCode.ListIndex = -1
txt_Tab0_FullName.Text = ""
txt_Tab0_Address.Text = ""
txt_Tab0_Contact.Text = ""
txt_Tab0_Phone.Text = ""
txt_Tab0_ShortName.Text = ""
txt_Tab0_GridCode.Text = ""
cmb_Tab0_VehicleType.ListIndex = -1
cmb_Tab0_ExtraDemand1.ListIndex = -1: cmb_Tab0_ExtraDemand2.ListIndex = -1
txt_Tab0_Channel.Text = ""
txt_Tab0_ChannelType.Text = ""
txt_Tab0_UnLoad.Text = ""
chk_Tab0_MultiCustomer.Value = vbUnchecked
chkDC.Value = vbUnchecked
txt_Tab0_Fax.Text = ""
txt_Tab0_CodeDate1.Text = ""
txt_Tab0_CodeDate2.Text = ""
txt_Tab0_CodeDate3.Text = ""
txt_Tab0_Stamp.Text = ""
txt_Tab0_Penalties.Text = ""
txt_Tab0_PalletType.Text = ""
txt_Tab0_PalletSpec.Text = ""
txt_Tab0_Notes.Text = ""
cmb_Tab0_Group.Text = ""
cmdCodeDateRate = ""
End Sub

Private Function Check_ComsigneeData() As Boolean
'�Ȥ�򥻸���ˮ�

Check_ComsigneeData = False
msg_text = ""

If cmb_Tab0_Zip.ListIndex = -1 Then
   If msg_text = "" Then
      msg_text = "����J�l���ϸ�"
   Else
      msg_text = msg_text & vbCrLf & "����J�l���ϸ�"
   End If
End If

If Len(Trim(txt_Tab0_FullName.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J�Ȥ�W��"
   Else
      msg_text = msg_text & vbCrLf & "����J�Ȥ�W��"
   End If
End If

If Len(Trim(txt_Tab0_ShortName.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J�Ȥ�²��"
   Else
      msg_text = msg_text & vbCrLf & "����J�Ȥ�²��"
   End If
End If

If Len(Trim(txt_Tab0_Address.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J�B�e�a�}"
   Else
      msg_text = msg_text & vbCrLf & "����J�B�e�a�}"
   End If
End If

If txt_Tab0_ConsigneeKey.Enabled = True And UCase(Left(Trim(txt_Tab0_ConsigneeKey.Text), 4)) = "BEST" Then MsgBox "�Ƚs�}�Y����ϥ�""BEST""�O�d�r!!", 16, "�`�N": Exit Function

If Trim(cmdCodeDateRate) = "1/2" Or Trim(cmdCodeDateRate) = "2/3" Or Trim(cmdCodeDateRate) = "" Then
Else
   If msg_text = "" Then
      msg_text = "���~��� [��������]�A���������u���\���1/2�B2/3�P�ť� "
   Else
      msg_text = msg_text & vbCrLf & "�����������~"
   End If
End If

If Len(Trim(txt_Tab0_CodeDate1.Text)) = 0 And mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
   If msg_text = "" Then
      msg_text = "����J [��s������]"
   Else
      msg_text = msg_text & vbCrLf & "����J�Ȥ᤹����(���~���������N�v�T�t�f���T��)"
   End If
End If

If Len(Trim(txt_Tab0_CodeDate2.Text)) = 0 And mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
   If msg_text = "" Then
      msg_text = "����J [�M�s������]"
   Else
      msg_text = msg_text & vbCrLf & "����J�Ȥ᤹����(���~���������N�v�T�t�f���T��)"
   End If
End If

If Len(Trim(txt_Tab0_CodeDate3.Text)) = 0 And mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
   If msg_text = "" Then
      msg_text = "����J [���Ƥ�����]"
   Else
      msg_text = msg_text & vbCrLf & "����J�Ȥ᤹����(���~���������N�v�T�t�f���T��)"
   End If
End If

If msg_text = "" Then
   Check_ComsigneeData = True
Else
   msg_text = "�Ȥ��Ʋ��`�A�Эץ���A���� [�s ��]�G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function

Private Sub Display_SelectedCarData(ByVal strCarID As String)
'��ܶǤJ�������򥻸��
Call Clear_CarData

str_SQL = "Select * From TRP09M Where Vehicle_ID_No = '" & strCarID & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧨����򥻸��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab1_CarID.Text = Trim(tmp_Rs.Fields("Vehicle_ID_No").Value)
txt_Tab1_Driver.Text = IIf(IsNull(tmp_Rs.Fields("Driver").Value), "", Trim(tmp_Rs.Fields("Driver").Value))
txt_Tab1_Phone.Text = IIf(IsNull(tmp_Rs.Fields("DRIVER_PHONE").Value), "", Trim(tmp_Rs.Fields("DRIVER_PHONE").Value))
If IsNull(tmp_Rs.Fields("ZIP").Value) Then
   cmb_Tab1_ZIP.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_ZIP.ListCount - 1
       If arZip(i) = Trim(tmp_Rs.Fields("ZIP").Value) Then
          cmb_Tab1_ZIP.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Area_Code").Value) Then
   cmb_Tab1_AreaCode.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_AreaCode.ListCount - 1
       If arAreaCode(i) = Trim(tmp_Rs.Fields("Area_Code").Value) Then
          cmb_Tab1_AreaCode.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Vehicle_Type").Value) Then
   cmb_Tab1_VehicleType.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_VehicleType.ListCount - 1
       If arVehicleType(i) = Trim(tmp_Rs.Fields("Vehicle_Type").Value) Then
          cmb_Tab1_VehicleType.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("TRP_COMPANY_CODE").Value) Then
   cmb_Tab1_Company.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_Company.ListCount - 1
       If arCompany(i) = Trim(tmp_Rs.Fields("TRP_COMPANY_CODE").Value) Then
          cmb_Tab1_Company.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab1_CarWeight.Text = IIf(IsNull(tmp_Rs.Fields("CAR_WIEGHT").Value), "", Trim(tmp_Rs.Fields("CAR_WIEGHT").Value))
txt_Tab1_CarHeight.Text = IIf(IsNull(tmp_Rs.Fields("CAR_HEIGHT").Value), "", Trim(tmp_Rs.Fields("CAR_HEIGHT").Value))
If IsNull(tmp_Rs.Fields("CARBOX_TYPE").Value) Then
   cmb_Tab1_CarBox.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_CarBox.ListCount - 1
   DoEvents: DoEvents
       If arCarBox(i) = Trim(tmp_Rs.Fields("CARBOX_TYPE").Value) Then
          cmb_Tab1_CarBox.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(tmp_Rs.Fields("EMPLOY_TYPE").Value) Then
   cmb_Tab1_EmployType.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_EmployType.ListCount - 1
       If arEmployType(i) = Trim(tmp_Rs.Fields("EMPLOY_TYPE").Value) Then
          cmb_Tab1_EmployType.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(tmp_Rs.Fields("UNLAODING_TYPE").Value) Then
   cmb_Tab1_UnloadType.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_UnloadType.ListCount - 1
       If arUnloadType(i) = Trim(tmp_Rs.Fields("UNLAODING_TYPE").Value) Then
          cmb_Tab1_UnloadType.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(tmp_Rs.Fields("PND").Value) Then
   chkPND.Value = vbUnchecked
Else
   If Trim(tmp_Rs.Fields("pnd").Value) = "N" Then
      chkPND.Value = vbUnchecked
   Else
      chkPND.Value = vbChecked
   End If
End If

txtAPFix = (tmp_Rs.Fields("APFix").Value)

txt_Tab1_WeightCapacity.Text = IIf(IsNull(tmp_Rs.Fields("LOADING_SIZE").Value), "", Trim(tmp_Rs.Fields("LOADING_SIZE").Value))
txt_Tab1_VolumnCapacity.Text = IIf(IsNull(tmp_Rs.Fields("MAX_CUBIC_CAPACITY").Value), "", Trim(tmp_Rs.Fields("MAX_CUBIC_CAPACITY").Value))
txt_Tab1_PalletCapacity.Text = IIf(IsNull(tmp_Rs.Fields("PALLET_CAPACITY").Value), "", Trim(tmp_Rs.Fields("PALLET_CAPACITY").Value))
txt_Tab1_Description.Text = IIf(IsNull(tmp_Rs.Fields("Description").Value), "", Trim(tmp_Rs.Fields("Description").Value))
cmb_Tab1_CarType.Text = IIf(IsNull(tmp_Rs.Fields("CAR_TYPE").Value), "", Trim(tmp_Rs.Fields("CAR_TYPE").Value))
txt_Tab1_Receiver = tmp_Rs("receiver") & ""
txt_Tab1_Receiver = tmp_Rs("receiver") & ""
txtAdd = tmp_Rs("addwho") & " / " & tmp_Rs("adddate")
txtEdit = tmp_Rs("editwho") & " / " & tmp_Rs("editdate")
tmp_Rs.Close

End Sub
Private Sub Display_SelectedCompanyData(ByVal strCompanyCode As String)
'��ܶǤJ���f�B���q�򥻸��
Call Clear_CompanyData

str_SQL = "Select * From TRP08M Where Company_Code = '" & strCompanyCode & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧳f�B���q�򥻸��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab2_CompanyCode.Text = Trim(tmp_Rs.Fields("company_code").Value)
txt_Tab2_CName.Text = IIf(IsNull(tmp_Rs.Fields("C_NAME").Value), "", Trim(tmp_Rs.Fields("C_NAME").Value))
txt_Tab2_EName.Text = IIf(IsNull(tmp_Rs.Fields("E_NAME").Value), "", Trim(tmp_Rs.Fields("E_NAME").Value))
txt_Tab2_Address.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
txt_Tab2_Descr.Text = IIf(IsNull(tmp_Rs.Fields("Description").Value), "", Trim(tmp_Rs.Fields("Description").Value))
txt_Tab2_Contact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_Tab2_Phone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_Tab2_ShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))

tmp_Rs.Close

End Sub

Private Sub Display_SelectedSkuData(ByVal strSkuCode As String)
'��ܶǤJ���f���򥻸��
Call Clear_SkuData

str_SQL = "Select StorerKey,Sku,DESCR, STDGROSSWGT, busr4," & _
    "SKUGROUP,BUSR1,NOTES1,NOTES2 From gv_Skuxpack Where Sku = '" & strSkuCode & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab3_StorerKey.Text = Trim(tmp_Rs.Fields("StorerKey").Value)
txt_Tab3_Sku.Text = IIf(IsNull(tmp_Rs.Fields("Sku").Value), "", Trim(tmp_Rs.Fields("Sku").Value))
txt_Tab3_DESCR.Text = IIf(IsNull(tmp_Rs.Fields("DESCR").Value), "", Trim(tmp_Rs.Fields("DESCR").Value))
txt_Tab3_STDGROSSWGT.Text = IIf(IsNull(tmp_Rs.Fields("STDGROSSWGT").Value), "", Trim(tmp_Rs.Fields("STDGROSSWGT").Value))
txt_Tab3_BUSR4.Text = IIf(IsNull(tmp_Rs.Fields("busr4").Value), "", Trim(tmp_Rs.Fields("busr4").Value))
txt_Tab3_SKUGROUP.Text = IIf(IsNull(tmp_Rs.Fields("SKUGROUP").Value), "", Trim(tmp_Rs.Fields("SKUGROUP").Value))
txt_Tab3_BUSR1.Text = IIf(IsNull(tmp_Rs.Fields("BUSR1").Value), "", Trim(tmp_Rs.Fields("BUSR1").Value))
txt_Tab3_NOTES1.Text = IIf(IsNull(tmp_Rs.Fields("NOTES1").Value), "", Trim(tmp_Rs.Fields("NOTES1").Value))
txt_Tab3_NOTES2.Text = IIf(IsNull(tmp_Rs.Fields("NOTES2").Value), "", Trim(tmp_Rs.Fields("NOTES2").Value))
tmp_Rs.Close

End Sub

Private Sub Clear_CarData()
'�M�� ������� �e��������
txt_Tab1_CarID.Text = ""
txt_Tab1_Driver.Text = ""
txt_Tab1_Phone.Text = ""
cmb_Tab1_ZIP.ListIndex = -1
cmb_Tab1_AreaCode.ListIndex = -1
cmb_Tab1_Company.ListIndex = -1
txt_Tab1_CarWeight.Text = ""
txt_Tab1_CarHeight.Text = ""
cmb_Tab1_VehicleType.ListIndex = -1
cmb_Tab1_CarBox.ListIndex = -1
cmb_Tab1_EmployType.ListIndex = -1
cmb_Tab1_UnloadType.ListIndex = -1
txt_Tab1_VolumnCapacity.Text = ""
txt_Tab1_WeightCapacity.Text = ""
txt_Tab1_PalletCapacity.Text = ""
txt_Tab1_Receiver.Text = ""
cmb_Tab1_CarType = ""
chkPND = vbUnchecked
txtAPFix = ""
End Sub

Private Function Check_CarData() As Boolean
'�����򥻸���ˮ�
Check_CarData = False
msg_text = ""
If Len(Trim(txt_Tab1_CarID.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [���P���X]"
   Else
      msg_text = msg_text & vbCrLf & "����J���P���X"
   End If
End If

'����ˮ�
If Len(RTrim(cmb_Tab1_CarType)) = 0 Then msg_text = msg_text & vbCrLf & "�����""�p�O���O""�I"
If Len(RTrim(cmb_Tab1_Company)) = 0 Then msg_text = msg_text & vbCrLf & "�����""�f�B���q""�I"
If Len(RTrim(txt_Tab1_WeightCapacity)) = 0 Then msg_text = msg_text & vbCrLf & "����J""�i�˸����q""�I"
If Len(RTrim(txt_Tab1_VolumnCapacity)) = 0 Then msg_text = msg_text & vbCrLf & "����J""�i�˸����n""�I"
If IsNumeric(txtAPFix) = False Then msg_text = msg_text & vbCrLf & "����J�ή榡���~""�B�O�վ�%""�I"
If Val(txtAPFix) < 0 Then msg_text = msg_text & vbCrLf & "���o���t��""�B�O�վ�%""�I"

If msg_text = "" Then
   Check_CarData = True
Else
   msg_text = "������Ʋ��`�A�Эץ���A���� [�s ��]�G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

End Function

Private Sub Clear_CompanyData()
'�M�� �f�B���q �e��������
txt_Tab2_CompanyCode.Text = ""
txt_Tab2_CName.Text = ""
txt_Tab2_EName.Text = ""
txt_Tab2_Address.Text = ""
txt_Tab2_Descr.Text = ""
txt_Tab2_Contact.Text = ""
txt_Tab2_Phone.Text = ""
txt_Tab2_ShortName.Text = ""
End Sub

Private Sub Clear_SkuData()
'�M�� �f�� �e��������
txt_Tab3_StorerKey.Text = ""
txt_Tab3_Sku.Text = ""
txt_Tab3_DESCR.Text = ""
txt_Tab3_STDGROSSWGT.Text = ""
txt_Tab3_BUSR4.Text = ""
txt_Tab3_SKUGROUP.Text = ""
txt_Tab3_BUSR1.Text = ""
txt_Tab3_NOTES1.Text = ""
txt_Tab3_NOTES2.Text = ""
End Sub


Private Function Check_CompanyData() As Boolean
'�f�B���q�򥻸���ˮ�
Check_CompanyData = False
msg_text = ""
If Len(Trim(txt_Tab2_CompanyCode.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [���q�N�X]"
   Else
      msg_text = msg_text & vbCrLf & "����J���q�N�X"
   End If
End If
If Len(Trim(txt_Tab2_CName.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [���q����W��]"
   Else
      msg_text = msg_text & vbCrLf & "����J [���q����W��]"
   End If
End If

If msg_text = "" Then
   Check_CompanyData = True
Else
   msg_text = "�f�B���q��Ʋ��`�A�Эץ���A���� [�s ��]�G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function

Private Function Check_SkuData() As Boolean
'�f������ˮ�
Check_SkuData = False
msg_text = ""
If Len(Trim(txt_Tab3_Sku.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [�f��]"
   Else
      msg_text = msg_text & vbCrLf & "����J�f��"
   End If
End If
If Len(Trim(txt_Tab3_DESCR.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [����W��]"
   Else
      msg_text = msg_text & vbCrLf & "����J [����W��]"
   End If
End If
If Len(Trim(txt_Tab3_STDGROSSWGT.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [�C�c��]"
   Else
      msg_text = msg_text & vbCrLf & "����J [�C�c��]"
   End If
End If
If Len(Trim(txt_Tab3_BUSR4.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [�C�c��]"
   Else
      msg_text = msg_text & vbCrLf & "����J [�C�c��]"
   End If
End If

If msg_text = "" Then
   Check_SkuData = True
Else
   msg_text = "�ӫ~��Ʋ��`�A�Эץ���A���� [�s ��]�G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function

Public Sub frm_BaseData_Car_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
'��椽�ΰƵ{���A�� frm_RS_FilterAndSort ���I�s
'�ǤJ�ȡGstrCode      �ʧ@�ѧO�X
'                     [FILTER] �ۭq�z��    [SORT] �Ƨ�
'        strReturn    �z�� or �Ƨ� ���]�w�r��

Select Case strCode
       Case "FILTER"  '�ۭq�z��
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TAB0_CONSIGNEELIST"   '�Ȥ�򥻸��
                        blTab0ConsignEventEnable = False
                        rs_Tab0_ConsigneeList.Filter = adFilterNone
                        rs_Tab0_ConsigneeList.Filter = strReturn
                        If rs_Tab0_ConsigneeList.RecordCount = 0 Then
                           msg_text = "��p���A�䤣��ŦX���󪺸�Ƴ�"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab0_ConsigneeList.Filter = adFilterNone
                           rs_Tab0_ConsigneeList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
                           blTab0ConsignEventEnable = True
                           Exit Sub
                        End If
                        blTab0ConsignEventEnable = True
                   Case "RS_TAB1_CARLIST"         '�����򥻸��
                        blTab1CarEventEnable = False
                        rs_Tab1_CarList.Filter = adFilterNone
                        rs_Tab1_CarList.Filter = strReturn
                        If rs_Tab1_CarList.RecordCount = 0 Then
                           msg_text = "��p���A�䤣��ŦX���󪺸�Ƴ�"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab1_CarList.Filter = adFilterNone
                           rs_Tab1_CarList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
                           blTab1CarEventEnable = True
                           Exit Sub
                        End If
                        blTab1CarEventEnable = True
                   Case "RS_TAB2_TRPCOMPANYLIST"         '�B�餽�q�򥻸��
                        blTab2CompanyEventEnable = False
                        rs_Tab2_TRPCompanyList.Filter = adFilterNone
                        rs_Tab2_TRPCompanyList.Filter = strReturn
                        If rs_Tab2_TRPCompanyList.RecordCount = 0 Then
                           msg_text = "��p���A�䤣��ŦX���󪺸�Ƴ�"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab2_TRPCompanyList.Filter = adFilterNone
                           rs_Tab2_TRPCompanyList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
                           blTab2CompanyEventEnable = True
                           Exit Sub
                        End If
                        blTab2CompanyEventEnable = True
                    Case "RS_TAB4_ACCEPTABLELIST"   '�Ȥ᤹���Ѽư򥻸��
                        blTab4AcceptableEventEnable = False
                        rs_Tab4_AcceptableList.Filter = adFilterNone
                        rs_Tab4_AcceptableList.Filter = strReturn
                        If rs_Tab4_AcceptableList.RecordCount = 0 Then
                           msg_text = "��p���A�䤣��ŦX���󪺸�Ƴ�"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab4_AcceptableList.Filter = adFilterNone
                           rs_Tab4_AcceptableList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
                           blTab4AcceptableEventEnable = True
                           Exit Sub
                        End If
                        blTab4AcceptableEventEnable = True
                    Case "RS_TAB3_SKULIST"   '�f���򥻸��
                        blTab3skuEventEnable = False
                        rs_Tab3_SkuList.Filter = adFilterNone
                        rs_Tab3_SkuList.Filter = strReturn
                        If rs_Tab3_SkuList.RecordCount = 0 Then
                           msg_text = "��p���A�䤣��ŦX���󪺸�Ƴ�"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab3_SkuList.Filter = adFilterNone
                           rs_Tab3_SkuList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
                           blTab3skuEventEnable = True
                           Exit Sub
                        End If
                        blTab3skuEventEnable = True
            End Select
       Case "SORT"    '�Ƨ�
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TAB0_CONSIGNEELIST"   '�Ȥ�򥻸��
                        blTab0ConsignEventEnable = False
                        rs_Tab0_ConsigneeList.Sort = strReturn
                        blTab0ConsignEventEnable = True
                   Case "RS_TAB1_CARLIST"        '�����򥻸��
                        blTab1CarEventEnable = False
                        rs_Tab1_CarList.Sort = strReturn
                        blTab1CarEventEnable = True
                   Case "RS_TAB2_TRPCOMPANYLIST"    '�f�B���q�򥻸��
                        blTab2CompanyEventEnable = False
                        rs_Tab2_TRPCompanyList.Sort = strReturn
                        blTab2CompanyEventEnable = True
                   Case "RS_TAB4_ACCEPTABLELIST"   '�Ȥ᤹���Ѽư򥻸��
                        blTab4AcceptableEventEnable = False
                        rs_Tab4_AcceptableList.Sort = strReturn
                        blTab4AcceptableEventEnable = True
                   Case "RS_TAB3_SKULIST"   '�f���򥻸��
                        blTab3skuEventEnable = False
                        rs_Tab3_SkuList.Sort = strReturn
                        blTab3skuEventEnable = True
            End Select
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_Tab1_CarID_LostFocus()

Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select vehicle_id_no from trp09m where rtrim(vehicle_id_no) = '" & RTrim(txt_Tab1_CarID) & "' ", cn
If rsTmp.EOF = False Then
    MsgBox "�s�W��������!!", 64, "�`�N"
       txt_Tab1_CarID.SelStart = 0: txt_Tab1_CarID.SelLength = Len(txt_Tab1_CarID.Text)
   txt_Tab1_CarID.SetFocus

End If

End Sub
Sub SendMail(strConsigneeKey As String)

'LTKK01�Ȥ�D�ɲ��ʦ۰� Mail �q��
'Gary edit strto 20170424 irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;pinkhsu@mail.kirin.com.tw
If mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
    
    Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String
    
    'Ū��ini�Ѽ�
    Dim objIni As New vbIniFile
    objIni.FileName = App.Path & "/" & App.title & ".ini"
    
    strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
    strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
    strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
    strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
    strSubject = "�Ȥ�D�ɲ���(" & strConsigneeKey & "-" & txt_Tab0_ShortName & ")"
    strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
    strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
    strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
    strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")
    
    
    strFrom = "Tkedi@bestlog.com.tw"
    strTo = "ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;pinkhsu@mail.kirin.com.tw"
    strCC = "Tkedi@bestlog.com.tw"
'     strTo = "gemini@bestlog.com.tw"
     strCC = ""
    
    Set objIni = Nothing
    
    Dim rsTmp As New ADODB.Recordset
    
    If Len(RTrim(strFrom)) > 0 Then '���H���
    
        str_SQL = "select * from gv_webCustomer where �Ȥ�s�� + �a�}�O = '" & strConsigneeKey & "' "

        rsTmp.Open str_SQL, cn
        
        '�p�G�L��Ƥ]�nmail
        If Not rsTmp.EOF Or UCase(RTrim(strAlways)) = "YES" Then
            
            strAddAttachment = "C:\BEST\DYDC_Best\LTKK01\�Ȥ�D�ɲ���\�Ȥ�D�ɲ���_" & Format(Now, "yyyymmddhhMMss") & ".xls"
            
            Call Recordset2Excel("�Ȥ�D�ɲ���", rsTmp)
            If Dir("C:\BEST\DYDC_Best\LTKK01\�Ȥ�D�ɲ���", vbDirectory) = "" Then MkDirs "C:\BEST\DYDC_Best\LTKK01\�Ȥ�D�ɲ���"
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
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            'SMTP ���A���ݭn���Ү�
            If Len(RTrim(strEmailID)) > 0 Then
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
            End If
            objEmail.Configuration.Fields.Update
            objEmail.Send
        
            MsgBox "LTKK01�Ȥ�D�ɲ��ʡA�t�Τw�oMail�q���f�D!", , "�۰�Mail�q��"
        
            Set objEmail = Nothing
        End If
    End If
End If

End Sub


Private Sub cmd_Tab4_AcceptableReset_Click()
'�Ȥ᤹���Ѽư򥻸�� >> �����z��Ƨ�
'�����z�����A���]�ƧǨ̾�
If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
 blTab4AcceptableEventEnable = False
 rs_Tab4_AcceptableList.Filter = adFilterNone
 rs_Tab4_AcceptableList.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
 blTab4AcceptableEventEnable = True
End Sub

Private Sub cmd_Tab4_AcceptableShow_Click()
'�����ѼƸ�� >> ��ܩҦ��Ȥ�+�f��
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab4_AcceptableList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab4_AcceptableList)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select distinct Rtrim(a.storerkey) as �f�D,Rtrim(a.Customer) as �Ȥ�s��," & _
" (select top 1 Rtrim(t1m.Full_Name) from TRP01M t1m where Left(t1m.ConsigneeKey,8)=a.customer order by  Rtrim(t1m.Full_Name) desc)  as �Ȥ�W��,Rtrim(a.ItemNo) as ���~�s��,Rtrim(s.descr) as ���~�W��,a.allowdays as �����Ѽ�" & _
" from Acceptable a " & _
"inner join " & strWMSDB & "..sku s on s.sku=a.itemno " & _
"order by Rtrim(a.Customer),Rtrim(a.ItemNo) "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    msg_text = "��ƿ��~�G�d�ߵ��G�Ǧ^ 0 �C�Ȥ᤹���ѼƸ��"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Set rs_Tab4_AcceptableList = Nothing
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab4_AcceptableList)
tmp_Rs.Close

blTab4AcceptableEventEnable = False
With dg_Tab4_AcceptableList
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Tab4_AcceptableList.MoveFirst
Set dg_Tab4_AcceptableList.DataSource = rs_Tab4_AcceptableList
With dg_Tab4_AcceptableList
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 800        '�f�D
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '�Ȥ�s��
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '�Ȥ�W��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000       '���~�s��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 3000       '���~�W��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '�����Ѽ�
    .Columns(6).Alignment = dbgLeft
End With
blTab4AcceptableEventEnable = True
Call Clear_AcceptableData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�Ȥ᤹���ѼƸ��-��ܩҦ����", Me.Caption, "cmd_Tab4_AcceptableShow_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_AddNew_Click()
'�Ȥ᤹���ѼƸ�� >> �ഫ�ܷs�W�Ҧ�
If Not rs_Tab4_AcceptableList Is Nothing Then
    If dg_Tab4_AcceptableList.SelBookmarks.Count > 0 Then dg_Tab4_AcceptableList.SelBookmarks.Remove 0
End If
fam_Tab4_Acceptable.BackColor = &HC0FFC0
fam_Tab4_Acceptable.Enabled = True
txt_Tab4_ConsigneeKey.Enabled = True
txt_Tab4_Sku.Enabled = True
Call Clear_AcceptableData
cmd_Tab4_Save.Enabled = True
cmd_Tab4_Cancel.Enabled = True
cmd_Tab4_AddNew.Enabled = False
cmd_Tab4_Modify.Enabled = False
cmd_Tab4_Delete.Enabled = False
cmb_Tab4_Storer.ListIndex = 0
txt_Tab4_ConsigneeKey.BackColor = &HFFFFFF
txt_Tab4_Sku.BackColor = &HFFFFFF
dg_Tab4_AcceptableList.Enabled = False
cmd_Tab4_AcceptableShow.Enabled = False
cmd_Tab4_AcceptableQuery.Enabled = False
cmd_Tab4_AcceptableReset.Enabled = False
End Sub

Private Sub cmd_Tab4_Cancel_Click()
'�Ȥ᤹���ѼƸ�� >> �����ק�
Call Clear_AcceptableData
If txt_Tab4_ConsigneeKey.Enabled = False And txt_Tab4_Sku.Enabled = False Then
    If Not rs_Tab4_AcceptableList Is Nothing Then
        dg_Tab4_AcceptableList.SelBookmarks.Add rs_Tab4_AcceptableList.Bookmark
        Call Display_SelectedAcceptableData(rs_Tab4_AcceptableList.Fields("�Ȥ�s��").Value, rs_Tab4_AcceptableList.Fields("���~�s��").Value)
    End If
End If
fam_Tab4_Acceptable.BackColor = &H8000000C
fam_Tab4_Acceptable.Enabled = False
cmd_Tab4_Cancel.Enabled = False
cmd_Tab4_Save.Enabled = False
cmd_Tab4_AddNew.Enabled = True
cmd_Tab4_Modify.Enabled = True
cmd_Tab4_Delete.Enabled = True
txt_Tab4_ConsigneeKey.BackColor = &H8000000F
txt_Tab4_Sku.BackColor = &H8000000F
dg_Tab4_AcceptableList.Enabled = True
cmd_Tab4_AcceptableShow.Enabled = True
cmd_Tab4_AcceptableQuery.Enabled = True
cmd_Tab4_AcceptableReset.Enabled = True
End Sub

Private Sub cmd_Tab4_Delete_Click()
'�Ȥ᤹���ѼƸ�� >> �R��
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

If Len(RTrim(txt_Tab4_ConsigneeKey.Text)) = 0 Or Len(RTrim(txt_Tab4_Sku.Text)) = 0 Then
   msg_text = "�п�ܱ��R�����Ȥ���"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'�ˮ֬O�_���\�i��R���X�Э�
If blDelete = False Then
   msg_text = "�Ȥ��ƵL�k�R���G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
Else
   Dim CheckDelete As Integer
   msg_text = "�T�w�n�R�������Ȥ᤹���ѼƸ�ơH"
   CheckDelete = MsgBox(msg_text, vbOKCancel + vbQuestion, msg_title)
End If

If CheckDelete = 1 Then
    '���\�R��
    str_SQL = "Delete From Acceptable Where Customer = '" & Trim(txt_Tab4_ConsigneeKey.Text) & "' and ItemNo='" & Trim(txt_Tab4_Sku.Text) & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
Else
    Screen.MousePointer = vbDefault
    Exit Sub
End If

fam_Tab4_Acceptable.BackColor = &H8000000C
fam_Tab4_Acceptable.Enabled = False
cmd_Tab4_Cancel.Enabled = False
cmd_Tab4_Save.Enabled = False
cmd_Tab4_AddNew.Enabled = True
cmd_Tab4_Modify.Enabled = False
cmd_Tab4_Delete.Enabled = False
txt_Tab4_ConsigneeKey.BackColor = &H8000000F
txt_Tab4_Sku.BackColor = &H8000000F

'���s��ܩҦ��Ȥ᤹���ѼƸ��
Call cmd_Tab4_AcceptableShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�Ȥ᤹���ѼƸ��-�R��", Me.Caption, "cmd_Tab4_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_Modify_Click()
'�Ȥ᤹���ѼƸ�� >> �ഫ�ק�Ҧ�
'�T�{����Ȥ᤹���ѼƸ�Ƥ褹�\ [�ק�] �\��
If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
If dg_Tab4_AcceptableList.SelBookmarks.Count <> 0 Then
   fam_Tab4_Acceptable.BackColor = &HC0E0FF
   fam_Tab4_Acceptable.Enabled = True
   txt_Tab4_ConsigneeKey.Enabled = False
   txt_Tab4_Sku.Enabled = False
   cmd_Tab4_Save.Enabled = True
   cmd_Tab4_Cancel.Enabled = True
   cmd_Tab4_AddNew.Enabled = False
   cmd_Tab4_Modify.Enabled = False
   cmd_Tab4_Delete.Enabled = False
   txt_Tab4_ConsigneeKey.BackColor = &H8000000F
   txt_Tab4_Sku.BackColor = &H8000000F
End If
End Sub



Private Sub cmd_Tab4_Save_Click()
'�Ȥ᤹���ѼƸ�� >> �Ȥ᤹���ѼƸ�Ʀs��
On Error GoTo err_Handle

Dim rsTmp As New ADODB.Recordset

'�Ȥ�s�� �ˬd
If txt_Tab4_ConsigneeKey.Enabled = True And Len(Trim(txt_Tab4_Sku.Text)) <> 0 Then
    rsTmp.Open "select consigneekey,Rtrim(Isnull(Full_Name,'')) as full_name from trp01m where Left(rtrim(consigneekey),8) = '" & Trim(txt_Tab4_ConsigneeKey.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = True Then
       MsgBox "�u�Ȥ�s���v���s�b�t�ΡA�нT�{���!!", 64, "�`�N"
           txt_Tab4_ConsigneeKey.SelStart = 0: txt_Tab4_ConsigneeKey.SelLength = Len(txt_Tab4_ConsigneeKey.Text)
       txt_Tab4_ConsigneeKey.SetFocus
       rsTmp.Close
       Exit Sub
    Else
       rsTmp.Close
    End If
End If

'���~�s�� �ˬd
If txt_Tab4_Sku.Enabled = True And Len(Trim(txt_Tab4_Sku.Text)) <> 0 Then
    rsTmp.Open "select Sku,Rtrim(Isnull(DESCR,'')) as DESCR from " & strWMSDB & "..Sku where rtrim(Sku) = '" & Trim(txt_Tab4_Sku.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = True Then
       MsgBox "�u���~�s���v���s�b�t�ΡA�нT�{���!!", 64, "�`�N"
           txt_Tab4_Sku.SelStart = 0: txt_Tab4_Sku.SelLength = Len(txt_Tab4_Sku.Text)
       txt_Tab4_Sku.SetFocus
       rsTmp.Close
       Exit Sub
    Else
       rsTmp.Close
    End If
End If

'�Ȥ�s��+���~�s�� �����ˬd
If txt_Tab4_ConsigneeKey.Enabled = True And txt_Tab4_Sku.Enabled = True Then
    rsTmp.Open "select customer,itemno from Acceptable where rtrim(customer) = '" & Trim(txt_Tab4_ConsigneeKey.Text) & "' and rtrim(itemno)='" & Trim(txt_Tab4_Sku.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = False Then
        MsgBox "�P�@�f�D�A�s�W�u�Ȥ�s��+���~�s���v����!!", 64, "�`�N"
           txt_Tab4_ConsigneeKey.SelStart = 0: txt_Tab4_ConsigneeKey.SelLength = Len(txt_Tab4_ConsigneeKey.Text)
       txt_Tab4_ConsigneeKey.SetFocus
       rsTmp.Close
       Exit Sub
    Else
        rsTmp.Close
    End If
End If

'�s�ɸ���ˮ�
If Check_AcceptableData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '����ɶ��]�w�G�L��������
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_AcceptableData_Update"

'�f�D
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("StorerKey").Value = arStorer(cmb_Tab4_Storer.ListIndex)

'�Ȥ�s��
Set tmp_para = tmp_Cmd.CreateParameter("ConsigneeKey", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_Tab4_ConsigneeKey.Text)

'���~�s��
Set tmp_para = tmp_Cmd.CreateParameter("SKU", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("SKU").Value = Trim(txt_Tab4_Sku.Text)

'�����Ѽ�
Set tmp_para = tmp_Cmd.CreateParameter("AllowDays", adInteger, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("AllowDays").Value = Trim(txt_Tab4_AllowDays.Text)

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'�D�P�B����
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
Loop
Set tmp_Cmd = Nothing

fam_Tab4_Acceptable.BackColor = &H8000000C
fam_Tab4_Acceptable.Enabled = False
cmd_Tab4_Cancel.Enabled = False
cmd_Tab4_Save.Enabled = False
cmd_Tab4_AddNew.Enabled = True
cmd_Tab4_Modify.Enabled = False
cmd_Tab4_Delete.Enabled = False
dg_Tab4_AcceptableList.Enabled = True
cmd_Tab4_AcceptableShow.Enabled = True
cmd_Tab4_AcceptableQuery.Enabled = True
cmd_Tab4_AcceptableReset.Enabled = True

'���s��ܩҦ��Ȥ���
Call cmd_Tab4_AcceptableShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�Ȥ᤹���ѼƸ��-�s��", Me.Caption, "cmd_Tab4_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_SaveToExcel_Click()
'�d�ߵ��G>> �� EXCEL
If blTab4AcceptableEventEnable = False Then
    msg_text = "�L��Ƥ�����Excel��I"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
    
    If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
    If rs_Tab4_AcceptableList.RecordCount = 0 Then Exit Sub
    
blTab4AcceptableEventEnable = False '�קK�C���ƿ��

    rs_Tab4_AcceptableList.MoveFirst
    
    Recordset2ExcelV2 "�����ѼƸ��", "�����ѼƸ��", rs_Tab4_AcceptableList
    
    Set MyXlsAppV2 = Nothing
    
blTab4AcceptableEventEnable = True
End Sub

Private Sub cmdOpenFilesT5_Click()

On Error GoTo err_Handle
gd_Tab5_AaccessTable.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, str_soldcode As String, Str_Sku As String, str_allowdays As Integer
'�T�{���|�O�_�a"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If
'�إ����W�ٰ}�C
strFieldName = ""
If Right(filLocalFileT5.Path, 1) <> "\" Then
    strFilePath = filLocalFileT5.Path & "\"
Else
    strFilePath = filLocalFileT5.Path
End If
Set rsMain = New ADODB.Recordset
strSheetName = "DATA"
Call Excel2Recordset(strFilePath & filLocalFileT5.FileName, strSheetName, strFieldName, rsMain)
rsMain.MoveFirst
Set gd_Tab5_AaccessTable.DataSource = rsMain

'�YAcceptableTemp������table�R����
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "if object_id ('AcceptableTemp') is not null drop TABLE  AcceptableTemp"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If rsMain.EOF = False Then
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "CREATE TABLE AcceptableTemp ( ItemNo char(20) , Customer varchar(20) , AllowDays int, Storerkey char(15) );"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    End If
    Do While Not rsMain.EOF
        gd_Tab5_AaccessTable.Col = 0: str_storekey = Trim(gd_Tab5_AaccessTable.Text)
        gd_Tab5_AaccessTable.Col = 1: str_soldcode = Trim(gd_Tab5_AaccessTable.Text)
        gd_Tab5_AaccessTable.Col = 2: Str_Sku = Trim(gd_Tab5_AaccessTable.Text)
        gd_Tab5_AaccessTable.Col = 3: str_allowdays = Trim(gd_Tab5_AaccessTable.Text)
        str_SQL = "INSERT INTO AcceptableTemp (storerkey,Customer,ItemNo,allowdays) VALUES ('" & str_storekey & "','" & str_soldcode & "','" & Str_Sku & "','" & str_allowdays & "')"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        rsMain.MoveNext
Loop
rsMain.MoveFirst

'

If rsMain Is Nothing Then
    MsgBox "�d�L���!", 64, "Excel2Recordset"
Else
    SetDataGridColWidth Me.Caption, gd_Tab5_AaccessTable
    MsgBox "���u�@��@ " & rsMain.RecordCount & "����ơA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    cmdImportT5.Enabled = True
End If

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdImportT5_Click()

    On Error GoTo err_Handle
    gd_Tab5_AaccessTable.Enabled = False: cmdImportT5.Enabled = False
    Dim existmark As Integer
    
    Call Confirm_Recordset_Closed(tmp_Rs)
'    '�ˬd�f�D�O�_�OLCHF01
'    str_SQL = "select * from AcceptableTemp where Storerkey <> 'LCHF01'"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_Rs.EOF Then
'       msg_text = "�f�D����LCHF01:" & Trim(tmp_Rs.Fields("Customer").Value) & "," & Trim(tmp_Rs.Fields("ItemNo").Value) & "�A�нT�{��ơA���¡C"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
'        tmp_Rs.Close
'        Exit Sub
'    End If
    
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    '�ˬd�Ȥ�s���O�_�s�b
    str_SQL = "select  customer,* from AcceptableTemp where customer NOT in (select  distinct Left(rtrim(consigneekey),8) from trp01m where storerkey='LCHF01')"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
       msg_text = "���Ȥ�s��(soldcode):" & Trim(tmp_Rs.Fields("Customer").Value) & "���s�b�A�нT�{��ơA���¡C"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
        tmp_Rs.Close
        Exit Sub
    End If

    Call Confirm_Recordset_Closed(tmp_Rs)
    '�ˬdsku�O�_�s�b
    str_SQL = "select  ItemNo,* from AcceptableTemp where  ItemNo NOT in (select  distinct sku from " & strWMSDB & "..sku where storerkey='LCHF01')"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
       msg_text = "���~��(sku):" & Trim(tmp_Rs.Fields("ItemNo").Value) & "���s�b�A�нT�{��ơA���¡C"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
        tmp_Rs.Close
        Exit Sub
    End If

'    Call Confirm_Recordset_Closed(tmp_Rs)
'    '�ˬd�O�_��allowdays<=0
'    str_SQL = "select * from AcceptableTemp where allowdays=0"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_Rs.EOF Then
'       msg_text = "�������0:" & Trim(tmp_Rs.Fields("Customer").Value) & "," & Trim(tmp_Rs.Fields("ItemNo").Value) & "�A�нT�{��ơA���¡C"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
'        tmp_Rs.Close
'        Exit Sub
'    End If
    

    Call Confirm_Recordset_Closed(tmp_Rs)
    '�ˬd�O�_������soldcode+sku
    str_SQL = "select * from AcceptableTemp where ItemNo+Customer in (select ItemNo+Customer from AcceptableTemp group by Customer,ItemNo having count(*)>1) order by ItemNo+Customer "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        msg_text = "�����Ъ�SoldCode+SKU:" & Trim(tmp_Rs.Fields("Customer").Value) & "," & Trim(tmp_Rs.Fields("ItemNo").Value) & "�A�нT�{��ơA���¡C"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
        tmp_Rs.Close
        Exit Sub
    End If
'    Call Confirm_Recordset_Closed(tmp_rs)
'      str_SQL = "select * from  Acceptable where RTRim(Customer)+Rtrim(ItemNo) in (select Rtrim(u.customer)+Rtrim(u.ItemNo) from AcceptableTemp u inner join Acceptable a on u.customer=a.Customer and u.ItemNo=a.ItemNo)"
'    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_rs.EOF Then
'        msg_text = "�����Ъ�SoldCode+SKU2�A�нT�{��ơA���¡C"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
'        tmp_rs.Close
'        Exit Sub
'    End If
    existmark = 0
    '�ˬd�O�_���w�s�bAcceptable�����
    Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select a.*,u.* from AcceptableTemp u inner join Acceptable a on u.customer=a.Customer and u.ItemNo=a.ItemNo"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        existmark = 1
        tmp_Rs.Close
        'Exit Sub
    End If


        Tran_Level = cn.BeginTrans
        If existmark = 1 Then
            str_SQL = "delete Acceptable where RTRim(Customer)+Rtrim(ItemNo) in (select Rtrim(u.Customer)+Rtrim(u.ItemNo) from AcceptableTemp u inner join Acceptable a on u.Customer=a.Customer and u.ItemNo=a.ItemNo)"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
        str_SQL = "Insert Acceptable(Storerkey, Customer, ItemNo, AllowDays,addwho,adddate,editwho,editdate) select Storerkey,Customer,ItemNo,AllowDays,'" & User_id & "' , getdate() , '" & User_id & "' , getdate() from AcceptableTemp  where  ItemNo in (select  distinct sku from " & strWMSDB & "..sku where storerkey='LCHF01')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        cn.CommitTrans: Tran_Level = 0
        '�ƥ��ɮ�
        Dim fl_file As Scripting.File
        Set fso = New FileSystemObject
        Dim strExcelFileName As String, str_ArPath As String, str_file As String, arrstr_file
        strExcelFileName = filLocalFileT5.Path & "\" & filLocalFileT5.FileName
        If fso.FileExists(strExcelFileName) = True Then
        str_ArPath = "D:\LCHF01\AcceptTable\"
        str_file = filLocalFileT5.FileName
            If Dir(str_ArPath, vbDirectory) = "" Then MkDirs str_ArPath
            Set fl_file = fso.GetFile(strExcelFileName) '���ɮ׸��|
            arrstr_file = Split(str_file, ".")
            fl_file.copy (str_ArPath & arrstr_file(0) & "_" & Format(Now, "YYMMDDHHMMSS") & ".xls")

            If fso.FileExists(str_ArPath & arrstr_file(0) & "_" & Format(Now, "YYMMDDHHMMSS") & ".xls") = True Then
                    fl_file.Delete
            End If

        End If
        Set rsMain = Nothing
        filLocalFileT5.Refresh
        msg_text = "�פJ���\�A�ɮ׳ƥ���D:\LCHF01\AcceptTable�C"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
    Exit Sub

err_Handle:
    Dim tmpString As String
    gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
    If Tran_Level = 1 Then cn.RollbackTrans
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "�����ѼƶפJ-�s��", Me.Caption, "cmd_Tab0_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault

End Sub
Private Sub Display_SelectedAcceptableData(ByVal strConsigneeKey As String, strSku As String)
'��ܶǤJ���Ȥ᤹���ѼƸ��
Call Clear_AcceptableData

str_SQL = "select distinct Rtrim(a.storerkey) as �f�D,Rtrim(a.Customer) as �Ȥ�s��" & _
",(select top 1 Rtrim(t1m.Full_Name) from TRP01M t1m where Left(t1m.ConsigneeKey,8)=a.customer order by  Rtrim(t1m.Full_Name) desc)  as �Ȥ�W��,Rtrim(a.ItemNo) as ���~�s��,Rtrim(s.descr) as ���~�W��,a.allowdays as �����Ѽ� " & _
" from Acceptable a " & _
"inner join " & strWMSDB & "..sku s on s.sku=a.itemno " & _
"Where Customer = '" & strConsigneeKey & "' and ItemNo='" & strSku & "'" & _
" order by Rtrim(a.Customer),Rtrim(a.ItemNo) "

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫Ȥ᤹���Ѽư򥻸��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
For i = 0 To cmb_Tab4_Storer.ListCount - 1
    If arStorer(i) = Trim(tmp_Rs.Fields("�f�D").Value) Then
       cmb_Tab4_Storer.ListIndex = i
       Exit For
    End If
Next i

txt_Tab4_ConsigneeKey.Text = Trim(tmp_Rs.Fields("�Ȥ�s��").Value)
txt_Tab4_FullName.Text = IIf(IsNull(tmp_Rs.Fields("�Ȥ�W��").Value), "", Trim(tmp_Rs.Fields("�Ȥ�W��").Value))
txt_Tab4_Sku.Text = IIf(IsNull(tmp_Rs.Fields("���~�s��").Value), "", Trim(tmp_Rs.Fields("���~�s��").Value))
txt_Tab4_DESCR.Text = IIf(IsNull(tmp_Rs.Fields("���~�W��").Value), "", Trim(tmp_Rs.Fields("���~�W��").Value))
txt_Tab4_AllowDays.Text = IIf(IsNull(tmp_Rs.Fields("�����Ѽ�").Value), "", Trim(tmp_Rs.Fields("�����Ѽ�").Value))

tmp_Rs.Close

End Sub

Private Sub Clear_AcceptableData()
'�M�� �Ȥ᤹���ѼƸ�� �e��������
cmb_Tab4_Storer.ListIndex = -1
txt_Tab4_ConsigneeKey.Text = ""
txt_Tab4_FullName.Text = ""
txt_Tab4_Sku.Text = ""
txt_Tab4_DESCR.Text = ""
txt_Tab4_AllowDays.Text = ""
txt_Tab4_ConsigneeKey.BackColor = &H8000000F
txt_Tab4_Sku.BackColor = &H8000000F
End Sub
Private Function Check_AcceptableData() As Boolean
'�Ȥ᤹���Ѽư򥻸���ˮ�
Check_AcceptableData = False
msg_text = ""
If cmb_Tab4_Storer.ListIndex = -1 Then
   If msg_text = "" Then
      msg_text = "����J [�f�D]"
   Else
      msg_text = msg_text & vbCrLf & "����J [�f�D]"
   End If
End If
If Len(Trim(txt_Tab4_ConsigneeKey.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [�Ȥ�s��]"
   Else
      msg_text = msg_text & vbCrLf & "����J [�Ȥ�s��]"
   End If
End If
If Len(Trim(txt_Tab4_Sku.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [���~�s��]"
   Else
      msg_text = msg_text & vbCrLf & "����J [���~�s��]"
   End If
End If
If Len(Trim(txt_Tab4_AllowDays.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J [�����Ѽ�]"
   Else
      msg_text = msg_text & vbCrLf & "����J [�����Ѽ�]"
   End If
End If

If msg_text = "" Then
   Check_AcceptableData = True
Else
   msg_text = "�Ȥ᤹���ѼƸ�Ʋ��`�A�Эץ���A���� [�s ��]�G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function
Private Sub txt_Tab4_ConsigneeKey_LostFocus()
'�Ȥ�s�� �ˬd
If txt_Tab4_ConsigneeKey.Enabled = True And Len(Trim(txt_Tab4_ConsigneeKey.Text)) <> 0 Then
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Open "select top 1 consigneekey,Rtrim(Isnull(Full_Name,'')) as full_name from trp01m where Left(rtrim(consigneekey),8) = '" & RTrim(txt_Tab4_ConsigneeKey.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' order by Rtrim(Isnull(Full_Name,'')) desc", cn
    If rsTmp.EOF = True Then
       MsgBox "�u�Ȥ�s���v���s�b�t�ΡA�нT�{���!!", 64, "�`�N"
           txt_Tab4_ConsigneeKey.SelStart = 0: txt_Tab4_ConsigneeKey.SelLength = Len(txt_Tab4_ConsigneeKey.Text)
       txt_Tab4_FullName.Text = ""
       txt_Tab4_ConsigneeKey.SetFocus
       Exit Sub
    Else
       txt_Tab4_FullName.Text = rsTmp.Fields("full_name").Value
    End If
rsTmp.Close
End If
End Sub
Private Sub dirLocalDirT5_Change()
    filLocalFileT5.Path = dirLocalDirT5.Path
End Sub

Private Sub txt_Tab4_Sku_LostFocus()
'���~�s�� �ˬd
If txt_Tab4_Sku.Enabled = True And Len(Trim(txt_Tab4_Sku.Text)) <> 0 Then
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Open "select Sku,Rtrim(Isnull(DESCR,'')) as DESCR from " & strWMSDB & "..Sku where rtrim(Sku) = '" & RTrim(txt_Tab4_Sku.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = True Then
       MsgBox "�u���~�s���v���s�b�t�ΡA�нT�{���!!", 64, "�`�N"
           txt_Tab4_Sku.SelStart = 0: txt_Tab4_Sku.SelLength = Len(txt_Tab4_Sku.Text)
       txt_Tab4_DESCR.Text = ""
       txt_Tab4_Sku.SetFocus
       Exit Sub
    Else
       txt_Tab4_DESCR.Text = rsTmp.Fields("DESCR").Value
    End If
rsTmp.Close
End If
End Sub

Private Sub txt_Tab4_AllowDays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack        '0 - 9,BACKSPACE�B�z
'        Case vbKeyDelete, vbKeyDecimal          '�p���I�B�z
'            If InStr(1, txt_Tab4_AllowDays.Text, ".") <> 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Sub Recordset2ExcelV2(str As String, FileName As String, rs As Object)
On Error GoTo err_Handle
If rs Is Nothing Then MsgBox "�L��ƥi�����ɡI", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String
Dim tmp_col As Double, tmp_row As Double
Dim tmp_letter As String, tmp_RangNo As String, tmpI As Integer

    Dim ExcelTitle As String, saved As Integer, msg_text As String
    msg_text = ""
    saved = 0
    Call DocStoreDirectory(strDocPath)

    Dim strTranFileName As String           'Excel �ɮצW��
    CmnDialog.DialogTitle = "��s Excel ��"
    CmnDialog.InitDir = "C:\my documents"
    CmnDialog.FileName = FileName & "-" & Format(Now, "YYYYMMDDHHNNSS")
    CmnDialog.Filter = "Excel�ɮ�(*.xls)|*.xls"
    CmnDialog.FilterIndex = 1
    CmnDialog.CancelError = True
    CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
    On Error Resume Next
    CmnDialog.ShowSave

    If err.Number = cdlCancel Then          '�� [�s��] ��ܤ�����A���U [����] �s
       msg_text = "�z��� [����] ���s�A������ Excel ���ۦ�s�ɡI"
       MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
       strTranFileName = ""
    Else
       strTranFileName = CmnDialog.FileName
       If Dir(strTranFileName) <> "" Then
          Kill strTranFileName
       End If
    End If

Screen.MousePointer = 11
'�}��EXCEL����
Set MyXlsAppV2 = CreateObject("Excel.Application")

With MyXlsAppV2
    .Visible = False
    
    If Dir(App.Path & "\XLT\" & str & ".xlt") = "" Then '�䤣�쥻���d����
        
        '���d���ɸ��|
        Dim objIni As vbIniFile, arrTmp, strXltPath As String
        Set objIni = New vbIniFile
        
        With objIni
        
            .FileName = striniFileName_FullPath
            strXltPath = RTrim(.ReadData("EXCEL", "XLTPATH", "")) '���䴩�����Ƨ��W��
            
        End With
        Set objIni = Nothing

    End If

    '�L���w���|���ϥνd����
    If Len(RTrim(strXltPath)) = 0 Then strXltPath = App.Path & "\XLT"
    If Dir(strXltPath, vbDirectory) = "" Then GoTo Run
    
    '�d����
    If Dir(strXltPath & "\" & str & ".xlt") <> "" Then
'        If MsgBox("�O�_�ϥνd����?(" & strXltPath & "\" & str & ".xlt), vbQuestion + vbYesNo, "��Excel") = vbNo Then GoTo Run
        
        '�}�ҽd����
        .Workbooks.Open (strXltPath & "\" & str & ".xlt")
        
        '�M��DATA�u�@��
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets("Data").Select: Exit For '��wDATA�u�@��
        Next
        
        '�䤣��s�WDATA�u�@��
        If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then .Sheets.Add: .ActiveSheet.Name = "DATA":
        
        '�j�M�s���x�s��
        For k = 65 To 66 '90
            For j = 1 To 100
                tmp_row = j
                If UCase(.Range(Chr(k) & j).Value) = "BESTLOG" Then GoTo NextStep
                
            Next j
        Next k
        k = 65: j = 1

        '�g�J���D�C
        For i = 0 To rs.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(k + l) & j).Value = rs.Fields(i).Name
            '���W�L26
            If Chr(65 + l) = "Z" Then
                If strCol = "" Then
                    strCol = "A"
                Else
                    strCol = Chr(Asc(strCol) + 1)
                End If
            End If
        Next i
        
NextStep:

        '��Ƽg�J
        '.ActiveSheet.Cells(2, 1).CopyFromRecordset rs
        '.Range(Chr(k) & j).CopyFromRecordset rs
        
'        tmp_row = 2
        Do While Not rs.EOF
            DoEvents
            '�P�_�ϥΪ̬O�_�������ɧ@�~
            For tmp_col = 0 To rs.Fields.Count - 1
                tmp_letter = Chr(65 + tmp_col)      ' A �� ascii code
                If Asc(tmp_letter) > 90 Then        ' > Z �h�ܦ� AA �_�l
                   tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                End If
                tmp_RangNo = tmp_letter & (tmp_row)
                '�]�w�榡
'                With excelAP.Range(tmp_RangNo)
'                    .NumberFormatLocal = "@"      '�x�s��榡 >> �Ʀr >> ���O = ��r
'                    '.Font.Name = "�s�ө���"       '�x�s��榡 >> �r�� >> �r�� = Times New Roman
'                    '.Font.FontStyle = "�з�"      '�x�s��榡 >> �r�� >> �~���˦� = �з�
'                    '.Font.Size = 12               '�x�s��榡 >> �r�� >> �j�p = 12
'                End With
                .Range(tmp_RangNo) = Trim(rs.Fields(tmp_col).Value)
            Next tmp_col
            rs.MoveNext
            tmp_row = tmp_row + 1
        Loop
        
        
    Else '���ϥνd����

Run:
        '�s�WExcel
        .Workbooks.Add: .Sheets("Sheet1").Select: .Sheets("Sheet1").Name = str
        
        '�g�J���D�C
        For i = 0 To rs.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(65 + l) & "1").Value = rs.Fields(i).Name
            '���W�L26
            If Chr(65 + l) = "Z" Then
                If strCol = "" Then
                    strCol = "A"
                Else
                    strCol = Chr(Asc(strCol) + 1)
                End If
            End If
        Next i
        
        '��Ƽg�J
        '.Range("A2").CopyFromRecordset rs
        '.ActiveSheet.Cells(2, 1).CopyFromRecordset rs

        tmp_row = 2
        Do While Not rs.EOF
            DoEvents
            '�P�_�ϥΪ̬O�_�������ɧ@�~
            For tmp_col = 0 To rs.Fields.Count - 1
                tmp_letter = Chr(65 + tmp_col)      ' A �� ascii code
                If Asc(tmp_letter) > 90 Then        ' > Z �h�ܦ� AA �_�l
                   tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                End If
                tmp_RangNo = tmp_letter & (tmp_row)
                '�]�w�榡
'                With excelAP.Range(tmp_RangNo)
'                    .NumberFormatLocal = "@"      '�x�s��榡 >> �Ʀr >> ���O = ��r
'                    '.Font.Name = "�s�ө���"       '�x�s��榡 >> �r�� >> �r�� = Times New Roman
'                    '.Font.FontStyle = "�з�"      '�x�s��榡 >> �r�� >> �~���˦� = �з�
'                    '.Font.Size = 12               '�x�s��榡 >> �r�� >> �j�p = 12
'                End With
                .Range(tmp_RangNo) = Trim(rs.Fields(tmp_col).Value)
            Next tmp_col
            rs.MoveNext
            tmp_row = tmp_row + 1
        Loop
    
    End If
      
        If Len(RTrim(strTranFileName)) > 0 Then
           .ActiveWorkbook.Author = User_id
           .ActiveWorkbook.SaveAs FileName:=strTranFileName, FileFormat:=xlNormal
           .ActiveWindow.Close
           .Visible = False
        Else
           .ActiveWorkbook.Author = User_id
           .Visible = True
        End If

    saved = 1
    
End With

    If Len(RTrim(strTranFileName)) = 0 And saved = 1 Then
       Screen.MousePointer = vbDefault
    ElseIf Len(RTrim(strTranFileName)) > 0 And saved = 1 Then
       Screen.MousePointer = vbDefault
       msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
        
Exit Sub

err_Handle:
   If err.Number = 0 Then Exit Sub
   Dim tmpString As String
   Screen.MousePointer = vbDefault
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--�����Ѽ���EXCEL", Me.Caption, "cmd_SaveToExcel_Tab4", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub dg_Tab4_AcceptableList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�Ȥ᤹���ѼƸ�ƦC��G�����
If blTab4AcceptableEventEnable Then
   If Not rs_Tab4_AcceptableList.EOF Then
      dg_Tab4_AcceptableList.SelBookmarks.Add rs_Tab4_AcceptableList.Bookmark
      Call Display_SelectedAcceptableData(rs_Tab4_AcceptableList.Fields("�Ȥ�s��").Value, rs_Tab4_AcceptableList.Fields("���~�s��").Value)
      fam_Tab4_Acceptable.BackColor = &H8000000C
      fam_Tab4_Acceptable.Enabled = False
      cmd_Tab4_Cancel.Enabled = False
      cmd_Tab4_Save.Enabled = False
      cmd_Tab4_AddNew.Enabled = True
      cmd_Tab4_Modify.Enabled = True
      cmd_Tab4_Delete.Enabled = True
   End If
End If
End Sub

Private Sub cmd_Tab4_AcceptableQuery_Click()
'�Ȥ᤹���ѼƸ�� >> ������Ʒj�M
If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
If rs_Tab4_AcceptableList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab4_AcceptableList"

If ShowForm_RS_FilterAndSort(rs_Tab4_AcceptableList, "�Ȥ᤹���ѼƸ��", Me.Tag) = False Then
    MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
Me.WindowState = vbNormal
End Sub


