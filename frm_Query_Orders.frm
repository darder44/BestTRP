VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Query_Orders 
   Caption         =   "  �q   ��   ��   ��   �d   ��"
   ClientHeight    =   7725
   ClientLeft      =   150
   ClientTop       =   960
   ClientWidth     =   12570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   12570
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   5520
      TabIndex        =   1
      Top             =   4440
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
      StartOfWeek     =   98107393
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7320
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   12912
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "�d�߱���"
      TabPicture(0)   =   "frm_Query_Orders.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�����"
      TabPicture(1)   =   "frm_Query_Orders.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�d�ߵ��G"
      TabPicture(2)   =   "frm_Query_Orders.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "CmnDialog"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dg_Result"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmd_Exit(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmd_Tab2SavetoExcel"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame3 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   11415
         Begin VB.Frame Frame15 
            Height          =   885
            Left            =   120
            TabIndex        =   149
            Top             =   5880
            Width           =   4950
            Begin VB.ComboBox cmb_RSC 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1155
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   151
               Top             =   180
               Width           =   2250
            End
            Begin VB.ComboBox cmb_RBC 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1155
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   150
               Top             =   510
               Width           =   2250
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   34
               Left            =   285
               TabIndex        =   153
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�d���k��"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   35
               Left            =   285
               TabIndex        =   152
               Top             =   570
               Width           =   840
            End
         End
         Begin VB.Frame Frame14 
            Height          =   2595
            Left            =   5115
            TabIndex        =   134
            Top             =   4170
            Width           =   6135
            Begin VB.ComboBox cmb_Zip 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   141
               Top             =   195
               Width           =   2565
            End
            Begin VB.ComboBox cmb_AreaCode 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   140
               Top             =   525
               Width           =   4455
            End
            Begin VB.ComboBox cmb_ExtraDemand 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   139
               Top             =   855
               Width           =   4455
            End
            Begin VB.ComboBox cmb_VehicleType 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   138
               Top             =   1185
               Width           =   4455
            End
            Begin VB.ComboBox cmb_TRPCompany 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   137
               Top             =   1515
               Width           =   4455
            End
            Begin VB.ComboBox cmb_STRPCompany 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   136
               Top             =   2175
               Width           =   4455
            End
            Begin VB.ComboBox cmb_SVehicleType 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               Left            =   1530
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   135
               Top             =   1845
               Width           =   4455
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�l���ϸ�"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   14
               Left            =   660
               TabIndex        =   148
               Top             =   255
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�B�e�ϰ�"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   9
               Left            =   660
               TabIndex        =   147
               Top             =   585
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�S��ݨD"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   660
               TabIndex        =   146
               Top             =   915
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�B�e����"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   660
               TabIndex        =   145
               Top             =   1245
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�f�B���q"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   660
               TabIndex        =   144
               Top             =   1575
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�G���f�B���q"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   32
               Left            =   240
               TabIndex        =   143
               Top             =   2235
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�G���B�e����"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   33
               Left            =   240
               TabIndex        =   142
               Top             =   1905
               Width           =   1260
            End
         End
         Begin VB.Frame Frame13 
            Height          =   1815
            Left            =   120
            TabIndex        =   118
            Top             =   4155
            Width           =   4950
            Begin VB.TextBox txt_SRouteNo_End 
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
               Left            =   3075
               TabIndex        =   125
               Top             =   525
               Width           =   1290
            End
            Begin VB.TextBox txt_SRouteNo_Start 
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
               Left            =   1485
               TabIndex        =   124
               Top             =   525
               Width           =   1290
            End
            Begin VB.TextBox txt_SDeliveryDate_End 
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
               Left            =   3075
               TabIndex        =   123
               Top             =   840
               Width           =   1290
            End
            Begin VB.TextBox txt_SDeliveryDate_Start 
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
               Left            =   1485
               TabIndex        =   122
               Top             =   825
               Width           =   1290
            End
            Begin VB.TextBox txt_SVehicleID 
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
               Left            =   1485
               TabIndex        =   121
               Top             =   1125
               Width           =   1290
            End
            Begin VB.CheckBox chk_SecondPlan 
               Caption         =   "�z��i��G���ƨ�"
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
               Left            =   570
               TabIndex        =   120
               Top             =   225
               Width           =   2355
            End
            Begin VB.TextBox txt_SAddWho 
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
               Left            =   1485
               TabIndex        =   119
               Top             =   1425
               Width           =   1290
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
               Index           =   27
               Left            =   2820
               TabIndex        =   133
               Top             =   555
               Width           =   240
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�G���ƨ����s"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   180
               TabIndex        =   132
               Top             =   570
               Width           =   1260
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
               Index           =   28
               Left            =   2820
               TabIndex        =   131
               Top             =   855
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�G���X�����"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   29
               Left            =   180
               TabIndex        =   130
               Top             =   885
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�G�����P���X"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   30
               Left            =   180
               TabIndex        =   129
               Top             =   1185
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�G���ƨ���"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   31
               Left            =   390
               TabIndex        =   128
               Top             =   1485
               Width           =   1050
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   4
               Left            =   2790
               TabIndex        =   127
               Top             =   1170
               Width           =   240
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   5
               Left            =   2790
               TabIndex        =   126
               Top             =   1470
               Width           =   240
            End
         End
         Begin VB.Frame Frame12 
            Height          =   2370
            Left            =   120
            TabIndex        =   70
            Top             =   1875
            Width           =   8610
            Begin VB.TextBox txt_FReceiptNo 
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
               Left            =   6840
               TabIndex        =   94
               Top             =   1725
               Width           =   1425
            End
            Begin VB.TextBox txt_FAddWho 
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
               Left            =   6840
               TabIndex        =   93
               Top             =   525
               Width           =   1410
            End
            Begin VB.ComboBox cmb_EXEConfirm 
               BackColor       =   &H0080C0FF&
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
               Height          =   315
               ItemData        =   "frm_Query_Orders.frx":0054
               Left            =   6840
               List            =   "frm_Query_Orders.frx":0061
               Style           =   2  '��¤U�Ԧ�
               TabIndex        =   92
               Top             =   195
               Width           =   1440
            End
            Begin VB.TextBox txt_FVehicleID 
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
               Left            =   6840
               TabIndex        =   91
               Top             =   825
               Width           =   1410
            End
            Begin VB.TextBox txt_FDriver 
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
               Left            =   6840
               TabIndex        =   90
               Top             =   1125
               Width           =   1410
            End
            Begin VB.TextBox txt_FDockNo 
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
               Left            =   6840
               TabIndex        =   89
               Top             =   1425
               Width           =   1410
            End
            Begin VB.TextBox txt_FCheckout_End 
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
               Left            =   3240
               TabIndex        =   88
               Top             =   1350
               Width           =   1125
            End
            Begin VB.TextBox txt_FCheckout_Start 
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
               Left            =   1845
               TabIndex        =   87
               Top             =   1350
               Width           =   1125
            End
            Begin VB.TextBox txt_FCheckin_End 
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
               Left            =   3240
               TabIndex        =   86
               Top             =   1050
               Width           =   1125
            End
            Begin VB.TextBox txt_FCheckin_Start 
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
               Left            =   1845
               TabIndex        =   85
               Top             =   1050
               Width           =   1125
            End
            Begin VB.TextBox txt_FPlanCheckin_End 
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
               Left            =   3780
               TabIndex        =   84
               Top             =   1665
               Width           =   1125
            End
            Begin VB.TextBox txt_FPlanCheckin_Start 
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
               Left            =   1845
               TabIndex        =   83
               Top             =   1665
               Width           =   1125
            End
            Begin VB.TextBox txt_FDeliveryDate_End 
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
               Left            =   3240
               TabIndex        =   82
               Top             =   756
               Width           =   1125
            End
            Begin VB.TextBox txt_FDeliveryDate_Start 
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
               Left            =   1845
               TabIndex        =   81
               Top             =   756
               Width           =   1125
            End
            Begin VB.TextBox txt_FPlanDate_End 
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
               Left            =   3240
               TabIndex        =   80
               Top             =   453
               Width           =   1125
            End
            Begin VB.TextBox txt_FPlanDate_Start 
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
               Left            =   1845
               TabIndex        =   79
               Top             =   453
               Width           =   1125
            End
            Begin VB.TextBox txt_FRouteNo_End 
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
               Left            =   3495
               TabIndex        =   78
               Top             =   150
               Width           =   1300
            End
            Begin VB.TextBox txt_FRouteNo_Start 
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
               Left            =   1845
               TabIndex        =   77
               Top             =   150
               Width           =   1300
            End
            Begin VB.TextBox txt_FPlanCheckinTime_Start 
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
               Left            =   2985
               TabIndex        =   76
               Top             =   1665
               Width           =   525
            End
            Begin VB.TextBox txt_FPlanCheckinTime_End 
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
               Left            =   4920
               TabIndex        =   75
               Top             =   1665
               Width           =   525
            End
            Begin VB.TextBox txt_SDNTime_End 
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
               Left            =   4920
               TabIndex        =   74
               Top             =   1950
               Width           =   525
            End
            Begin VB.TextBox txt_SDNTime_Start 
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
               Left            =   2985
               TabIndex        =   73
               Top             =   1950
               Width           =   525
            End
            Begin VB.TextBox txt_SDNDate_Start 
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
               Left            =   1845
               TabIndex        =   72
               Top             =   1950
               Width           =   1125
            End
            Begin VB.TextBox txt_SDNDate_End 
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
               Left            =   3780
               TabIndex        =   71
               Top             =   1950
               Width           =   1125
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
               Height          =   195
               Index           =   4
               Left            =   5940
               TabIndex        =   117
               Top             =   1785
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�ƨ���"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   6150
               TabIndex        =   116
               Top             =   585
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�^�Ǫ��A"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   17
               Left            =   5940
               TabIndex        =   115
               Top             =   270
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���P���X"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   5940
               TabIndex        =   114
               Top             =   885
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�r�p�H"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   6150
               TabIndex        =   113
               Top             =   1185
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�X�Y�Ȧs"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   5940
               TabIndex        =   112
               Top             =   1485
               Width           =   840
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   0
               Left            =   8220
               TabIndex        =   111
               Top             =   555
               Width           =   240
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   2
               Left            =   8220
               TabIndex        =   110
               Top             =   885
               Width           =   240
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   3
               Left            =   8220
               TabIndex        =   109
               Top             =   1170
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
               Index           =   26
               Left            =   3000
               TabIndex        =   108
               Top             =   1140
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�������ܤ��"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   25
               Left            =   525
               TabIndex        =   107
               Top             =   1455
               Width           =   1260
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
               Index           =   24
               Left            =   3000
               TabIndex        =   106
               Top             =   1440
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "����������"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   23
               Left            =   525
               TabIndex        =   105
               Top             =   1155
               Width           =   1260
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
               Left            =   3525
               TabIndex        =   104
               Top             =   1725
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�w�p�������ɶ�"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   105
               TabIndex        =   103
               Top             =   1755
               Width           =   1680
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
               Index           =   15
               Left            =   3000
               TabIndex        =   102
               Top             =   840
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
               Height          =   195
               Index           =   11
               Left            =   945
               TabIndex        =   101
               Top             =   840
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
               Index           =   13
               Left            =   3000
               TabIndex        =   100
               Top             =   510
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�ƨ����"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   945
               TabIndex        =   99
               Top             =   540
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
               Index           =   10
               Left            =   3225
               TabIndex        =   98
               Top             =   210
               Width           =   240
            End
            Begin VB.Label Label2 
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
               Height          =   195
               Index           =   5
               Left            =   945
               TabIndex        =   97
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�z�f�T�{����ɶ�"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   36
               Left            =   120
               TabIndex        =   96
               Top             =   2040
               Width           =   1680
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
               Index           =   37
               Left            =   3525
               TabIndex        =   95
               Top             =   2010
               Width           =   360
            End
         End
         Begin VB.CommandButton cmd_Tab0_Reset 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�M���d�߱���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":0081
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   69
            Top             =   2220
            Width           =   1935
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
            Height          =   900
            Index           =   3
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":0393
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   68
            Top             =   3135
            Width           =   1935
         End
         Begin VB.CommandButton cmd_Tab0_SelectField 
            BackColor       =   &H00C0E0FF&
            Caption         =   "�^�������"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":07D5
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   67
            Top             =   390
            Width           =   1935
         End
         Begin VB.CommandButton cmd_Query 
            BackColor       =   &H008080FF&
            Caption         =   "�q �� �d ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   0
            Left            =   8880
            Picture         =   "frm_Query_Orders.frx":0ADF
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   66
            Top             =   1305
            Width           =   1935
         End
         Begin VB.Frame Frame11 
            Height          =   1845
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   8610
            Begin VB.TextBox txt_StorerKey 
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
               TabIndex        =   52
               Top             =   180
               Width           =   1125
            End
            Begin VB.TextBox txt_Extern_Start 
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
               TabIndex        =   51
               Top             =   495
               Width           =   1125
            End
            Begin VB.TextBox txt_Extern_End 
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
               TabIndex        =   50
               Top             =   495
               Width           =   1125
            End
            Begin VB.TextBox txt_OrderDate_Start 
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
               TabIndex        =   49
               Top             =   810
               Width           =   1125
            End
            Begin VB.TextBox txt_OrderDate_End 
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
               TabIndex        =   48
               Top             =   810
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
               TabIndex        =   47
               Top             =   1125
               Width           =   1125
            End
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
               TabIndex        =   46
               Top             =   1125
               Width           =   1125
            End
            Begin VB.TextBox txt_ConsigneeKey 
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
               Left            =   4680
               TabIndex        =   45
               Top             =   180
               Width           =   2640
            End
            Begin VB.TextBox txt_ConsigName 
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
               Left            =   4680
               TabIndex        =   44
               Top             =   495
               Width           =   1890
            End
            Begin VB.TextBox txt_SKU 
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
               TabIndex        =   43
               Top             =   1455
               Width           =   2520
            End
            Begin VB.CheckBox chk_OnlyExpireDate 
               Caption         =   "���w�����"
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
               Left            =   3825
               TabIndex        =   42
               Top             =   915
               Width           =   1830
            End
            Begin VB.CheckBox txt_OrderNotes 
               Caption         =   "�q��Ƶ�"
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
               Left            =   5820
               TabIndex        =   41
               Top             =   915
               Width           =   1455
            End
            Begin VB.CheckBox chk_NotImport 
               Caption         =   "����J�q��"
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
               Left            =   3825
               TabIndex        =   40
               Top             =   1215
               Width           =   1455
            End
            Begin VB.CheckBox chk_WaitPlan 
               Caption         =   "�w��J�q��ݱƨ�"
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
               Left            =   3825
               TabIndex        =   39
               Top             =   1515
               Width           =   2085
            End
            Begin VB.CheckBox chk_CancelOrder 
               Caption         =   "�����q��"
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
               Left            =   5820
               TabIndex        =   38
               Top             =   1215
               Width           =   1455
            End
            Begin VB.CheckBox chk_ExpectOrder 
               Caption         =   "ñ�����`�q��"
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
               Left            =   5820
               TabIndex        =   37
               Top             =   1500
               Width           =   1665
            End
            Begin VB.CheckBox chk_Ship_qty 
               Caption         =   "���z�f�q��"
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
               Left            =   7080
               TabIndex        =   36
               Top             =   915
               Width           =   1455
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�f�D"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   570
               TabIndex        =   65
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�q��s��"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   150
               TabIndex        =   64
               Top             =   555
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
               Index           =   1
               Left            =   2205
               TabIndex        =   63
               Top             =   525
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�q����"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   150
               TabIndex        =   62
               Top             =   870
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
               Index           =   3
               Left            =   2205
               TabIndex        =   61
               Top             =   840
               Width           =   240
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�e�f���"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   150
               TabIndex        =   60
               Top             =   1185
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
               Index           =   5
               Left            =   2205
               TabIndex        =   59
               Top             =   1155
               Width           =   240
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�Ȥ�s��"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   3780
               TabIndex        =   58
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�Ȥ�W��"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   3780
               TabIndex        =   57
               Top             =   555
               Width           =   840
            End
            Begin VB.Label Label2 
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
               Height          =   195
               Index           =   3
               Left            =   555
               TabIndex        =   56
               Top             =   1515
               Width           =   420
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   1
               Left            =   3570
               TabIndex        =   55
               Top             =   1485
               Width           =   240
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   6
               Left            =   7650
               TabIndex        =   54
               Top             =   225
               Width           =   240
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�A"
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
               Index           =   7
               Left            =   6915
               TabIndex        =   53
               Top             =   525
               Width           =   240
            End
         End
         Begin VB.Shape Shape11 
            BackColor       =   &H00004080&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000D&
            BorderWidth     =   2
            Height          =   3765
            Left            =   8820
            Top             =   330
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   525
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   9915
         Begin VB.TextBox txt_Tab2_srcTotal_DifSDN 
            Alignment       =   1  '�a�k���
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   8520
            TabIndex        =   32
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_DifPick 
            Alignment       =   1  '�a�k���
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   6700
            TabIndex        =   30
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_Case 
            Alignment       =   1  '�a�k���
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1335
            TabIndex        =   26
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_PickCase 
            Alignment       =   1  '�a�k���
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   2940
            TabIndex        =   25
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox txt_Tab2_srcTotal_SDNCase 
            Alignment       =   1  '�a�k���
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4545
            TabIndex        =   24
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ñ���t����"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   4
            Left            =   7575
            TabIndex        =   33
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�z�f�t����"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   3
            Left            =   5760
            TabIndex        =   31
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "ñ����"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   9
            Left            =   3915
            TabIndex        =   29
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�z�f��"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   10
            Left            =   2325
            TabIndex        =   28
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�`�p�G�ƨ���"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   11
            Left            =   195
            TabIndex        =   27
            Top             =   210
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmd_Tab2SavetoExcel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�� Excel"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   10275
         Picture         =   "frm_Query_Orders.frx":0DE9
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   22
         Top             =   495
         Width           =   1065
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
         Height          =   1020
         Index           =   0
         Left            =   10275
         Picture         =   "frm_Query_Orders.frx":19AB
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   21
         Top             =   2385
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         Height          =   6315
         Left            =   -74685
         TabIndex        =   3
         Top             =   495
         Width           =   10785
         Begin VB.CommandButton cmd_Query 
            BackColor       =   &H008080FF&
            Caption         =   "�d  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Index           =   1
            Left            =   7830
            Picture         =   "frm_Query_Orders.frx":1DED
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   17
            Top             =   360
            Width           =   2385
         End
         Begin VB.ListBox lst_AllFields 
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
            Height          =   5580
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   2800
         End
         Begin VB.ListBox lst_SelectedFields 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   5580
            Left            =   4125
            TabIndex        =   15
            Top             =   600
            Width           =   2800
         End
         Begin VB.CommandButton cmd_Tab1_Add 
            BackColor       =   &H008080FF&
            Caption         =   "�֡�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   14
            ToolTipText     =   "�����"
            Top             =   3270
            Width           =   855
         End
         Begin VB.CommandButton cmd_Tab1_Remove 
            BackColor       =   &H0080C0FF&
            Caption         =   "�ա�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   13
            ToolTipText     =   "�������"
            Top             =   3720
            Width           =   855
         End
         Begin VB.CommandButton cmd_Tab1_Down 
            BackColor       =   &H00FF80FF&
            Caption         =   "�U��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   3570
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   12
            Top             =   1425
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_Up 
            BackColor       =   &H00FF8080&
            Caption         =   "�W��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   3570
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   11
            Top             =   675
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_Reload 
            BackColor       =   &H00C0C0C0&
            Caption         =   "���s���J"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   2940
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   10
            ToolTipText     =   "���J�d�߳]�w��"
            Top             =   5010
            Width           =   1185
         End
         Begin VB.ListBox lst_OrderBy 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   2940
            Left            =   7785
            TabIndex        =   9
            Top             =   2550
            Width           =   2800
         End
         Begin VB.CommandButton cmd_Tab1_OrderRemove 
            BackColor       =   &H0080C0FF&
            Caption         =   "�ա�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7035
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   8
            Top             =   4740
            Width           =   690
         End
         Begin VB.CommandButton cmd_Tab1_OrderAdd 
            BackColor       =   &H008080FF&
            Caption         =   "�֡�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7035
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   7
            Top             =   4290
            Width           =   690
         End
         Begin VB.CommandButton cmd_Tab1_OrderByUp 
            BackColor       =   &H00FF8080&
            Caption         =   "�W��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   7215
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   6
            Top             =   2625
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_OrderByDown 
            BackColor       =   &H00FF80FF&
            Caption         =   "�U��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   7215
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   5
            Top             =   3375
            Width           =   510
         End
         Begin VB.CommandButton cmd_Tab1_Reset 
            BackColor       =   &H008080FF&
            Caption         =   "�M  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2925
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   4
            ToolTipText     =   "�M���Ҧ��]�w��"
            Top             =   5700
            Width           =   1200
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H80000001&
            BackStyle       =   1  '���z��
            BorderColor     =   &H0000C000&
            BorderWidth     =   2
            Height          =   1140
            Left            =   7770
            Top             =   315
            Width           =   2505
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�� �� �� �� �C ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Index           =   0
            Left            =   615
            TabIndex        =   20
            Top             =   285
            Width           =   1905
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�w �� �� �� �C ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   1
            Left            =   4605
            TabIndex        =   19
            Top             =   285
            Width           =   1905
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00C0C000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   0
            Left            =   3510
            Top             =   615
            Width           =   630
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00400040&
            BackStyle       =   1  '���z��
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            Height          =   1005
            Index           =   1
            Left            =   3060
            Top             =   3210
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�� �� �] �w"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   240
            Index           =   2
            Left            =   8610
            TabIndex        =   18
            Top             =   2145
            Width           =   1245
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00400040&
            BackStyle       =   1  '���z��
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            Height          =   1005
            Index           =   2
            Left            =   6975
            Top             =   4230
            Width           =   810
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00C0C000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   3
            Left            =   7155
            Top             =   2565
            Width           =   630
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            Height          =   465
            Index           =   0
            Left            =   495
            Top             =   180
            Width           =   2145
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            Height          =   465
            Index           =   1
            Left            =   4485
            Top             =   180
            Width           =   2145
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404000&
            BackStyle       =   1  '���z��
            Height          =   465
            Index           =   2
            Left            =   8145
            Top             =   2040
            Width           =   2145
         End
      End
      Begin MSDataGridLib.DataGrid dg_Result 
         Height          =   5925
         Left            =   195
         TabIndex        =   2
         Top             =   990
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   10451
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
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   10485
         Top             =   1740
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frm_Query_Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private iLoop As Double

Private arZip() As String            '�l���ϸ�
Private arAreaCode() As String       '�B�e�ϰ�
Private arExtraDemand() As String    '�S��ݨD
Private arVehicleType() As String    '����
Private arTRPCompany() As String     '�f�B���q
Private arRSC() As String            '���`��]
Private arRBC() As String            '�d���k��
Private MyXlsApp As Excel.Application

Private rs_Result As ADODB.Recordset

Private Sub cmd_Tab2SavetoExcel_Click()
'�q��d�� >> �� EXCEL
Recordset2Excel Me.Caption, rs_Result
'..�b���s��EXCEL
Set MyXlsApp = Nothing

'If rs_Result Is Nothing Then Exit Sub
'If rs_Result.RecordCount = 0 Then Exit Sub
'
'Dim ExcelTitle As String
'Call DocStoreDirectory(strDocPath)
'
'Dim strTranFileName As String           'Excel �ɮצW��
'CmnDialog.DialogTitle = "��s Excel ��"
'CmnDialog.InitDir = "c:\my documents"
'CmnDialog.FileName = "�q��d��_" & Format(Now, "YYYYMMDDHHNNSS")
'CmnDialog.Filter = "Excel�ɮ�(*.xls)|*.xls"
'CmnDialog.FilterIndex = 1
'CmnDialog.CancelError = True
'On Error Resume Next
'CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
'CmnDialog.ShowOpen
'If Err.Number = cdlCancel Then          '�� [�}������] ��ܤ�����A���U [����] �s
'   msg_text = "��� [����] ���s�A������ Excel ���ۦ�s��"
'   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
'   strTranFileName = ""
'Else
'   strTranFileName = CmnDialog.FileName
'   If Dir(strTranFileName) <> "" Then
'      Kill strTranFileName
'   End If
'End If
'
'On Error GoTo err_Handle
'Screen.MousePointer = vbHourglass
'If SaveTo_ExcelFile(strTranFileName, rs_Result) = 1 Then
'   Screen.MousePointer = vbDefault
'   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
'Else
'   Screen.MousePointer = vbDefault
'   If Len(strTranFileName) > 0 Then
'      msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
'      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   End If
'End If
'If rs_Result Is Nothing Then Exit Sub
'rs_Result.MoveFirst
'Exit Sub

'err_Handle:
'   Dim tmpString As String
'   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'   CreateErrorLog Me.Name & "-�� EXCEL", Me.Caption, "cmd_Tab2SavetoExcel_Click", tmpString
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Query_Click(Index As Integer)
'�ˬd�Ƨ���즳���h���i�X�{�b�ݱƧǤ�
Dim i As Integer: Dim j As Integer
For i = 0 To lst_OrderBy.ListCount
    For j = 0 To lst_AllFields.ListCount
         If lst_OrderBy.List(i) = lst_AllFields.List(j) And lst_OrderBy.List(i) <> "" Then
               msg_text = "���ˬd'�Ƨ����'���i�X�{�b'�ݿ����'��!" & vbCrLf & "���~���:" & lst_OrderBy.List(i)
               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
               GoTo err:
         End If
    Next
Next

' �d��
If lst_SelectedFields.ListCount = 0 Then
   msg_text = "�@�~�{�ǿ��~�G�å�����d�ߵ��G�^�Ǫ����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Screen.MousePointer = vbHourglass
DoEvents: DoEvents
On Error GoTo err_Handle

'�ϥΪ̬d�����ҳ]�w�Ȧs��
Call SaveQueryEnv
Set dg_Result.DataSource = Nothing
Set rs_Result = Nothing

'�զX������
Dim strSelect As String, strOrderBy As String
strSelect = ""
For iLoop = 0 To lst_SelectedFields.ListCount - 1
    If strSelect = "" Then
       strSelect = lst_SelectedFields.List(iLoop)
    Else
       strSelect = strSelect & "," & lst_SelectedFields.List(iLoop)
    End If
Next iLoop
strSelect = "Select Distinct " & strSelect & " From Query_OrdersData "
strOrderBy = ""
For iLoop = 0 To lst_OrderBy.ListCount - 1
    If strOrderBy = "" Then
       strOrderBy = lst_OrderBy.List(iLoop)
    Else
       strOrderBy = strOrderBy & "," & lst_OrderBy.List(iLoop)
    End If
Next iLoop

'�զX�d�߱���
Dim str_Where As String, strSubwhere As String, intloop As Integer, tmp_data() As String
str_Where = ""
'Storer
txt_StorerKey.Text = Trim(txt_StorerKey.Text)
strSubwhere = ""
If txt_StorerKey.Text <> "" Then
   strSubwhere = " �f�D = '" & txt_StorerKey.Text & "' "
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�f�D�渹
txt_Extern_Start.Text = Trim(txt_Extern_Start.Text)
txt_Extern_End.Text = Trim(txt_Extern_End.Text)
strSubwhere = ""
If Len(txt_Extern_Start.Text) > 0 And Len(txt_Extern_End.Text) > 0 Then
   strSubwhere = " �f�D�渹 Between '" & txt_Extern_Start.Text & "' and '" & txt_Extern_End.Text & "' "
ElseIf Len(txt_Extern_Start.Text) > 0 And Len(txt_Extern_End.Text) = 0 Then
   strSubwhere = " �f�D�渹 = '" & txt_Extern_Start.Text & "' "
ElseIf Len(txt_Extern_Start.Text) = 0 And Len(txt_Extern_End.Text) > 0 Then
   strSubwhere = " �f�D�渹 = '" & txt_Extern_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�q����
txt_OrderDate_Start.Text = Trim(txt_OrderDate_Start.Text)
txt_OrderDate_End.Text = Trim(txt_OrderDate_End.Text)
strSubwhere = ""
If Len(txt_OrderDate_Start.Text) > 0 And Len(txt_OrderDate_End.Text) > 0 Then
   strSubwhere = " �q���� Between '" & txt_OrderDate_Start.Text & "' and '" & txt_OrderDate_End.Text & "' "
ElseIf Len(txt_OrderDate_Start.Text) > 0 And Len(txt_OrderDate_End.Text) = 0 Then
   strSubwhere = " �q���� = '" & txt_Extern_Start.Text & "' "
ElseIf Len(txt_OrderDate_Start.Text) = 0 And Len(txt_OrderDate_End.Text) > 0 Then
   strSubwhere = " �q���� = '" & txt_Extern_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�e�f���
txt_DeliveryDate_Start.Text = Trim(txt_DeliveryDate_Start.Text)
txt_DeliveryDate_End.Text = Trim(txt_DeliveryDate_End.Text)
strSubwhere = ""
If Len(txt_DeliveryDate_Start.Text) > 0 And Len(txt_DeliveryDate_End.Text) > 0 Then
   strSubwhere = " �e�f��� Between '" & txt_DeliveryDate_Start.Text & "' and '" & txt_DeliveryDate_End.Text & "' "
ElseIf Len(txt_DeliveryDate_Start.Text) > 0 And Len(txt_DeliveryDate_End.Text) = 0 Then
   strSubwhere = " �e�f��� = '" & txt_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_DeliveryDate_Start.Text) = 0 And Len(txt_DeliveryDate_End.Text) > 0 Then
   strSubwhere = " �e�f��� = '" & txt_DeliveryDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�f��
txt_SKU.Text = Trim(txt_SKU.Text)
strSubwhere = ""
If Len(txt_SKU.Text) > 0 Then
   If InStr(txt_SKU.Text, ",") > 0 Then
      tmp_data = Split(txt_SKU.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " �f�� in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and �f�� in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " �f�� like '%" & txt_SKU.Text & "%' "
      Else
         str_Where = str_Where & " and �f�� like '%" & txt_SKU.Text & "%' "
      End If
   End If
End If
'�Ȥ�s��
txt_ConsigneeKey.Text = Trim(txt_ConsigneeKey.Text)
strSubwhere = ""
If Len(txt_ConsigneeKey.Text) > 0 Then
   If InStr(txt_ConsigneeKey.Text, ",") > 0 Then
      tmp_data = Split(txt_ConsigneeKey.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " �Ȥ�s�� in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and �Ȥ�s�� in (" & strSubwhere & ") "
      End If
  Else
      If Len(str_Where) = 0 Then
         str_Where = " �Ȥ�s�� like '%" & txt_ConsigneeKey.Text & "%' "
      Else
         str_Where = str_Where & " and �Ȥ�s�� like '%" & txt_ConsigneeKey.Text & "%' "
      End If
  End If
End If
'�Ȥ�W��
txt_ConsigName.Text = Trim(txt_ConsigName.Text)
If txt_ConsigName.Text <> "" Then
   If Len(str_Where) = 0 Then
      str_Where = " �Ȥ�W�� like '%" & txt_ConsigName.Text & "%' "
   Else
      str_Where = str_Where & " and �Ȥ�W�� like '%" & strSubwhere & "%' "
   End If
End If
'���w�����
If chk_OnlyExpireDate.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " ���O <> '' "
   Else
      str_Where = str_Where & " and ���O <> ''"
   End If
End If
'�q��Ƶ�
If txt_OrderNotes.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " �q��Ƶ� <> '' "
   Else
      str_Where = str_Where & " and �q��Ƶ� <> '' "
   End If
End If
'�����q��
If chk_CancelOrder.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " ñ�����O = '�����q��' "
   Else
      str_Where = str_Where & " and ñ�����O = '���X�q��' "
   End If
End If
'ñ�����`�q��
If chk_ExpectOrder.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " ñ�����O = '���`�q��' "
   Else
      str_Where = str_Where & " and ñ�����O = '���`�q��' "
   End If
End If
'���z�f�q��chk_Ship_qty
If chk_Ship_qty.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " �z�f�q = '0' "
   Else
      str_Where = str_Where & " and �z�f�q = '0' "
   End If
End If
'��J�ƨ��t���ѧO���GOrders.B_PHONE2 >> 00 �w��J
If chk_NotImport.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " ��J�ѧO = '' "
   Else
      str_Where = str_Where & " and  ��J�ѧO = '' "
   End If
End If
'�w��J�A�|���ƨ�(�|�����͸��u�s��)
If chk_WaitPlan.Value = vbChecked Then
   If Len(str_Where) = 0 Then
      str_Where = " (��J�ѧO = 'V' and ���u�s�� = '') "
   Else
      str_Where = str_Where & " and (��J�ѧO = 'V' and ���u�s�� = '') "
   End If
End If
'�@���ƨ����u�s��
txt_FRouteNo_Start.Text = Trim(txt_FRouteNo_Start.Text)
txt_FRouteNo_End.Text = Trim(txt_FRouteNo_End.Text)
strSubwhere = ""
If Len(txt_FRouteNo_Start.Text) > 0 And Len(txt_FRouteNo_End.Text) > 0 Then
   strSubwhere = " ���u�s�� Between '" & txt_FRouteNo_Start.Text & "' and '" & txt_FRouteNo_End.Text & "' "
ElseIf Len(txt_FRouteNo_Start.Text) > 0 And Len(txt_FRouteNo_End.Text) = 0 Then
   strSubwhere = " ���u�s�� = '" & txt_FRouteNo_Start.Text & "' "
ElseIf Len(txt_FRouteNo_Start.Text) = 0 And Len(txt_FRouteNo_End.Text) > 0 Then
   strSubwhere = " ���u�s�� = '" & txt_FRouteNo_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�@���ƨ����
txt_FPlanDate_Start.Text = Trim(txt_FPlanDate_Start.Text)
txt_FPlanDate_End.Text = Trim(txt_FPlanDate_End.Text)
strSubwhere = ""
If Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   strSubwhere = " �ƨ���� Between '" & txt_FPlanDate_Start.Text & "' and '" & txt_FPlanDate_End.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) > 0 And Len(txt_FPlanDate_End.Text) = 0 Then
   strSubwhere = " �ƨ���� = '" & txt_FPlanDate_Start.Text & "' "
ElseIf Len(txt_FPlanDate_Start.Text) = 0 And Len(txt_FPlanDate_End.Text) > 0 Then
   strSubwhere = " �ƨ���� = '" & txt_FPlanDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�@���X�����
txt_FDeliveryDate_Start.Text = Trim(txt_FDeliveryDate_Start.Text)
txt_FDeliveryDate_End.Text = Trim(txt_FDeliveryDate_End.Text)
strSubwhere = ""
If Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   strSubwhere = " �X����� Between '" & txt_FDeliveryDate_Start.Text & "' and '" & txt_FDeliveryDate_End.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) > 0 And Len(txt_FDeliveryDate_End.Text) = 0 Then
   strSubwhere = " �X����� = '" & txt_FDeliveryDate_Start.Text & "' "
ElseIf Len(txt_FDeliveryDate_Start.Text) = 0 And Len(txt_FDeliveryDate_End.Text) > 0 Then
   strSubwhere = " �X����� = '" & txt_FDeliveryDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�@���w�p������<daniel 20041005>
txt_FPlanCheckin_Start.Text = Trim(txt_FPlanCheckin_Start.Text)
txt_FPlanCheckin_End.Text = Trim(txt_FPlanCheckin_End.Text)
txt_FPlanCheckinTime_Start.Text = Trim(txt_FPlanCheckinTime_Start.Text)
txt_FPlanCheckinTime_End.Text = Trim(txt_FPlanCheckinTime_End.Text)
strSubwhere = ""
If Len(txt_FPlanCheckin_Start.Text) > 0 And Len(txt_FPlanCheckin_End.Text) > 0 And Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
   strSubwhere = " �w�p������ Between '" & txt_FPlanCheckin_Start.Text & "' and '" & txt_FPlanCheckin_End.Text & "' "
ElseIf Len(txt_FPlanCheckin_Start.Text) > 0 And Len(txt_FPlanCheckin_End.Text) = 0 And Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
   strSubwhere = " �w�p������ = '" & txt_FPlanCheckin_Start.Text & "' "
ElseIf Len(txt_FPlanCheckin_Start.Text) = 0 And Len(txt_FPlanCheckin_End.Text) > 0 And Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
   strSubwhere = " �w�p������ = '" & txt_FPlanCheckin_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�@���ƨ��G�w�p����ɶ�<daniel 20041005>
If Len(Trim(txt_FPlanCheckinTime_Start.Text)) <> 0 Then
    If Len(txt_FPlanCheckinTime_Start.Text) <> 4 Then
        msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_FPlanCheckinTime_Start.Text, 2)
        Case "00" To "23"
        Case Else
             msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             Screen.MousePointer = vbDefault
             txt_FPlanCheckinTime_Start.SetFocus
             Exit Sub
     End Select
     Select Case Right(txt_FPlanCheckinTime_Start.Text, 2)
        Case "00" To "59"
        Case Else
             msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_FPlanCheckinTime_Start.SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
     End Select
End If
If Len(Trim(txt_FPlanCheckinTime_End.Text)) <> 0 Then
    If Len(txt_FPlanCheckinTime_End.Text) <> 4 Then
        msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_FPlanCheckinTime_End.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_FPlanCheckinTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
    Select Case Right(txt_FPlanCheckinTime_End.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_FPlanCheckinTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
End If
strSubwhere = ""
If Len(txt_FPlanCheckinTime_Start.Text) > 0 And Len(txt_FPlanCheckinTime_End.Text) > 0 Then
    If Len(txt_FPlanCheckin_Start.Text) = 0 And Len(txt_FPlanCheckin_End.Text) = 0 Then
        msg_text = "�п�J�w�p������"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " �w�p������+�w�p����ɶ� Between '" & txt_FPlanCheckin_Start.Text & txt_FPlanCheckinTime_Start.Text & "' and '" & txt_FPlanCheckin_End.Text & txt_FPlanCheckinTime_End.Text & "' "
ElseIf Len(txt_FPlanCheckinTime_Start.Text) > 0 And Len(txt_FPlanCheckinTime_End.Text) = 0 Then
    If Len(txt_FPlanCheckin_Start.Text) = 0 Then
        msg_text = "�п�J�w�p������"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " �w�p������+�w�p����ɶ� = '" & txt_FPlanCheckin_Start.Text & txt_FPlanCheckinTime_Start.Text & "' "
ElseIf Len(txt_FPlanCheckinTime_Start.Text) = 0 And Len(txt_FPlanCheckinTime_End.Text) > 0 Then
    If Len(txt_FPlanCheckin_End.Text) = 0 Then
        msg_text = "�п�J�w�p������"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " �w�p������+�w�p����ɶ� = '" & txt_FPlanCheckin_End.Text & txt_FPlanCheckinTime_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If

'�z�f�T�{���<daniel 20041005>
txt_SDNDate_Start.Text = Trim(txt_SDNDate_Start.Text)
txt_SDNDate_End.Text = Trim(txt_SDNDate_End.Text)
strSubwhere = ""
If Len(txt_SDNDate_Start.Text) > 0 And Len(txt_SDNDate_End.Text) > 0 And Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) = 0 Then
    strSubwhere = " left(�z�f�T�{�ɶ�,8) Between '" & txt_SDNDate_Start.Text & "' and '" & txt_SDNDate_End.Text & "' "
ElseIf Len(txt_SDNDate_Start.Text) > 0 And Len(txt_SDNDate_End.Text) = 0 And Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) = 0 Then
    strSubwhere = " left(�z�f�T�{�ɶ�,8) = '" & txt_SDNDate_Start.Text & "' "
ElseIf Len(txt_SDNDate_Start.Text) = 0 And Len(txt_SDNDate_End.Text) > 0 And Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) = 0 Then
    strSubwhere = " left(�z�f�T�{�ɶ�,8) = '" & txt_SDNDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�z�f�T�{�ɶ�<daniel 20041005>
txt_SDNTime_Start.Text = Trim(txt_SDNTime_Start.Text)
txt_SDNTime_End.Text = Trim(txt_SDNTime_End.Text)
If Len(Trim(txt_SDNTime_Start.Text)) <> 0 Then
    If Len(txt_SDNTime_Start.Text) <> 4 Then
        msg_text = "�z�f�T�{�ɶ��G��Ʈ榡 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_SDNTime_Start.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "�z�f�T�{�ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_Start.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
    Select Case Right(txt_SDNTime_Start.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "�z�f�T�{�ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_Start.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
End If
If Len(Trim(txt_SDNTime_End.Text)) <> 0 Then
    If Len(txt_SDNTime_End.Text) <> 4 Then
        msg_text = "�z�f�T�{�ɶ��G��Ʈ榡 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case Left(txt_SDNTime_End.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "�z�f�T�{�ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
    Select Case Right(txt_SDNTime_End.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "�z�f�T�{�ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_SDNTime_End.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
    End Select
End If
strSubwhere = ""
If Len(txt_SDNTime_Start.Text) > 0 And Len(txt_SDNTime_End.Text) > 0 Then
    If Len(txt_SDNDate_Start.Text) = 0 Or Len(txt_SDNDate_End.Text) = 0 Then
        msg_text = "�п�J�z�f�T�{���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " �z�f�T�{�ɶ�  Between '" & txt_SDNDate_Start.Text & txt_SDNTime_Start.Text & "' and '" & txt_SDNDate_End.Text & txt_SDNTime_End.Text & "' "
ElseIf Len(txt_SDNTime_Start.Text) > 0 And Len(txt_SDNTime_End.Text) = 0 Then
    If Len(txt_SDNDate_Start.Text) = 0 Then
        msg_text = "�п�J�z�f�T�{���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " �z�f�T�{�ɶ� = '" & txt_SDNDate_Start.Text & txt_SDNTime_Start.Text & "' "
ElseIf Len(txt_SDNTime_Start.Text) = 0 And Len(txt_SDNTime_End.Text) > 0 Then
    If Len(txt_SDNDate_End.Text) = 0 Then
        msg_text = "�п�J�z�f�T�{���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strSubwhere = " �z�f�T�{�ɶ� = '" & txt_SDNDate_End.Text & txt_SDNTime_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If


'�@���ƨ��G����������
txt_FCheckin_Start.Text = Trim(txt_FCheckin_Start.Text)
txt_FCheckin_End.Text = Trim(txt_FCheckin_End.Text)
strSubwhere = ""
If Len(txt_FCheckin_Start.Text) > 0 And Len(txt_FCheckin_End.Text) > 0 Then
   strSubwhere = " ������ Between '" & txt_FCheckin_Start.Text & "' and '" & txt_FCheckin_End.Text & "' "
ElseIf Len(txt_FCheckin_Start.Text) > 0 And Len(txt_FCheckin_End.Text) = 0 Then
   strSubwhere = " ������ = '" & txt_FCheckin_Start.Text & "' "
ElseIf Len(txt_FCheckin_Start.Text) = 0 And Len(txt_FCheckin_End.Text) > 0 Then
   strSubwhere = " ������ = '" & txt_FCheckin_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If

'�@���ƨ��G�������ܤ��
txt_FCheckout_Start.Text = Trim(txt_FCheckout_Start.Text)
txt_FCheckout_End.Text = Trim(txt_FCheckout_End.Text)
strSubwhere = ""
If Len(txt_FCheckout_Start.Text) > 0 And Len(txt_FCheckout_End.Text) > 0 Then
   strSubwhere = " ���ܤ�� Between '" & txt_FCheckout_Start.Text & "' and '" & txt_FCheckout_End.Text & "' "
ElseIf Len(txt_FCheckout_Start.Text) > 0 And Len(txt_FCheckout_End.Text) = 0 Then
   strSubwhere = " ���ܤ�� = '" & txt_FCheckout_Start.Text & "' "
ElseIf Len(txt_FCheckout_Start.Text) = 0 And Len(txt_FCheckout_End.Text) > 0 Then
   strSubwhere = " ���ܤ�� = '" & txt_FCheckout_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�@���ƨ��Gexe�^�Ǫ��A
If cmb_EXEConfirm.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " �^�Ǫ��A = '" & cmb_EXEConfirm.List(cmb_EXEConfirm.ListIndex) & "' "
   Else
      str_Where = str_Where & " �^�Ǫ��A = '" & cmb_EXEConfirm.List(cmb_EXEConfirm.ListIndex) & "' "
   End If
End If
'�@���ƨ��G�ƨ��H��
txt_FAddWho.Text = Trim(txt_FAddWho.Text)
strSubwhere = ""
If Len(txt_FAddWho.Text) > 0 Then
   If InStr(txt_FAddWho.Text, ",") > 0 Then
      tmp_data = Split(txt_FAddWho.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " �ƨ��� in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and �ƨ��� in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " �ƨ��� like '%" & txt_FAddWho.Text & "%' "
      Else
         str_Where = str_Where & " and �ƨ��� like '%" & txt_FAddWho.Text & "%' "
      End If
   End If
End If
'�@���ƨ��G���P���X
txt_FVehicleID.Text = Trim(txt_FVehicleID.Text)
strSubwhere = ""
If Len(txt_FVehicleID.Text) > 0 Then
   If InStr(txt_FVehicleID.Text, ",") > 0 Then
      tmp_data = Split(txt_FVehicleID.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " ���P���X in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and ���P���X in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " ���P���X like '%" & txt_FVehicleID.Text & "' "
      Else
         str_Where = str_Where & " and ���P���X like '%" & txt_FVehicleID.Text & "' "
      End If
   End If
End If
'�@���ƨ��G�r�p�H
txt_FDriver.Text = Trim(txt_FDriver.Text)
strSubwhere = ""
If Len(txt_FDriver.Text) > 0 Then
   If InStr(txt_FDriver.Text, ",") > 0 Then
      tmp_data = Split(txt_FDriver.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " �r�p�H in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and �r�p�H in (" & strSubwhere & ") "
      End If
   Else   '�S��J�r�I���j�A�� Like �i��d��
      If Len(str_Where) = 0 Then
         str_Where = " �r�p�H like '%" & txt_FDriver.Text & "%' "
      Else
         str_Where = str_Where & " and �r�p�H like '%" & txt_FDriver.Text & "%' "
      End If
   End If
End If
'�@���ƨ��G�X�Y�Ȧs
txt_FDockNo.Text = Trim(txt_FDockNo.Text)
strSubwhere = ""
If Len(txt_FDockNo.Text) > 0 Then
   If InStr(txt_FDockNo.Text, ",") > 0 Then
      tmp_data = Split(txt_FDockNo.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " �X�Y�Ȧs in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and �X�Y�Ȧs in (" & strSubwhere & ") "
      End If
   Else   '�S��J�r�I���j�A�� Like �i��d��
      If Len(str_Where) = 0 Then
         str_Where = " �X�Y�Ȧs like '%" & txt_FDockNo.Text & "%' "
      Else
         str_Where = str_Where & " and �X�Y�Ȧs like '%" & txt_FDockNo.Text & "%' "
      End If
   End If
End If
'�@���ƨ��GTMS�渹
txt_FReceiptNo.Text = Trim(txt_FReceiptNo.Text)
strSubwhere = ""
If Len(txt_FReceiptNo.Text) > 0 Then
   If InStr(txt_FReceiptNo.Text, ",") > 0 Then
      tmp_data = Split(txt_FReceiptNo.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " TMS�渹 in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and TMS�渹 in (" & strSubwhere & ") "
      End If
   Else   '�S��J�r�I���j�A�� Like �i��d��
      If Len(str_Where) = 0 Then
         str_Where = " TMS�渹 like '%" & txt_FReceiptNo.Text & "%' "
      Else
         str_Where = str_Where & " and TMS�渹 like '%" & txt_FReceiptNo.Text & "%' "
      End If
   End If
End If
'�z��i��G���ƨ�
If chk_SecondPlan.Value = vbChecked Then
      If Len(str_Where) = 0 Then
         str_Where = " �G�����u�s�� <> '' "
      Else
         str_Where = str_Where & " and �G�����u�s�� <> '' "
      End If
End If
'�G���ƨ��G���u�s��
txt_SRouteNo_Start.Text = Trim(txt_SRouteNo_Start.Text)
txt_SRouteNo_End.Text = Trim(txt_SRouteNo_End.Text)
strSubwhere = ""
If Len(txt_SRouteNo_Start.Text) > 0 And Len(txt_SRouteNo_End.Text) > 0 Then
   strSubwhere = "  �G�����u�s�� Between '" & txt_SRouteNo_Start.Text & "' and '" & txt_SRouteNo_End.Text & "' "
ElseIf Len(txt_SRouteNo_Start.Text) > 0 And Len(txt_SRouteNo_End.Text) = 0 Then
   strSubwhere = "  �G�����u�s�� = '" & txt_SRouteNo_Start.Text & "' "
ElseIf Len(txt_SRouteNo_Start.Text) = 0 And Len(txt_SRouteNo_End.Text) > 0 Then
   strSubwhere = "  �G�����u�s�� = '" & txt_SRouteNo_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If

'�G���ƨ��G�X�����
txt_SDeliveryDate_Start.Text = Trim(txt_SDeliveryDate_Start.Text)
txt_SDeliveryDate_End.Text = Trim(txt_SDeliveryDate_End.Text)
strSubwhere = ""
If Len(txt_SDeliveryDate_Start.Text) > 0 And Len(txt_SDeliveryDate_End.Text) > 0 Then
   strSubwhere = "  �G���X����� Between '" & txt_SDeliveryDate_Start.Text & "' and '" & txt_SDeliveryDate_End.Text & "' "
ElseIf Len(txt_SDeliveryDate_Start.Text) > 0 And Len(txt_SDeliveryDate_End.Text) = 0 Then
   strSubwhere = "  �G���X����� = '" & txt_SDeliveryDate_Start.Text & "' "
ElseIf Len(txt_SDeliveryDate_Start.Text) = 0 And Len(txt_SDeliveryDate_End.Text) > 0 Then
   strSubwhere = "  �G���X����� = '" & txt_SDeliveryDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�G���ƨ��G���P���X
txt_SVehicleID.Text = Trim(txt_SVehicleID.Text)
strSubwhere = ""
If Len(txt_SVehicleID.Text) > 0 Then
   If InStr(txt_SVehicleID.Text, ",") > 0 Then
      tmp_data = Split(txt_SVehicleID.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " �G�����P���X in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and �G�����P���X in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " �G�����P���X like '%" & txt_SVehicleID.Text & "' "
      Else
         str_Where = str_Where & " and �G�����P���X like '%" & txt_SVehicleID.Text & "' "
      End If
   End If
End If
'�G���ƨ��G�ƨ��H��
txt_SAddWho.Text = Trim(txt_SAddWho.Text)
strSubwhere = ""
If Len(txt_SAddWho.Text) > 0 Then
   If InStr(txt_SAddWho.Text, ",") > 0 Then
      tmp_data = Split(txt_SAddWho.Text, ",", -1, vbTextCompare)
      For intloop = LBound(tmp_data) To UBound(tmp_data)
          If Len(strSubwhere) = 0 Then
             strSubwhere = "'" & tmp_data(intloop) & "'"
          Else
             strSubwhere = strSubwhere & ",'" & tmp_data(intloop) & "'"
          End If
      Next intloop
      If Len(str_Where) = 0 Then
         str_Where = " �G���ƨ��� in (" & strSubwhere & ") "
      Else
         str_Where = str_Where & " and �G���ƨ��� in (" & strSubwhere & ") "
      End If
   Else
      If Len(str_Where) = 0 Then
         str_Where = " �G���ƨ��� like '%" & txt_SAddWho.Text & "%' "
      Else
         str_Where = str_Where & " and �G���ƨ��� like '%" & txt_SAddWho.Text & "%' "
      End If
   End If
End If
'�l���ϸ�
If cmb_ZIP.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " �l���ϸ� = '" & arZip(cmb_ZIP.ListIndex) & "' "
   Else
      str_Where = str_Where & " and �l���ϸ� = '" & arZip(cmb_ZIP.ListIndex) & "' "
   End If
End If
'�B�e�ϰ�
If cmb_AreaCode.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " Area = '" & arAreaCode(cmb_AreaCode.ListIndex) & "' "
   Else
      str_Where = str_Where & " and Area = '" & arAreaCode(cmb_AreaCode.ListIndex) & "' "
   End If
End If
'�S��ݨD
If cmb_ExtraDemand.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " (�S��ݨD�X1 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "' OR �S��ݨD�X2 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "') "
   Else
      str_Where = str_Where & " and (�S��ݨD�X1 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "' OR �S��ݨD�X2 = '" & arExtraDemand(cmb_ExtraDemand.ListIndex) & "') "
   End If
End If
'�@���ƨ��G�B�e����
If cmb_VehicleType.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " ���إN�X = '" & arVehicleType(cmb_VehicleType.ListIndex) & "' "
   Else
      str_Where = str_Where & " and ���إN�X = '" & arVehicleType(cmb_VehicleType.ListIndex) & "' "
   End If
End If
'�@���ƨ��G�f�B���q
If cmb_TRPCompany.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " �f�B���q�N�X = '" & arTRPCompany(cmb_TRPCompany.ListIndex) & "' "
   Else
      str_Where = str_Where & " and �f�B���q�N�X = '" & arTRPCompany(cmb_TRPCompany.ListIndex) & "' "
   End If
End If
'�G���ƨ��G�B�e����
If cmb_SVehicleType.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " �G�����إN�X = '" & arVehicleType(cmb_SVehicleType.ListIndex) & "' "
   Else
      str_Where = str_Where & " and �G�����إN�X = '" & arVehicleType(cmb_SVehicleType.ListIndex) & "' "
   End If
End If
'�G���ƨ��G�f�B���q
If cmb_STRPCompany.ListIndex <> -1 Then
   If Len(str_Where) = 0 Then
      str_Where = " �G���f�B���q�N�X = '" & arTRPCompany(cmb_STRPCompany.ListIndex) & "' "
   Else
      str_Where = str_Where & " and �G���f�B���q�N�X = '" & arTRPCompany(cmb_STRPCompany.ListIndex) & "' "
   End If
End If

If Len(str_Where) = 0 Then
   Call Unload_RunLogForm
   Screen.MousePointer = vbDefault
   msg_text = "�`�N�G��ƶq�Ӥj�A�п�J�d�߱���H��ָ�ƶq"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
str_SQL = strSelect & " Where " & str_Where
If strOrderBy <> "" Then
   str_SQL = str_SQL & " Order by " & strOrderBy
End If

SSTab1.Tab = 2
DoEvents: DoEvents

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   Call Unload_RunLogForm
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_Tab2_srcTotal_Case.Text = "0"
   txt_Tab2_srcTotal_PickCase.Text = "0"
   txt_Tab2_srcTotal_SDNCase.Text = "0"
   txt_Tab2_srcTotal_DifPick = "0"
   txt_Tab2_srcTotal_DifSDN = "0"
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Result)
tmp_Rs.Close

With dg_Result
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 2                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 300                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Result.MoveFirst
Set dg_Result.DataSource = rs_Result
With dg_Result
    .RowHeight = 250
    For iLoop = 0 To .Columns.Count - 1
'        .Columns(iLoop).Width = GetFieldWidth(rs_Result.Fields(iLoop).Name)
        .Columns(iLoop).Alignment = GetFieldAlignment(rs_Result.Fields(iLoop).Name)
    Next iLoop
End With

'�]�w��e
SetDataGridColWidth "�q���Ƭd�ߵ��G", dg_Result

'�έp<daniel 20041005>
str_SQL = "select sum(�ƨ��q),sum(�z�f�q),sum(ñ���q) from Query_OrdersData Where " & str_Where
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   txt_Tab2_srcTotal_Case.Text = "0"
   txt_Tab2_srcTotal_PickCase.Text = "0"
   txt_Tab2_srcTotal_SDNCase.Text = "0"
   txt_Tab2_srcTotal_DifPick = "0"
   txt_Tab2_srcTotal_DifSDN = "0"
   tmp_Rs.Close
   Exit Sub
End If
txt_Tab2_srcTotal_Case.Text = tmp_Rs.Fields(0)
txt_Tab2_srcTotal_PickCase.Text = tmp_Rs.Fields(1)
txt_Tab2_srcTotal_SDNCase.Text = tmp_Rs.Fields(2)
txt_Tab2_srcTotal_DifPick = Val(tmp_Rs.Fields(0)) - Val(tmp_Rs.Fields(1))
txt_Tab2_srcTotal_DifSDN = tmp_Rs.Fields(0) - Val(tmp_Rs.Fields(2))
tmp_Rs.Close

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��d��", Me.Caption, "cmd_Query", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
      
err:
End Sub

Private Sub cmd_Tab0_Reset_Click()
'�M���d�߱���
Set dg_Result.DataSource = Nothing
Set rs_Result = Nothing

Call ClearForm_AllField(Me)

End Sub

Private Sub cmd_Tab1_OrderAdd_Click()
'����� >> �ƧǤ覡-�[�J
If lst_SelectedFields.SelCount > 0 Then
   lst_OrderBy.AddItem lst_SelectedFields.List(lst_SelectedFields.ListIndex)
End If

End Sub

Private Sub cmd_Tab1_OrderByDown_Click()
'����� >> �Ƨ���춶�ǤU��
If lst_OrderBy.ListCount = 0 Then Exit Sub
If lst_OrderBy.SelCount > 0 Then
   Dim afItem As String, selItem As String
   Dim afIndex As Double, selIndex As Double
   selItem = lst_OrderBy.List(lst_OrderBy.ListIndex)
   selIndex = lst_OrderBy.ListIndex
   If (lst_OrderBy.ListIndex + 1) > (lst_OrderBy.ListCount - 1) Then
      afItem = lst_OrderBy.List(0)
      afIndex = 0
   Else
      afItem = lst_OrderBy.List(lst_OrderBy.ListIndex + 1)
      afIndex = (lst_OrderBy.ListIndex + 1)
   End If
   lst_OrderBy.List(afIndex) = selItem
   lst_OrderBy.List(selIndex) = afItem
   lst_OrderBy.Selected(afIndex) = True
End If

End Sub

Private Sub cmd_Tab1_OrderByUp_Click()
'����� >> �Ƨ���춶�ǤW��
If lst_OrderBy.ListCount = 0 Then Exit Sub
If lst_OrderBy.SelCount > 0 Then
   Dim preItem As String, selItem As String
   Dim preIndex As Double, selIndex As Double
   selItem = lst_OrderBy.List(lst_OrderBy.ListIndex)
   selIndex = lst_OrderBy.ListIndex
   If (lst_OrderBy.ListIndex - 1) < 0 Then
      preItem = lst_OrderBy.List(lst_OrderBy.ListCount - 1)
      preIndex = (lst_OrderBy.ListCount - 1)
   Else
      preItem = lst_OrderBy.List(lst_OrderBy.ListIndex - 1)
      preIndex = (lst_OrderBy.ListIndex - 1)
   End If
   lst_OrderBy.List(preIndex) = selItem
   lst_OrderBy.List(selIndex) = preItem
   lst_OrderBy.Selected(preIndex) = True
End If

End Sub

Private Sub cmd_Tab1_OrderRemove_Click()
'����� >> DoubleClick ��������ƧǤ覡���
If lst_OrderBy.SelCount > 0 Then
   lst_OrderBy.RemoveItem lst_OrderBy.ListIndex
End If
End Sub

Private Sub cmd_Tab0_SelectField_Click()
'�d�߱��� >> �����
SSTab1.Tab = 1
End Sub

Private Sub cmd_Tab1_Add_Click()
'����� >> �[�J
If lst_AllFields.SelCount > 0 Then
   lst_SelectedFields.AddItem lst_AllFields.List(lst_AllFields.ListIndex)
   lst_AllFields.RemoveItem lst_AllFields.ListIndex
End If
End Sub

Private Sub cmd_Tab1_Down_Click()
'����� >> ��춶�ǤU��
If lst_SelectedFields.ListCount = 0 Then Exit Sub
If lst_SelectedFields.SelCount > 0 Then
   Dim afItem As String, selItem As String
   Dim afIndex As Double, selIndex As Double
   selItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   selIndex = lst_SelectedFields.ListIndex
   If (lst_SelectedFields.ListIndex + 1) > (lst_SelectedFields.ListCount - 1) Then
      afItem = lst_SelectedFields.List(0)
      afIndex = 0
   Else
      afItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex + 1)
      afIndex = (lst_SelectedFields.ListIndex + 1)
   End If
   lst_SelectedFields.List(afIndex) = selItem
   lst_SelectedFields.List(selIndex) = afItem
   lst_SelectedFields.Selected(afIndex) = True
End If

End Sub

Private Sub cmd_Tab1_Reload_Click()
'�����>>���s���J
'���o�Ҧ��i�Ϊ��d�����
Call GetAllFields
End Sub

Private Sub cmd_Tab1_Remove_Click()
'����� >> ����
If lst_SelectedFields.SelCount > 0 Then
   lst_AllFields.AddItem lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   lst_SelectedFields.RemoveItem lst_SelectedFields.ListIndex

End If

End Sub

Private Sub cmd_Tab1_Reset_Click()
'����� >> �M��
On Error GoTo err_Handle
Tran_Level = 0
Tran_Level = cn.BeginTrans
str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYFIELDS' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYORDER' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans
Tran_Level = 0
'���^�Ҧ����
Call GetAllFields
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��d��-�]�w�ȲM��", Me.Caption, "cmd_Tab1_Reset", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Up_Click()
'����� >> ��춶�ǤW��
If lst_SelectedFields.ListCount = 0 Then Exit Sub
If lst_SelectedFields.SelCount > 0 Then
   Dim preItem As String, selItem As String
   Dim preIndex As Double, selIndex As Double
   selItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   selIndex = lst_SelectedFields.ListIndex
   If (lst_SelectedFields.ListIndex - 1) < 0 Then
      preItem = lst_SelectedFields.List(lst_SelectedFields.ListCount - 1)
      preIndex = (lst_SelectedFields.ListCount - 1)
   Else
      preItem = lst_SelectedFields.List(lst_SelectedFields.ListIndex - 1)
      preIndex = (lst_SelectedFields.ListIndex - 1)
   End If
   lst_SelectedFields.List(preIndex) = selItem
   lst_SelectedFields.List(selIndex) = preItem
   lst_SelectedFields.Selected(preIndex) = True
End If
End Sub

Private Sub dg_Result_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

Dim objDataGrid As Object: Set objDataGrid = dg_Result
If Len(objDataGrid.Columns(ColIndex).DataField) = 0 Or objDataGrid.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, "�q���Ƭd�ߵ��G" & objDataGrid.Name, objDataGrid.Columns(ColIndex).DataField, objDataGrid.Columns(ColIndex).Width

End Sub

Private Sub lst_AllFields_DblClick()
'����� >> DoubleClick �[�J���
If lst_AllFields.SelCount > 0 Then
   lst_SelectedFields.AddItem lst_AllFields.List(lst_AllFields.ListIndex)
   lst_AllFields.RemoveItem lst_AllFields.ListIndex
End If
End Sub

Private Sub lst_OrderBy_DblClick()
'����� >> DoubleClick ��������ƧǤ覡���
If lst_OrderBy.SelCount > 0 Then
   lst_OrderBy.RemoveItem lst_OrderBy.ListIndex
End If
End Sub

Private Sub lst_SelectedFields_DblClick()
'����� >> DoubleClick �������
If lst_SelectedFields.SelCount > 0 Then
   lst_AllFields.AddItem lst_SelectedFields.List(lst_SelectedFields.ListIndex)
   lst_SelectedFields.RemoveItem lst_SelectedFields.ListIndex
End If
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�q��d�ߧ@�~"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
Me.Height = 7650: Me.Width = 11595
SSTab1.Tab = 0
If SSTab1.Tab = 0 Then cmd_Tab2SavetoExcel.Visible = False: cmd_Exit(0).Visible = False
'���o�Ҧ��i�Ϊ��d�����
Call GetAllFields

'�d�߱���ݿ�M��إ�
Dim dbZip As Double, dbAreaCode As Double, dbExtraDemand As Double, dbVehicleType As Double, dbTRPCompany As Double
Dim dbRSC As Double, dbRBC As Double
cmb_ZIP.Clear: cmb_AreaCode.Clear: cmb_ExtraDemand.Clear
cmb_VehicleType.Clear: cmb_TRPCompany.Clear
cmb_RSC.Clear: cmb_RBC.Clear
str_SQL = "Select �Ϥ�,�N�X,���� From Query_OrdersBaseData Order by �Ϥ�,�N�X"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arZip(1) As String
ReDim arAreaCode(1) As String
ReDim arExtraDemand(1) As String
ReDim arVehicleType(1) As String
ReDim arTRPCompany(1) As String
ReDim arRSC(1) As String
ReDim arRBC(1) As String
If Not tmp_Rs.EOF Then
   dbZip = 0: dbAreaCode = 0: dbExtraDemand = 0: dbVehicleType = 0: dbTRPCompany = 0
   Do While Not tmp_Rs.EOF
      Select Case tmp_Rs.Fields("�Ϥ�").Value
         Case "�l���ϸ�"
              arZip(dbZip) = tmp_Rs.Fields("�N�X").Value
              cmb_ZIP.AddItem tmp_Rs.Fields("�N�X").Value & Space(6 - Len(Trim(tmp_Rs.Fields("�N�X").Value))) & tmp_Rs.Fields("����").Value
              dbZip = dbZip + 1
              If dbZip = UBound(arZip) Then
                 ReDim Preserve arZip(UBound(arZip) + 2) As String
              End If
         Case "�B�e�ϰ�"
              arAreaCode(dbAreaCode) = tmp_Rs.Fields("�N�X").Value
              cmb_AreaCode.AddItem tmp_Rs.Fields("�N�X").Value & Space(6 - Len(Trim(tmp_Rs.Fields("�N�X").Value))) & tmp_Rs.Fields("����").Value
              dbAreaCode = dbAreaCode + 1
              If dbAreaCode = UBound(arAreaCode) Then
                 ReDim Preserve arAreaCode(UBound(arAreaCode) + 2) As String
              End If
         Case "�S��ݨD"
              arExtraDemand(dbExtraDemand) = tmp_Rs.Fields("�N�X").Value
              cmb_ExtraDemand.AddItem tmp_Rs.Fields("����").Value
              dbExtraDemand = dbExtraDemand + 1
              If dbExtraDemand = UBound(arExtraDemand) Then
                 ReDim Preserve arExtraDemand(UBound(arExtraDemand) + 2) As String
              End If
         Case "����"
              arVehicleType(dbVehicleType) = tmp_Rs.Fields("�N�X").Value
              cmb_VehicleType.AddItem tmp_Rs.Fields("�N�X").Value & Space(6 - Len(Trim(tmp_Rs.Fields("�N�X").Value))) & tmp_Rs.Fields("����").Value
              cmb_SVehicleType.AddItem tmp_Rs.Fields("�N�X").Value & Space(6 - Len(Trim(tmp_Rs.Fields("�N�X").Value))) & tmp_Rs.Fields("����").Value
              dbVehicleType = dbVehicleType + 1
              If dbVehicleType = UBound(arVehicleType) Then
                 ReDim Preserve arVehicleType(UBound(arVehicleType) + 2) As String
              End If
         Case "�f�B���q"
              arTRPCompany(dbTRPCompany) = tmp_Rs.Fields("�N�X").Value
              cmb_TRPCompany.AddItem tmp_Rs.Fields("����").Value
              cmb_STRPCompany.AddItem tmp_Rs.Fields("����").Value
              dbTRPCompany = dbTRPCompany + 1
              If dbTRPCompany = UBound(arTRPCompany) Then
                 ReDim Preserve arTRPCompany(UBound(arTRPCompany) + 2) As String
              End If
         Case "���`��]"
              arRSC(dbRSC) = tmp_Rs.Fields("�N�X").Value
              cmb_RSC.AddItem tmp_Rs.Fields("�N�X").Value & "  " & tmp_Rs.Fields("����").Value
              dbRSC = dbRSC + 1
              If dbRSC = UBound(arRSC) Then
                 ReDim Preserve arRSC(UBound(arRSC) + 2) As String
              End If
         Case "�d���k��"
              arRBC(dbRBC) = tmp_Rs.Fields("�N�X").Value
              cmb_RBC.AddItem tmp_Rs.Fields("�N�X").Value & "  " & tmp_Rs.Fields("����").Value
              dbRBC = dbRBC + 1
              If dbRBC = UBound(arRBC) Then
                 ReDim Preserve arRBC(UBound(arRBC) + 2) As String
              End If
              
      End Select
      tmp_Rs.MoveNext
   Loop
End If
tmp_Rs.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�d�I��Ӫ����L����ƥ�
'�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������������
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
End If
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    
    Frame2.Top = Me.Top + (Me.ScaleHeight - Frame2.Height) / 2 + 240
    Frame3.Top = Me.Top + (Me.ScaleHeight - Frame3.Height) / 2 + 240
    Frame2.Left = Me.Left + (Me.ScaleWidth - Frame2.Width) / 2 + 240
    Frame3.Left = Me.Left + (Me.ScaleWidth - Frame3.Width) / 2 + 240
    
    If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
        SSTab1.Height = Me.ScaleHeight
        dg_Result.Height = Me.ScaleHeight - cmd_Tab2SavetoExcel.Width - 360

    End If
    
    If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
        SSTab1.Width = Me.ScaleWidth
        dg_Result.Width = Me.ScaleWidth - cmd_Tab2SavetoExcel.Width - 360
        cmd_Tab2SavetoExcel.Left = dg_Result.Width + 240
        cmd_Tab2SavetoExcel.Top = dg_Result.Top
        cmd_Exit(0).Left = dg_Result.Width + 240
    End If

End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_Query_Orders = Nothing
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub GetAllFields()
'���o�Ҧ��i�ϥ����
On Error GoTo err_Handle
lst_AllFields.Clear
lst_SelectedFields.Clear
lst_OrderBy.Clear

'�ϥΪ̿�������Ȧs
Dim rs_UserSelectedFields As ADODB.Recordset
Call ReDim_Recordset(rs_UserSelectedFields)
With rs_UserSelectedFields
     .Fields.Append "�s��", adDouble
     .Fields.Append "���W��", adVarChar, 40
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
'�d�ߵ��G�����ƧǼȦs
Dim rs_OrderByFields As ADODB.Recordset
Call ReDim_Recordset(rs_OrderByFields)
With rs_OrderByFields
     .Fields.Append "�s��", adDouble
     .Fields.Append "���W��", adVarChar, 40
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
End With
'���^�q��d�ߩҦ����
str_SQL = "Select ac.FieldName as ���W�� , Isnull(Cast( cd1.Description as integer),0) as SeqNo,Isnull(Cast( cd2.Description as integer),0) as OrderNo  " & _
          "From Query_UserSelectedField ac " & _
          "Left outer join CodeLKUP cd1 on cd1.ListName = 'ORDERSQUERYFIELDS' and cd1.Code = ac.ViewName and cd1.Long = ac.FieldName and cd1.Short = '" & User_id & "' " & _
          "Left outer join CodeLKUP cd2 on cd2.ListName = 'ORDERSQUERYORDER' and cd2.Code = ac.ViewName and cd2.Long = ac.FieldName and cd2.Short = '" & User_id & "' " & _
          "Where ac.ViewName = 'Query_OrdersData' Order by ac.ColIndex"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
Do While Not tmp_Rs.EOF
   If tmp_Rs.Fields("SeqNo").Value = 0 Then
      lst_AllFields.AddItem tmp_Rs.Fields("���W��").Value
   Else
      rs_UserSelectedFields.AddNew
      rs_UserSelectedFields.Fields("�s��").Value = tmp_Rs.Fields("SeqNo").Value
      rs_UserSelectedFields.Fields("���W��").Value = tmp_Rs.Fields("���W��").Value
      rs_UserSelectedFields.Update
   End If
   If tmp_Rs.Fields("OrderNo").Value <> 0 Then
      rs_OrderByFields.AddNew
      rs_OrderByFields.Fields("�s��").Value = tmp_Rs.Fields("OrderNo").Value
      rs_OrderByFields.Fields("���W��").Value = tmp_Rs.Fields("���W��").Value
      rs_OrderByFields.Update
   End If
   tmp_Rs.MoveNext
Loop
Set tmp_Rs = Nothing

'�d�ߵ��G���
If rs_UserSelectedFields.EOF Then
   Set rs_UserSelectedFields = Nothing
   Exit Sub
Else
   rs_UserSelectedFields.Sort = " �s�� "
   rs_UserSelectedFields.MoveFirst
   Do While Not rs_UserSelectedFields.EOF
      lst_SelectedFields.AddItem rs_UserSelectedFields.Fields("���W��").Value
      rs_UserSelectedFields.MoveNext
   Loop
   Set rs_UserSelectedFields = Nothing
End If

'�ƧǨ̾�
If rs_OrderByFields.EOF Then
   Set rs_OrderByFields = Nothing
   Exit Sub
Else
   rs_OrderByFields.Sort = " �s�� "
   rs_OrderByFields.MoveFirst
   Do While Not rs_OrderByFields.EOF
      lst_OrderBy.AddItem rs_OrderByFields.Fields("���W��").Value
      rs_OrderByFields.MoveNext
   Loop
   Set rs_OrderByFields = Nothing
End If

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��d��-���J���", Me.Caption, "From ���� Subprogram GetAllFields", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub SaveQueryEnv()
'�x�s�ϥΪ̬d�߳]�w��
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

On Error GoTo err_Handle
Tran_Level = 0
Tran_Level = cn.BeginTrans
'�d�ߵ��G������Ȧs��
If lst_SelectedFields.ListCount <> 0 Then
   str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYFIELDS' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   For iLoop = 0 To lst_SelectedFields.ListCount - 1
       str_SQL = "Insert into Codelkup (ListName,Code,Long,Short,Description,AddWho,EditWho) Values ('ORDERSQUERYFIELDS','Query_OrdersData','" & _
                 lst_SelectedFields.List(iLoop) & "','" & User_id & "'," & iLoop + 1 & ",'" & User_id & "','" & User_id & "')"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   Next iLoop
End If
'�d�ߵ��G�ƧǨ̾ڦs��
If lst_OrderBy.ListCount <> 0 Then
   str_SQL = "Delete From Codelkup Where ListName = 'ORDERSQUERYORDER' and Code = 'Query_OrdersData' and Short = '" & User_id & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   For iLoop = 0 To lst_OrderBy.ListCount - 1
       str_SQL = "Insert into Codelkup (ListName,Code,Long,Short,Description,AddWho,EditWho) Values ('ORDERSQUERYORDER','Query_OrdersData','" & _
                 lst_OrderBy.List(iLoop) & "','" & User_id & "'," & iLoop + 1 & ",'" & User_id & "','" & User_id & "')"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   Next iLoop
End If
cn.CommitTrans
Tran_Level = 0
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��d��-�]�w�Ȧs��", Me.Caption, "From ���� Subprogram [SaveQueryEnv]", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
       Case "�q����.�_"
            txt_OrderDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�q����.��"
            txt_OrderDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�e�f���.�_"
            txt_DeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�e�f���.��"
            txt_DeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�ƨ����.�_"
            txt_FPlanDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�ƨ����.��"
            txt_FPlanDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�X�����.�_"
            txt_FDeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�X�����.��"
            txt_FDeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�w�p������.�_"
            txt_FPlanCheckin_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�w�p������.��"
            txt_FPlanCheckin_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.����������.�_"
            txt_FCheckin_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.����������.��"
            txt_FCheckin_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�������ܤ��.�_"
            txt_FCheckout_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�@���ƨ�.�������ܤ��.��"
            txt_FCheckout_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�G���ƨ�.�����X�����.�_"
            txt_SDeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�G���ƨ�.�����X�����.��"
            txt_SDeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�z�f�T�{���.�_"
            txt_SDNDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
       Case "�z�f�T�{���.��"
            txt_SDNDate_End.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then cmd_Tab2SavetoExcel.Visible = False: cmd_Exit(0).Visible = False: Frame3.Visible = True
If SSTab1.Tab = 1 Then cmd_Tab2SavetoExcel.Visible = False: cmd_Exit(0).Visible = False: Frame2.Visible = True: Frame3.Visible = False
If SSTab1.Tab = 2 Then cmd_Tab2SavetoExcel.Visible = True: cmd_Exit(0).Visible = True: Frame3.Visible = False: Frame2.Visible = False
End Sub

Private Sub txt_SDNDate_End_Click()
'�z�f�T�{����G��
If Trim(txt_SDNDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDNDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDNDate_End.Text, 4) & "/" & Mid(txt_SDNDate_End.Text, 5, 2) & "/" & Right(txt_SDNDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_SDNDate_End.Left
mvDate.Top = Frame12.Top + txt_SDNDate_End.Top + txt_SDNDate_End.Height
mvDate.Tag = "�z�f�T�{���.��"
mvDate.Visible = True
End Sub

Private Sub txt_SDNDate_Start_Click()
'�z�f�T�{����G�_
If Trim(txt_SDNDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDNDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDNDate_Start.Text, 4) & "/" & Mid(txt_SDNDate_Start.Text, 5, 2) & "/" & Right(txt_SDNDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_SDNDate_Start.Left
mvDate.Top = Frame12.Top + txt_SDNDate_Start.Top + txt_SDNDate_Start.Height
mvDate.Tag = "�z�f�T�{���.�_"
mvDate.Visible = True
End Sub

Private Sub txt_StorerKey_KeyPress(KeyAscii As Integer)
'�f�D
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_Extern_Start.SelStart = 0: txt_Extern_Start.SelLength = Len(txt_Extern_Start.Text)
         txt_Extern_Start.SetFocus
End Select
End Sub

Private Sub txt_Extern_Start_KeyPress(KeyAscii As Integer)
'�f�D�渹�G�_
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_Extern_End.SelStart = 0: txt_Extern_End.SelLength = Len(txt_Extern_End.Text)
         txt_Extern_End.SetFocus
End Select
End Sub

Private Sub txt_Extern_End_KeyPress(KeyAscii As Integer)
'�f�D�渹�G��
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_OrderDate_Start.SelStart = 0: txt_OrderDate_Start.SelLength = Len(txt_OrderDate_Start.Text)
         txt_OrderDate_Start.SetFocus
End Select
End Sub

Private Sub txt_OrderDate_Start_Click()
'�q�����G�_
If Trim(txt_OrderDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_OrderDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_OrderDate_Start.Text, 4) & "/" & Mid(txt_OrderDate_Start.Text, 5, 2) & "/" & Right(txt_OrderDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_OrderDate_Start.Left
mvDate.Top = Frame11.Top + txt_OrderDate_Start.Top + txt_OrderDate_Start.Height
mvDate.Tag = "�q����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_OrderDate_Start_KeyPress(KeyAscii As Integer)
'�q�����G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_OrderDate_End.SelStart = 0: txt_OrderDate_End.SelLength = Len(txt_OrderDate_End.Text)
         txt_OrderDate_End.SetFocus
End Select
End Sub

Private Sub txt_OrderDate_End_Click()
'�q�����G��
If Trim(txt_OrderDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_OrderDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_OrderDate_End.Text, 4) & "/" & Mid(txt_OrderDate_End.Text, 5, 2) & "/" & Right(txt_OrderDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_OrderDate_End.Left
mvDate.Top = Frame11.Top + txt_OrderDate_End.Top + txt_OrderDate_End.Height
mvDate.Tag = "�q����.��"
mvDate.Visible = True
End Sub

Private Sub txt_OrderDate_End_KeyPress(KeyAscii As Integer)
'�q�����G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_DeliveryDate_Start.SelStart = 0: txt_DeliveryDate_Start.SelLength = Len(txt_DeliveryDate_Start.Text)
         txt_DeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_DeliveryDate_Start_Click()
'�e�f����G�_
If Trim(txt_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_DeliveryDate_Start.Left
mvDate.Top = Frame11.Top + txt_DeliveryDate_Start.Top + txt_DeliveryDate_Start.Height
mvDate.Tag = "�e�f���.�_"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�e�f����G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_DeliveryDate_End.SelStart = 0: txt_DeliveryDate_End.SelLength = Len(txt_DeliveryDate_End.Text)
         txt_DeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_DeliveryDate_End_Click()
'�e�f����G��
If Trim(txt_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate_End.Text, 4) & "/" & Mid(txt_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame11.Left + txt_DeliveryDate_End.Left
mvDate.Top = Frame11.Top + txt_DeliveryDate_End.Top + txt_DeliveryDate_End.Height
mvDate.Tag = "�e�f���.��"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_End_KeyPress(KeyAscii As Integer)
'�e�f����G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_SKU.SelStart = 0: txt_SKU.SelLength = Len(txt_SKU.Text): txt_SKU.SetFocus
End Select
End Sub

Private Sub txt_SKU_KeyPress(KeyAscii As Integer)
'�f��
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_ConsigneeKey.SelStart = 0: txt_ConsigneeKey.SelLength = Len(txt_ConsigneeKey.Text): txt_ConsigneeKey.SetFocus
End Select
End Sub

Private Sub txt_ConsigneeKey_KeyPress(KeyAscii As Integer)
'�Ȥ�s��
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_ConsigName.SelStart = 0: txt_ConsigName.SelLength = Len(txt_ConsigName.Text): txt_ConsigName.SetFocus
End Select
End Sub

Private Sub txt_ConsigName_KeyPress(KeyAscii As Integer)
'�Ȥ�W��
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FRouteNo_Start.SelStart = 0: txt_FRouteNo_Start.SelLength = Len(txt_FRouteNo_Start.Text)
         txt_FRouteNo_Start.SetFocus
End Select
End Sub

Private Sub txt_FRouteNo_Start_KeyPress(KeyAscii As Integer)
'�@���ƨ��G���u�s���G�_
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FRouteNo_End.SelStart = 0: txt_FRouteNo_End.SelLength = Len(txt_FRouteNo_End.Text)
         txt_FRouteNo_End.SetFocus
End Select
End Sub

Private Sub txt_FRouteNo_End_KeyPress(KeyAscii As Integer)
'�@���ƨ��G���u�s���G��
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FPlanDate_Start.SelStart = 0: txt_FPlanDate_Start.SelLength = Len(txt_FPlanDate_Start.Text)
         txt_FPlanDate_Start.SetFocus
End Select
End Sub

Private Sub txt_FPlanDate_Start_Click()
'�@���ƨ��G�ƨ�����G�_
If Trim(txt_FPlanDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_Start.Text, 4) & "/" & Mid(txt_FPlanDate_Start.Text, 5, 2) & "/" & Right(txt_FPlanDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanDate_Start.Left
mvDate.Top = Frame12.Top + txt_FPlanDate_Start.Top + txt_FPlanDate_Start.Height
mvDate.Tag = "�@���ƨ�.�ƨ����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_Start_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�ƨ�����G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanDate_End.SelStart = 0: txt_FPlanDate_End.SelLength = Len(txt_FPlanDate_End.Text)
         txt_FPlanDate_End.SetFocus
End Select
End Sub

Private Sub txt_FPlanDate_End_Click()
'�@���ƨ��G�ƨ�����G��
If Trim(txt_FPlanDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanDate_End.Text, 4) & "/" & Mid(txt_FPlanDate_End.Text, 5, 2) & "/" & Right(txt_FPlanDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanDate_End.Left
mvDate.Top = Frame12.Top + txt_FPlanDate_End.Top + txt_FPlanDate_End.Height
mvDate.Tag = "�@���ƨ�.�ƨ����.��"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanDate_End_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�ƨ�����G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_Start.SelStart = 0: txt_FDeliveryDate_Start.SelLength = Len(txt_FDeliveryDate_Start.Text)
         txt_FDeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_Start_Click()
'�@���ƨ��G�X������G�_
If Trim(txt_FDeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_Start.Text, 4) & "/" & Mid(txt_FDeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FDeliveryDate_Start.Left
mvDate.Top = Frame12.Top + txt_FDeliveryDate_Start.Top + txt_FDeliveryDate_Start.Height
mvDate.Tag = "�@���ƨ�.�X�����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�X������G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FDeliveryDate_End.SelStart = 0: txt_FDeliveryDate_End.SelLength = Len(txt_FDeliveryDate_End.Text)
         txt_FDeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_FDeliveryDate_End_Click()
'�@���ƨ��G�X������G��
If Trim(txt_FDeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FDeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FDeliveryDate_End.Text, 4) & "/" & Mid(txt_FDeliveryDate_End.Text, 5, 2) & "/" & Right(txt_FDeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FDeliveryDate_End.Left
mvDate.Top = Frame12.Top + txt_FDeliveryDate_End.Top + txt_FDeliveryDate_End.Height
mvDate.Tag = "�@���ƨ�.�X�����.��"
mvDate.Visible = True
End Sub

Private Sub txt_FDeliveryDate_End_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�X������G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanCheckin_Start.SelStart = 0: txt_FPlanCheckin_Start.SelLength = Len(txt_FPlanCheckin_Start.Text)
         txt_FPlanCheckin_Start.SetFocus
End Select
End Sub

Private Sub txt_FPlanCheckin_Start_Click()
'�@���ƨ��G�w�p�������G�_
If Trim(txt_FPlanCheckin_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanCheckin_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanCheckin_Start.Text, 4) & "/" & Mid(txt_FPlanCheckin_Start.Text, 5, 2) & "/" & Right(txt_FPlanCheckin_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanCheckin_Start.Left
mvDate.Top = Frame12.Top + txt_FPlanCheckin_Start.Top + txt_FPlanCheckin_Start.Height
mvDate.Tag = "�@���ƨ�.�w�p������.�_"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanCheckin_Start_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�w�p�������G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FPlanCheckin_End.SelStart = 0: txt_FPlanCheckin_End.SelLength = Len(txt_FPlanCheckin_End.Text)
         txt_FPlanCheckin_End.SetFocus
End Select
End Sub

Private Sub txt_FPlanCheckin_End_Click()
'�@���ƨ��G�w�p�������G��
If Trim(txt_FPlanCheckin_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FPlanCheckin_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FPlanCheckin_End.Text, 4) & "/" & Mid(txt_FPlanCheckin_End.Text, 5, 2) & "/" & Right(txt_FPlanCheckin_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FPlanCheckin_End.Left
mvDate.Top = Frame12.Top + txt_FPlanCheckin_End.Top + txt_FPlanCheckin_End.Height
mvDate.Tag = "�@���ƨ�.�w�p������.��"
mvDate.Visible = True
End Sub

Private Sub txt_FPlanCheckin_End_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�w�p�������G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckin_Start.SelStart = 0: txt_FCheckin_Start.SelLength = Len(txt_FCheckin_Start.Text)
         txt_FCheckin_Start.SetFocus
End Select
End Sub

Private Sub txt_FCheckin_Start_Click()
'�@���ƨ��G�����������G�_
If Trim(txt_FCheckin_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckin_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckin_Start.Text, 4) & "/" & Mid(txt_FCheckin_Start.Text, 5, 2) & "/" & Right(txt_FCheckin_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckin_Start.Left
mvDate.Top = Frame12.Top + txt_FCheckin_Start.Top + txt_FCheckin_Start.Height
mvDate.Tag = "�@���ƨ�.����������.�_"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckin_Start_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�����������G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckin_End.SelStart = 0: txt_FCheckin_End.SelLength = Len(txt_FCheckin_End.Text)
         txt_FCheckin_End.SetFocus
End Select
End Sub

Private Sub txt_FCheckin_End_Click()
'�@���ƨ��G�����������G��
If Trim(txt_FCheckin_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckin_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckin_End.Text, 4) & "/" & Mid(txt_FCheckin_End.Text, 5, 2) & "/" & Right(txt_FCheckin_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckin_End.Left
mvDate.Top = Frame12.Top + txt_FCheckin_End.Top + txt_FCheckin_End.Height
mvDate.Tag = "�@���ƨ�.����������.��"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckin_End_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�����������G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckout_Start.SelStart = 0: txt_FCheckout_Start.SelLength = Len(txt_FCheckout_Start.Text)
         txt_FCheckout_Start.SetFocus
End Select
End Sub

Private Sub txt_FCheckout_Start_Click()
'�@���ƨ��G�������ܤ���G�_
If Trim(txt_FCheckout_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckout_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckout_Start.Text, 4) & "/" & Mid(txt_FCheckout_Start.Text, 5, 2) & "/" & Right(txt_FCheckout_Start.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckout_Start.Left
mvDate.Top = Frame12.Top + txt_FCheckout_Start.Top + txt_FCheckout_Start.Height
mvDate.Tag = "�@���ƨ�.�������ܤ��.�_"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckout_Start_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�������ܤ���G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_FCheckout_End.SelStart = 0: txt_FCheckout_End.SelLength = Len(txt_FCheckout_End.Text)
         txt_FCheckout_End.SetFocus
End Select
End Sub

Private Sub txt_FCheckout_End_Click()
'�@���ƨ��G�������ܤ���G��
If Trim(txt_FCheckout_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_FCheckout_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_FCheckout_End.Text, 4) & "/" & Mid(txt_FCheckout_End.Text, 5, 2) & "/" & Right(txt_FCheckout_End.Text, 2))
   End If
End If
mvDate.Left = Frame12.Left + txt_FCheckout_End.Left
mvDate.Top = Frame12.Top + txt_FCheckout_End.Top + txt_FCheckout_End.Height
mvDate.Tag = "�@���ƨ�.�������ܤ��.��"
mvDate.Visible = True
End Sub

Private Sub txt_FCheckout_End_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�������ܤ���G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
End Select
End Sub

Private Sub txt_FVehicleID_KeyPress(KeyAscii As Integer)
'�@���ƨ��G���P���X
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_FDriver.SetFocus
End Select
End Sub

Private Sub txt_FReceiptNo_KeyPress(KeyAscii As Integer)
'�@���ƨ��G�ƨ��q��s��
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
End Select
End Sub

Private Sub txt_SRouteNo_Start_KeyPress(KeyAscii As Integer)
'�G���ƨ��G���u�s���G�_
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_SRouteNo_End.SelStart = 0: txt_SRouteNo_End.SelLength = Len(txt_SRouteNo_End.Text)
         txt_SRouteNo_End.SetFocus
End Select
End Sub

Private Sub txt_SRouteNo_End_KeyPress(KeyAscii As Integer)
'�G���ƨ��G���u�s���G��
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_SDeliveryDate_Start.SelStart = 0: txt_SDeliveryDate_Start.SelLength = Len(txt_SDeliveryDate_Start.Text)
         txt_SDeliveryDate_Start.SetFocus
End Select
End Sub

Private Sub txt_SDeliveryDate_Start_Click()
'�G���ƨ��G�����X������G�_
If Trim(txt_SDeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDeliveryDate_Start.Text, 4) & "/" & Mid(txt_SDeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_SDeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = Frame13.Left + txt_SDeliveryDate_Start.Left
mvDate.Top = Frame13.Top + (txt_SDeliveryDate_Start.Top + txt_SDeliveryDate_Start.Height) - (mvDate.Height + txt_SDeliveryDate_Start.Height)
mvDate.Tag = "�G���ƨ�.�����X�����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_SDeliveryDate_Start_KeyPress(KeyAscii As Integer)
'�G���ƨ��G�����X������G�_
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_SDeliveryDate_End.SelStart = 0: txt_SDeliveryDate_End.SelLength = Len(txt_SDeliveryDate_End.Text)
         txt_SDeliveryDate_End.SetFocus
End Select
End Sub

Private Sub txt_SDeliveryDate_End_Click()
'�G���ƨ��G�����X������G��
If Trim(txt_SDeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_SDeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_SDeliveryDate_End.Text, 4) & "/" & Mid(txt_SDeliveryDate_End.Text, 5, 2) & "/" & Right(txt_SDeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = Frame13.Left + txt_SDeliveryDate_End.Left
mvDate.Top = Frame13.Top + (txt_SDeliveryDate_End.Top + txt_SDeliveryDate_End.Height) - (mvDate.Height + txt_SDeliveryDate_End.Height)
mvDate.Tag = "�G���ƨ�.�����X�����.��"
mvDate.Visible = True
End Sub

Private Sub txt_SDeliveryDate_End_KeyPress(KeyAscii As Integer)
'�G���ƨ��G�����X������G��
Select Case KeyAscii
    Case 97 To 122, 65 To 90   '�����\��J�r��
         KeyAscii = 0
    Case vbKeyReturn
         txt_SVehicleID.SelStart = 0: txt_SVehicleID.SelLength = Len(txt_SVehicleID.Text)
         txt_SVehicleID.SetFocus
End Select
End Sub

Private Sub txt_SVehicleID_KeyPress(KeyAscii As Integer)
'�G���ƨ��G���P���X
Select Case KeyAscii
    Case 97 To 122     '�p�g�r���אּ�j�g�r��
         KeyAscii = KeyAscii - 32
    Case vbKeyReturn
         txt_SAddWho.SelStart = 0: txt_SAddWho.SelLength = Len(txt_SAddWho.Text)
         txt_SAddWho.SetFocus
End Select
End Sub

Private Function GetFieldWidth(ByVal strFieldName) As Double
'���o�d�ߵ��G���e��
Select Case strFieldName
       Case "�s��", "�f�D", "����", "Area", "ZIP", "���O", "����"
            GetFieldWidth = 500
       Case "�f��", "�X����", "�x�}�X", "�q��q", "�z�f�q", "��c��", "PalletTI", "PalletHI", "�r�p�H", "�˸��q", _
            "ñ���q", "���`�X", "�d�ݽX"
            GetFieldWidth = 800
       Case "OrderKey", "�f�D�渹", "�q����", "�e�f���", "�p���H", "�Ȥ�q��", "�z�f�O��", "�z�f���q", "�z�f���n", _
            "���P���X", "�q��", "�X�Y�Ȧs", "�˸��O��", "�˸����q", "�˸����n", "������", "����ɶ�", "���ܤ�� ", _
            "���ܮɶ�", "���إN�X", "ñ�檬�A", "ñ�����", "ñ�����O"
            GetFieldWidth = 1000
       Case "��c���n", "��J�ѧO", "�ƨ��q", "�ƨ��O��", "�ƨ����q", "�ƨ����n", "�ƨ����", "�ƨ��ɶ�", "�ƨ���", _
            "�^�Ǥ��", "�^�Ǯɶ�", "�^�Ǫ��A", "�X�����", "�G���ƨ���", "�G������", "�G���q��", "ñ���O��", "ñ�����q", _
            "ñ�����n"
            GetFieldWidth = 1000
       Case "�Ȥ�s��", "�ƨ��q��s��", "���u�s��", "�w�p������", "�w�p����ɶ�", "�G�����u�s��", "�G���X�����", _
            "�G�����P���X", "�G���r�p�H", "�G���t�e����", "�G���X�Y�Ȧs", "�G��������", "�G������ɶ�", "�G�����ܤ��", _
            "�G�����ܮɶ�", "�S��ݨD�X1", "�S��ݨD�X2", "�f�B���q�N�X", "�G�����إN�X"
            GetFieldWidth = 1200
       Case "�l���ϸ�", "�q��Ƶ�", "���O���e", "�G���w�p������", "�G���w������ɶ�", "�G���f�B���q�N�X", _
            "���`��]", "�d���k��", "ñ���J�ɶ�", "ñ���J�H��"
            GetFieldWidth = 1600
       Case "�Ȥ�²��", "�G���B�餽�q"
            GetFieldWidth = 2000
       Case "�Ȥ�W��", "�B�e�ϰ�", "�a�}", "�Ȥ�t�e����", "�S��ݨD1", "�S��ݨD2", "�~�W", "�f�B���q", "�t�e����"
            GetFieldWidth = 2500
       Case Else
            GetFieldWidth = 1000
End Select
End Function

Private Function GetFieldAlignment(ByVal strFieldName) As Double
'���o�d�ߵ��G���e��
Select Case strFieldName
       Case "���O", "�f��", "�r�p�H", "�p���H", "�Ȥ�q��", "�q��", "�X�Y�Ȧs", "���ܮɶ�", "���إN�X", "�ƨ���", "�^�Ǫ��A", _
            "�G���ƨ���", "�G������", "�G���q��", "�G�����P���X", "�G���r�p�H", "�G���t�e����", "�G���X�Y�Ȧs", _
            "�S��ݨD�X1", "�S��ݨD�X2", "�f�B���q�N�X", "�G�����إN�X", "�l���ϸ�", "�q��Ƶ�", "���O���e", _
            "�G���w�p������", "�G���w������ɶ�", "�G���f�B���q�N�X", "�Ȥ�²��", "�G���B�餽�q", "�Ȥ�W��", _
            "�B�e�ϰ�", "�a�}", "�Ȥ�t�e����", "�S��ݨD1", "�S��ݨD2", "�~�W", "�f�B���q", "�t�e����", _
            "���`��]", "�d���k��", "ñ���J�H��", "ñ�����O"
            GetFieldAlignment = dbgLeft
       Case "����", "�X����", "�q��q", "�z�f�q", "��c��", "PalletTI", "PalletHI", "�˸��O��", "�˸����q", "�˸����n", _
            "�ƨ��q", "�ƨ��O��", "�ƨ����q", "�ƨ����n", "�˸��q", "�z�f�O��", "�z�f���q", "�z�f���n", "��c���n", _
            "ñ���q", "ñ���O��", "ñ�����q", "ñ�����n"
            GetFieldAlignment = dbgRight
       Case "�s��", "�f�D", "Area", "ZIP", "����", "�x�}�X", "OrderKey", "�f�D�渹", "�q����", "�e�f���", "���P���X", _
            "������", "����ɶ�", "���ܤ�� ", "��J�ѧO", "�ƨ����", "�ƨ��ɶ�", "�^�Ǥ��", "�^�Ǯɶ�", "�X�����", _
            "�Ȥ�s��", "�ƨ��q��s��", "���u�s��", "�w�p������", "�w�p����ɶ�", "�G�����u�s��", "�G���X�����", _
            "�G��������", "�G������ɶ�", "�G�����ܤ��", "�G�����ܮɶ�", "���`�X", "�d�ݽX", "ñ�檬�A", "ñ�����", _
            "ñ���J�ɶ�"
            GetFieldAlignment = dbgCenter
       Case Else
            GetFieldAlignment = dbgGeneral
End Select
End Function


