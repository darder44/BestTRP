VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_OP_RouteData 
   Caption         =   "���u�s�����@�@�~"
   ClientHeight    =   7140
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11535
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   1320
      TabIndex        =   21
      Top             =   2895
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
      StartOfWeek     =   32768001
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin VB.Frame fra_ExtraQuery 
      Appearance      =   0  '����
      BackColor       =   &H00E0E0E0&
      Caption         =   "�d�߱���]�w"
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   120
      TabIndex        =   58
      Top             =   1200
      Visible         =   0   'False
      Width           =   3600
      Begin VB.CheckBox chk_AddWho 
         Caption         =   "�ƨ��H���z��"
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
         Left            =   1035
         TabIndex        =   61
         Top             =   630
         Value           =   1  '�֨�
         Width           =   1875
      End
      Begin VB.TextBox txt_ExternOrderKey 
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
         TabIndex        =   59
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�f�D�渹"
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
         Left            =   135
         TabIndex        =   60
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '����
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   75
      TabIndex        =   17
      Top             =   1170
      Width           =   2610
      Begin VB.CommandButton cmd_Cancel 
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
         Left            =   1320
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   19
         Top             =   165
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Modify 
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
         TabIndex        =   18
         Top             =   165
         Width           =   1200
      End
   End
   Begin VB.Frame fam_Header 
      Appearance      =   0  '����
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   75
      TabIndex        =   0
      Top             =   -30
      Width           =   11400
      Begin VB.CommandButton cmd_Tab0_ShowQuery 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�H"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4155
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   57
         Top             =   720
         Width           =   360
      End
      Begin VB.CommandButton cmd_Query 
         BackColor       =   &H00FF8080&
         Caption         =   "���u�s���d��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4815
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   20
         Top             =   225
         Width           =   1650
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
         Height          =   750
         Index           =   0
         Left            =   10050
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   16
         Top             =   225
         Width           =   1170
      End
      Begin VB.CommandButton cmd_Delete 
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
         Height          =   750
         Left            =   8310
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   15
         Top             =   225
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.CommandButton cmd_Save 
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
         Height          =   750
         Left            =   6570
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   225
         Width           =   1650
      End
      Begin VB.CommandButton cmd_Reset 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�M��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3570
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   12
         Top             =   495
         Width           =   420
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
         Left            =   1035
         TabIndex        =   6
         Top             =   165
         Width           =   1635
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
         Left            =   2970
         TabIndex        =   5
         Top             =   165
         Width           =   1635
      End
      Begin VB.TextBox txt_PlanDate_Start 
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
         Left            =   1035
         TabIndex        =   4
         Top             =   480
         Width           =   1125
      End
      Begin VB.TextBox txt_PlanDate_End 
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
         Left            =   2430
         TabIndex        =   3
         Top             =   480
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
         Left            =   1035
         TabIndex        =   2
         Top             =   795
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
         Left            =   2430
         TabIndex        =   1
         Top             =   795
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
         Index           =   1
         Left            =   2715
         TabIndex        =   62
         Top             =   225
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
         Left            =   135
         TabIndex        =   11
         Top             =   225
         Width           =   840
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
         Left            =   135
         TabIndex        =   10
         Top             =   525
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
         Left            =   2190
         TabIndex        =   9
         Top             =   510
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
         Left            =   135
         TabIndex        =   8
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
         Index           =   15
         Left            =   2190
         TabIndex        =   7
         Top             =   825
         Width           =   240
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Height          =   885
         Left            =   4725
         Top             =   165
         Width           =   6600
      End
   End
   Begin VB.Frame fam_RouteData 
      Enabled         =   0   'False
      Height          =   1680
      Left            =   75
      TabIndex        =   13
      Top             =   1065
      Width           =   11400
      Begin VB.TextBox txt_DockNo 
         Alignment       =   2  '�m�����
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   7275
         TabIndex        =   31
         Top             =   180
         Width           =   570
      End
      Begin VB.TextBox txt_CarCheckInTime 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10620
         TabIndex        =   28
         Top             =   180
         Width           =   660
      End
      Begin VB.TextBox txt_CarCheckInDate 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8670
         TabIndex        =   27
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox txt_DeliveryDate 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3255
         TabIndex        =   24
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox txt_VehicleNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   5055
         TabIndex        =   23
         Top             =   180
         Width           =   1080
      End
      Begin VB.CommandButton cmd_SelectCar 
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
         Height          =   375
         Left            =   6150
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   22
         Top             =   165
         Width           =   330
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '����
         BackColor       =   &H008080FF&
         BorderStyle     =   0  '�S���ؽu
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   15
         TabIndex        =   34
         Top             =   645
         Width           =   11355
         Begin VB.TextBox txt_Phone 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
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
            Height          =   285
            Left            =   7605
            TabIndex        =   46
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_TRPCompany 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
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
            Height          =   285
            Left            =   4680
            TabIndex        =   45
            Top             =   315
            Width           =   1725
         End
         Begin VB.TextBox txt_Driver 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
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
            Height          =   285
            Left            =   6420
            TabIndex        =   44
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_VehicleType 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
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
            Height          =   285
            Left            =   8790
            TabIndex        =   43
            Top             =   315
            Width           =   2520
         End
         Begin VB.TextBox txt_Weight 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
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
            Height          =   285
            Left            =   10380
            TabIndex        =   42
            Top             =   615
            Width           =   930
         End
         Begin VB.TextBox txt_Volumn 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
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
            Height          =   285
            Left            =   9435
            TabIndex        =   41
            Top             =   615
            Width           =   930
         End
         Begin VB.TextBox txt_PalletQty 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
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
            Height          =   285
            Left            =   8490
            TabIndex        =   40
            Top             =   615
            Width           =   930
         End
         Begin VB.TextBox txt_CaseQty 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
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
            Height          =   285
            Left            =   7545
            TabIndex        =   39
            Top             =   615
            Width           =   930
         End
         Begin VB.TextBox txt_SecondRouteNo 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   960
            TabIndex        =   38
            Top             =   300
            Width           =   1425
         End
         Begin VB.TextBox txt_PlanDate 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   960
            TabIndex        =   37
            Top             =   600
            Width           =   2070
         End
         Begin VB.TextBox txt_Status 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3000
            TabIndex        =   36
            Top             =   300
            Width           =   1245
         End
         Begin VB.TextBox txt_DriveTimes 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4230
            TabIndex        =   35
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lab_RouteNo 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            BackStyle       =   0  '�z��
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
            Left            =   2760
            TabIndex        =   56
            Top             =   45
            Width           =   1875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q  ��"
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
            Index           =   17
            Left            =   7950
            TabIndex        =   55
            Top             =   90
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�餽�q"
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
            Index           =   16
            Left            =   5085
            TabIndex        =   54
            Top             =   90
            Width           =   840
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   15
            Left            =   6705
            TabIndex        =   53
            Top             =   90
            Width           =   630
         End
         Begin VB.Label Label1 
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
            Index           =   14
            Left            =   9750
            TabIndex        =   52
            Top             =   90
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����"
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
            Height          =   195
            Index           =   0
            Left            =   3750
            TabIndex        =   51
            Top             =   645
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�c�� / �O�� / ���n / ���q"
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
            Index           =   1
            Left            =   5265
            TabIndex        =   50
            Top             =   675
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�G�����s"
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
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   49
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ƨ��ɶ�"
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
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   48
            Top             =   645
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���A"
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
            Height          =   195
            Index           =   5
            Left            =   2535
            TabIndex        =   47
            Top             =   345
            Width           =   420
         End
      End
      Begin VB.Label Label1 
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
         Height          =   435
         Index           =   2
         Left            =   6810
         TabIndex        =   32
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '�z��
         Caption         =   "�w�p����ɶ�"
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
         Index           =   19
         Left            =   9945
         TabIndex        =   30
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '�z��
         Caption         =   "�w�p������"
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
         Index           =   20
         Left            =   7995
         TabIndex        =   29
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label1 
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
         Height          =   435
         Index           =   12
         Left            =   2775
         TabIndex        =   26
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label1 
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
         Height          =   390
         Index           =   13
         Left            =   4590
         TabIndex        =   25
         Top             =   180
         Width           =   420
      End
   End
   Begin MSDataGridLib.DataGrid dg_Route 
      Height          =   4320
      Left            =   75
      TabIndex        =   33
      Top             =   2760
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   7620
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
End
Attribute VB_Name = "frm_OP_RouteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private iLoop As Double

Private blRouteEventEnable As Boolean
Private strRouteNo As String         '��������u�s��
Private rs_Route As ADODB.Recordset

Private Sub cmd_Cancel_Click()
'����
If dg_Route.SelBookmarks.Count = 0 Then Exit Sub
fam_RouteData.Enabled = False
fam_RouteData.BackColor = &H8000000F
cmd_Modify.Enabled = True
cmd_Cancel.Enabled = False
cmd_Save.Enabled = False
cmd_Delete.Enabled = True
           
Call ClearData_RouteData
lab_RouteNo.Caption = "���s�G" & rs_Route.Fields("���u�s��").Value
lab_RouteNo.AutoSize = True
txt_DeliveryDate.Text = rs_Route.Fields("�X�����").Value
txt_VehicleNo.Text = rs_Route.Fields("���P���X").Value
txt_DockNo.Text = rs_Route.Fields("�X�Y�Ȧs").Value
txt_CarCheckInDate.Text = rs_Route.Fields("�w�p������").Value
txt_CarCheckInTime.Text = rs_Route.Fields("�w�p����ɶ�").Value
txt_SecondRouteNo.Text = rs_Route.Fields("�G���ƨ����s").Value
txt_Status.Text = rs_Route.Fields("EXE�^��").Value
txt_PlanDate.Text = rs_Route.Fields("�ƨ����").Value & " " & rs_Route.Fields("�ƨ��ɶ�").Value
txt_DriveTimes.Text = rs_Route.Fields("����").Value
txt_TRPCompany.Text = rs_Route.Fields("�B�餽�q").Value
txt_Driver.Text = rs_Route.Fields("�r�p�H").Value
txt_Phone.Text = rs_Route.Fields("�q��").Value
txt_VehicleType.Text = rs_Route.Fields("����").Value
txt_CaseQty.Text = rs_Route.Fields("�c��").Value
txt_PalletQty.Text = rs_Route.Fields("�O��").Value
txt_Volumn.Text = rs_Route.Fields("���n").Value
txt_Weight.Text = rs_Route.Fields("���q").Value

End Sub

Private Sub cmd_Delete_Click()
'�R��
If rs_Route Is Nothing Then Exit Sub
If rs_Route.RecordCount = 0 Then Exit Sub
If dg_Route.SelBookmarks.Count = 0 Then Exit Sub
If strRouteNo = "" Then Exit Sub

If dg_Route.SelBookmarks.Count = 0 Then
   msg_text = "�{�ǿ��~�G��������R�������u�s��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
strDeleteRouteNo = strRouteNo
strCarno = txt_VehicleNo.Text
dbDriveTimes = Val(txt_DriveTimes.Text)

''���R�������s: �O�_�w�^��WMS
'str_SQL = "Select EXE_CONFIRM From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
'tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_rs.Fields("EXE_CONFIRM").Value = "1" Or tmp_rs.Fields("EXE_CONFIRM").Value = "2" Then
'    tmp_rs.Close
'    msg_text = "ĵ�i�G�����u�s���w�^��WMS!"
'    MsgBox msg_text, 64, msg_title
'    Exit Sub
'End If

str_SQL = "Select isnull(Route,'') From " & strWMSDB & "..orders Where Route = '" & strDeleteRouteNo & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
    msg_text = "�`�N�GWMS�t�Φ������u�s���ɡA�L�k�ק�ΧR��!"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
tmp_Rs.Close

'���R�������s�G�O�_�w�X���T�{
str_SQL = "Select c_Route_No  From SDN01T Where c_Route_No = '" & strDeleteRouteNo & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
    tmp_Rs.Close
    msg_text = "�`�N�G�����u�s���w�X���T�{�A�L�k�R��! "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If

msg_text = "�T�{�R�����u�s���G" & strDeleteRouteNo
'If rs_Tab1_Route.Fields("EXE�^��").Value = "�w�^��" Then
'   msg_text = msg_text & vbCrLf & "�`�N�G�w�^�ǭq��A���s�ƨ��ɡA�N�����\�P��L���^�ǭq���ܬ��P�@���u�s��"
'End If
If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub

Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
'���ұ��R�������s�A�ƨ��̬O�_�����ɵn�J���ϥΪ�

If Left(strRouteNo, 1) = "R" Then
    str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From ORT01T Where Route_No = '" & strDeleteRouteNo & "'"
Else
    str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
End If

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "��Ʋ��`�G�䤣����R�������u�s��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
Else
   If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl = True Then
      tmp_Rs.Close
      msg_text = "�v�����ޡG���u�s�����R���u���\�ѭ�Ʃw�̰���"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      Exit Sub
   End If
End If
tmp_Rs.Close

'���R�������s�G��������B���ܮɶ��O�_�w�n��
If Left(strRouteNo, 1) = "R" Then
    str_SQL = "Select Convert(varchar(8),Vehicle_Check_in,112) as Checkin,Convert(varchar(8),Vehicle_Check_out,112) as Checkout From ORT05T Where Route_No = '" & strDeleteRouteNo & "'"
Else
    str_SQL = "Select Convert(varchar(8),Vehicle_Check_in,112) as Checkin,Convert(varchar(8),Vehicle_Check_out,112) as Checkout From TRP05T Where Route_No = '" & strDeleteRouteNo & "'"
End If

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("Checkin").Value <> "" Or tmp_Rs.Fields("CheckOut").Value <> "" Then
   tmp_Rs.Close
   msg_text = "��Ʋ��`�G�����u�s���w���� [��������] �� [��������]�A���R�������s�A�вM�������i�X������A�i��R��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
tmp_Rs.Close

Screen.MousePointer = vbHourglass
blRouteEventEnable = False
Tran_Level = 0
Tran_Level = cn.BeginTrans

'�R�� TRP01T ���u�s���D��
Call DB_CheckConnectStatus

If Left(strRouteNo, 1) = "R" Then

    '(1).�N ORT03T �g�^ ORT03W >> �R�� ORT03T
    str_SQL = "Insert into ORT03W(" & _
              "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From ORT03T A Where a.Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).�N ORT02T �g�^ ORT02W >> �R�� ORT02T
    str_SQL = "Insert into ORT02W(" & _
              "   RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no) " & _
              "Select RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no " & _
              "From ORT02T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(3).�R�� ORT02T & ORT03T
    str_SQL = "Delete From ORT03T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Delete From ORT02T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
              
    '(4).�R�� ORT05T
    str_SQL = "Delete From ORT05T Where Route_No = '" & strDeleteRouteNo & "' and Vehicle_ID_No = '" & strCarno & "' and Drive_Times = " & dbDriveTimes
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(5).�R�� TRP01T
    str_SQL = "Delete From ORT01T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

Else

    '(1).�N TRP03T �g�^ TRP03W >> �R�� TRP03T
    str_SQL = "Insert into TRP03W(" & _
              "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From TRP03T A Where a.Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).�N TRP02T �g�^ TRP02W >> �R�� TRP02T
    str_SQL = "Insert into TRP02W(" & _
              "   RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no) " & _
              "Select RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,0,Priority,c_receipt_no " & _
              "From TRP02T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(3).�R�� TRP02T & TRP03T
    str_SQL = "Delete From TRP03T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Delete From TRP02T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
              
    '(4).�R�� TRP05T
    str_SQL = "Delete From TRP05T Where Route_No = '" & strDeleteRouteNo & "' and Vehicle_ID_No = '" & strCarno & "' and Drive_Times = " & dbDriveTimes
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(5).�R�� TRP01T
    str_SQL = "Delete From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

End If

'(6).�R���d�ߵ��G���ӵ����u�s��
rs_Route.Delete
If Not rs_Route.EOF Then rs_Route.MoveFirst

blRouteEventEnable = True
cn.CommitTrans
Tran_Level = 0

fam_RouteData.Enabled = False
fam_RouteData.BackColor = &H8000000F
cmd_Modify.Enabled = True
cmd_Cancel.Enabled = False
cmd_Save.Enabled = False
cmd_Delete.Enabled = False

If dg_Route.SelBookmarks.Count > 0 Then dg_Route.SelBookmarks.Remove 0
Call ClearData_RouteData

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
   CreateErrorLog Me.Name & "-�R��", Me.Caption, "cmd_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_Modify_Click()

'�ק�
If rs_Route Is Nothing Then Exit Sub
If rs_Route.RecordCount = 0 Then Exit Sub
If dg_Route.SelBookmarks.Count = 0 Then Exit Sub

'�O�_�X���T�{
str_SQL = "select * from trp05t where route_no in ('" & RTrim(rs_Route("���u�s��")) & "','" & RTrim(rs_Route("�G���ƨ����s")) & "') and sdnstatus > 0"

Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then MsgBox "�����s�w�X���T�{�A�ЦP�B�ק�X���T�{�᪺�����C", 16, "�`�N"
tmp_Rs.Close

fam_RouteData.Enabled = True
fam_RouteData.BackColor = &HC0E0FF
cmd_Modify.Enabled = False
cmd_Cancel.Enabled = True
cmd_Save.Enabled = True
cmd_Delete.Enabled = False
End Sub

Private Sub cmd_Query_Click()
'���u�s���d��
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Route.DataSource = Nothing
Set rs_Route = Nothing
fra_ExtraQuery.Visible = False

Call ClearData_RouteData
DoEvents

'���^�ݱƨ��q��
str_SQL = "Select ���u�s��,�X�����,���P���X,����,�r�p�H,�q��,�B�餽�q,�c��,�O��,���n,���q,����,�X�Y�Ȧs," & _
          "   �w�p������, �w�p����ɶ�, EXE�^��, �ƨ����, �ƨ��ɶ�, �ƨ���, �G���ƨ����s " & _
          "From RouteData_Maintain "
Dim str_Where As String, strSubwhere As String, intloop As Integer
str_Where = ""
'���u�s��
strSubwhere = ""
If Len(txt_RouteNo_Start.Text) > 0 And Len(txt_RouteNo_End.Text) > 0 Then
   strSubwhere = " ���u�s�� Between '" & txt_RouteNo_Start.Text & "' and '" & txt_RouteNo_End.Text & "' "
ElseIf Len(txt_RouteNo_Start.Text) > 0 And Len(txt_RouteNo_End.Text) = 0 Then
   strSubwhere = " ���u�s�� = '" & txt_RouteNo_Start.Text & "' "
ElseIf Len(txt_RouteNo_Start.Text) = 0 And Len(txt_RouteNo_End.Text) > 0 Then
   strSubwhere = " ���u�s�� = '" & txt_RouteNo_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�ƨ����
strSubwhere = ""
If Len(txt_PlanDate_Start.Text) > 0 And Len(txt_PlanDate_End.Text) > 0 Then
   strSubwhere = " �ƨ���� Between '" & txt_PlanDate_Start.Text & "' and '" & txt_PlanDate_End.Text & "' "
ElseIf Len(txt_PlanDate_Start.Text) > 0 And Len(txt_PlanDate_End.Text) = 0 Then
   strSubwhere = " �ƨ���� = '" & txt_PlanDate_Start.Text & "' "
ElseIf Len(txt_PlanDate_Start.Text) = 0 And Len(txt_PlanDate_End.Text) > 0 Then
   strSubwhere = " �ƨ���� = '" & txt_PlanDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If
'�X�����
strSubwhere = ""
If Len(txt_DeliveryDate_Start.Text) > 0 And Len(txt_DeliveryDate_End.Text) > 0 Then
   strSubwhere = " �X����� Between '" & txt_DeliveryDate_Start.Text & "' and '" & txt_DeliveryDate_End.Text & "' "
ElseIf Len(txt_DeliveryDate_Start.Text) > 0 And Len(txt_DeliveryDate_End.Text) = 0 Then
   strSubwhere = " �X����� = '" & txt_DeliveryDate_Start.Text & "' "
ElseIf Len(txt_DeliveryDate_Start.Text) = 0 And Len(txt_DeliveryDate_End.Text) > 0 Then
   strSubwhere = " �X����� = '" & txt_DeliveryDate_End.Text & "' "
End If
If Len(strSubwhere) > 0 Then
   If Len(str_Where) = 0 Then
      str_Where = strSubwhere
   Else
      str_Where = str_Where & " and " & strSubwhere
   End If
End If

'�ƨ��H���z��
'If chk_AddWho.Value = vbChecked Then
'   strSubwhere = " �ƨ��� = '" & user_id & "' "
'   If Len(strSubwhere) > 0 Then
'      If Len(str_Where) = 0 Then
'         str_Where = strSubwhere
'      Else
'         str_Where = str_Where & " and " & strSubwhere
'      End If
'   End If
'End If

'�f�D�渹
If Len(Trim(txt_ExternOrderKey.Text)) > 0 Then
   strSubwhere = " ���u�s�� in (Select Distinct Route_No From TRP02T Where Extern = '" & txt_ExternOrderKey.Text & "') "
   If Len(strSubwhere) > 0 Then
      If Len(str_Where) = 0 Then
         str_Where = strSubwhere
      Else
         str_Where = str_Where & " and " & strSubwhere
      End If
   End If
End If

If Len(str_Where) > 0 Then
   str_SQL = str_SQL & " Where " & str_Where
End If
str_SQL = str_SQL & " Order by ���u�s�� "

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Route)
tmp_Rs.Close

blRouteEventEnable = False
With dg_Route
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Route.MoveFirst
Set dg_Route.DataSource = rs_Route
With dg_Route
    .RowHeight = 250
    .Columns(0).Width = 500         '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000         '���u�s��
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900         '�X�����
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 800        '���P���X
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500         '����
    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 800         '�r�p�H
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1300         '�r�p�q��
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1600        '�B�餽�q
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 800         '�c��
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 800         '�O��
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 800        '���n
    .Columns(10).Alignment = dbgRight
    .Columns(11).Width = 800         '���q
    .Columns(11).Alignment = dbgRight
    .Columns(12).Width = 1500       '����
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 800        '�X�Y�Ȧs
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1300       '�w�p������
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1300       '�w�p����ɶ�
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1000       'exe�^��
    .Columns(16).Alignment = dbgLeft
    .Columns(17).Width = 1000       '�ƨ����
    .Columns(17).Alignment = dbgCenter
    .Columns(18).Width = 1000       '�ƨ��ɶ�
    .Columns(18).Alignment = dbgCenter
    .Columns(19).Width = 800        '�ƨ���
    .Columns(19).Alignment = dbgLeft
    .Columns(20).Width = 1300        '�G���ƨ����s
    .Columns(20).Alignment = dbgLeft
End With
If dg_Route.SelBookmarks.Count > 0 Then dg_Route.SelBookmarks.Remove 0
strRouteNo = ""
Call ClearData_RouteData
fam_RouteData.Enabled = False
fam_RouteData.BackColor = &H8000000F
cmd_Modify.Enabled = True
cmd_Cancel.Enabled = False
cmd_Save.Enabled = False
cmd_Delete.Enabled = True
blRouteEventEnable = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-���u�s���d��", Me.Caption, "cmd_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Reset_Click()
'�M��
txt_RouteNo_Start.Text = "": txt_RouteNo_End.Text = ""
txt_PlanDate_Start.Text = "": txt_PlanDate_End.Text = ""
txt_DeliveryDate_Start.Text = "": txt_DeliveryDate_End.Text = ""
Set dg_Route.DataSource = Nothing
Set rs_Route = Nothing
Call ClearData_RouteData
txt_ExternOrderKey.Text = ""
fra_ExtraQuery.Visible = False
End Sub

Private Sub cmd_Save_Click()
'�s��
If strRouteNo = "" Then Exit Sub

'1.���ҥ��n����ƬO�_��J
If Len(Trim(txt_DeliveryDate.Text)) = 0 Then
   msg_text = "��ƿ��~�G����J�X�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_DeliveryDate.SetFocus
   Exit Sub
End If
If Len(Trim(txt_VehicleNo.Text)) = 0 Then
   msg_text = "��ƿ��~�G����J���P���X"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_VehicleNo.SetFocus
   Exit Sub
End If

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
 
'����ˮ�
'a.�X������G�榡 yyyymmdd
txt_DeliveryDate.Text = Trim(txt_DeliveryDate.Text)
If Fun_ChkDateFormat(txt_DeliveryDate.Text) = 1 Then
   msg_text = "�X������G" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_DeliveryDate.SelStart = 0: txt_DeliveryDate.SelLength = Len(txt_DeliveryDate.Text): txt_DeliveryDate.SetFocus
   Exit Sub
End If

'b.�ˮ� [���P���X] �O�_����
txt_VehicleNo.Text = Trim(txt_VehicleNo.Text)
str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_VehicleNo.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "��ƿ��~�G���P���X " & txt_VehicleNo.Text & " ������"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   txt_VehicleNo.SelStart = 0: txt_VehicleNo.SelLength = Len(txt_VehicleNo.Text)
   txt_VehicleNo.SetFocus
   Exit Sub
End If
tmp_Rs.Close

'���w�X�Y�Ȧs�G������J
txt_DockNo.Text = Trim(txt_DockNo.Text)
If Len(Trim(txt_DockNo.Text)) = 0 Then
   msg_text = "��ƿ��~�G[�X�Y�Ȧs] ������J"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_DockNo.SetFocus
   Exit Sub
End If

'�w�p������
txt_CarCheckInDate.Text = Trim(txt_CarCheckInDate.Text)
If Len(Trim(txt_CarCheckInDate.Text)) <> 8 Then
   msg_text = "�w�p�������G��Ʈ榡 yyyymmdd "
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_CarCheckInDate.SelStart = 0: txt_CarCheckInDate.SelLength = Len(txt_CarCheckInDate.Text)
   txt_CarCheckInDate.SetFocus
   Exit Sub
End If
If Fun_ChkDateFormat(txt_CarCheckInDate.Text) = 1 Then
   msg_text = "�w�p�������G��ƿ��~ yyyymmdd�A" & funRtn_msg
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_CarCheckInDate.SelStart = 0: txt_CarCheckInDate.SelLength = Len(txt_CarCheckInDate.Text)
   txt_CarCheckInDate.SetFocus
   Exit Sub
End If

'�w�p����ɶ�
txt_CarCheckInTime.Text = Trim(txt_CarCheckInTime.Text)
If Len(Trim(txt_CarCheckInTime.Text)) <> 4 Then
   msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_CarCheckInTime.SelStart = 0: txt_CarCheckInTime.SelLength = Len(txt_CarCheckInTime.Text)
   txt_CarCheckInTime.SetFocus
   Exit Sub
End If
Select Case Left(txt_CarCheckInTime.Text, 2)
       Case "00" To "24"
       Case Else
            msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            txt_CarCheckInTime.SelStart = 0: txt_CarCheckInTime.SelLength = Len(txt_CarCheckInTime.Text)
            txt_CarCheckInTime.SetFocus
            Exit Sub
End Select
Select Case Right(txt_CarCheckInTime.Text, 2)
       Case "00" To "59"
       Case Else
            msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            txt_CarCheckInTime.SelStart = 0: txt_CarCheckInTime.SelLength = Len(txt_CarCheckInTime.Text)
            txt_CarCheckInTime.SetFocus
            Exit Sub
End Select

'1.���ұ��R�������s�A�ƨ��̬O�_�����ɵn�J���ϥΪ�
If Left(strRouteNo, 1) = "R" Then
    str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From ORT01T Where Route_No = '" & strRouteNo & "'"
Else
    str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From TRP01T Where Route_No = '" & strRouteNo & "'"
End If

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "��Ʋ��`�G�䤣����R�������u�s��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
Else
   If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl = True Then
      tmp_Rs.Close
      msg_text = "�v�����ޡG���u�s�����R���u���\�ѭ�Ʃw�̰���"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      Exit Sub
   End If
End If
tmp_Rs.Close

Screen.MousePointer = vbHourglass
cmd_Save.Enabled = False

'2.���s�p�⨮�� >>
'  �Y���� [�X�����]�B[����] ���ܡA���ͷs����
'  �Y [�X�����]�B[����] ����A�u�έ쨮��
Dim intDriveTimes As Double
If txt_DeliveryDate.Text <> rs_Route.Fields("�X�����").Value And _
   txt_VehicleNo.Text <> rs_Route.Fields("���P���X").Value Then
   
   If Left(strRouteNo, 1) = "R" Then
   str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
             "From ORT05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_DeliveryDate.Text & "' and Vehicle_ID_No = '" & txt_VehicleNo.Text & "'"
   Else
      str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
             "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_DeliveryDate.Text & "' and Vehicle_ID_No = '" & txt_VehicleNo.Text & "'"
   End If
      
   tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
   intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
   tmp_Rs.Close
Else
   intDriveTimes = Val(txt_DriveTimes.Text)
End If

'3.��s TRP05T & TRP01T & TRP03T add TRP02T by Gemini @ 20080313
Tran_Level = 0
Tran_Level = cn.BeginTrans

If Left(strRouteNo, 1) = "R" Then
    str_SQL = "Update ORT01T Set Delivery_Date = '" & Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "' " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update ORT05T Set Delivery_Date = '" & Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "', " & _
              "   Vehicle_ID_No = '" & txt_VehicleNo.Text & "',Drive_Times = " & intDriveTimes & ",Dock_No = '" & txt_DockNo.Text & "',Expect_Date = '" & txt_CarCheckInDate.Text & "'," & _
              "   Expect_Time = '" & txt_CarCheckInTime.Text & "' " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update ORT02T Set Vehicle_ID_No = '" & txt_VehicleNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update ORT03T Set Vehicle_ID_No = '" & txt_VehicleNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�G�����s�ݧ�s�@�����s C_Vehicle_ID_No ����
    If Left(strRouteNo, 1) = "S" Then
        str_SQL = "Update ORT05T Set C_Vehicle_ID_No = '" & txt_VehicleNo.Text & "' Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Update ORT01T Set C_Vehicle_ID_No = '" & txt_VehicleNo.Text & "' Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    End If
Else

    str_SQL = "Update TRP01T Set Delivery_Date = '" & Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "' " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP05T Set Delivery_Date = '" & Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "', " & _
              "   Vehicle_ID_No = '" & txt_VehicleNo.Text & "',Drive_Times = " & intDriveTimes & ",Dock_No = '" & txt_DockNo.Text & "',Expect_Date = '" & txt_CarCheckInDate.Text & "'," & _
              "   Expect_Time = '" & txt_CarCheckInTime.Text & "' " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP02T Set Vehicle_ID_No = '" & txt_VehicleNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP03T Set Vehicle_ID_No = '" & txt_VehicleNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�G�����s�ݧ�s�@�����s C_Vehicle_ID_No ����
    If Left(strRouteNo, 1) = "S" Then
        str_SQL = "Update TRP05T Set C_Vehicle_ID_No = '" & txt_VehicleNo.Text & "' Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Update TRP01T Set C_Vehicle_ID_No = '" & txt_VehicleNo.Text & "' Where C_Route_No = '" & strRouteNo & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    End If
    
End If

'�Ѩ����D�ɧ�s�����������
If Left(strRouteNo, 1) = "R" Then
    str_SQL = "Update ORT05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From ORT05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_VehicleNo.Text & "' and Route_No = '" & strRouteNo & "' "
Else
    str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_VehicleNo.Text & "' and Route_No = '" & strRouteNo & "' "
End If
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'3.Update �e�� Recordset �����
blRouteEventEnable = False
rs_Route.Fields("�X�����").Value = txt_DeliveryDate.Text
rs_Route.Fields("���P���X").Value = txt_VehicleNo.Text
rs_Route.Fields("�X�Y�Ȧs").Value = txt_DockNo.Text
rs_Route.Fields("�w�p������").Value = txt_CarCheckInDate.Text
rs_Route.Fields("�w�p����ɶ�").Value = txt_CarCheckInTime.Text
rs_Route.Fields("����").Value = intDriveTimes
rs_Route.Fields("�B�餽�q").Value = txt_TRPCompany.Text
rs_Route.Fields("�r�p�H").Value = txt_Driver.Text
rs_Route.Fields("�q��").Value = txt_Phone.Text
rs_Route.Fields("����").Value = txt_VehicleType.Text
blRouteEventEnable = True

fam_RouteData.Enabled = False
fam_RouteData.BackColor = &H8000000F
cmd_Modify.Enabled = True
cmd_Cancel.Enabled = False
cmd_Save.Enabled = False
cmd_Delete.Enabled = False

If dg_Route.SelBookmarks.Count > 0 Then dg_Route.SelBookmarks.Remove 0
Call ClearData_RouteData
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
   CreateErrorLog Me.Name & "-�s��", Me.Caption, "cmd_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Save.Enabled = True
End Sub

Private Sub cmd_SelectCar_Click()
'�q�����
If Len(txt_DeliveryDate.Text) = 0 Then
   msg_text = "�Х���J�G�X�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_DeliveryDate.SetFocus
   Exit Sub
Else
   Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_SelectCar.Name)
End If
End Sub

Private Sub cmd_Tab0_ShowQuery_Click()
fra_ExtraQuery.Visible = Not fra_ExtraQuery.Visible
End Sub

Private Sub dg_Route_HeadClick(ByVal ColIndex As Integer)
'�H�ƹ��I�� [�ݱƨ��q��] dg_TRP02W �����D�ϡG�Ƨ������
Dim OrderFieldName As String
If TypeName(rs_Route) <> "Nothing" Then
   OrderFieldName = "[" & dg_Route.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_Route.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_Route.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

Private Sub dg_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If blRouteEventEnable Then
      With dg_Route
        '�ϥ���ܿ������ƦC
        If Not rs_Route.EOF Then
           dg_Route.SelBookmarks.Add rs_Route.Bookmark
           Call ClearData_RouteData
           fam_RouteData.Enabled = False
           cmd_Modify.Enabled = True
           cmd_Cancel.Enabled = False
           cmd_Save.Enabled = False
           cmd_Delete.Enabled = True
           
           lab_RouteNo.Caption = "���s�G" & rs_Route.Fields("���u�s��").Value
           strRouteNo = rs_Route.Fields("���u�s��").Value
           lab_RouteNo.AutoSize = True
           txt_DeliveryDate.Text = rs_Route.Fields("�X�����").Value
           txt_VehicleNo.Text = rs_Route.Fields("���P���X").Value
           txt_DockNo.Text = rs_Route.Fields("�X�Y�Ȧs").Value
           txt_CarCheckInDate.Text = rs_Route.Fields("�w�p������").Value
           txt_CarCheckInTime.Text = rs_Route.Fields("�w�p����ɶ�").Value
           txt_SecondRouteNo.Text = rs_Route.Fields("�G���ƨ����s").Value
           txt_Status.Text = rs_Route.Fields("EXE�^��").Value
           txt_PlanDate.Text = rs_Route.Fields("�ƨ����").Value & " " & rs_Route.Fields("�ƨ��ɶ�").Value
           txt_DriveTimes.Text = rs_Route.Fields("����").Value
           txt_TRPCompany.Text = rs_Route.Fields("�B�餽�q").Value
           txt_Driver.Text = rs_Route.Fields("�r�p�H").Value
           txt_Phone.Text = rs_Route.Fields("�q��").Value
           txt_VehicleType.Text = rs_Route.Fields("����").Value
           txt_CaseQty.Text = rs_Route.Fields("�c��").Value
           txt_PalletQty.Text = rs_Route.Fields("�O��").Value
           txt_Volumn.Text = rs_Route.Fields("���n").Value
           txt_Weight.Text = rs_Route.Fields("���q").Value
        End If
     End With
End If
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "���u�s����ƺ��@�@�~"
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
txt_DeliveryDate_Start = Format(Now(), "YYYYMMDD")

blRouteEventEnable = False

End Sub

Private Sub Form_Resize()
'�����j�p�ܰ�
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   'SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   'SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   fam_Header.Left = fam_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   Frame2.Left = Frame2.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_RouteData.Left = fam_RouteData.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   
   dg_Route.Width = dg_Route.Width - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Route.Height = dg_Route.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   'SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   'SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   fam_Header.Left = fam_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   Frame2.Left = Frame2.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_RouteData.Left = fam_RouteData.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)

   dg_Route.Width = dg_Route.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Route.Height = dg_Route.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
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
Set frm_OP_RouteData = Nothing
End Sub

Private Sub ClearData_RouteData()
'�M�� ���u�s����ư�
strRouteNo = ""
lab_RouteNo.Caption = ""
txt_DeliveryDate.Text = ""
txt_VehicleNo.Text = ""
txt_DockNo.Text = ""
txt_CarCheckInDate.Text = "": txt_CarCheckInTime.Text = ""
txt_SecondRouteNo.Text = "": txt_Status.Text = "": txt_PlanDate.Text = ""
txt_DriveTimes.Text = "": txt_TRPCompany.Text = "": txt_Driver.Text = "": txt_Phone.Text = "": txt_VehicleType.Text = ""
txt_CaseQty.Text = "": txt_PalletQty.Text = "": txt_Volumn.Text = "": txt_Weight.Text = ""
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
   Case "�ƨ����.�_"
        txt_PlanDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
   Case "�ƨ����.��"
        txt_PlanDate_End.Text = Format(mvDate.Value, "yyyymmdd")
   Case "�X�����.�_"
        txt_DeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
   Case "�X�����.��"
        txt_DeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
   Case "�X�����"
        txt_DeliveryDate.Text = Format(mvDate.Value, "yyyymmdd")
   Case "�w�p������"
        txt_CarCheckInDate.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_CarCheckInDate_Click()
'�w�p������
If Trim(txt_CarCheckInDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_CarCheckInDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_CarCheckInDate.Text, 4) & "/" & Mid(txt_CarCheckInDate.Text, 5, 2) & "/" & Right(txt_CarCheckInDate.Text, 2))
   End If
End If
mvDate.Left = fam_RouteData.Left + txt_CarCheckInDate.Left
mvDate.Top = fam_RouteData.Top + txt_CarCheckInDate.Top + txt_CarCheckInDate.Height
mvDate.Tag = "�w�p������"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_Click()
'�ƨ����-�_
If Trim(txt_DeliveryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2))
   End If
End If
mvDate.Left = fam_RouteData.Left + txt_DeliveryDate.Left
mvDate.Top = fam_RouteData.Top + txt_DeliveryDate.Top + txt_DeliveryDate.Height
mvDate.Tag = "�X�����"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_End_Click()
'�ƨ����-�_
If Trim(txt_DeliveryDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate_End.Text, 4) & "/" & Mid(txt_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_DeliveryDate_End.Text, 2))
   End If
End If
mvDate.Left = fam_Header.Left + txt_DeliveryDate_End.Left
mvDate.Top = fam_Header.Top + txt_DeliveryDate_End.Top + txt_DeliveryDate_End.Height
mvDate.Tag = "�X�����.��"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_Start_Click()
'�ƨ����-�_
If Trim(txt_DeliveryDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_DeliveryDate_Start.Text, 2))
   End If
End If
mvDate.Left = fam_Header.Left + txt_DeliveryDate_Start.Left
mvDate.Top = fam_Header.Top + txt_DeliveryDate_Start.Top + txt_DeliveryDate_Start.Height
mvDate.Tag = "�X�����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_PlanDate_End_Click()
'�ƨ����-�_
If Trim(txt_PlanDate_End.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_PlanDate_End.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_PlanDate_End.Text, 4) & "/" & Mid(txt_PlanDate_End.Text, 5, 2) & "/" & Right(txt_PlanDate_End.Text, 2))
   End If
End If
mvDate.Left = fam_Header.Left + txt_PlanDate_End.Left
mvDate.Top = fam_Header.Top + txt_PlanDate_End.Top + txt_PlanDate_End.Height
mvDate.Tag = "�ƨ����.��"
mvDate.Visible = True
End Sub

Private Sub txt_PlanDate_Start_Click()
'�ƨ����-�_
If Trim(txt_PlanDate_Start.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_PlanDate_Start.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_PlanDate_Start.Text, 4) & "/" & Mid(txt_PlanDate_Start.Text, 5, 2) & "/" & Right(txt_PlanDate_Start.Text, 2))
   End If
End If
mvDate.Left = fam_Header.Left + txt_PlanDate_Start.Left
mvDate.Top = fam_Header.Top + txt_PlanDate_Start.Top + txt_PlanDate_Start.Height
mvDate.Tag = "�ƨ����.�_"
mvDate.Visible = True
End Sub

Private Sub txt_VehicleNo_LostFocus()
    If Len(txt_VehicleNo) = 0 Then Exit Sub
    str_SQL = "Select Vehicle_ID_No from trp09m where Vehicle_ID_No='" & Trim(txt_VehicleNo) & "' "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        'tmp_rs.Close
        msg_text = "�L���������"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_VehicleNo.SetFocus
    End If
    tmp_Rs.Close

End Sub
