VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_MBO_SDNReturnList 
   Caption         =   "MBO�^���ˮ֪�"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14235
   BeginProperty Font 
      Name            =   "�ө���"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_Report_MBO_SDNReturnList.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   14235
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
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
      StartOfWeek     =   135593985
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
      TabIndex        =   6
      Top             =   2280
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   9
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
      Height          =   2295
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14175
      Begin VB.ListBox List5 
         Columns         =   3
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   7680
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   35
         ToolTipText     =   "A:���_�`���q�AB:�x�_�����q�AC:���������q�AD:�x�������q"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox optPodReturn 
         Caption         =   "POD�w�^��"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Value           =   1  '�֨�
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox optSdnReturn 
         Caption         =   "ñ��w�^��"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox optSdnback 
         Caption         =   "ñ��w�^"
         Height          =   255
         Left            =   3480
         TabIndex        =   32
         Top             =   1560
         Value           =   1  '�֨�
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Columns         =   1
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         ItemData        =   "frm_Report_MBO_SDNReturnList.frx":0342
         Left            =   9960
         List            =   "frm_Report_MBO_SDNReturnList.frx":0344
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   31
         ToolTipText     =   "�t�e�ܧO"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Columns         =   3
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         ItemData        =   "frm_Report_MBO_SDNReturnList.frx":0346
         Left            =   4800
         List            =   "frm_Report_MBO_SDNReturnList.frx":0348
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   30
         ToolTipText     =   "�q�����O"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ListBox List2 
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
         ItemData        =   "frm_Report_MBO_SDNReturnList.frx":034A
         Left            =   7440
         List            =   "frm_Report_MBO_SDNReturnList.frx":034C
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   29
         ToolTipText     =   "�f�B���q"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_Report_MBO_SDNReturnList.frx":034E
         Left            =   1200
         List            =   "frm_Report_MBO_SDNReturnList.frx":035B
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   27
         Top             =   1920
         Width           =   2325
      End
      Begin VB.CheckBox optNotYet 
         Caption         =   "���T�{ñ��"
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox optAbnormal 
         Caption         =   "���`ñ��"
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox optNormal 
         Caption         =   "���`ñ��"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveToText 
         BackColor       =   &H00C0E0FF&
         Caption         =   "�ˮ֪�"
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
         Left            =   12120
         Picture         =   "frm_Report_MBO_SDNReturnList.frx":0393
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ListBox List1 
         Columns         =   3
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   6240
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   21
         ToolTipText     =   "�ϽX"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPreView 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�w���C�L"
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
         Left            =   4560
         Picture         =   "frm_Report_MBO_SDNReturnList.frx":069D
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FF8080&
         Caption         =   "����C�L"
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
         Picture         =   "frm_Report_MBO_SDNReturnList.frx":09A7
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   16
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   15
         Top             =   960
         Width           =   1485
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
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
         Left            =   1200
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateE 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateS 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   0
         Top             =   600
         Width           =   1485
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
         Left            =   9960
         Picture         =   "frm_Report_MBO_SDNReturnList.frx":0CB1
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
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
         Left            =   12120
         Picture         =   "frm_Report_MBO_SDNReturnList.frx":1FAB
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
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
         Left            =   11040
         Picture         =   "frm_Report_MBO_SDNReturnList.frx":2BBBD
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   240
         Width           =   1065
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
         Left            =   8880
         Picture         =   "frm_Report_MBO_SDNReturnList.frx":2BECF
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ƨ�"
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
         Left            =   360
         TabIndex        =   28
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ݧ@���X���T�{"
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
         Left            =   2760
         TabIndex        =   22
         Top             =   240
         Width           =   1680
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
         TabIndex        =   18
         Top             =   1020
         Width           =   360
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
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   960
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
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���@���"
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
         TabIndex        =   12
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
         Index           =   1
         Left            =   2655
         TabIndex        =   11
         Top             =   660
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   6030
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   476
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
            Object.Width           =   18468
            MinWidth        =   2646
            Object.ToolTipText     =   "��Ƶ���"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.ToolTipText     =   "�ϥΪ�"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Report_MBO_SDNReturnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmdPreView_Click()

Dim i As Integer, j As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "�L��ƥi�ѦC�L�I", vbOKOnly + vbInformation, "����C�L": Exit Sub
Screen.MousePointer = 11

'��Ƽg�J Access ��Ʈw
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From MBO�^���ˮ֪�", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "MBO�^���ˮ֪�", cnAccess, adOpenStatic, adLockOptimistic

With rsMain
    .MoveFirst
    Do While Not .EOF
       rs_Access.AddNew
       For i = 0 To .Fields.Count - 1
           rs_Access.Fields(i).Value = .Fields(i).Value
       Next i
       rs_Access.Update
       .MoveNext
    Loop
    .MoveFirst
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
End With

strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
With MSAccessAP
    .Visible = False
    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
    .DoCmd.Maximize
    
    '�g�JUSER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    .Reports(Me.Caption).[User_id].Caption = User_id
    .DoCmd.Close

    .DoCmd.OpenReport "MBO�^���ˮ֪�", acViewPreview
    .Visible = True

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdPrint_Click()
Dim i As Integer, j As Integer, Str_Orders As String, bl_Sdn As Boolean, bl_SdnR As Boolean
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "�L��ƥi�ѦC�L�I", vbOKOnly + vbInformation, "����C�L": Exit Sub
Screen.MousePointer = 11

bl_Sdn = False
bl_SdnR = False
'
'rsMain.MoveFirst
'Do While Not rsMain.EOF
'        str_SQL = "select ���TMS = s2.receipt_no,�^�Ǫ��A = s2.returnstatus from sdn02t s2 where s2.c_receipt_no = '" & rsMain.Fields("TMS�渹") & "'"
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        tmp_Rs.MoveFirst
'        Do While Not tmp_Rs.EOF
'            If tmp_Rs.Fields("�^�Ǫ��A") = "0" Then
'                MsgBox "TMS�渹:" & rsMain.Fields("TMS�渹") & "�������渹:" & tmp_Rs.Fields("���TMS") & "�|����POD�^�ǡA�нT�{��A�A�i��^���ˮ֪�@�~�A�^�ǲפ�", vbCritical + vbOKOnly, "�^���ˬd"
'                tmp_Rs.Close: Screen.MousePointer = 0: Exit Sub
'            End If
'            tmp_Rs.MoveNext
'        Loop
'        rsMain.MoveNext
'Loop

'��Ƽg�J Access ��Ʈw
Call AccessDB_Connect
cnAccess.BeginTrans
Tran_Level = cn.BeginTrans

cnAccess.Execute "Delete From MBO�^���ˮ֪�", RowsAffect, adExecuteNoRecords
cnAccess.Execute "Delete From MBO�h�f�^���ˮ֪�", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset

'rs_Access.Open "MBO�^���ˮ֪�", cnAccess, adOpenStatic, adLockOptimistic
'rs_Access.Open "MBO�h�f�^���ˮ֪�", cnAccess, adOpenStatic, adLockOptimistic

'�ˬd�^�Ǫ�ñ��A�O�_����POD�^��


rsMain.MoveFirst
With rsMain
    .MoveFirst
    Do While Not .EOF
      If .Fields("��O") <> "R" Then
        bl_Sdn = True
        If Str_Orders <> .Fields("TMS�渹") Then
            '��s�^�Ǫ��A
            Str_Orders = .Fields("TMS�渹")
           ' cn.Execute "update sdn02t set ReturnStatus = '2' where c_receipt_no = '" & RTrim(rsMain.Fields("TMS�渹")) & "'", RowsAffect, adExecuteNoRecords
            '�Ĥ@�����g�L�h
            If Val(.Fields("�q���q")) <> Val(.Fields("�ꦬ�q")) Then
                str_SQL = "Insert into MBO�^���ˮ֪� (�s��,���@���,�����q,���q�W��,�^���,TMS�渹,�f�D�渹,�q�櫬�A,�Ȥ�N��,�Ȥ�W��,�~��,�q���q,�ꦬ�q,�o�����B,�o�����X,���A,����,�N�����B,�w�p�X�f��,�X����,��,�^�Ǫ��A,User_id) " & _
                "values ('" & .Fields("�s��") & "','" & .Fields("���@���") & "','" & .Fields("�����q") & "','" & .Fields("���q�W��") & "','" & .Fields("�^���") & "','" & .Fields("TMS�渹") & "','" & .Fields("�f�D�渹") & "','" & .Fields("�q�����O") & "','" & .Fields("�Ȥ�N��") & "','" & .Fields("�Ȥ�W��") & _
                "','" & .Fields("�~��") & "','" & .Fields("�q���q") & "','" & .Fields("�ꦬ�q") & "','" & .Fields("�o�����B") & "','" & .Fields("�o�����X") & "','" & .Fields("���A") & "','" & .Fields("����") & "','" & .Fields("�N�����B") & "','" & .Fields("�w�p�X�f��") & _
                "','" & .Fields("�X����") & "','" & .Fields("��") & "','" & .Fields("�^�Ǫ��A") & "','" & User_id & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Else
                str_SQL = "Insert into MBO�^���ˮ֪� (�s��,���@���,�����q,���q�W��,�^���,TMS�渹,�f�D�渹,�q�櫬�A,�Ȥ�N��,�Ȥ�W��,�~��,�q���q,�ꦬ�q,�o�����B,�o�����X,���A,����,�N�����B,�w�p�X�f��,�X����,��,�^�Ǫ��A,User_id) " & _
                "values ('" & .Fields("�s��") & "','" & .Fields("���@���") & "','" & .Fields("�����q") & "','" & .Fields("���q�W��") & "','" & .Fields("�^���") & "','" & .Fields("TMS�渹") & "','" & .Fields("�f�D�渹") & "','" & .Fields("�q�����O") & "','" & .Fields("�Ȥ�N��") & "','" & .Fields("�Ȥ�W��") & _
                "','','','','" & .Fields("�o�����B") & "','" & .Fields("�o�����X") & "','" & .Fields("���A") & "','" & .Fields("����") & "','" & .Fields("�N�����B") & "','" & .Fields("�w�p�X�f��") & _
                "','" & .Fields("�X����") & "','" & .Fields("��") & "','" & .Fields("�^�Ǫ��A") & "','" & User_id & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
        End If
            '���P�_�q���q�O�_����ꦬ�q�A���P�h�n�g�ƶq,�~���L�h
            If Val(.Fields("�q���q")) <> Val(.Fields("�ꦬ�q")) Then
                str_SQL = "Insert into MBO�^���ˮ֪� (�s��,���@���,�����q,���q�W��,�^���,TMS�渹,�f�D�渹,�q�櫬�A,�Ȥ�N��,�Ȥ�W��,�~��,�q���q,�ꦬ�q,�o�����B,�o�����X,���A,����,�N�����B,�w�p�X�f��,�X����,��,�^�Ǫ��A,User_id) " & _
                "values ('" & .Fields("�s��") & "','" & .Fields("���@���") & "','" & .Fields("�����q") & "','" & .Fields("���q�W��") & "','" & .Fields("�^���") & "','" & .Fields("TMS�渹") & "','" & .Fields("�f�D�渹") & "','" & .Fields("�q�����O") & "','" & .Fields("�Ȥ�N��") & "','" & .Fields("�Ȥ�W��") & _
                "','" & .Fields("�~��") & "','" & .Fields("�q���q") & "','" & .Fields("�ꦬ�q") & "','" & .Fields("�o�����B") & "','" & .Fields("�o�����X") & "','" & .Fields("���A") & "','" & .Fields("����") & "','" & .Fields("�N�����B") & "','" & .Fields("�w�p�X�f��") & _
                "','" & .Fields("�X����") & "','" & .Fields("��") & "','" & .Fields("�^�Ǫ��A") & "','" & User_id & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
     Else
'''''''''''''''''''''''''''''''''''
        bl_SdnR = True
        If Str_Orders <> .Fields("TMS�渹") Then
            '��s�^�Ǫ��A
            Str_Orders = .Fields("TMS�渹")
            'cn.Execute "update sdn02t set ReturnStatus = '2' where c_receipt_no = '" & RTrim(rsMain.Fields("TMS�渹")) & "'", RowsAffect, adExecuteNoRecords
            '�Ĥ@�����g�L�h
            If Val(.Fields("�q���q")) <> Val(.Fields("�ꦬ�q")) Then
                str_SQL = "Insert into MBO�h�f�^���ˮ֪� (�s��,���@���,�����q,���q�W��,�^���,TMS�渹,�f�D�渹,�q�櫬�A,�Ȥ�N��,�Ȥ�W��,�~��,�q���q,�ꦬ�q,�o�����B,�o�����X,���A,����,�N�����B,�w�p�X�f��,�X����,��,�^�Ǫ��A,User_id) " & _
                "values ('" & .Fields("�s��") & "','" & .Fields("���@���") & "','" & .Fields("�����q") & "','" & .Fields("���q�W��") & "','" & .Fields("�^���") & "','" & .Fields("TMS�渹") & "','" & .Fields("�f�D�渹") & "','" & .Fields("�q�����O") & "','" & .Fields("�Ȥ�N��") & "','" & .Fields("�Ȥ�W��") & _
                "','" & .Fields("�~��") & "','" & .Fields("�q���q") & "','" & .Fields("�ꦬ�q") & "','" & .Fields("�o�����B") & "','" & .Fields("�o�����X") & "','" & .Fields("���A") & "','" & .Fields("����") & "','" & .Fields("�N�����B") & "','" & .Fields("�w�p�X�f��") & _
                "','" & .Fields("�X����") & "','" & .Fields("��") & "','" & .Fields("�^�Ǫ��A") & "','" & User_id & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Else
                str_SQL = "Insert into MBO�h�f�^���ˮ֪� (�s��,���@���,�����q,���q�W��,�^���,TMS�渹,�f�D�渹,�q�櫬�A,�Ȥ�N��,�Ȥ�W��,�~��,�q���q,�ꦬ�q,�o�����B,�o�����X,���A,����,�N�����B,�w�p�X�f��,�X����,��,�^�Ǫ��A,User_id) " & _
                "values ('" & .Fields("�s��") & "','" & .Fields("���@���") & "','" & .Fields("�����q") & "','" & .Fields("���q�W��") & "','" & .Fields("�^���") & "','" & .Fields("TMS�渹") & "','" & .Fields("�f�D�渹") & "','" & .Fields("�q�����O") & "','" & .Fields("�Ȥ�N��") & "','" & .Fields("�Ȥ�W��") & _
                "','','','','" & .Fields("�o�����B") & "','" & .Fields("�o�����X") & "','" & .Fields("���A") & "','" & .Fields("����") & "','" & .Fields("�N�����B") & "','" & .Fields("�w�p�X�f��") & _
                "','" & .Fields("�X����") & "','" & .Fields("��") & "','" & .Fields("�^�Ǫ��A") & "','" & User_id & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
        End If
            '���P�_�q���q�O�_����ꦬ�q�A���P�h�n�g�ƶq,�~���L�h
            If Val(.Fields("�q���q")) <> Val(.Fields("�ꦬ�q")) Then
                str_SQL = "Insert into MBO�h�f�^���ˮ֪� (�s��,���@���,�����q,���q�W��,�^���,TMS�渹,�f�D�渹,�q�櫬�A,�Ȥ�N��,�Ȥ�W��,�~��,�q���q,�ꦬ�q,�o�����B,�o�����X,���A,����,�N�����B,�w�p�X�f��,�X����,��,�^�Ǫ��A,User_id) " & _
                "values ('" & .Fields("�s��") & "','" & .Fields("���@���") & "','" & .Fields("�����q") & "','" & .Fields("���q�W��") & "','" & .Fields("�^���") & "','" & .Fields("TMS�渹") & "','" & .Fields("�f�D�渹") & "','" & .Fields("�q�����O") & "','" & .Fields("�Ȥ�N��") & "','" & .Fields("�Ȥ�W��") & _
                "','" & .Fields("�~��") & "','" & .Fields("�q���q") & "','" & .Fields("�ꦬ�q") & "','" & .Fields("�o�����B") & "','" & .Fields("�o�����X") & "','" & .Fields("���A") & "','" & .Fields("����") & "','" & .Fields("�N�����B") & "','" & .Fields("�w�p�X�f��") & _
                "','" & .Fields("�X����") & "','" & .Fields("��") & "','" & .Fields("�^�Ǫ��A") & "','" & User_id & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
    End If
       .MoveNext
    Loop
    .MoveFirst
    cnAccess.CommitTrans
    cn.CommitTrans
    Call DB_Disconnect(cnAccess)
End With

strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
With MSAccessAP
    .Visible = False
    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
    
    '�g�JUSER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    '.Reports(Me.Caption).[User_id].Caption = User_id
    '.DoCmd.Close
    
    '�����C�L�ܦL���
    .Visible = True
    If bl_Sdn = True Then
        .DoCmd.OpenReport "MBO�^���ˮ֪�", acViewPreview
    End If
    If bl_SdnR = True Then
        .DoCmd.OpenReport "MBO�h�f�^���ˮ֪�", acViewPreview
    End If
    
    '.CloseCurrentDatabase
    '.Quit: Set MSAccessAP = Nothing

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel "LMBO01�^���ˮ֪�", rsMain

'..�b���s��EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
'        .Columns("L").Select
'        .Selection.ClearContents
        .Range("B3").Value = Combo1
        If Len(RTrim(txtOrderDateS)) > 0 Then .Range("C4").Value = .Range("C4").Value & "���@��:" & RTrim(txtOrderDateS) & "  "
        If Len(RTrim(txtDeliveryDateS)) > 0 Then .Range("C4").Value = .Range("C4").Value & "��f��:" & RTrim(txtDeliveryDateS) & "  "
        .Range("K4").Value = Format(Now(), "YYYY/MM/DD hh:mm:ss")
        .Range("A1").Select
        '�ƥ��ɮ�
        '    If Dir("C:\LTKK01\DelievryTrack", vbDirectory) = "" Then MkDirs "C:\LTKK01\DelievryTrack"
        '    .ActiveWorkbook.SaveAs "C:\LTKK01\DelievryTrack\DelievryTrack" & Format(Now, "yyyymmddhhMMss") & ".xls"
                
    End With
End If
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String, chc_DeliveryDate As String, i As Integer, strPriority As String, strArea As String, strBranchid As String
            
'�ϽX
strArea = ""
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then strArea = strArea & "'" & Left(List1.List(i), 2) & "',"
Next

If Len(RTrim(strArea)) > 0 Then strArea = " and t1m.area_code in ( " & strArea & "'') "

'�ϽX
strBranchid = ""
For i = 0 To List5.ListCount - 1
    If List5.Selected(i) Then strBranchid = strBranchid & "'" & Left(List5.List(i), 1) & "',"
Next

If Len(RTrim(strBranchid)) > 0 Then strBranchid = " and co.branchid in (" & strBranchid & "'') "

''�f�B���q
'strSelected = ""
'For i = 0 To List2.ListCount - 1
'    If List2.Selected(i) Then strSelected = strSelected & "'" & mySplit(List2.List(i), "_", 0) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t8m.company_code in ( " & strSelected & "'') "

'��O
strPriority = ""
For i = 0 To List3.ListCount - 1
    If List3.Selected(i) Then strPriority = strPriority & "'" & mySplit(List3.List(i), "_", 0) & "',"
Next

If Len(RTrim(strPriority)) > 0 Then strPriority = " and isnull(s2.priority,'') in ( " & strPriority & "'') "
'
''�t�e�ܧO
'strSelected = ""
'For i = 0 To List4.ListCount - 1
'    If List4.Selected(i) Then strSelected = strSelected & "'" & List4.List(i) & "',"
'Next
'
'If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and isnull(o.facility,'') in (" & Left(strSelected, Len(strSelected) - 1) & ") "

'���@���
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateE.Text & "' "
End If

'��f���
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(Char(8),s2.arrive_date,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_DeliveryDate = "and convert(Char(8),s2.arrive_date,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(Char(8),s2.arrive_date,112) = '" & txtDeliveryDateE.Text & "' "
End If

'ñ�����O
If optNormal = 0 And optAbnormal = 0 And optNotYet = 0 Then GoTo NextStep
Dim strStatus As String

strStatus = "and s2.confirm_notes in ("

If optNormal = 1 Then strStatus = strStatus & "'���`�q��',"
If optAbnormal = 1 Then strStatus = strStatus & "'���`�q��','���X�q��',"
If optNotYet = 1 Then strStatus = strStatus & "'',"

strStatus = Left(strStatus, Len(strStatus) - 1) & ") "

NextStep:


If optNotYet = 1 Then
'�����@ñ��A�����
    If strStatus = "and s2.confirm_notes in ('') " Then
        '�u���Ŀ良���@,�hunion �t���~��
        str_SQL = "select �� =case when (select top 1 cod.InvoicePCode from custorderdetail cod where cod.orderkey = co.orderkey and cod.InvoicePCode = 'N') ='N'  then '��' else '' end " & _
                ",�q�����O = co.ordertype ,�q�渹�X = rtrim(s2.extern),���q�W�� = case when  co.BranchId = 'A' then '���_-�`���q' when  co.BranchId = 'B' then '���_-�x�_' when  co.BranchId = 'C' then '���_-����' when  co.BranchId = 'D' then '���_-�x��' end ,�Ȥ�s�� = t1m.consigneekey ,�Ȥ�W�� = t1m.short_name " & _
                ",�w�p�X�f�� =  isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') " & _
                ",'�o�����B/�~��' = cast(cast(co.Amount as float) as char),'�o�����X/�q���q' = rtrim(co.Invoice),'����/�ꦬ�q' = rtrim(isnull(s2.customerorderkey1,'')),���`���p = case when s2.confirm_notes = '���`�q��' then 'N' when len(rtrim(isnull(s2.confirm_notes,''))) = 0 then '�����@' else 'Y' end " & _
                ",�禬�渹 = ' ',�N�����B = rtrim(cast(case when co.Payment = '.' then o.receivecash else '' end as char)) " & _
                ",�����q = co.BranchId ,TMS�渹 = co.orderkey " & _
                ",�^��� = isnull(rtrim(convert(char(4),getdate(),112) - 1911) + '/' + substring(rtrim(convert(char(8),getdate(),112)),5,2) + '/' + rtrim(right(convert(char(8),getdate(),112),2)),'') ,ñ��T�{�ɶ� = isnull(convert(char(19),s2.confirm_date,121),''),�G������ = rtrim(s1.c_vehicle_id_no),�@������ = rtrim(s2.vehicle_id_no) ,���� = '1',�o���q��渹���O  = co.Invoice+co.externorderkey+co.ordertype " & _
                "from sdn02t s2 join sdn01t s1 on s1.c_route_no = s2.c_route_no " & _
                "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
                "join orders o on o.orderkey = s2.c_receipt_no " & _
                "join trp01m t1m on t1m.storerkey = o.storerkey  and  case when rtrim(isnull(s2.priority,'')) = 'A2B' then o.b_company else o.consigneekey end = t1m.consigneekey " & _
                "where 1=1 " & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' " & strStatus & strBranchid & strArea & strPriority & _
                "group by  convert(char(8),s2.confirm_date,112),co.BranchId ,case when  co.BranchId = 'A' then '���_-�`���q' when  co.BranchId = 'B' then '���_-�x�_' when  co.BranchId = 'C' then '���_-����' when  co.BranchId = 'D' then '���_-�x��' end , co.orderkey , " & _
                "s2.priority , rtrim(s2.extern) , co.ordertype , t1m.consigneekey , t1m.short_name ,co.Amount , co.Invoice , case when s2.confirm_notes = '���`�q��' then '�X�f' when len(rtrim(isnull(s2.confirm_notes,''))) = 0 then '�����@' else '�����X�f' end , rtrim(isnull(s2.customerorderkey1,'')) , rtrim(cast(case when co.Payment = '.' then o.receivecash else '' end as char)),isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') , isnull(rtrim(convert(char(4),s1.delivery_date,112) - 1911) + '/' + substring(rtrim(convert(char(8),s1.delivery_date,112)),5,2) + '/' + rtrim(right(convert(char(8),s1.delivery_date,112),2)),'') , s2.confirm_notes,s2.confirm_date,co.Invoice+co.externorderkey+co.ordertype,rtrim(s1.c_vehicle_id_no), rtrim(s2.vehicle_id_no)  order by  �w�p�X�f��,�o���q��渹���O,���� asc"
    Else
        '���]�t�����@��ñ��A�t���~����where����]�ws2.confirm_notes <> ''
        str_SQL = "select �� =case when (select top 1 cod.InvoicePCode from custorderdetail cod where cod.orderkey = co.orderkey and cod.InvoicePCode = 'N') ='N'  then '��' else '' end " & _
                ",�q�����O = co.ordertype ,�q�渹�X = rtrim(s2.extern),���q�W�� = case when  co.BranchId = 'A' then '���_-�`���q' when  co.BranchId = 'B' then '���_-�x�_' when  co.BranchId = 'C' then '���_-����' when  co.BranchId = 'D' then '���_-�x��' end ,�Ȥ�s�� = t1m.consigneekey ,�Ȥ�W�� = t1m.short_name " & _
                ",�w�p�X�f�� =  isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') " & _
                ",'�o�����B/�~��' = cast(cast(co.Amount as float) as char),'�o�����X/�q���q' = rtrim(co.Invoice),'����/�ꦬ�q' = rtrim(isnull(s2.customerorderkey1,'')),���`���p = case when s2.confirm_notes = '���`�q��' then 'N' when len(rtrim(isnull(s2.confirm_notes,''))) = 0 then '�����@' else 'Y' end " & _
                ",�禬�渹 = ' ',�N�����B = rtrim(cast(case when co.Payment = '.' then o.receivecash else '' end as char)) " & _
                ",�����q = co.BranchId ,TMS�渹 = co.orderkey " & _
                ",�^��� = isnull(rtrim(convert(char(4),getdate(),112) - 1911) + '/' + substring(rtrim(convert(char(8),getdate(),112)),5,2) + '/' + rtrim(right(convert(char(8),getdate(),112),2)),'') ,ñ��T�{�ɶ� = isnull(convert(char(19),s2.confirm_date,121),''),�G������ = rtrim(s1.c_vehicle_id_no),�@������ = rtrim(s2.vehicle_id_no) ,���� = '1',�o���q��渹���O  = co.Invoice+co.externorderkey+co.ordertype " & _
                "from sdn02t s2 join sdn01t s1 on s1.c_route_no = s2.c_route_no " & _
                "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
                "join orders o on o.orderkey = s2.c_receipt_no " & _
                "join trp01m t1m on t1m.storerkey = o.storerkey  and  case when rtrim(isnull(s2.priority,'')) = 'A2B' then o.b_company else o.consigneekey end = t1m.consigneekey " & _
                "where 1=1 " & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' " & strStatus & strBranchid & strArea & strPriority & _
                "group by  convert(char(8),s2.confirm_date,112),co.BranchId ,case when  co.BranchId = 'A' then '���_-�`���q' when  co.BranchId = 'B' then '���_-�x�_' when  co.BranchId = 'C' then '���_-����' when  co.BranchId = 'D' then '���_-�x��' end , co.orderkey , " & _
                "s2.priority , rtrim(s2.extern) , co.ordertype , t1m.consigneekey , t1m.short_name ,co.Amount , co.Invoice , case when s2.confirm_notes = '���`�q��' then '�X�f' when len(rtrim(isnull(s2.confirm_notes,''))) = 0 then '�����@' else '�����X�f' end , rtrim(isnull(s2.customerorderkey1,'')) , rtrim(cast(case when co.Payment = '.' then o.receivecash else '' end as char)),isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') , isnull(rtrim(convert(char(4),s1.delivery_date,112) - 1911) + '/' + substring(rtrim(convert(char(8),s1.delivery_date,112)),5,2) + '/' + rtrim(right(convert(char(8),s1.delivery_date,112),2)),'') , s2.confirm_notes,s2.confirm_date,co.Invoice+co.externorderkey+co.ordertype,rtrim(s1.c_vehicle_id_no), rtrim(s2.vehicle_id_no)  " & _
                "Union All " & _
                "select �� =' ',�q�����O =' ',�q�渹�X =' ' ,���q�W�� =' ',�Ȥ�s�� =' ',�Ȥ�W�� =' ',�w�p�X�f�� =isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') " & _
                ",�~�� = rtrim(s3.product_no) ,�q���q = rtrim(cast(sum(s3.order_qty) as char)),�ꦬ�q = rtrim(cast(sum(s3.sign_qty) as char)),���`���p =' ' " & _
                ",�禬�渹 =' ',�N�����B =' ',�����q =' ',TMS�渹 = co.orderkey ,�^��� =' ',ñ��T�{�ɶ� = ' ',�G������ = ' ',�@������ = ' ',���� = '2',�o���q��渹���O  = co.Invoice+co.externorderkey+co.ordertype " & _
                "from sdn02t s2 join sdn01t s1 on s1.c_route_no = s2.c_route_no " & _
                "join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
                "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
                "join orders o on o.orderkey = s2.c_receipt_no " & _
                "join trp01m t1m on t1m.storerkey = o.storerkey  and  case when rtrim(isnull(s2.priority,'')) = 'A2B' then o.b_company else o.consigneekey end = t1m.consigneekey " & _
                "where 1=1 " & Mid(strStatus, 1, Len(strStatus) - 5) & ")" & " and sign_qty <> order_qty " & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' " & strBranchid & strStatus & strArea & strPriority & _
                "group by isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),''),co.orderkey , s3.product_no,co.Invoice+co.externorderkey+co.ordertype order by  �w�p�X�f��,�o���q��渹���O,���� asc"
        End If
Else
'���X�Ҧ�ñ��A�A���X���`ñ��A�worderkey�Ƨ�
str_SQL = "select �� =case when (select top 1 cod.InvoicePCode from custorderdetail cod where cod.orderkey = co.orderkey and cod.InvoicePCode = 'N') ='N'  then '��' else '' end " & _
        ",�q�����O = co.ordertype ,�q�渹�X = rtrim(s2.extern) ,���q�W�� = case when  co.BranchId = 'A' then '���_-�`���q' when  co.BranchId = 'B' then '���_-�x�_' when  co.BranchId = 'C' then '���_-����' when  co.BranchId = 'D' then '���_-�x��' end ,�Ȥ�s�� = t1m.consigneekey ,�Ȥ�W�� = t1m.short_name " & _
        ",�w�p�X�f�� =  isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') " & _
        ",'�o�����B/�~��' = cast(cast(co.Amount as float) as char),'�o�����X/�q���q' = rtrim(co.Invoice),'����/�ꦬ�q' = rtrim(isnull(s2.customerorderkey1,'')),���`���p = case when s2.confirm_notes = '���`�q��' then 'N' when len(rtrim(isnull(s2.confirm_notes,''))) = 0 then '�����@' else 'Y' end " & _
        ",�禬�渹 = ' ',�N�����B = rtrim(cast(case when co.Payment = '.' then o.receivecash else '' end as char)) " & _
        ",�����q = co.BranchId ,TMS�渹 = co.orderkey " & _
        ",�^��� = isnull(rtrim(convert(char(4),getdate(),112) - 1911) + '/' + substring(rtrim(convert(char(8),getdate(),112)),5,2) + '/' + rtrim(right(convert(char(8),getdate(),112),2)),'') ,ñ��T�{�ɶ� = isnull(convert(char(19),s2.confirm_date,121),''),�G������ = rtrim(s1.c_vehicle_id_no),�@������ = rtrim(s2.vehicle_id_no) ,���� = '1',�o���q��渹���O  = co.Invoice+co.externorderkey+co.ordertype " & _
        "from sdn02t s2 join sdn01t s1 on s1.c_route_no = s2.c_route_no " & _
        "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
        "join orders o on o.orderkey = s2.c_receipt_no " & _
        "join trp01m t1m on t1m.storerkey = o.storerkey  and  case when rtrim(isnull(s2.priority,'')) = 'A2B' then o.b_company else o.consigneekey end = t1m.consigneekey " & _
        "where 1=1 " & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' " & strStatus & strBranchid & strArea & strPriority & _
        "group by  convert(char(8),s2.confirm_date,112),co.BranchId ,case when  co.BranchId = 'A' then '���_-�`���q' when  co.BranchId = 'B' then '���_-�x�_' when  co.BranchId = 'C' then '���_-����' when  co.BranchId = 'D' then '���_-�x��' end , co.orderkey , " & _
        "s2.priority ,rtrim(s2.extern) , co.ordertype , t1m.consigneekey , t1m.short_name ,co.Amount , co.Invoice , case when s2.confirm_notes = '���`�q��' then '�X�f' when len(rtrim(isnull(s2.confirm_notes,''))) = 0 then '�����@' else '�����X�f' end , rtrim(isnull(s2.customerorderkey1,'')) , rtrim(cast(case when co.Payment = '.' then o.receivecash else '' end as char)),isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') , isnull(rtrim(convert(char(4),s1.delivery_date,112) - 1911) + '/' + substring(rtrim(convert(char(8),s1.delivery_date,112)),5,2) + '/' + rtrim(right(convert(char(8),s1.delivery_date,112),2)),'') , s2.confirm_notes,s2.confirm_date,co.Invoice+co.externorderkey+co.ordertype,rtrim(s1.c_vehicle_id_no), rtrim(s2.vehicle_id_no)  " & _
        "Union All " & _
        "select �� =' ',�q�����O =' ',�q�渹�X =' ' ,���q�W�� =' ',�Ȥ�s�� =' ',�Ȥ�W�� =' ',�w�p�X�f�� =isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),'') " & _
        ",�~�� = rtrim(s3.product_no) ,�q���q = rtrim(cast(sum(s3.order_qty) as char)),�ꦬ�q = rtrim(cast(sum(s3.sign_qty) as char)),���`���p =' ' " & _
        ",�禬�渹 =' ',�N�����B =' ',�����q =' ',TMS�渹 = co.orderkey ,�^��� =' ',ñ��T�{�ɶ� = ' ',�G������ = ' ',�@������ = ' ',���� = '2',�o���q��渹���O  = co.Invoice+co.externorderkey+co.ordertype " & _
        "from sdn02t s2 join sdn01t s1 on s1.c_route_no = s2.c_route_no " & _
        "join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
        "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
        "join orders o on o.orderkey = s2.c_receipt_no " & _
        "join trp01m t1m on t1m.storerkey = o.storerkey  and  case when rtrim(isnull(s2.priority,'')) = 'A2B' then o.b_company else o.consigneekey end = t1m.consigneekey " & _
        "where 1=1 and sign_qty <> order_qty " & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' " & strStatus & strBranchid & strArea & strPriority & _
        "group by isnull(rtrim(convert(char(4),o.DeliveryDate,112) - 1911) + '/' + substring(rtrim(convert(char(8),o.DeliveryDate,112)),5,2) + '/' + rtrim(right(convert(char(8),o.DeliveryDate,112),2)),''),co.orderkey , s3.product_no,co.Invoice+co.externorderkey+co.ordertype order by  �w�p�X�f��,�o���q��渹���O,���� asc"
End If

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Call Replication_Recordset(tmp_Rs, rsMain)

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

End With

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Screen.MousePointer = 0
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdSaveToText_Click()
'��ƱƧ�
Recordset2Excel "�ըƹF���yMBO�^���ˮ֪�", rsMain

'..�b���s��EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
'        .Columns("L").Select
'        .Selection.ClearContents
        .Range("B3").Value = Combo1
        .Range("A1").Select
        '�ƥ��ɮ�
        '    If Dir("C:\LTKK01\DelievryTrack", vbDirectory) = "" Then MkDirs "C:\LTKK01\DelievryTrack"
        '    .ActiveWorkbook.SaveAs "C:\LTKK01\DelievryTrack\DelievryTrack" & Format(Now, "yyyymmddhhMMss") & ".xls"
                
    End With
End If
Set MyXlsApp = Nothing
    
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    dgMain.Height = Frame2.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth - 120
    dgMain.Width = Frame2.Width - 240
End If

End Sub

Private Sub cmdReset_Click()

'���]
Call ClearForm_AllField(Me)

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

'�f�D
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(storerkey) from trp16M", cn, adOpenKeyset, adLockPessimistic

If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        Combo1.AddItem tmp_Rs("storerkey")
        tmp_Rs.MoveNext
    Next
    Combo1.Text = "LMBO01"
End If
tmp_Rs.Close


'�ϰ�
With tmp_Rs
    .Open "select area_code from trp03m order by area_code ", cn
    
    If Not .EOF Then
        .MoveFirst
        For i = 0 To .RecordCount - 1
            List1.AddItem RTrim(tmp_Rs("area_code"))
            .MoveNext
        Next
    
    End If
    .Close
End With

'�����q�N�X
With tmp_Rs
    .Open "select distinct branchid=isnull(branchid,'') from custorders(nolock) order by isnull(branchid,'') ", cn
    
    If Not .EOF Then
        .MoveFirst
        For i = 0 To .RecordCount - 1
            List5.AddItem RTrim(tmp_Rs("branchid"))
            .MoveNext
        Next
    
    End If
    .Close
End With
'
''�f�B���q
'    .Open "select company_code,short_name from trp08m order by company_code ", cn
'
'If Not .EOF Then
'    .MoveFirst
'    For i = 0 To .RecordCount - 1
'        List2.AddItem RTrim(tmp_Rs("company_code")) & "_" & RTrim(tmp_Rs("short_name"))
'        .MoveNext
'    Next
'End If
'.Close

''��O
'    .Open "select distinct rtrim(isnull(priority,'')) as Priority from sdn02t order by priority ", cn
'
'If Not .EOF Then
'    .MoveFirst
'    For i = 0 To .RecordCount - 1
'        List3.AddItem RTrim(tmp_Rs("Priority"))
'        .MoveNext
'    Next
'End If
'.Close

List3.AddItem "I"
List3.AddItem "R"
List3.AddItem "A2B"
List3.AddItem "RC"

''�t�e�ܧO
'    .Open "select distinct rtrim(isnull(facility,'')) as facility from Orders order by facility ", cn
'
'If Not .EOF Then
'    .MoveFirst
'    For i = 0 To .RecordCount - 1
'        List4.AddItem RTrim(tmp_Rs("facility"))
'        .MoveNext
'    Next
'End If
'.Close

'End With

Combo2.ListIndex = 0
optNormal = 1
optAbnormal = 1
txtDeliveryDateS = Format(Now - 1, "YYYYMMDD")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryDateS_Click()
Set objMvdateTarget = txtDeliveryDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub

Private Sub txtDeliveryDateE_Click()
Set objMvdateTarget = txtDeliveryDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub
Private Sub txtOrderDateS_Click()

Set objMvdateTarget = txtOrderDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateE_Click()

Set objMvdateTarget = txtOrderDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
