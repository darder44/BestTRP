VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_OP_RSDNConfirm 
   Caption         =   "�h�fñ����@"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
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
   ScaleHeight     =   6300
   ScaleWidth      =   8370
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3600
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
      StartOfWeek     =   104202241
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
      TabIndex        =   12
      Top             =   2160
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   25
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
      Height          =   2175
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0E0FF&
         Caption         =   "��ƽT�{"
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
         Picture         =   "frm_OP_RSDNConfirm.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   23
         Top             =   1560
         Width           =   3375
         Begin VB.OptionButton optYes 
            Caption         =   "�w�T�{"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optNo 
            Caption         =   "���T�{"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
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
         Left            =   5880
         Picture         =   "frm_OP_RSDNConfirm.frx":1CFA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '�m�����
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
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   2
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
         Alignment       =   2  '�m�����
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
         TabIndex        =   3
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyS 
         Alignment       =   2  '�m�����
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
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   4
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyE 
         Alignment       =   2  '�m�����
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
         MaxLength       =   12
         TabIndex        =   5
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateS 
         Alignment       =   2  '�m�����
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateE 
         Alignment       =   2  '�m�����
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
         TabIndex        =   1
         Top             =   240
         Width           =   1485
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
         Left            =   7080
         Picture         =   "frm_OP_RSDNConfirm.frx":2FF4
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         Top             =   1200
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
         Left            =   7080
         Picture         =   "frm_OP_RSDNConfirm.frx":2CC06
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
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
         Left            =   4680
         Picture         =   "frm_OP_RSDNConfirm.frx":2CF18
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "ñ�檬�A"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1800
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
         Index           =   23
         Left            =   2655
         TabIndex        =   22
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label1 
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
         Index           =   22
         Left            =   120
         TabIndex        =   21
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
         Index           =   3
         Left            =   2655
         TabIndex        =   20
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1005
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
         TabIndex        =   18
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "Key����"
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
         TabIndex        =   17
         Top             =   285
         Width           =   1110
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   26
      Top             =   6030
      Width           =   8370
      _ExtentX        =   14764
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
            Object.Width           =   8123
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
Attribute VB_Name = "frm_OP_RSDNConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmd2Excel_Click()
Recordset2Excel Me.Caption, rsMain
'..�b���s��EXCEL
Set MyXlsApp = Nothing
End Sub

Private Sub cmdOK_Click()
strOtQtyFixOrderkey = rsMain("�q�渹�X")
frm_OTQtyFix.Show vbModal

'��sDatagrid
str_SQL = "select �T�{��� = convert(char(20) , o.effectivedate,120) " & _
", �ꦬ��� = case when s02t.sdnstatus > 0 then o.billedcontainerqty else 0 end " & _
", ñ�檬�A = case " & _
"        when s02t.sdnstatus > 0 then '�w�T�{' " & _
"        Else '���T�{' " & _
"        End " & _
"from sdn01t s01t join sdn02t s02t on s01t.c_route_no = s02t.c_route_no " & _
"join orders o on s02t.extern = o.externorderkey " & _
"where rtrim(s02t.extern) = '" & rsMain("�q�渹�X") & "'"

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn

rsMain("�T�{���") = tmp_Rs("�T�{���") & ""
rsMain("�ꦬ���") = tmp_Rs("�ꦬ���")
rsMain("ñ�檬�A") = tmp_Rs("ñ�檬�A")
rsMain.Update

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_DeliveryDate As String, chc_Orderdate As String, chc_ExternOrderkey, chc_Status As String

str_SQL = "select [Key����] = convert(char(8),o.orderdate,112) " & _
", �ݨD��� = s02t.arrive_date " & _
", �X����� = convert(char(8) , s01t.delivery_date,112) " & _
", �N���^��� = ' ' " & _
", �q�渹�X = rtrim(s02t.extern) " & _
", �y���s�� = ' ' " & _
", �Ȥ�W�� = rtrim(s02t.cust_name) " & _
", �ϰ� = case (select t1m.area_code from trp01m t1m where t1m.consigneekey = o.consigneekey) " & _
"            when 'N1' then '�_��' " & _
"            when 'N2' then '�_��' " & _
"            when 'C' then '����' " & _
"            when 'S' then '���̰�' " & _
"            when 'S1' then '�x�n��' " & _
"            when 'W' then '�g�P�N�e' " & _
"            Else '��L' end " & _
", �T�{��� = convert(char(20) , o.effectivedate,120) " & _
", �ꦬ��� = case when s02t.sdnstatus > 0 then o.billedcontainerqty else 0 end " & _
", ñ�檬�A = case " & _
"        when s02t.sdnstatus > 0 then '�w�T�{' " & _
"        Else '���T�{' " & _
"        End " & _
"from sdn01t s01t join sdn02t s02t on s01t.c_route_no = s02t.c_route_no " & _
"join orders o on s02t.extern = o.externorderkey " & _
"where o.priority = 'R' "

'�q����
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(char(8) , o.orderdate,112) between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and convert(char(8) , o.orderdate,112) = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(char(8) , o.orderdate,112) = '" & txtOrderDateE.Text & "' "
End If

'�X�����
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_Orderdate = "and convert(char(8) , s01t.delivery_date,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_Orderdate = "and convert(char(8) , s01t.delivery_date,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_Orderdate = "and convert(char(8) , s01t.delivery_date,112) = '" & txtDeliveryDateE.Text & "' "
End If

'�q�渹�X
chc_ExternOrderkey = ""
If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chc_ExternOrderkey = "and rtrim(s02t.extern) between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
   chc_ExternOrderkey = "and rtrim(s02t.extern) = '" & txtExternOrderkeyS.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chc_ExternOrderkey = "and rtrim(s02t.extern) = '" & txtExternOrderkeyE.Text & "' "
End If

'�q�檬�A
chc_Status = ""
If optNo = True Then chc_Status = "and s02t.sdnstatus = 0 "
If optYes = True Then chc_Status = "and s02t.sdnstatus > 0 "

'�զX�r��
str_SQL = str_SQL & chc_DeliveryDate & chc_Orderdate & chc_Status & chc_ExternOrderkey

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: cmdOK.Enabled = False: Exit Sub
tmp_Rs.Sort = "KEY����,�q�渹�X"
Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

    .ColumnHeaders = True        '���D�����
    .RowHeight = 300
    .Columns(0).Alignment = dbgCenter
    .Columns(10).Alignment = dbgRight

End With
SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If dgMain.DataSource Is Nothing Then Exit Sub
On Error GoTo err_Handle
If LastRow = Empty Then Exit Sub
'mvDate.Visible = False
'If Left(rsMain("�q�渹�X"), 1) = "R" Then cmdOK.Enabled = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
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
txtDeliveryDateS.Text = "": txtDeliveryDateE.Text = ""
txtOrderDateS.Text = "": txtOrderDateE.Text = ""
txtExternOrderkeyS.Text = "": txtExternOrderkeyE.Text = ""
optNo.Value = True

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

optNo.Value = True
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

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
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
