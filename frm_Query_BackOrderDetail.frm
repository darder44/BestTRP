VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Query_BackOrderDetail 
   Caption         =   "�h���f�P�ڵu������"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
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
   ScaleWidth      =   11160
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3240
      TabIndex        =   8
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
      StartOfWeek     =   93454337
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
      Top             =   2160
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
      Height          =   2175
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdBatchConfirm 
         Caption         =   "�妸���f�T�{"
         Height          =   375
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtExternorderkeyE 
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
         TabIndex        =   24
         Top             =   1320
         Width           =   1485
      End
      Begin VB.TextBox txtExternorderkeyS 
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
         TabIndex        =   23
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "���A�ק�"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cboExpectReceiptOK 
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
         TabIndex        =   21
         Text            =   "Combo2"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   4560
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDeliveryDateS 
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
         MaxLength       =   10
         TabIndex        =   16
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
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
         MaxLength       =   10
         TabIndex        =   15
         Top             =   960
         Width           =   1485
      End
      Begin VB.ComboBox Combo1 
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
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   13
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateE 
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
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateS 
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
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
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
         Left            =   7560
         Picture         =   "frm_Query_BackOrderDetail.frx":0000
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
         Left            =   8760
         Picture         =   "frm_Query_BackOrderDetail.frx":12FA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
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
         Left            =   8760
         Picture         =   "frm_Query_BackOrderDetail.frx":2AF0C
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
         Left            =   6360
         Picture         =   "frm_Query_BackOrderDetail.frx":2B21E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   240
         Width           =   1065
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
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   1365
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
         Index           =   6
         Left            =   2655
         TabIndex        =   25
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���f���A"
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
         TabIndex        =   20
         Top             =   240
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
         Left            =   2655
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
         Caption         =   "�q����"
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
         Visible         =   0   'False
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
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   6030
      Width           =   11160
      _ExtentX        =   19685
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
            Object.Width           =   13044
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
Attribute VB_Name = "frm_Query_BackOrderDetail"
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

'��ƱƧ�
Recordset2Excel Me.Caption, rsMain
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String, chc_DeliveryDate As String, chc_txtExternorderkey As String, chc_Storerkey As String, strTmp As String, i As Integer

str_SQL = "select * from gv_BackOrderDetail where 1 = 1 "

'�q����
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and �q���� between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and �q���� = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and �q���� = '" & txtOrderDateE.Text & "' "
End If

'��f���
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and ��f��� between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_DeliveryDate = "and ��f��� = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and ��f��� = '" & txtDeliveryDateE.Text & "' "
End If

'�q�渹�X
chc_txtExternorderkey = ""
If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chc_txtExternorderkey = "and �q�渹�X between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
   chc_txtExternorderkey = "and �q�渹�X = '" & txtExternOrderkeyS.Text & "' "
ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
   chc_txtExternorderkey = "and �q�渹�X = '" & txtExternOrderkeyE.Text & "' "
End If

'�f�D�s��
If Len(RTrim(Combo1.Text)) > 0 Then chc_Storerkey = " and �f�D�s�� ='" & Combo1.Text & "' "

'���A�����
strTmp = ""
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
            strTmp = strTmp & "'" & List1.List(i) & "',"
    End If
Next
If Len(strTmp) > 0 Then str_SQL = str_SQL & "and �J�w���� in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'�զX�r��
str_SQL = str_SQL & chc_Orderdate & chc_DeliveryDate & chc_Storerkey & chc_txtExternorderkey & " order by �f�D�s�� , ���O , ��f��� , �q�渹�X , �~�� "

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "�f�D�s��,�q�渹�X,����"

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
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdSave_Click()
If rsMain.RecordCount < 1 Then Exit Sub
If (rsMain("���O") = "�h���f" Or rsMain("���O") = "�ڵu��") = False Then MsgBox "�u��TMS���h���f�P�ڵu�����O�A�~���ʧ�怜�f���A!", 64, "�`�N": Exit Sub

'��s���A
cn.Execute "update sdn02t set ExpectReceiptOK = '" & Trim(cboExpectReceiptOK) & "' where extern = '" & rsMain("�q�渹�X") & "' and receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords

If RowsAffect > 0 Then
    rsMain("�J�w����") = Trim(cboExpectReceiptOK)
    MsgBox "TMS�渹 " & rsMain("TMS�渹") & "�A�J�w���A�w��s��" & Trim(cboExpectReceiptOK) & "�C", 64, "���A�ק�"
Else
    MsgBox "��s���ѡA�нT�{���q��O�_�w�����X���T�{?", 16, "���A�ק�"
End If
End Sub

Private Sub cmdBatchConfirm_Click()
On Error GoTo err_Handle

If RTrim(Combo1.Text) = "" Then MsgBox "�п���f�D", 64, cmdBatchConfirm.Caption: Exit Sub

If MsgBox("1.WMS�w���f�������禬��" & vbCrLf & "2.TMS�w�����X���T�{" & vbCrLf & "3.��B�渹(WMS)=�q�渹�X(TMS)", vbOKCancel, "��s����") = vbOK Then

cn.Execute "exec gs_ExpectReceiptOK '" & RTrim(Combo1.Text) & "' ", RowsAffect, adExecuteNoRecords
MsgBox "��s " & RowsAffect & "���q����", , cmdBatchConfirm.Caption
End If

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 200 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'�P�@����
If LastRow = Empty Then Exit Sub

cboExpectReceiptOK = Trim(rsMain("�J�w����"))

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
Call ClearForm_AllField(Me)

Dim i As Integer

'�f�D
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
    'tmp_rs.Close: Set tmp_rs = Nothing
    Combo1.ListIndex = -1
End If

'���A
Call Confirm_Recordset_Closed(tmp_Rs)
With tmp_Rs
    .CursorLocation = 3
    .Open "select distinct isnull(expectReceiptOK,'') as expectReceiptOK from sdn02t order by isnull(expectReceiptOK,'') ", cn
    
If Not tmp_Rs.EOF Then
    
    .MoveFirst
    For i = 0 To .RecordCount - 1
        List1.AddItem Trim(tmp_Rs("expectReceiptOK"))
        cboExpectReceiptOK.AddItem Trim(tmp_Rs("expectReceiptOK"))
        .MoveNext
    Next
    .Close: Set tmp_Rs = Nothing
    cboExpectReceiptOK.ListIndex = 0
End If
    
End With

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
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub
Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
