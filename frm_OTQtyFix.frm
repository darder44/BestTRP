VERSION 5.00
Begin VB.Form frm_OTQtyFix 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�X�f��ƽT�{"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3240
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmdOKPrint 
         Caption         =   "�x�s�C�L"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtOTconfirmuser 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtReceipt_no 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtDeliveryDate 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtCompany 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "���}"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtExternOrderkey 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtStorerkey 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtOTQty 
         Alignment       =   1  '�a�k���
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   0
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "�x�s"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "�T�{�H��"
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "�Ȥ�W��"
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "�f�D�s��"
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
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "�X�f���"
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_OTQtyFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As New ADODB.Recordset
Dim strSql As String

Private Sub cmdCancel_Click()

Unload Me
If frm_OP_CaseConfirm.chkScan.Value = 1 Then Call frm_OP_CaseConfirm.chkScan_Click

End Sub

Private Sub cmdOK_Click()
If Val(txtOTQty) < 0 Then MsgBox "��Ƥ��o���t�ơI", 16, "�`�N": Exit Sub

If Val(txtOTQty) = 0 Then
If MsgBox("��Ƭ� 0�H", vbYesNo, Me.Caption) <> vbYes Then: Exit Sub
End If

Screen.MousePointer = 11
Tran_Level = cn.BeginTrans

Dim i As Integer
i = 0

'��sTRP
strSql = "update trp02t set otqty = '" & txtOTQty.Text & "', otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where receipt_no = '" & RTrim(txtReceipt_no.Text) & "'"
cn.Execute strSql, RowsAffect, adExecuteNoRecords

i = i + RowsAffect

strSql = "update trp02w set otqty = '" & txtOTQty.Text & "', otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where receipt_no = '" & RTrim(txtReceipt_no.Text) & "'"
cn.Execute strSql, RowsAffect, adExecuteNoRecords

i = i + RowsAffect

'��sORT
strSql = "update ort02t set otqty = '" & txtOTQty.Text & "', otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where receipt_no = '" & RTrim(txtReceipt_no.Text) & "'"
cn.Execute strSql, RowsAffect, adExecuteNoRecords

i = i + RowsAffect

strSql = "update ort02w set otqty = '" & txtOTQty.Text & "', otconfirmdate = getdate () , otconfirmuser = '" & User_id & "' where receipt_no = '" & RTrim(txtReceipt_no.Text) & "'"
cn.Execute strSql, RowsAffect, adExecuteNoRecords

i = i + RowsAffect

If i <> 1 Then cn.RollbackTrans: Tran_Level = 0: MsgBox "�s�ɥ��ѡA�Э���!", 16, Me.Caption: Screen.MousePointer = 0: Exit Sub

cn.CommitTrans: Tran_Level = 0

Call cmdCancel_Click
Screen.MousePointer = 0
End Sub

Private Sub cmdOKPrint_Click()

If Val(txtOTQty) > 10000 Then
If MsgBox("��Ƥj��10000�H", vbYesNo + vbDefaultButton2, "���" & Val(txtOTQty) & "?") <> vbYes Then: Exit Sub
End If

Call cmdOK_Click

'��sDataGrid
Call frm_OP_CaseConfirm.UpdateDatagrid

'�C�L
If Val(txtOTQty) > 200 Then
If MsgBox("��Ƥj��200�O�_�T�w�C�L�H", vbYesNo + vbDefaultButton2, "���" & Val(txtOTQty) & "?") <> vbYes Then: GoTo NoPrint
End If

Call frm_OP_CaseConfirm.cmdPrintReport_Click

NoPrint:

Call cmdCancel_Click
Screen.MousePointer = 0
If frm_OP_CaseConfirm.chkScan.Value = 1 Then Call frm_OP_CaseConfirm.chkScan_Click

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Call cmdCancel_Click
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
txtCompany.Text = ""
txtExternOrderkey.Text = ""
txtStorerkey.Text = ""
txtOTQty.Text = ""

strSql = "select Storerkey = t2.Storerkey , receipt_no = rtrim(t2.receipt_no) , Externorderkey = rtrim(t2.extern) " & _
        ", Company = rtrim(t1m.full_name), DeliveryDate = convert(char(8),t2.arrive_Date,112) " & _
        ", OrderOT = sum(case when sp.casecnt = 0 then 1 else ceiling(t3.order_qty /sp.casecnt) end) " & _
        ", OT = isnull(t2.otqty,0) , effectivedate = t2.otconfirmdate ,���ʤH�� = isnull(t2.OTconfirmuser,'���T�{') " & _
        "from trp01m t1m join trp02t t2 on t2.consigneekey = t1m.consigneekey and t1m.storerkey = t2.storerkey and t2.receipt_no = '" & strOtQtyFixOrderkey & "' " & _
        "join trp03t t3 on t3.receipt_no = t2.receipt_no join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey group by t2.STORERKEY,t2.receipt_no,t2.extern,t1m.full_name,convert(char(8),t2.arrive_Date,112),t2.otconfirmdate,isnull(t2.OTconfirmuser,'���T�{'),isnull(t2.otqty,0) " & _
        "Union select Storerkey = t2.Storerkey , receipt_no = rtrim(t2.receipt_no) , Externorderkey = rtrim(t2.extern) " & _
        ", Company = rtrim(t1m.full_name), DeliveryDate = convert(char(8),t2.arrive_Date,112) " & _
        ", OrderOT = sum(case when sp.casecnt = 0 then 1 else ceiling(t3.order_qty /sp.casecnt) end) " & _
        ", OT = isnull(t2.otqty,0) , effectivedate = t2.otconfirmdate ,���ʤH�� = isnull(t2.OTconfirmuser,'���T�{')" & _
        "from trp01m t1m join trp02w t2 on t2.consigneekey = t1m.consigneekey and t1m.storerkey = t2.storerkey and t2.receipt_no = '" & strOtQtyFixOrderkey & "' " & _
        "join trp03w t3 on t3.receipt_no = t2.receipt_no join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey group by t2.STORERKEY,t2.receipt_no,t2.extern,t1m.full_name,convert(char(8),t2.arrive_Date,112),t2.otconfirmdate,isnull(t2.OTconfirmuser,'���T�{'),isnull(t2.otqty,0) " & _
        "Union select Storerkey = t2.Storerkey , receipt_no = rtrim(t2.receipt_no) , Externorderkey = rtrim(t2.extern) " & _
        ", Company = rtrim(t1m.full_name), DeliveryDate = convert(char(8),t2.arrive_Date,112) " & _
        ", OrderOT = sum(case when sp.casecnt = 0 then 1 else ceiling(t3.order_qty /sp.casecnt) end) " & _
        ", OT = isnull(t2.otqty,0) , effectivedate = t2.otconfirmdate ,���ʤH�� = isnull(t2.OTconfirmuser,'���T�{')" & _
        "from trp01m t1m join ort02t t2 on t2.consigneekey = t1m.consigneekey and t1m.storerkey = t2.storerkey and t2.receipt_no = '" & strOtQtyFixOrderkey & "' " & _
        "join ort03t t3 on t3.receipt_no = t2.receipt_no join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey group by t2.STORERKEY,t2.receipt_no,t2.extern,t1m.full_name,convert(char(8),t2.arrive_Date,112),t2.otconfirmdate,isnull(t2.OTconfirmuser,'���T�{'),isnull(t2.otqty,0) " & _
        "Union select Storerkey = t2.Storerkey , receipt_no = rtrim(t2.receipt_no) , Externorderkey = rtrim(t2.extern) " & _
        ", Company = rtrim(t1m.full_name), DeliveryDate = convert(char(8),t2.arrive_Date,112) " & _
        ", OrderOT = sum(case when sp.casecnt = 0 then 1 else ceiling(t3.order_qty /sp.casecnt) end) " & _
        ", OT = isnull(t2.otqty,0) , effectivedate = t2.otconfirmdate ,���ʤH�� = isnull(t2.OTconfirmuser,'���T�{')" & _
        "from trp01m t1m join ort02w t2 on t2.consigneekey = t1m.consigneekey and t1m.storerkey = t2.storerkey and t2.receipt_no = '" & strOtQtyFixOrderkey & "' " & _
        "join ort03w t3 on t3.receipt_no = t2.receipt_no join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey group by t2.STORERKEY,t2.receipt_no,t2.extern,t1m.full_name,convert(char(8),t2.arrive_Date,112),t2.otconfirmdate,isnull(t2.OTconfirmuser,'���T�{'),isnull(t2.otqty,0) "

rsMain.Open strSql, cn

txtStorerkey.Text = rsMain("Storerkey")
txtReceipt_no.Text = rsMain("receipt_no")
txtExternOrderkey.Text = rsMain("Externorderkey")
txtDeliveryDate.Text = rsMain("DeliveryDate") & ""
txtCompany = rsMain("Company") & ""

If rsMain("���ʤH��") = "���T�{" Then

    '��ƹw��
'    str_SQL = "select sp.sku ,sp.casecnt " & _
'        ", OrderCS = sum(case when sp.casecnt = 0 then 1 else floor(t3.order_qty /sp.casecnt) end) " & _
'        ", OrderEA = sum(case when sp.casecnt = 0 then 0 else (cast(t3.order_qty as int) % cast(sp.casecnt as int)/sp.casecnt) end) " & _
'        "from trp02t t2 join trp03t t3 on t3.receipt_no = t2.receipt_no and t2.receipt_no = '" & strOtQtyFixOrderkey & "' " & _
'        "join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey " & _
'        "group by sp.sku,sp.casecnt " & _
'        "union select sp.sku ,sp.casecnt " & _
'        ", OrderCS = sum(case when sp.casecnt = 0 then 1 else floor(t3.order_qty /sp.casecnt) end) " & _
'        ", OrderEA = sum(case when sp.casecnt = 0 then 0 else (cast(t3.order_qty as int) % cast(sp.casecnt as int)/sp.casecnt) end) " & _
'        "from ort02w t2 join ort03w t3 on t3.receipt_no = t2.receipt_no and t2.receipt_no = '" & strOtQtyFixOrderkey & "' " & _
'        "join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey " & _
'        "group by sp.sku,sp.casecnt " & _
'        "union select sp.sku,sp.casecnt " & _
'        ", OrderCS = sum(case when sp.casecnt = 0 then 1 else floor(t3.order_qty /sp.casecnt) end) " & _
'        ", OrderEA = sum(case when sp.casecnt = 0 then 0 else (cast(t3.order_qty as int) % cast(sp.casecnt as int)/sp.casecnt) end) " & _
'        "from ort02t t2 join ort03t t3 on t3.receipt_no = t2.receipt_no and t2.receipt_no = '" & strOtQtyFixOrderkey & "' " & _
'        "join gv_skuxpack sp on sp.sku = t3.product_no and sp.storerkey = t3.storerkey " & _
'        "group by sp.sku,sp.casecnt "

    str_SQL = "exec gs_ProOTQty '" & strOtQtyFixOrderkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
    txtOTQty.Text = tmp_Rs("preotqty")
 
    tmp_Rs.Close
Else
    txtOTQty.Text = Val(rsMain("ot")) & ""
End If

txtOTconfirmuser.Text = rsMain("���ʤH��") & ""

'��������
If Val(txtDeliveryDate) > lngDueDate Then
    cmdOK.Enabled = True
    cmdOKPrint.Enabled = True
Else
    cmdOK.Enabled = False
    cmdOKPrint.Enabled = False
End If

txtOTQty.SelStart = 0: txtOTQty.SelLength = Len(txtOTQty)

rsMain.Close: Set rsMain = Nothing
Screen.MousePointer = 0
End Sub

Private Sub txtOTQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And cmdOKPrint.Enabled And frm_OP_CaseConfirm.chkScan.Value = 1 Then
    '���y�Ҧ�
    Call cmdOKPrint_Click
Else
    'Call cmdOK_Click
End If

End Sub
