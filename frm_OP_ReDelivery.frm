VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_OP_ReDelivery 
   Caption         =   "�����q��A�t�e�@�~"
   ClientHeight    =   7140
   ClientLeft      =   255
   ClientTop       =   885
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11475
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   1110
      TabIndex        =   19
      Top             =   1380
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
      StartOfWeek     =   106561537
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_OrderDetail 
      Height          =   3165
      Left            =   90
      TabIndex        =   2
      Top             =   3900
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5583
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSDataGridLib.DataGrid dg_Orders 
      Height          =   2505
      Left            =   75
      TabIndex        =   1
      Top             =   1365
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4419
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
   Begin VB.Frame fam_Header 
      Height          =   1365
      Left            =   90
      TabIndex        =   0
      Top             =   -15
      Width           =   11300
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
         Height          =   990
         Index           =   0
         Left            =   10050
         Picture         =   "frm_OP_ReDelivery.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   17
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton cmd_ReBuildOrders 
         BackColor       =   &H008080FF&
         Caption         =   "�q��A�t�e"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   6885
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   16
         Top             =   540
         Width           =   2490
      End
      Begin VB.CommandButton cmd_OrdersQuery 
         BackColor       =   &H0080C0FF&
         Caption         =   "�q��d��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   5310
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   15
         Top             =   120
         Width           =   720
      End
      Begin VB.TextBox txt_OrderKey 
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
         Height          =   360
         Left            =   3720
         TabIndex        =   14
         Top             =   915
         Width           =   1560
      End
      Begin VB.TextBox txt_ConsigneeKey 
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
         Height          =   360
         Left            =   3720
         TabIndex        =   12
         Top             =   540
         Width           =   1560
      End
      Begin VB.TextBox txt_Extern 
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
         Height          =   360
         Left            =   3720
         TabIndex        =   10
         Top             =   165
         Width           =   1560
      End
      Begin VB.TextBox txt_CarID 
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
         Height          =   360
         Left            =   1065
         TabIndex        =   8
         Top             =   915
         Width           =   1560
      End
      Begin VB.TextBox txt_DeliveryDate 
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
         Height          =   360
         Left            =   1065
         TabIndex        =   6
         Top             =   540
         Width           =   1560
      End
      Begin VB.TextBox txt_RouteNo 
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
         Height          =   360
         Left            =   1065
         TabIndex        =   4
         Top             =   165
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�`�N�G�C���ȳB�z�@���A�t�e�q��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   195
         Left            =   6555
         TabIndex        =   18
         Top             =   210
         Width           =   3150
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         Height          =   825
         Left            =   6825
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   2835
         TabIndex        =   13
         Top             =   1005
         Width           =   840
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   2835
         TabIndex        =   11
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   2835
         TabIndex        =   9
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   7
         Top             =   990
         Width           =   840
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   5
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   255
         Width           =   840
      End
   End
End
Attribute VB_Name = "frm_OP_ReDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private iLoop As Double              '�j��p��

Private blOrdersEvent As Boolean
Private rs_Orders As ADODB.Recordset

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_OrdersQuery_Click()
'�q��d��
Screen.MousePointer = vbHourglass
cmd_OrdersQuery.Enabled = False
DoEvents: DoEvents
Set dg_Orders.DataSource = Nothing
'�q��W��
Call SetGDFormat_OrderDetail

Dim strWhere As String, strTmp As String, tmp_data() As String, intloop As Integer
strWhere = ""
'���u�s��
strTmp = ""
If Len(txt_RouteNo.Text) > 0 Then
   strTmp = " ���u�s�� like '" & Trim(txt_RouteNo.Text) & "%' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'�X�����
strTmp = ""
If Len(txt_DeliveryDate.Text) > 0 Then
   strTmp = " �X����� like '" & Trim(txt_DeliveryDate.Text) & "%' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'���P���X
strTmp = ""
If Len(txt_CarID.Text) > 0 Then
   strTmp = " ���P���X like '" & Trim(txt_CarID.Text) & "%' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'�f�D�渹
strTmp = ""
If Len(txt_Extern.Text) > 0 Then
   strTmp = " �f�D�渹 like '" & Trim(txt_Extern.Text) & "%' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'�Ȥ�s��
strTmp = ""
If Len(txt_ConsigneeKey.Text) > 0 Then
   strTmp = " �Ȥ�s�� like '" & Trim(txt_ConsigneeKey.Text) & "%' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If
'�q��s��
strTmp = ""
If Len(txt_OrderKey.Text) > 0 Then
   strTmp = " �q��s�� like '" & Trim(txt_OrderKey.Text) & "%' "
   If Len(strWhere) = 0 Then
      strWhere = strTmp
   Else
      strWhere = strWhere & " and " & strTmp
   End If
End If

str_SQL = "select �X�����,���P���X,����,���u�s��,�q��s��,�f�D�渹,�Ȥ�s��,�Ȥ�W��,�r�p�H," & _
          " �q��,�@��h��,�B�餽�q,�q���,�X�f��,�q��Ƶ�,�f�D,�A�t�e " & _
          "from RejectOrder_Orders "

If Len(strWhere) > 0 Then
   str_SQL = str_SQL & " Where " & strWhere
End If
str_SQL = str_SQL & " Order by �X�����,���P���X,���u�s��,�q��s��"

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_OrdersQuery.Enabled = True
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Orders)
tmp_Rs.Close

With dg_Orders
     .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
     .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
     .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
     .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
End With
rs_Orders.MoveFirst
blOrdersEvent = False
Set dg_Orders.DataSource = rs_Orders
With dg_Orders
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '�X�����
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900        '���P���X
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 500        '����
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 1100       '���u�s��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1100       '�q��s��
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 900        '�f�D�渹
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1100       '�Ȥ�s��
    .Columns(7).Alignment = dbgCenter
    .Columns(8).Width = 2000       '�Ȥ�W��
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 800        '�r�p�H
    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 1100      '�q��
    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 900       '�@��h��
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1600      '�B�餽�q
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 1000      '�q���
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1000      '�X�f��
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1000      '�q��Ƶ�
    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 600      '�f�D
    .Columns(16).Alignment = dbgLeft
    .Columns(17).Width = 700      '�A�t�e����
    .Columns(17).Alignment = dbgLeft
End With
rs_Orders.MoveFirst
blOrdersEvent = True
cmd_OrdersQuery.Enabled = True
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��d��", Me.Caption, "cmd_OrdersQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_OrdersQuery.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_ReBuildOrders_Click()
'�q��A�t�e
If rs_Orders Is Nothing Then Exit Sub
If dg_Orders.SelBookmarks.Count = 0 Then
   msg_text = "������q��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
'�ˬd�O�_������q��Ӷ�
Dim dbCount As Double
dbCount = 0
With dg_OrderDetail
     For iLoop = 1 To .Rows - 2
         .Row = iLoop
         .Col = 1       '����ѧO
         If Trim(.Text) <> "" Then
            dbCount = dbCount + 1
            Exit For
         End If
     Next iLoop
End With
If dbCount = 0 Then
   msg_text = "��ƿ��~�G������A�t�e���q��Ӷ�"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

On Error GoTo err_Handle
cmd_ReBuildOrders.Enabled = False
Screen.MousePointer = vbHourglass
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
Dim strNewOrderKey As String, strSrcOrderKey As String

Tran_Level = 0
Tran_Level = cn.BeginTrans

'1.�����s�q�椧�q��s��
strSrcOrderKey = rs_Orders.Fields("�q��s��").Value
str_SQL = "Select Cast(Code as integer) as AvailNo From CodeLKUP Where ListName = 'RETURNORDER'  "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   strNewOrderKey = "RD" & Format(1, "00000000")
   str_SQL = "Insert into CodeLKUP (ListName,Code,Description,AddWho,EditWho) Values ('RETURNORDER',2,'�����q��A�t�e���q��s�����X','" & User_id & "','" & User_id & "')"
Else
   strNewOrderKey = "RD" & Format(tmp_Rs.Fields("AvailNo").Value, "00000000")
   str_SQL = "Update CodeLKUP Set Code = " & (tmp_Rs.Fields("AvailNo").Value + 1) & " Where ListName = 'CUTORDERSNO'"
End If
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
tmp_Rs.Close

'2. ���� TRP02W
str_SQL = "Insert into TRP02W(" & _
          "   RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
          "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM) " & _
          "Select '" & strNewOrderKey & "',RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
          "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM " & _
          "From TRP02T Where Receipt_No = '" & strSrcOrderKey & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "Update TRP02W Set Description = Isnull(Rtrim(Description),'')+'[�����q��A�t�e]' Where Receipt_No = '" & strNewOrderKey & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'3.���� TRP03W
With dg_OrderDetail
     For iLoop = 1 To .Rows - 2
         .Row = iLoop
         .Col = 1       '����ѧO
         If Trim(.Text) <> "" Then
            .Col = 2    '�����s��
            str_SQL = "Insert into TRP03W(" & _
                      "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
                      "Select A.STORERKEY,'" & strNewOrderKey & "',A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
                      "From TRP03T A Where a.Receipt_No = '" & strSrcOrderKey & "' and a.Seq_No = " & .Text
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
         End If
     Next iLoop
End With

'4.Update TRP02T Case_cnt , Pallet_Qty , Weight , Volumn_Weight
str_SQL = "update TRP02W set CASE_CNT=(" & _
          "  select sum(A.ORDER_QTY/B.CaseCnt) from TRP03W A,Pack B,Sku C " & _
          "   where TRP02W.RECEIPT_NO=A.RECEIPT_NO and A.PRODUCT_NO=C.Sku and A.StorerKey=C.StorerKey and C.PackKey=B.PackKey and TRP02W.Receipt_No = '" & strNewOrderKey & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "update TRP02W set WEIGHT=(" & _
          "  select sum(TRP03W.WEIGHT) from TRP03W where TRP02W.RECEIPT_NO=TRP03W.RECEIPT_NO and TRP02W.Receipt_No = '" & strNewOrderKey & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "update TRP02W set VOLUMN_WEIGHT=(" & _
          "  select sum(TRP03W.VOLUMN_WEIGHT) from TRP03W where TRP02W.RECEIPT_NO=TRP03W.RECEIPT_NO and TRP02W.Receipt_No = '" & strNewOrderKey & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

str_SQL = "update TRP02W set Pallet_Qty=(" & _
          "select sum(TRP03W.Pallet_Qty) from TRP03W where TRP02W.RECEIPT_NO=TRP03W.RECEIPT_NO and TRP02W.Receipt_No = '" & strNewOrderKey & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'5.�g�J trp04t �A�t�e����
str_SQL = "Insert into TRP04T (Receipt_No,Receipt_Old_No,StorerKey,Extern,AddWho) Values ('" & _
          strNewOrderKey & "','" & strSrcOrderKey & "','" & rs_Orders.Fields("�f�D").Value & "','" & rs_Orders.Fields("�f�D�渹").Value & "','" & User_id & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

If Tran_Level <> 0 Then
   cn.CommitTrans
   Tran_Level = 0
End If

dg_Orders.SelBookmarks.Remove 0
msg_text = "�q��A�t�e�B�z�����A�s�q��s���G" & strNewOrderKey
MsgBox msg_text, vbOKOnly + vbInformation, msg_title
cmd_ReBuildOrders.Enabled = True
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
    CreateErrorLog Me.Name & "-�q��A�t�e", Me.Caption, "cmd_RebuildOrders_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmd_ReBuildOrders.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub dg_OrderDetail_Click()
'�q��W��
'�I�@���G����A�I�ĤG���G�������
Dim i As Double
With dg_OrderDetail
     .Col = 3   '�f��
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 1
     If Len(.Text) = 0 Then
        .Text = "V"
     Else
        .Text = ""
     End If
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub dg_Orders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'�������� >> �ݳ��쨮���C�� >> ���
If blOrdersEvent Then
   With dg_Orders
        '�ϥ���ܿ������ƦC
        If Not rs_Orders.EOF Then
           dg_Orders.SelBookmarks.Add rs_Orders.Bookmark
           Call Display_OrderDetail(rs_Orders.Fields("�q��s��").Value, rs_Orders.Fields("�A�t�e").Value)
        End If
   End With
End If
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�����i�X�ި�@�~"
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

'�q��W��
Call SetGDFormat_OrderDetail
End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   'SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   'SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   fam_Header.Left = fam_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Orders.Height = dg_Orders.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Orders.Width = dg_Orders.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_OrderDetail.Top = dg_OrderDetail.Top - (dbsrcFormHeight - Me.ScaleHeight)
   dg_OrderDetail.Width = dg_OrderDetail.Width - (dbsrcFormWidth - Me.ScaleWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   'SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   'SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   fam_Header.Left = fam_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Orders.Height = dg_Orders.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Orders.Width = dg_Orders.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_OrderDetail.Top = dg_OrderDetail.Top + (Me.ScaleHeight - dbsrcFormHeight)
   dg_OrderDetail.Width = dg_OrderDetail.Width + (Me.ScaleWidth - dbsrcFormWidth)
   
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
Set frm_OP_ReDelivery = Nothing
End Sub

Private Sub SetGDFormat_OrderDetail()
'�W�١GSetGDFormat_OrderDetail
'���O�G�Ƶ{��
'�\��G�M���ó]�w [�q����Ӹ��] ��ܮ榡
'�ѼơG�ǤJ�ȡG�L
Dim sub_var1 As Integer, sub_var2 As Integer
dg_OrderDetail.Visible = False
With dg_OrderDetail
     .FixedRows = 1: .Cols = 13
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
     .ColWidth(0) = 500
     .ColWidth(1) = 400
     .ColWidth(2) = 500
     .ColWidth(3) = 700
     .ColWidth(4) = 2500
     .ColWidth(5) = 900
     .ColWidth(6) = 900
     .ColWidth(7) = 900
     .ColWidth(8) = 900
     .ColWidth(9) = 900
     .ColWidth(10) = 900
     .ColWidth(11) = 900
     .ColWidth(12) = 900
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "�s��"
     .Col = 1: .Text = "��"
     .Col = 2: .Text = "����"
     .Col = 3: .Text = "�f��"
     .Col = 4: .Text = "�~�W"
     .Col = 5: .Text = "�z�f�c��"
     .Col = 6: .Text = "�z�f�O��"
     .Col = 7: .Text = "�z�f���q"
     .Col = 8: .Text = "�z�f���n"
     .Col = 9: .Text = "�q��c��"
     .Col = 10: .Text = "�q��O��"
     .Col = 11: .Text = "�q�歫�q"
     .Col = 12: .Text = "�q����n"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignRightCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignLeftCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignRightCenter
     .ColAlignment(10) = flexAlignRightCenter
     .ColAlignment(11) = flexAlignRightCenter
     .ColAlignment(12) = flexAlignRightCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_OrderDetail.Visible = True
End Sub

Private Sub Display_OrderDetail(ByVal strOrderkey As String, ByVal dbRDCount As Double)
'��ܭq��W��
Screen.MousePointer = vbHourglass
'�q��W��
Call SetGDFormat_OrderDetail
str_SQL = "Select ����,�f��,�~�W,�z�f�q,�z�f���q,�z�f���n,�z�f�O��,�q��q,�q�歫�q,�q����n,�q��O�� " & _
          "From RejectOrder_OrderDetail Where �q��s�� = '" & strOrderkey & "' Order by ����"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX���󪺭q��W�Ӹ��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
dg_OrderDetail.Visible = False
Do While Not tmp_Rs.EOF
   With dg_OrderDetail
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '�Ǹ�
       .Text = .Rows - 2
       .Col = 1    '����ѧO
        If dbRDCount = 0 Then
          .Text = "V"
        Else
          .Text = ""
        End If
       .Col = 2    '����
       .Text = tmp_Rs.Fields("����").Value
       .Col = 3    '�f��
       .Text = tmp_Rs.Fields("�f��").Value
       .Col = 4    '�~�W
       .Text = tmp_Rs.Fields("�~�W").Value
       .Col = 5    '�z�f�c��
       .Text = tmp_Rs.Fields("�z�f�q").Value
       .Col = 6    '�z�f�O��
       .Text = tmp_Rs.Fields("�z�f�O��").Value
       .Col = 7    '�z�f���q
       .Text = tmp_Rs.Fields("�z�f���q").Value
       .Col = 8    '�z�f���n
       .Text = tmp_Rs.Fields("�z�f���n").Value
       .Col = 9    '�q��c��
       .Text = tmp_Rs.Fields("�q��q").Value
       .Col = 10   '�q��O��
       .Text = tmp_Rs.Fields("�q��O��").Value
       .Col = 11   '�q�歫�q
       .Text = tmp_Rs.Fields("�q�歫�q").Value
       .Col = 12   '�q����n
       .Text = tmp_Rs.Fields("�q����n").Value
  End With
  tmp_Rs.MoveNext
Loop
dg_OrderDetail.Visible = True
tmp_Rs.Close
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-��ܭq��W��", Me.Caption, "Form ���� SubOrigram Display_OrderDetail", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'������
Select Case mvDate.Tag
    Case "�X�����"
       txt_DeliveryDate.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_DeliveryDate_Click()
'�X�����
If Trim(txt_DeliveryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2))
   End If
End If
mvDate.Tag = "�X�����"
mvDate.Top = fam_Header.Top + txt_DeliveryDate.Top + txt_DeliveryDate.Height
mvDate.Left = fam_Header.Left + txt_DeliveryDate.Left
mvDate.Visible = True
End Sub
