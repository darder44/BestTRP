VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_OP_BacktoEXE 
   Caption         =   "�ƨ���Ʀ^�ǳ]�w"
   ClientHeight    =   7140
   ClientLeft      =   270
   ClientTop       =   1170
   ClientWidth     =   13455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   13455
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   2265
      TabIndex        =   8
      Top             =   1410
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
      StartOfWeek     =   92667905
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin VB.Frame fra_ExtraQuery 
      Appearance      =   0  '����
      BackColor       =   &H80000003&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   90
      TabIndex        =   9
      Top             =   660
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox chk_AddWho 
         BackColor       =   &H80000003&
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
         Left            =   195
         TabIndex        =   11
         Top             =   510
         Value           =   1  '�֨�
         Width           =   1815
      End
      Begin VB.CheckBox chk_Status 
         BackColor       =   &H80000003&
         Caption         =   "�^�Ǫ��A�z��(�s�ظ��s)"
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
         Left            =   210
         TabIndex        =   10
         Top             =   240
         Value           =   1  '�֨�
         Width           =   2685
      End
   End
   Begin MSDataGridLib.DataGrid dg_Orders 
      Height          =   3285
      Left            =   90
      TabIndex        =   4
      Top             =   3780
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5794
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
      Height          =   705
      Left            =   90
      TabIndex        =   1
      Top             =   -45
      Width           =   12615
      Begin VB.CheckBox Chk_all 
         Caption         =   "����"
         Height          =   375
         Left            =   9960
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_UTL 
         BackColor       =   &H008080FF&
         Caption         =   "�^�Ƕ}�o��"
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
         Height          =   525
         Left            =   8595
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd_ShowQuery 
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
         TabIndex        =   12
         Top             =   285
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmd_Update 
         BackColor       =   &H008080FF&
         Caption         =   "�^�ǳ]�w�T�{"
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
         Left            =   7200
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
         Top             =   120
         Width           =   1395
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
         Height          =   525
         Index           =   0
         Left            =   11280
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmd_RouteList 
         BackColor       =   &H8000000A&
         Caption         =   "���s�d��"
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
         Left            =   2655
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   120
         Width           =   1485
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
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1290
         TabIndex        =   2
         Top             =   225
         Width           =   1350
      End
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   120
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000001&
         BackStyle       =   1  '���z��
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         Height          =   615
         Left            =   7080
         Top             =   75
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�X�����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�s�ɩ�C:\Best\Order2WMS\"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   17
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   2475
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Route 
      Height          =   3045
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5371
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm_OP_BacktoEXE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Dim i As Double

Private rs_Orders As ADODB.Recordset   '������s���q����
Private rs_Src As ADODB.Recordset           '��l�q����

Private Sub Chk_all_Click()
    If Chk_all.Value = 1 Then
        '����
        For i = 1 To dg_Route.Rows - 2
            dg_Route.Row = i
            dg_Route.Col = 5
            If dg_Route.Text = UCase(User_id) Then
               dg_Route.Col = 1
               dg_Route.Text = "V"
               dg_Route.Col = 2
               Call Display_RouteOrders(dg_Route.Text)
            Else
               dg_Route.Col = 1
               dg_Route.Text = ""
            End If
            dg_Route.Col = 0
        Next
    Else
        '������
        For i = 1 To dg_Route.Rows - 2
            dg_Route.Row = i
            dg_Route.Col = 5
            If dg_Route.Text = UCase(User_id) Then
               dg_Route.Col = 1
               dg_Route.Text = " "
               dg_Route.Col = 2
               Call Display_RouteOrders(dg_Route.Text)
            Else
               dg_Route.Col = 1
               dg_Route.Text = ""
            End If
            dg_Route.Col = 0
        Next
    End If
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_RouteList_Click()
'���s�d��
On Error GoTo err_Handle

'�ˬd���`���(�w�]�w�^�ǡA���O���^��WMS)
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open "select route_no from trp01t where exe_confirm = 2 and uploadwho is null ", cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then MsgBox "�o�{���`���s���(" & tmp_Rs("route_no") & ")�A�гq����T�H���B�z�I", 16, "�`�N�I"
tmp_Rs.Close

Screen.MousePointer = vbHourglass
Call SetGridFormat_Route
Set dg_Orders.DataSource = Nothing
Set rs_Orders = Nothing
fra_ExtraQuery.Visible = False


'�w�^�Ǹ��s�w���� View �z��簣--
'���^��ids�����s���]�t���Ƨ����q��(�@��h��)--daniel 20041129
str_SQL = "SELECT ���u�s��,�X�����,�^�Ǫ��A,�ƨ���,�^�Ǫ�,�c��,�O��,���n,���q,���P���X,����,�r�p�H,�w�p������," & _
        "�w�p����ɶ� , �X�Y�Ȧs, �G���ƨ����s, �G���ƨ�����, �G���ƨ����� " & _
        "FROM BacktoEXE_srcRoute where 1 = 1 "
        
'�ư��P�@�Ȥ�q��s�������B�����ƨ����--���������� 20090121
'        "where ���u�s�� not in ( select route_no  from trp02t where extern In ( select extern from trp02w where extern in ( select extern from trp02t  group by extern  having count(*)>1 )))"

Dim strWhere As String
strWhere = ""

'�z�� [�X�����] ���u�s��
If Len(Trim(txt_DeliveryDate.Text)) > 0 Then
    strWhere = " And �X����� = '" & txt_DeliveryDate.Text & "' "
End If

'�z�� [�^�Ǫ��A=0] ���u�s��
'If chk_Status.Value = vbChecked Then
'    strWhere = strWhere & " And �^�ǽX = '0' "
'End If

'���w�n�J�ϥΪ̥u�i�d�ߦۤv�Ʃw�����u�s��
'If chk_AddWho.Value = vbChecked Then
'   If Len(strWhere) = 0 Then
'      strWhere = " �ƨ��� = '" & User_id & "' "
'   Else
'      strWhere = strWhere & " and �ƨ��� = '" & User_id & "' "
'   End If
'End If

str_SQL = str_SQL & strWhere & " Order by ���u�s�� "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX���󪺸��u�s�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

dg_Route.Visible = False
Do While Not tmp_Rs.EOF
   With dg_Route
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '�Ǹ�
       .Text = .Rows - 2
       .Col = 1    '����ѧO
       .Text = ""
       .Col = 2    '���u�s��
       .Text = tmp_Rs.Fields("���u�s��").Value
       .Col = 3    '�X�����
       .Text = tmp_Rs.Fields("�X�����").Value
       .Col = 4    '�^�Ǫ��A
       .Text = tmp_Rs.Fields("�^�Ǫ��A").Value
       .Col = 5    '���u�s���ƨ���
       .Text = tmp_Rs.Fields("�ƨ���").Value
       .Col = 6    '�^�ǳ]�w�H��
       .Text = tmp_Rs.Fields("�^�Ǫ�").Value
       .Col = 7    '�c��
       .Text = tmp_Rs.Fields("�c��").Value
       .Col = 8    '�O��
       .Text = tmp_Rs.Fields("�O��").Value
       .Col = 9    '���n
       .Text = tmp_Rs.Fields("���n").Value
       .Col = 10    '���q
       .Text = tmp_Rs.Fields("���q").Value
       .Col = 11    '���P���X
       .Text = tmp_Rs.Fields("���P���X").Value
       .Col = 12   '����
       .Text = tmp_Rs.Fields("����").Value
       .Col = 13   '�r�p�H
       .Text = tmp_Rs.Fields("�r�p�H").Value
       .Col = 14   '�w�p������
       .Text = tmp_Rs.Fields("�w�p������").Value
       .Col = 15   '�w�p����ɶ�
       .Text = tmp_Rs.Fields("�w�p����ɶ�").Value
       .Col = 16   '�X�Y�Ȧs
       .Text = tmp_Rs.Fields("�X�Y�Ȧs").Value
       .Col = 17   '�G���ƨ����s
       .Text = tmp_Rs.Fields("�G���ƨ����s").Value
       .Col = 18   '�G���ƨ�����
       .Text = tmp_Rs.Fields("�G���ƨ�����").Value
       .Col = 17   '�G���ƨ�����
       .Text = tmp_Rs.Fields("�G���ƨ�����").Value
  End With
  tmp_Rs.MoveNext
Loop
dg_Route.Visible = True
tmp_Rs.Close
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-���J�ݦ^�Ǹ��s", Me.Caption, "cmd_RouteList_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_ShowQuery_Click()
'����B�~���d�߱���
fra_ExtraQuery.Visible = Not fra_ExtraQuery.Visible
End Sub

Private Sub cmd_Update_Click()

On Error GoTo err_Handle
With dg_Route
     If .Rows = 2 Then Exit Sub
     
    '��Ʈw���ʥ��--�_�I
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '�^�ǳ]�w�T�{
    Screen.MousePointer = 11: cmd_Update.Enabled = False

     Dim strRouteNo As String, strTMSorderkey As String, intLineNumber As Integer, str As String, strRoutekey As String, strKeycount As String, strWMSorderkeyS As String
     strRouteNo = ""
     For i = 1 To .Rows - 2
        .Row = i
        .Col = 1   '����ѧO
        If Trim(.Text) <> "" Then
           .Col = 2   '���u�s��
           If strRouteNo = "" Then
              strRouteNo = "'" & RTrim(.Text) & "'"
           Else
              strRouteNo = strRouteNo & ",'" & RTrim(.Text) & "'"
           End If
           '��s�@��h�����O & �p����έq�涵���s��
           'CALL SQL Server Stored Procedure �B�z
           str_SQL = "exec TRPPlan_BacktoEXE " & .Text & ""
           cn.CommandTimeout = 120
           cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
           .Col = 4
           .Text = "�]�w�^��"
        End If
     Next i
End With

If strRouteNo <> "" Then

    Dim rsTmp As New ADODB.Recordset
    rsTmp.CursorLocation = 3
    '�ˬdTMS�渹�O�_������J
    str_SQL = "select �f�D = rtrim(Storerkey) " & _
                ",���u�s�� = rtrim(route) " & _
                ",TMS�渹 = updatesource " & _
                ",WMS�渹 = orderkey " & _
                ",�f�D�渹 = rtrim(externorderkey) " & _
                ",���O = rtrim(type) " & _
                ",�q���� = orderdate " & _
                ",��f��� = deliverydate " & _
                ",�Ȥ�W�� = rtrim(c_company) " & _
                ",�a�} = rtrim(c_address1) " & _
                ",�Ƶ� = notes " & _
                "from " & strWMSDB & "..orders (nolock) " & _
                "where updatesource in ( select receipt_no from trp02t where route_no in (" & strRouteNo & ")) " & _
                "order by updatesource , orderkey "
                
    rsTmp.Open str_SQL, cn
    If Not rsTmp.EOF Then
        cn.RollbackTrans: Tran_Level = 0
        cmd_Update.Enabled = True
        MsgBox "WMS�t�θ̵o�{TMS�渹���ơA�нT�{�q��O�_������J!!", 16, "��J�@�~�פ�"
        Call Recordset2Excel("Order2WMS-�渹����", rsTmp)
        If Dir("C:\BEST\Order2WMS", vbDirectory) = "" Then MkDirs "C:\BEST\Order2WMS"
        MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\Order2WMS\Order2WMS-�渹����_" & Format(Now, "yyyymmddhhMMss") & ".xls"
        Set MyXlsApp = Nothing
        rsTmp.Close: Set rsTmp = Nothing
        Screen.MousePointer = 0
        Exit Sub
    End If
    rsTmp.Close
    
    '�]�w�^��
    str_SQL = "Update TRP01T Set EXE_CONFIRM = '1',UploadWho='" & User_id & "' Where Route_No in (" & strRouteNo & ") and EXE_CONFIRM not in ('2','9')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    '���ƨ����
    str_SQL = "select * from gv_order2wms Where Route_No in (" & strRouteNo & ") order by route_no , receipt_no , OrderLineNumber "
    rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockPessimistic
    
    If rsTmp.EOF Then
'        MsgBox "�L�ݭn�^��WMS���q����!", 64, Me.Caption--Mark by Gemini @20150915 4 �קKUser�����T�{����w��Ʈw
    Else
        Dim rsKeycount As New ADODB.Recordset
        rsKeycount.CursorLocation = 3
             
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
                     
            If Trim(rsTmp("receipt_no")) <> strTMSorderkey Then
            
                '��WMS�q��渹
                rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='order' ", cn, adOpenForwardOnly, adLockPessimistic
                '�q��渹+1
                cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'order'", RowsAffect, adExecuteNoRecords
                strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                rsKeycount.Close
                
                '�g�J���Y
                If Trim(rsTmp("storerkey")) = "LLFA01" Then
                    '�Q�צh���
                    str_SQL = "insert into " & strWMSDB & "..orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,route,DeliveryDate,stop,consigneekey,c_contact1,c_Company,C_Address1,C_Address2,C_Zip,C_Phone1,type,Notes,updatesource,door,customerorderkey,b_company,incoterm) " & _
                            "values( '" & strKeycount & "','" & Trim(rsTmp("StorerKey")) & "','" & Trim(rsTmp("ExternOrderKey")) & "','" & Format(rsTmp("OrderDate"), "yyyy/mm/dd HH:mm:ss") & "','" & rsTmp("route_no") & "','" & Format(rsTmp("DeliveryDate"), "yyyy/mm/dd HH:mm:ss") & "','" & rsTmp("VEHICLE_ID_NO") & "','" & rsTmp("consigneekey") & "','" & rsTmp("c_contact1") & "','" & rsTmp("c_Company") & "','" & rsTmp("C_Address1") & "','" & rsTmp("C_Address2") & "','" & Trim(rsTmp("C_Zip")) & "','" & Trim(rsTmp("C_Phone1")) & "','" & Trim(rsTmp("priority")) & "','" & rsTmp("notes") & "','" & rsTmp("receipt_no") & "','" & rsTmp("dock") & "','" & GetWordNew(rsTmp("customerorderkey"), 1, 35) & "','" & Trim(rsTmp("b_company")) & "','" & Trim(rsTmp("incoterm")) & "') "
                Else
                    '�D�Q��
                    str_SQL = "insert into " & strWMSDB & "..orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,route,DeliveryDate,stop,consigneekey,c_contact1,c_Company,C_Address1,C_Address2,C_Zip,C_Phone1,type,Notes,updatesource,door,customerorderkey) " & _
                            "values( '" & strKeycount & "','" & Trim(rsTmp("StorerKey")) & "','" & Trim(rsTmp("ExternOrderKey")) & "','" & Format(rsTmp("OrderDate"), "yyyy/mm/dd HH:mm:ss") & "','" & rsTmp("route_no") & "','" & Format(rsTmp("DeliveryDate"), "yyyy/mm/dd HH:mm:ss") & "','" & rsTmp("VEHICLE_ID_NO") & "','" & rsTmp("consigneekey") & "','" & rsTmp("c_contact1") & "','" & rsTmp("c_Company") & "','" & rsTmp("C_Address1") & "','" & rsTmp("C_Address2") & "','" & Trim(rsTmp("C_Zip")) & "','" & Trim(rsTmp("C_Phone1")) & "','" & Trim(rsTmp("priority")) & "','" & rsTmp("notes") & "','" & rsTmp("receipt_no") & "','" & rsTmp("dock") & "','" & GetWordNew(rsTmp("customerorderkey"), 1, 35) & "') "
                End If
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                      
                intLineNumber = 1
                strTMSorderkey = Trim(rsTmp("receipt_no"))
                str = "WMS: " & strKeycount & ",TMS: " & rsTmp("route_no") & "," & rsTmp("StorerKey") & "," & rsTmp("receipt_no") & "," & rsTmp("ExternOrderKey")
                
                '�g�J����
                Call WriteLog(err.Number & Chr(9) & "�ƨ���q��" & Chr(9) & str)
                If strWMSorderkeyS = "" Then strWMSorderkeyS = strKeycount
                strRoutekey = strRoutekey & "'" & rsTmp("route_no") & "',"
                
            End If
                '�g�J��
                If Trim(rsTmp("storerkey")) = "LLFA01" Then
                '�Q��
                str_SQL = "insert into " & strWMSDB & "..orderdetail (OrderKey,OrderLineNumber,ExternOrderKey,ExternLineno,SKU,StorerKey,OpenQty,UOM,Packkey,lottable03,lottable05,lottable06,updatesource,retailsku) "
                        If Not rsTmp("lottable05") Then
                            str_SQL = str_SQL & "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Trim(rsTmp("ExternOrderKey")) & "','" & Trim(rsTmp("OrderLineNumber")) & "','" & RTrim(rsTmp("product_no")) & "','" & rsTmp("StorerKey") & "'," & rsTmp("order_qty") & ",'" & RTrim(rsTmp("otheruom")) & "','" & rsTmp("packkey") & "','" & rsTmp("lottable03") & "','" & Format(rsTmp("lottable05"), "YYYY/MM/DD hh:mm:ss") & "','" & RTrim(rsTmp("lottable06")) & "','" & RTrim(rsTmp("updatesource")) & "','" & RTrim(rsTmp("retailsku")) & "') "
                        Else
                            str_SQL = str_SQL & "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Trim(rsTmp("ExternOrderKey")) & "','" & Trim(rsTmp("OrderLineNumber")) & "','" & RTrim(rsTmp("product_no")) & "','" & rsTmp("StorerKey") & "'," & rsTmp("order_qty") & ",'" & RTrim(rsTmp("otheruom")) & "','" & rsTmp("packkey") & "','" & rsTmp("lottable03") & "',null,'" & RTrim(rsTmp("lottable06")) & "','" & RTrim(rsTmp("updatesource")) & "','" & RTrim(rsTmp("retailsku")) & "') "
                        End If
                Else
                '�D�Q��
                str_SQL = "insert into " & strWMSDB & "..orderdetail (OrderKey,OrderLineNumber,ExternOrderKey,ExternLineno,SKU,StorerKey,OpenQty,UOM,Packkey,lottable03,lottable05,lottable06,updatesource) "
                        If Not rsTmp("lottable05") Then
                            str_SQL = str_SQL & "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Trim(rsTmp("ExternOrderKey")) & "','" & Trim(rsTmp("OrderLineNumber")) & "','" & RTrim(rsTmp("product_no")) & "','" & rsTmp("StorerKey") & "'," & rsTmp("order_qty") & ",'EA','" & rsTmp("packkey") & "','" & rsTmp("lottable03") & "','" & Format(rsTmp("lottable05"), "YYYY/MM/DD hh:mm:ss") & "','" & RTrim(rsTmp("lottable06")) & "','" & RTrim(rsTmp("updatesource")) & "') "
                        Else
                            str_SQL = str_SQL & "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & Trim(rsTmp("ExternOrderKey")) & "','" & Trim(rsTmp("OrderLineNumber")) & "','" & RTrim(rsTmp("product_no")) & "','" & rsTmp("StorerKey") & "'," & rsTmp("order_qty") & ",'EA','" & rsTmp("packkey") & "','" & rsTmp("lottable03") & "',null,'" & RTrim(rsTmp("lottable06")) & "','" & RTrim(rsTmp("updatesource")) & "') "
                        End If
                End If
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                intLineNumber = intLineNumber + 1
                
            rsTmp.MoveNext
        
        Loop
        rsTmp.Close
        
        '�ˬd��J�O�_���T
        str_SQL = "select TMSQty = isnull(sum(order_qty),0) , WMSQty= isnull((select sum(od.openqty) from " & strWMSDB & "..orderdetail od join " & strWMSDB & "..orders o on o.orderkey = od.orderkey where o.route in (" & Mid(strRoutekey, 1, Len(RTrim(strRoutekey)) - 1) & ")),0) " & _
                    "from trp01t t1 join trp02t t2 on t1.route_no = t2.route_no and t1.exe_confirm = '1' and t2.priority <> 'C' " & _
                    "join trp03t t3 on t2.receipt_no = t3.receipt_no " & _
                    "where t2.route_no in (" & Mid(strRoutekey, 1, Len(RTrim(strRoutekey)) - 1) & ") "
        
        rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If rsTmp("TMSQty") <> rsTmp("WMSQty") Or rsTmp("TMSQty") = 0 Then
        
            cn.RollbackTrans: Tran_Level = 0
            MsgBox "����XTMS�t�μƶq " & rsTmp("TMSQty") & "�A��JWMS�t�μƶq " & rsTmp("WMSQty") & "�A�ƶq���~!�C", vbOKOnly, App.EXEName & "�@�~����"
            rsTmp.Close: Screen.MousePointer = 0: cmd_Update.Enabled = True
            Exit Sub
            
        End If
        rsTmp.Close
        
        '���פJ���q������EXCEL
        str_SQL = "exec gs_excorder2wms '" & strWMSorderkeyS & "','" & strKeycount & "' "
        
        cn.CommandTimeout = 0
        rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
        '��Excel
        Call Recordset2Excel("Order2WMS", rsTmp)
        
        '�榡�]�w
        MyXlsApp.Cells.Select
        MyXlsApp.Cells.EntireRow.AutoFit
        MyXlsApp.Range("A1").Select
        
        If Dir("C:\BEST\Order2WMS", vbDirectory) = "" Then MkDirs "C:\BEST\Order2WMS"
        MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\Order2WMS\Order2WMS_" & Format(Now, "yyyymmddhhMMss") & ".xls"
        
        Set MyXlsApp = Nothing
        rsTmp.Close: Set rsTmp = Nothing
    
        '��s���w�^��
        cn.Execute "update trp01t set exe_confirm = '2' where route_no in (" & Mid(strRoutekey, 1, Len(RTrim(strRoutekey)) - 1) & ") ", RowsAffect, adExecuteNoRecords
    
    End If
    
    '�]�w��Ӹ��s���OC�w�^��
    str_SQL = "update trp01t set EXE_CONFIRM = '2' from trp01t t1 join trp02t t2 on t1.route_no = t2.route_no and t1.exe_confirm = '1' and t2.priority not in ('I','B','A') Where t2.Route_No in (" & strRouteNo & ") "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
End If
     
cn.CommitTrans: Tran_Level = 0
Screen.MousePointer = 0: cmd_Update.Enabled = True
Call cmd_RouteList_Click

Exit Sub
err_Handle:
   Screen.MousePointer = 0: cmd_Update.Enabled = True
   Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "�^�ǳ]�w�T�{")
End Sub


Private Sub cmd_UTL_Click()
'�^�ǳ]�w�T�{
Dim txtpath, FileTime As String          'Excel �ɮצW��
CmnDialog.DialogTitle = "�� �¤�r��"
CmnDialog.InitDir = "c:\my documents"
CmnDialog.FileName = "UTL_" & Format(Now, "YYYYMMDDHHNNSS")
CmnDialog.Filter = "�¤�r��(*.txt)|*.txt"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
CmnDialog.ShowOpen
If err.Number = cdlCancel Then          '�� [�}������] ��ܤ�����A���U [����] �s
   Exit Sub
Else
   txtpath = CmnDialog.FileName
   If Dir(txtpath) <> "" Then
      Kill txtpath
   End If
End If
On Error GoTo err_Handle
With dg_Route
     If .Rows = 2 Then Exit Sub
     
     Screen.MousePointer = vbHourglass
     cmd_Update.Enabled = False

     '��Ʈw���ʥ��--�_�I
     Tran_Level = 0
     Tran_Level = cn.BeginTrans

     Dim strRouteNo As String
     strRouteNo = ""
     For i = 1 To .Rows - 2
        .Row = i
        .Col = 1   '����ѧO
        If Trim(.Text) <> "" Then
           .Col = 2   '���u�s��
           If strRouteNo = "" Then
              strRouteNo = "'" & RTrim(.Text) & "'"
           Else
              strRouteNo = strRouteNo & ",'" & RTrim(.Text) & "'"
           End If
           '��s�@��h�����O & �p����έq�涵���s��
           str_SQL = "exec TRPPlan_BacktoEXE " & .Text & ""
           cn.CommandTimeout = 120
           cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
           .Col = 4
           .Text = "�]�w�^��"
        End If
     Next i
     If strRouteNo <> "" Then
        '�w�^�Ƕ}�o�������s,���b�^�ǵ�ids>>EXE_CONFIRM = '2'
        str_SQL = "Update TRP01T Set EXE_CONFIRM = '2',UploadWho='" & User_id & "' Where Route_No in (" & strRouteNo & ")"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '���r��
        str_SQL = "select (select top 1 route_no  from trp02t where EXTERN=t3.EXTERN),t3.EXTERN,convert(char(8),t2.RECEIPT_DATE,112),convert(char(8),t2.ARRIVE_DATE,112)," & _
            "t2.CONSIGNEEKEY,o.C_Company,'','','','','',o.C_Address1,o.C_Address2,o.C_Address3,o.C_Zip,t3.PRODUCT_NO,s.DESCR," & _
            "sum(t3.ORDER_QTY),'0000000','0000000',od.Lottable03,'PC','','PK',sum(t3.SHIP_QTY),'0000000','0000000',od.Lottable03," & _
            "Rtrim(Cast(o.Notes as varchar(300))),od.Lottable02,t3.SEQ_NO,t3.EXTERN,'N','N','',t3.route_no,t3.VEHICLE_ID_NO," & _
            "case when len(rtrim(od.Lottable02))=0 then 'P' else 'S' End as Lottable02," & _
            "'' as PO,o.Priority as OrderType,'0000000',t3.SEQ_NO as orderline,'' as UPO,'" & User_id & "' as Users,convert(char(8),getdate(),112),getdate() " & _
            "from trp03t t3 " & _
            "inner join trp02t t2 on t3.route_no=t2.route_no and t2.EXTERN=t3.EXTERN " & _
            "inner join orders o on o.ExternOrderKey=t3.EXTERN " & _
            "inner join ORDERDETAIL od on od.ExternOrderKey=t3.EXTERN and t3.SEQ_NO=od.OrderLineNumber " & _
            "inner join sku s on s.sku=t3.PRODUCT_NO and s.storerkey=t3.storerkey " & _
            "where t3.route_no in (" & strRouteNo & ") " & _
            "GROUP BY t3.route_no,t3.EXTERN,t2.RECEIPT_DATE,t2.ARRIVE_DATE,t2.CONSIGNEEKEY,o.C_Company, " & _
            "o.C_Address1,o.C_Address2,o.C_Address3,o.C_Zip,o.Priority,t3.PRODUCT_NO,s.DESCR, " & _
            "t3.route_no,Rtrim(Cast(o.Notes as varchar(300))),od.Lottable02,od.Lottable03,t3.VEHICLE_ID_NO,t3.SEQ_NO order by t3.EXTERN,t3.SEQ_NO"
        Set rs_Src = New Recordset
        rs_Src.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        rs_Src.MoveFirst
        FileTime = Format(Now, "HHNNSS")
        Open txtpath For Append As #1
        Do While Not rs_Src.EOF
            If Not IsNull(rs_Src.Fields(0)) Then Print #1, StrPadRightC(rs_Src.Fields(0), 12); Else Print #1, StrPadRightC(" ", 12);
            If Not IsNull(rs_Src.Fields(1)) Then Print #1, StrPadRightC(rs_Src.Fields(1), 12); Else Print #1, StrPadRightC(" ", 12);
            If Not IsNull(rs_Src.Fields(2)) Then Print #1, StrPadRightC(rs_Src.Fields(2), 8); Else Print #1, StrPadRightC(" ", 8);
            If Not IsNull(rs_Src.Fields(3)) Then Print #1, StrPadRightC(rs_Src.Fields(3), 8); Else Print #1, StrPadRightC(" ", 8);
            If Not IsNull(rs_Src.Fields(4)) Then Print #1, StrPadRightC(rs_Src.Fields(4), 10); Else Print #1, StrPadRightC(" ", 10);
            If Not IsNull(rs_Src.Fields(5)) Then Print #1, StrPadRightC(rs_Src.Fields(5), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(6)) Then Print #1, StrPadRightC(rs_Src.Fields(6), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(7)) Then Print #1, StrPadRightC(rs_Src.Fields(7), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(8)) Then Print #1, StrPadRightC(rs_Src.Fields(8), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(9)) Then Print #1, StrPadRightC(rs_Src.Fields(9), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(10)) Then Print #1, StrPadRightC(rs_Src.Fields(10), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(11)) Then Print #1, StrPadRightC(rs_Src.Fields(11), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(12)) Then Print #1, StrPadRightC(rs_Src.Fields(12), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(13)) Then Print #1, StrPadRightC(rs_Src.Fields(13), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(14)) Then Print #1, StrPadRightC(rs_Src.Fields(14), 3); Else Print #1, StrPadRightC(" ", 3);
            If Not IsNull(rs_Src.Fields(15)) Then Print #1, StrPadRightC(rs_Src.Fields(15), 14); Else Print #1, StrPadRightC(" ", 14);
            If Not IsNull(rs_Src.Fields(16)) Then Print #1, StrPadRightC(rs_Src.Fields(16), 30); Else Print #1, StrPadRightC(" ", 30);
            If Not IsNull(rs_Src.Fields(17)) Then Print #1, StrPadLeft(rs_Src.Fields(17), 7, 0); Else Print #1, StrPadLeft(" ", 7, 0);
            If Not IsNull(rs_Src.Fields(18)) Then Print #1, StrPadRightC(rs_Src.Fields(18), 7); Else Print #1, StrPadRightC(" ", 7);
            If Not IsNull(rs_Src.Fields(19)) Then Print #1, StrPadRightC(rs_Src.Fields(19), 7); Else Print #1, StrPadRightC(" ", 7);
            If Not IsNull(rs_Src.Fields(20)) Then Print #1, StrPadRightC(rs_Src.Fields(20), 7); Else Print #1, StrPadRightC(" ", 7);
            If Not IsNull(rs_Src.Fields(21)) Then Print #1, StrPadRightC(rs_Src.Fields(21), 4); Else Print #1, StrPadRightC(" ", 4);
            If Not IsNull(rs_Src.Fields(22)) Then Print #1, StrPadRightC(rs_Src.Fields(22), 4); Else Print #1, StrPadRightC(" ", 4);
            If Not IsNull(rs_Src.Fields(23)) Then Print #1, StrPadRightC(rs_Src.Fields(23), 4); Else Print #1, StrPadRightC(" ", 4);
            If Not IsNull(rs_Src.Fields(24)) Then Print #1, StrPadLeft(rs_Src.Fields(24), 7, 0); Else Print #1, StrPadLeft(" ", 7, 0);
            If Not IsNull(rs_Src.Fields(25)) Then Print #1, StrPadRightC(rs_Src.Fields(25), 7); Else Print #1, StrPadRightC(" ", 7);
            If Not IsNull(rs_Src.Fields(26)) Then Print #1, StrPadRightC(rs_Src.Fields(26), 7); Else Print #1, StrPadRightC(" ", 7);
            If Not IsNull(rs_Src.Fields(27)) Then Print #1, StrPadRightC(rs_Src.Fields(27), 7); Else Print #1, StrPadRightC(" ", 7);
            If Not IsNull(rs_Src.Fields(28)) Then Print #1, StrPadRightC(rs_Src.Fields(28), 60); Else Print #1, StrPadRightC(" ", 60);
            If Not IsNull(rs_Src.Fields(29)) Then Print #1, StrPadRightC(rs_Src.Fields(29), 8); Else Print #1, StrPadRightC(" ", 8);
            If Not IsNull(rs_Src.Fields(30)) Then Print #1, StrPadLeft(rs_Src.Fields(30), 2, 0); Else Print #1, StrPadLeft(" ", 2, 0);
            If Not IsNull(rs_Src.Fields(31)) Then Print #1, StrPadRightC(rs_Src.Fields(31), 12); Else Print #1, StrPadRightC(" ", 12);
            If Not IsNull(rs_Src.Fields(32)) Then Print #1, StrPadRightC(rs_Src.Fields(32), 1); Else Print #1, StrPadRightC(" ", 1);
            If Not IsNull(rs_Src.Fields(33)) Then Print #1, StrPadRightC(rs_Src.Fields(33), 1); Else Print #1, StrPadRightC(" ", 1);
            If Not IsNull(rs_Src.Fields(34)) Then Print #1, StrPadRightC(rs_Src.Fields(34), 11); Else Print #1, StrPadRightC(" ", 1);
            If Not IsNull(rs_Src.Fields(35)) Then Print #1, StrPadRightC(rs_Src.Fields(35), 11); Else Print #1, StrPadRightC(" ", 11);
            If Not IsNull(rs_Src.Fields(36)) Then Print #1, StrPadRightC(rs_Src.Fields(36), 10); Else Print #1, StrPadRightC(" ", 10);
            If Not IsNull(rs_Src.Fields(37)) Then Print #1, StrPadRightC(rs_Src.Fields(37), 1); Else Print #1, StrPadRightC(" ", 1);
            If Not IsNull(rs_Src.Fields(38)) Then Print #1, StrPadRightC(rs_Src.Fields(38), 12); Else Print #1, StrPadRightC(" ", 12);
            If Not IsNull(rs_Src.Fields(39)) Then Print #1, StrPadRightC(rs_Src.Fields(39), 1); Else Print #1, StrPadRightC(" ", 1);
            If Not IsNull(rs_Src.Fields(40)) Then Print #1, StrPadRightC(rs_Src.Fields(40), 7); Else Print #1, StrPadRightC(" ", 7);
            If Not IsNull(rs_Src.Fields(41)) Then Print #1, StrPadLeft(rs_Src.Fields(41), 4, 0); Else Print #1, StrPadLeft(" ", 4, 0);
            If Not IsNull(rs_Src.Fields(42)) Then Print #1, StrPadRightC(rs_Src.Fields(42), 16); Else Print #1, StrPadRightC(" ", 16);
            If Not IsNull(rs_Src.Fields(43)) Then Print #1, StrPadRightC(rs_Src.Fields(43), 10); Else Print #1, StrPadRightC(" ", 10);
            If Not IsNull(rs_Src.Fields(44)) Then Print #1, StrPadRightC(rs_Src.Fields(44), 8); Else Print #1, StrPadRightC(" ", 8);
            Print #1, FileTime
            rs_Src.MoveNext
        Loop
        Close #1
     End If
     cn.CommitTrans
     Tran_Level = 0
     cmd_Update.Enabled = True
     Screen.MousePointer = vbDefault
     Call cmd_RouteList_Click
End With
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�^�ǳ]�w�T�{", Me.Caption, "cmd_Update_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_Update.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub dg_Route_Click()
'�Ҧ����]�w�^�Ǫ����u�s���C��
'�I�@���G����A�I�ĤG���G�������
Screen.MousePointer = 11
With dg_Route
     .Col = 2   '���u�s��
     If Len(Trim(.Text)) = 0 Then Exit Sub
     
     .Col = 5 '�ƨ���
     If UCase(.Text) <> UCase(User_id) And blRouteModifyControl = True Then Screen.MousePointer = 0: Exit Sub
     
     .Col = 1
     If Len(.Text) = 0 Then
        .Text = "V"
        .Col = 2
        Call Display_RouteOrders(.Text)
     Else
        .Text = ""
     End If
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
Screen.MousePointer = 0
End Sub




Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�q��ƨ���Ʀ^�ǳ]�w"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'�^�Ǳ��]�w�����u�s�����
Call SetGridFormat_Route

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�d�I��Ӫ����L����ƥ�
'�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������������
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
   fra_ExtraQuery.Visible = False
End If
End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   'SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   'SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   fam_Header.Left = fam_Header.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fra_ExtraQuery.Left = fra_ExtraQuery.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   dg_Route.Height = dg_Route.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Route.Width = dg_Route.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_Orders.Top = dg_Orders.Top - (dbsrcFormHeight - Me.ScaleHeight)
   dg_Orders.Width = dg_Orders.Width - (dbsrcFormWidth - Me.ScaleWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   'SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   'SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   fam_Header.Left = fam_Header.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fra_ExtraQuery.Left = fra_ExtraQuery.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   dg_Route.Height = dg_Route.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Route.Width = dg_Route.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_Orders.Top = dg_Orders.Top + (Me.ScaleHeight - dbsrcFormHeight)
   dg_Orders.Width = dg_Orders.Width + (Me.ScaleWidth - dbsrcFormWidth)
   
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
Set frm_OP_BacktoEXE = Nothing
End Sub

Private Sub SetGridFormat_Route()
'�^�Ǳ��]�w�����u�s�����
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Route.Visible = False
With dg_Route
     .Rows = 2: .Cols = 20
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
     .ColWidth(0) = 300
     .ColWidth(1) = 400
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1000
     .ColWidth(5) = 700
     .ColWidth(6) = 700
     .ColWidth(7) = 700
     .ColWidth(8) = 700
     .ColWidth(9) = 700
     .ColWidth(10) = 700
     .ColWidth(11) = 900
     .ColWidth(12) = 500
     .ColWidth(13) = 1000
     .ColWidth(14) = 1300
     .ColWidth(15) = 1300
     .ColWidth(16) = 800
     .ColWidth(17) = 1300
     .ColWidth(18) = 1300
     .ColWidth(19) = 1300
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "No"
     .Col = 1: .Text = "��"
     .Col = 2: .Text = "���u�s��"
     .Col = 3: .Text = "�X�����"
     .Col = 4: .Text = "�^�Ǫ��A"
     .Col = 5: .Text = "�ƨ��H"
     .Col = 6: .Text = "�^�ǤH"
     .Col = 7: .Text = "�c��"
     .Col = 8: .Text = "�O��"
     .Col = 9: .Text = "���n"
     .Col = 10: .Text = "���q"
     .Col = 11: .Text = "���P���X"
     .Col = 12: .Text = "����"
     .Col = 13: .Text = "�r�p�H"
     .Col = 14: .Text = "�w�p������"
     .Col = 15: .Text = "�w�p����ɶ�"
     .Col = 16: .Text = "�X�Y�Ȧs"
     .Col = 17: .Text = "�G���ƨ����s"
     .Col = 18: .Text = "�G���ƨ�����"
     .Col = 19: .Text = "�G���ƨ�����"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignCenterCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignLeftCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignRightCenter
     .ColAlignment(10) = flexAlignRightCenter
     .ColAlignment(11) = flexAlignCenterCenter
     .ColAlignment(12) = flexAlignCenterCenter
     .ColAlignment(13) = flexAlignLeftCenter
     .ColAlignment(14) = flexAlignLeftCenter
     .ColAlignment(15) = flexAlignLeftCenter
     .ColAlignment(16) = flexAlignLeftCenter
     .ColAlignment(17) = flexAlignCenterCenter
     .ColAlignment(18) = flexAlignCenterCenter
     .ColAlignment(19) = flexAlignCenterCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Route.Visible = True
End Sub

Private Sub Display_RouteOrders(strRouteNo As String)
'��ܶǤJ���s���q����
Screen.MousePointer = 11
str_SQL = "Select ���u�s��,�q��s��,�e�f��,�Ȥ�s��,�f�D�渹,�c��,�O��,���n,���q,ZIP,�ϽX,�Ȥ�W��,�q���,�f�D,�Ȥ�Ƶ� " & _
          "From BacktoEXE_RouteOrders " & _
          "Where ���u�s�� = '" & strRouteNo & "' order by �f�D�渹 "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '�L��������
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   Screen.MousePointer = 0
   msg_text = "�d�ߵ��G�G�L���w���u�s�����q����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
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
Set dg_Orders.DataSource = rs_Orders
With dg_Orders
    .RowHeight = 250
    .Columns(0).Width = 500        '�Ǹ�
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000       '���u�s��
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1100       '�q��s��
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1000       '�e�f��
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1200       '�Ȥ�s��
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 900        '�f�D�渹
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 700        '�c��
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 700        '�O��
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 700        '���n
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 700        '���q
    .Columns(9).Alignment = dbgRight
    .Columns(10).Width = 500       'zip
    .Columns(10).Alignment = dbgCenter
    .Columns(11).Width = 500       '�ϽX
    .Columns(11).Alignment = dbgCenter
    .Columns(12).Width = 1600      '�Ȥ�W��
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 1000      '�q���
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 500       '�f�D
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1400      '�Ȥ�Ƶ� Orders.Notes
    .Columns(15).Alignment = dbgLeft
End With
Screen.MousePointer = 0
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���s�C��-�q��d��", Me.Caption, "Form ���� SubProgram�GDisplay_RouteOrders", tmpString
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
mvDate.Left = fam_Header.Left + txt_DeliveryDate.Left
mvDate.Top = fam_Header.Top + txt_DeliveryDate.Top + txt_DeliveryDate.Height
mvDate.Tag = "�X�����"
mvDate.Visible = True
End Sub

Private Sub txt_DeliveryDate_KeyPress(KeyAscii As Integer)
'�X���������ˮ�
Select Case KeyAscii
     Case 97 To 122, 65 To 90   '�����\��J�r��
          KeyAscii = 0
     Case vbKeyReturn
          If Fun_ChkDateFormat(txt_DeliveryDate.Text) = 1 Then
             msg_text = "�X���������ˮֿ��~�G" & vbCrLf & funRtn_msg
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_DeliveryDate.SelStart = 0: txt_DeliveryDate.SelLength = Len(txt_DeliveryDate.Text): txt_DeliveryDate.SetFocus
             Exit Sub
          Else
             cmd_RouteList.SetFocus
          End If
End Select
End Sub
