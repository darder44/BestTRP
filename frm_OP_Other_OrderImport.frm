VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_Other_OrderImport 
   Caption         =   "�䥦�q����J�ΫȤᲧ�ʺ��@"
   ClientHeight    =   7140
   ClientLeft      =   270
   ClientTop       =   990
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11280
   WindowState     =   2  '�̤j��
   Begin VB.Frame fam_Command 
      Height          =   720
      Left            =   4260
      TabIndex        =   53
      Top             =   -75
      Width           =   7155
      Begin VB.CommandButton cmd_OrderImport 
         BackColor       =   &H00FF8080&
         Caption         =   "�q��ΫȤ�����J"
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
         Left            =   60
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   31
         Top             =   120
         Width           =   2250
      End
      Begin VB.CommandButton cmd_Update 
         BackColor       =   &H8000000B&
         Caption         =   "�T�{�s��"
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
         Left            =   2790
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   32
         Top             =   135
         Width           =   1860
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
         Height          =   495
         Index           =   0
         Left            =   5220
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   33
         Top             =   150
         Width           =   1860
      End
   End
   Begin VB.Frame fam_ConsignHead 
      BackColor       =   &H8000000B&
      Height          =   3390
      Left            =   4260
      TabIndex        =   34
      Top             =   555
      Width           =   7155
      Begin VB.TextBox txt_Storer_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   1020
         TabIndex        =   5
         Top             =   180
         Width           =   705
      End
      Begin VB.TextBox txt_Storer 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   1020
         TabIndex        =   6
         Top             =   450
         Width           =   705
      End
      Begin VB.TextBox txt_ConsigneeKey_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   2715
         TabIndex        =   7
         Top             =   180
         Width           =   1380
      End
      Begin VB.TextBox txt_ConsigneeKey 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   2715
         TabIndex        =   8
         Top             =   450
         Width           =   1380
      End
      Begin VB.TextBox txt_AreaCode_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   5130
         TabIndex        =   13
         Top             =   780
         Width           =   705
      End
      Begin VB.TextBox txt_AreaCode 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   5130
         TabIndex        =   14
         Top             =   1065
         Width           =   705
      End
      Begin VB.ComboBox cmb_Zip_New 
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
         TabIndex        =   11
         Top             =   780
         Width           =   1995
      End
      Begin VB.ComboBox cmb_Zip 
         BackColor       =   &H8000000A&
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
         TabIndex        =   12
         Top             =   1125
         Width           =   1995
      End
      Begin VB.TextBox txt_FullName_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   1020
         TabIndex        =   15
         Top             =   1515
         Width           =   6000
      End
      Begin VB.TextBox txt_FullName 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   1020
         TabIndex        =   16
         Top             =   1785
         Width           =   6000
      End
      Begin VB.TextBox txt_Address_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   1020
         TabIndex        =   17
         Top             =   2130
         Width           =   6000
      End
      Begin VB.TextBox txt_Address 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   1020
         TabIndex        =   18
         Top             =   2400
         Width           =   6000
      End
      Begin VB.TextBox txt_Class 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   5130
         TabIndex        =   10
         Top             =   450
         Width           =   705
      End
      Begin VB.TextBox txt_Class_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   5130
         TabIndex        =   9
         Top             =   180
         Width           =   705
      End
      Begin VB.TextBox txt_Contact 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   1020
         TabIndex        =   20
         Top             =   3015
         Width           =   1575
      End
      Begin VB.TextBox txt_Contact_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   1020
         TabIndex        =   19
         Top             =   2745
         Width           =   1575
      End
      Begin VB.TextBox txt_Phone 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   3270
         TabIndex        =   22
         Top             =   3015
         Width           =   1575
      End
      Begin VB.TextBox txt_Phone_New 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   3270
         TabIndex        =   21
         Top             =   2745
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�f        �D"
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   43
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�s��"
         Height          =   180
         Index           =   1
         Left            =   1950
         TabIndex        =   42
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�l���ϸ�"
         Height          =   180
         Index           =   2
         Left            =   255
         TabIndex        =   41
         Top             =   870
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�B�e�ϽX"
         Height          =   180
         Index           =   3
         Left            =   4335
         TabIndex        =   40
         Top             =   855
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�W��"
         Height          =   180
         Index           =   4
         Left            =   255
         TabIndex        =   39
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�B�e�a�}"
         Height          =   180
         Index           =   5
         Left            =   255
         TabIndex        =   38
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ӽh"
         Height          =   180
         Index           =   7
         Left            =   4335
         TabIndex        =   37
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�p���H"
         Height          =   180
         Index           =   8
         Left            =   435
         TabIndex        =   36
         Top             =   2790
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q��"
         Height          =   180
         Index           =   9
         Left            =   2850
         TabIndex        =   35
         Top             =   2790
         Width           =   360
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_ORT01W 
      Height          =   6945
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   12250
      _Version        =   393216
      Cols            =   5
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Frame fam_ConsignDetail 
      BackColor       =   &H8000000B&
      Height          =   3165
      Left            =   4260
      TabIndex        =   44
      Top             =   3870
      Width           =   7155
      Begin VB.ComboBox cmb_PickTool 
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
         Left            =   4755
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   30
         Top             =   2085
         Width           =   1995
      End
      Begin VB.TextBox txt_ShortName 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   1290
         TabIndex        =   23
         Top             =   210
         Width           =   1545
      End
      Begin VB.ComboBox cmb_VehicleType 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   1290
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   25
         Top             =   570
         Width           =   5445
      End
      Begin VB.ComboBox cmb_ExtraDemand1 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   1290
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   26
         Top             =   915
         Width           =   5445
      End
      Begin VB.ComboBox cmb_ExtraDemand2 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   1290
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   27
         Top             =   1275
         Width           =   5445
      End
      Begin VB.TextBox txt_ChannelType 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   1290
         TabIndex        =   28
         Top             =   1680
         Width           =   1200
      End
      Begin VB.TextBox txt_UnLoad 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   6315
         TabIndex        =   29
         Top             =   1680
         Width           =   420
      End
      Begin VB.CheckBox chk_MultiCustomer 
         BackColor       =   &H8000000C&
         Caption         =   "���e�Ȥ�"
         Height          =   180
         Left            =   465
         TabIndex        =   45
         Top             =   2070
         Width           =   1260
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  '����
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2445
         Width           =   705
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  '����
         BackColor       =   &H8000000F&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   480
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   2730
         Width           =   705
      End
      Begin VB.TextBox txt_GridCode 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   3780
         TabIndex        =   24
         Top             =   210
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�h�B�u��"
         Height          =   180
         Index           =   19
         Left            =   3975
         TabIndex        =   54
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�²��"
         Height          =   180
         Index           =   6
         Left            =   495
         TabIndex        =   52
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���إN�X"
         Height          =   180
         Index           =   10
         Left            =   495
         TabIndex        =   51
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�S��ݨD 1"
         Height          =   180
         Index           =   11
         Left            =   360
         TabIndex        =   50
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�S��ݨD 2"
         Height          =   180
         Index           =   12
         Left            =   360
         TabIndex        =   49
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q�����A"
         Height          =   180
         Index           =   13
         Left            =   495
         TabIndex        =   48
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���f������"
         Height          =   180
         Index           =   15
         Left            =   5370
         TabIndex        =   47
         Top             =   1755
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "  �ݨϥΪ̽T�{���Ȥ���  "
         Height          =   180
         Index           =   16
         Left            =   1245
         TabIndex        =   3
         Top             =   2505
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "  �w���ɤ����Ȥ���         "
         Height          =   180
         Index           =   17
         Left            =   1245
         TabIndex        =   2
         Top             =   2790
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�x�}�ϽX"
         Height          =   180
         Index           =   18
         Left            =   3000
         TabIndex        =   46
         Top             =   270
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm_OP_Other_OrderImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e

Private intloop As Double
Private ZipQueryAreaCode As Boolean
Private intGridRow As Double

Private arZip() As String
Private arVehicleType() As String
Private arExtraDemand() As String
Private arPickTool() As String        '�h�B�u��

Private rs_ORT01W As ADODB.Recordset

Private Sub cmb_Zip_New_Change()
'���^ �l���ϸ� ���ݤ� �B�e�ϰ�N�X
If ZipQueryAreaCode = False Then Exit Sub
If cmb_Zip_New.ListIndex = -1 Then Exit Sub

'���o����� Company ���Ҧ� Branch
str_SQL = "SELECT RTRIM(Area_Code) AS AreaCode From TRP02M Where ZIP = '" & arZip(cmb_Zip_New.ListIndex) & "'"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   txt_AreaCode_New.Text = tmp_Rs.Fields("AreaCode").Value
End If
tmp_Rs.Close

End Sub

Private Sub cmb_Zip_New_Click()
'���^ �l���ϸ� ���ݤ� �B�e�ϰ�N�X
'If ZipQueryAreaCode = False Then Exit Sub
If cmb_Zip_New.ListIndex = -1 Then Exit Sub

'���o����� Company ���Ҧ� Branch
str_SQL = "SELECT RTRIM(Area_Code) AS AreaCode From TRP02M Where ZIP = '" & arZip(cmb_Zip_New.ListIndex) & "'"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   txt_AreaCode_New.Text = tmp_Rs.Fields("AreaCode").Value
End If
tmp_Rs.Close

End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'���}
Unload Me
End Sub

Private Sub cmd_OrderImport_Click()
'�q��ΫȤ�����J

On Error GoTo err_Handle

    Tran_Level = cn.BeginTrans
'�h�f�����JWMS�A�n�ư��Q�ת��A�Q�װh�f�g�JASN Table
str_SQL = "select od.storerkey " & _
            ",o.orderkey " & _
            ", o.externorderkey " & _
            ", o.priority " & _
            ", orderlinenumber = (select top 1 orderdetail.orderlinenumber from orderdetail where orderdetail.sku = od.sku and orderdetail.orderkey = o.orderkey order by orderdetail.sku) " & _
            ", od.sku " & _
            ", s.descr " & _
            ", s.packkey " & _
            ", openqty = sum(od.openqty) " & _
            ", notes = cast(o.notes as varchar(300)) " & _
            ", o.consigneekey " & _
            ", o.c_company " & _
            "from orders o join orderdetail od on o.orderkey = od.orderkey " & _
            "join gv_skuxpack s on s.sku = od.sku and s.storerkey = od.storerkey " & _
            "where o.B_PHONE2 is null and o.priority in ('R','RC','A2B') and o.storerkey not in ('LLFA01','LMBO01','LPSI01','LCHF01','LKYF01', 'LNCE01') and o.type <> '�R��' " & _
            "group by od.storerkey ,o.orderkey , o.externorderkey , o.priority ,od.storerkey ,o.orderkey , od.externorderkey , o.priority , o.consigneekey , o.c_company ,od.sku, s.descr,cast(o.notes as varchar(300)),s.packkey " & _
            "order by od.storerkey , o.orderkey "

Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = 3
Dim rsKeycount As New ADODB.Recordset
rsTmp.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
    
'�S���h�f����
If rsTmp.EOF Then GoTo LFMBO

    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    Dim strKeycount As String, strOrderkey As String, intLineNumber As Integer
    
    'Ū��ini�ѼơA��L�ƨ��O�_�JWMS�t��
    Dim objIni As New vbIniFile, strOtherOrder2WMS As String
    objIni.FileName = App.Path & "/" & App.title & ".ini"
    
    strOtherOrder2WMS = objIni.ReadData("OPTION", "OtherOrder2WMS", "YES")
    Set objIni = Nothing
    
    If UCase(strOtherOrder2WMS) = "YES" Then 'WMS�s�W���ʳ�
    
        rsTmp.MoveFirst
        rsTmp.Filter = "Priority = 'R' Or Priority = 'RC'"
        Do While Not rsTmp.EOF
        '�Ȱ����q��ư����g�JExceed
        If RTrim(rsTmp("StorerKey")) = "LABT01" Then GoTo NextRow
        
        '�g�JWMS
        If Trim(rsTmp("orderkey")) <> strOrderkey Then
    
            '���t�έq��渹
            rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
            '�渹+1
            cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
            strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
            rsKeycount.Close

            '�g�J���Y
            str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,BuyersReference , sellername,selleraddress1,externpokey,potype,notes) " & _
                      "values( '" & strKeycount & "','" & rsTmp("StorerKey") & "','" & RTrim(GetWord(Trim(rsTmp("ExternOrderKey")), 1, 18)) & "','" & rsTmp("consigneekey") & "','" & rsTmp("C_company") & "','" & rsTmp("OrderKey") & "','" & rsTmp("priority") & "','" & rsTmp("notes") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
            intLineNumber = 1
            strOrderkey = Trim(rsTmp("orderkey"))
    
        End If
    
            '�g�J��
            str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,Skudescription,StorerKey,QtyOrdered,packkey) " & _
                    "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & rsTmp("OrderLineNumber") & "','" & rsTmp("SKU") & "','" & rsTmp("descr") & "','" & rsTmp("StorerKey") & "','" & rsTmp("openqty") & "','" & rsTmp("packkey") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
            intLineNumber = intLineNumber + 1
NextRow:
        rsTmp.MoveNext

        Loop
        rsTmp.Close: Set rsTmp = Nothing
        
LFMBO:
    '�g�J�Q�ת�
    Call Confirm_Recordset_Closed(rsTmp)
        str_SQL = "select " & _
                "od.storerkey , " & _
                "o.orderkey , " & _
                "o.externorderkey , " & _
                "o.b_company, " & _
                "o.customerorderkey, " & _
                "orderdate = convert(char(8),o.orderdate,112), " & _
                "deliverydate = convert(char(8),o.deliverydate,112), " & _
                "priority = rtrim(o.priority) ,ordertype = rtrim(isnull(o.externordertype,'')), " & _
                "od.orderlinenumber, " & _
                "od.externlineno, " & _
                "od.otheruom, " & _
                "od.lottable05, " & _
                "od.retailsku, " & _
                "od.sku , " & _
                "s.descr , " & _
                "s.packkey , " & _
                "openqty = sum(od.openqty) , " & _
                "notes = cast(o.notes as varchar(300)) , " & _
                "o.consigneekey , " & _
                "o.c_company, od.lottable06, od.lottable03 " & _
                "from orders o join orderdetail od on o.orderkey = od.orderkey join gv_skuxpack s on s.sku = od.sku and s.storerkey = od.storerkey " & _
                "where  o.priority in ('R','RC','A2B') and o.storerkey in ('LLFA01','LMBO01','LPSI01','LCHF01','LKYF01', 'LNCE01') and o.B_PHONE2 is null and o.type <> '�R��' " & _
                "group by od.storerkey ,o.orderkey , o.externorderkey , o.priority ,od.storerkey ,o.orderkey , od.externorderkey , o.priority , rtrim(isnull(o.externordertype,'')),o.consigneekey , o.c_company ,od.sku, s.descr,cast(o.notes as varchar(300)),s.packkey ,orderlinenumber,o.b_company,o.buyerpo,o.CustomerOrderkey,convert(char(8),o.orderdate,112),convert(char(8),o.deliverydate,112),od.externlineno,od.otheruom,od.lottable05,od.retailsku,od.lottable06,od.lottable03 " & _
                "order by o.orderkey ,orderlinenumber "

            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
            
    If rsTmp.EOF Then GoTo final  ': MsgBox "�d�ߵ��G�G�S���ݺ��@���Ȥ��ƶǦ^�A���~��i�� [�ƨ��@�~]", vbOKOnly, Me.Caption:
        
    rsTmp.MoveFirst
    rsTmp.Filter = "Priority = 'R' Or Priority = 'RC'"
    strOrderkey = ""
    '�g�JASN�����
    Do While Not rsTmp.EOF
            '�g�JWMS
            If Trim(rsTmp("orderkey")) <> strOrderkey Then
                '�ˬd�O�_�f�D�渹���ơA���ƫh���g�J
                str_SQL = "select externasnkey from " & strWMSDB & "..asn where asntype = 'R' and storerkey = '" & Trim(rsTmp("storerkey")) & "' and externasnkey = '" & Trim(rsTmp("externorderkey")) & "'"
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.CursorLocation = 3
                tmp_Rs.Open str_SQL, cn
                    If tmp_Rs.EOF Then
                        tmp_Rs.Close
                        '���t�έq��渹
                        rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='ASN' ", cn
                        '�渹+1
                        cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'ASN'", RowsAffect, adExecuteNoRecords
                        strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                        rsKeycount.Close
                        
                        '�g�J���Y
                        If Trim(rsTmp("StorerKey")) = "LLFA01" Then
                            str_SQL = "insert into " & strWMSDB & "..asn (asnKey,StorerKey,externasnkey , sellersreference,asntype,notes,SellersReference2,OtherReference,ASNDate,VesselDate) " & _
                                      "values( '" & strKeycount & "','" & rsTmp("StorerKey") & "','" & RTrim(GetWord(Trim(rsTmp("ExternOrderKey")), 1, 18)) & "','" & rsTmp("consigneekey") & "','" & rsTmp("priority") & "','" & rsTmp("notes") & "','" & _
                                      rsTmp("b_company") & "','" & rsTmp("customerorderkey") & "','" & rsTmp("orderdate") & "','" & rsTmp("deliverydate") & "') "
                                                              
                        Else
                            '���_
                            str_SQL = "insert into " & strWMSDB & "..asn (asnKey,StorerKey,externasnkey , sellersreference,asntype,notes,SellersReference2,OtherReference,ASNDate,VesselDate) " & _
                                      "values( '" & strKeycount & "','" & rsTmp("StorerKey") & "','" & RTrim(GetWord(Trim(rsTmp("ExternOrderKey")), 1, 18)) & "','" & rsTmp("consigneekey") & "','" & rsTmp("priority") & "','" & rsTmp("notes") & "','" & _
                                      rsTmp("b_company") & "','" & rsTmp("customerorderkey") & "','" & rsTmp("orderdate") & "','" & rsTmp("deliverydate") & "') "
                        End If
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                        intLineNumber = 1
                        strOrderkey = Trim(rsTmp("orderkey"))
                    Else
                        tmp_Rs.Close
                        GoTo NextRow1
                    End If
                End If
                '�g�J��
                If Trim(rsTmp("StorerKey")) = "LCHF01" Then
                    str_SQL = "insert into " & strWMSDB & "..asndetail (asnKey,asnLineNumber,ExternLineNo,externasnkey,SKU,Skudescription,StorerKey,QtyOrdered,packkey,OtherUOM,RetailSku,UOM,Lottable06,Effectivedate,Lottable03) " & _
                            "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & rsTmp("ExternLineNo") & "','" & RTrim(GetWord(Trim(rsTmp("ExternOrderKey")), 1, 18)) & "','" & rsTmp("SKU") & "','" & rsTmp("descr") & "','" & rsTmp("StorerKey") & "','" & rsTmp("openqty") & "','" & rsTmp("packkey") & "','" & rsTmp("otheruom") & "','" & rsTmp("retailsku") & "','EA','R01','" & rsTmp("Lottable03") & "','" & rsTmp("Lottable03") & "') "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                Else
                    str_SQL = "insert into " & strWMSDB & "..asndetail (asnKey,asnLineNumber,ExternLineNo,externasnkey,SKU,Skudescription,StorerKey,QtyOrdered,packkey,OtherUOM,RetailSku,UOM,Lottable06,Effectivedate,Lottable03) " & _
                            "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & rsTmp("ExternLineNo") & "','" & RTrim(GetWord(Trim(rsTmp("ExternOrderKey")), 1, 18)) & "','" & rsTmp("SKU") & "','" & rsTmp("descr") & "','" & rsTmp("StorerKey") & "','" & rsTmp("openqty") & "','" & rsTmp("packkey") & "','" & rsTmp("otheruom") & "','" & rsTmp("retailsku") & "','EA','" & rsTmp("Lottable06") & "','" & rsTmp("Lottable05") & "','" & rsTmp("Lottable03") & "') "
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                End If
        
                intLineNumber = intLineNumber + 1
NextRow1:

            rsTmp.MoveNext
    
            Loop
            'Close RecordSet
            rsTmp.Close: Set rsTmp = Nothing
            
        End If

final:

' call stored procedure ORTD11_IMPORTN
Screen.MousePointer = vbHourglass
cmd_OrderImport.Enabled = False
cmd_Update.Enabled = False
Call SetGrid_Format_ORT01W

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
tmp_Cmd.CommandText = "ORTD11_IMPORTN"
Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'��� [���椤] �T��
Load frm_WaitWindows
frm_WaitWindows.Tag = Me.Name
frm_WaitWindows.ZOrder
frm_WaitWindows.Refresh
DoEvents: DoEvents

cn.CommitTrans: Tran_Level = 0

'�D�P�B����
'On Error GoTo err_Handle
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
Loop

Me.WindowState = 2
 
If tmp_Rs.EOF Then
   'Release [���椤] �T������
   Unload frm_WaitWindows
   Set frm_WaitWindows = Nothing
   tmp_Rs.Close
   Screen.MousePointer = vbDefault
   msg_text = "�d�ߵ��G�G�S���ݺ��@���Ȥ��ƶǦ^�A���~��i�� [�ƨ��@�~]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_OrderImport.Enabled = True
   cmd_Update.Enabled = True
   Exit Sub
End If

Do While Not tmp_Rs.EOF
   With dg_ORT01W
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0    '�Ǹ�
        .Text = .Row
        .Col = 1    '���@�@�~�O
        .Text = tmp_Rs.Fields("�������O").Value
        .Col = 2    '�f�D
        .Text = tmp_Rs.Fields("�f�D").Value
        .Col = 3    '�Ȥ�s��
        .Text = tmp_Rs.Fields("�Ȥ�s��").Value
        .Col = 4    '�Ȥ�W��
        .Text = tmp_Rs.Fields("�Ȥ�W��").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close

'�p����J�q��ƶq
str_SQL = "Select Count(*) as RecCount From ORT02W"
Set tmp_Rs = Nothing

'��ܥثe��m���Ȥ���
dg_ORT01W.Row = 1
Call dg_ORT01W_Click

'Release [���椤] �T������
Unload frm_WaitWindows
Set frm_WaitWindows = Nothing
cmd_OrderImport.Enabled = True
cmd_Update.Enabled = True
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Unload frm_WaitWindows
   Set frm_WaitWindows = Nothing

   If Tran_Level <> 0 Then cn.RollbackTrans
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��ΫȤ�����J", Me.Caption, "cmd_OrderImport_Click", tmpString & str_SQL
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_OrderImport.Enabled = True
   cmd_Update.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Update_Click()
'�T�{�s��

'�M���S��r��
Call myFormExCharFilter(Me)

On Error GoTo err_Handle

'�s�ɸ���ˮ�
If CheckOP_ComsigneeData = False Then Exit Sub

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
tmp_Cmd.CommandText = "Master_ConsigneeData_Other_ImportUpdate"
'�f�D
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Storer_New.Text) > 0 Then
   tmp_Cmd.Parameters("StorerKey").Value = Trim(txt_Storer_New.Text)
Else
   tmp_Cmd.Parameters("StorerKey").Value = Trim(txt_Storer.Text)
End If

'�Ȥ�s��
Set tmp_para = tmp_Cmd.CreateParameter("ConsigneeKey", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_ConsigneeKey_New.Text) > 0 Then
   tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_ConsigneeKey_New.Text)
Else
   tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_ConsigneeKey.Text)
End If

'�l���ϸ�
Set tmp_para = tmp_Cmd.CreateParameter("ZIP", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Zip_New.ListIndex <> -1 Then
   tmp_Cmd.Parameters("ZIP").Value = arZip(cmb_Zip_New.ListIndex)
Else
   If cmb_ZIP.ListIndex <> -1 Then
      tmp_Cmd.Parameters("ZIP").Value = arZip(cmb_ZIP.ListIndex)
   Else
      tmp_Cmd.Parameters("ZIP").Value = ""
   End If
End If

'�B�e�ϽX�ˬd daniel
str_SQL = "select * from dbo.TRP03M Where AREA_CODE = '" & Trim(txt_AreaCode_New.Text) & "'"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧹B�e�ϽX"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    txt_AreaCode_New.SetFocus: Screen.MousePointer = 0
    Exit Sub
End If
tmp_Rs.Close

'�B�e�ϽX
Set tmp_para = tmp_Cmd.CreateParameter("Area_Code", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_AreaCode_New.Text) > 0 Then
   tmp_Cmd.Parameters("Area_Code").Value = Trim(txt_AreaCode_New.Text)
Else
   If Trim(txt_AreaCode.Text) = "" Then
      tmp_Cmd.Parameters("Area_Code").Value = Null
   Else
      tmp_Cmd.Parameters("Area_Code").Value = Trim(txt_AreaCode.Text)
   End If
End If

'�B�e�a�}
Set tmp_para = tmp_Cmd.CreateParameter("Address", adVarChar, adParamInput, 200)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Address_New.Text) > 0 Then
   tmp_Cmd.Parameters("Address").Value = Trim(txt_Address_New.Text)
Else
   If Trim(txt_Address.Text) = "" Then
      tmp_Cmd.Parameters("Address").Value = ""
   Else
      tmp_Cmd.Parameters("Address").Value = Trim(txt_Address.Text)
   End If
End If

'�p���H
Set tmp_para = tmp_Cmd.CreateParameter("Contact", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Contact_New.Text) > 0 Then
   tmp_Cmd.Parameters("Contact").Value = Trim(txt_Contact_New.Text)
Else
   If Trim(txt_Contact.Text) = "" Then
      tmp_Cmd.Parameters("Contact").Value = ""
   Else
      tmp_Cmd.Parameters("Contact").Value = Trim(txt_Contact.Text)
   End If
End If

'�q��
Set tmp_para = tmp_Cmd.CreateParameter("Phone", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Phone_New.Text) > 0 Then
   tmp_Cmd.Parameters("Phone").Value = Trim(txt_Phone_New.Text)
Else
   If Trim(txt_Phone.Text) = "" Then
      tmp_Cmd.Parameters("Phone").Value = ""
   Else
      tmp_Cmd.Parameters("Phone").Value = Trim(txt_Phone.Text)
   End If
End If

'�Ȥᵥ��
Set tmp_para = tmp_Cmd.CreateParameter("Class", adDouble, adParamInput, 2, 0)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Class_New.Text) > 0 Then
   tmp_Cmd.Parameters("Class").Value = RTrim(Val(txt_Class_New.Text))
Else
   If txt_Class.Text = "" Then
      tmp_Cmd.Parameters("Class").Value = Null
   Else
      tmp_Cmd.Parameters("Class").Value = RTrim(Val(txt_Class.Text))
   End If
End If

'�S��ݨD 1
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_ExtraDemand1.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = arExtraDemand(cmb_ExtraDemand1.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = arExtraDemand(0)
End If

'�S��ݨD 2
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code2", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_ExtraDemand2.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = arExtraDemand(cmb_ExtraDemand2.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = arExtraDemand(0)
End If

'�Ȥ�W��
Set tmp_para = tmp_Cmd.CreateParameter("Full_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_FullName_New.Text) > 0 Then
   tmp_Cmd.Parameters("Full_Name").Value = Trim(txt_FullName_New.Text)
Else
   tmp_Cmd.Parameters("Full_Name").Value = Trim(txt_FullName.Text)
End If

'�Ȥ�²��
Set tmp_para = tmp_Cmd.CreateParameter("Short_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_ShortName.Text) > 0 Then
   tmp_Cmd.Parameters("Short_Name").Value = Trim(txt_ShortName.Text)
Else
   tmp_Cmd.Parameters("Short_Name").Value = ""
End If

'�q�����A
Set tmp_para = tmp_Cmd.CreateParameter("Channel_Type", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_ChannelType.Text) > 0 Then
   tmp_Cmd.Parameters("Channel_Type").Value = Trim(txt_ChannelType.Text)
Else
   tmp_Cmd.Parameters("Channel_Type").Value = Null
End If

'���d������
Set tmp_para = tmp_Cmd.CreateParameter("Unload_Type", adVarChar, adParamInput, 3)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_UnLoad.Text) > 0 Then
   tmp_Cmd.Parameters("Unload_Type").Value = Trim(txt_UnLoad.Text)
Else
   tmp_Cmd.Parameters("Unload_Type").Value = Null
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
If chk_MultiCustomer.Value = vbChecked Then
   tmp_Cmd.Parameters("Multi_Customer").Value = "Y"
Else
   tmp_Cmd.Parameters("Multi_Customer").Value = "N"
End If

'Grid_Code �x�}�ϽX
Set tmp_para = tmp_Cmd.CreateParameter("Grid_Code", adVarChar, adParamInput, 5)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_UnLoad.Text) > 0 Then
   tmp_Cmd.Parameters("Grid_Code").Value = Trim(txt_GridCode.Text)
Else
   '����J�x�}�ϽX�G�H�l���ϸ��X�[�@�Ӧr�� + [1]
   If cmb_ZIP.ListIndex <> -1 Then
      tmp_Cmd.Parameters("Grid_Code").Value = arZip(cmb_ZIP.ListIndex) & "1"
   Else
      tmp_Cmd.Parameters("Grid_Code").Value = Null
   End If
End If
'���إN�X
Set tmp_para = tmp_Cmd.CreateParameter("Vehicle_Type", adVarChar, adParamInput, 2)
tmp_Cmd.Parameters.Append tmp_para
If cmb_VehicleType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Vehicle_Type").Value = arVehicleType(cmb_VehicleType.ListIndex)
Else
   tmp_Cmd.Parameters("Vehicle_Type").Value = Null
End If
'�h�B�u��
Set tmp_para = tmp_Cmd.CreateParameter("PICK_TOOL", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_PickTool.ListIndex <> -1 Then
   tmp_Cmd.Parameters("PICK_TOOL").Value = arPickTool(cmb_PickTool.ListIndex)
Else
   tmp_Cmd.Parameters("PICK_TOOL").Value = Null
End If

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'�D�P�B����
cmd_Update.Enabled = False
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   'Debug.Print tmp_cmd.State
   DoEvents: DoEvents  '�� [���椤] �T�������� [��s] �ɶ�
Loop
cmd_Update.Enabled = True

'�ݺ��@�Ȥ��� >> �w�s�ɤ���Ʀ�R��
If intGridRow = 0 Then Exit Sub
dg_ORT01W.Visible = False

Dim i As Integer, j As Integer

'1. �N�R���C��ƥѤU�@�C��ƨ��N
'   �ӫ᪺��ƦC���W���@�C
With dg_ORT01W
     For i = intGridRow To .Rows - 2   '�|���h�@��ťզC
         .Row = i
         For j = 0 To .Cols - 1
             .Col = j
             .Text = .TextArray((.Row + 1) * .Cols + .Col)
         Next j
         DoEvents
         '����̫�Ĥ@�C���W�����̫�ĤG�C�ɡA�|�O�˥ո�ƦC�A[�Ǹ�] ��줣�঳��
         '����ƪ��C�A[�Ǹ�] �������s�s��
         .Col = 0
         If Val(.Text) = 0 Then .Text = "" Else .Text = .Row
     Next i
'2. Grid �`�C�� - 1
     .Rows = .Rows - 1
     .Row = 1
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
'3. Reset �ܼ�
intGridRow = 0
dg_ORT01W.Visible = True

'4. ��ܥثe��Ʀ椧�Ȥ���
Call dg_ORT01W_Click

'��ܩҦ��ݽT�{���Ȥ���
Call Display_ORT01W
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�q��ΫȤ�����J-�T�{�s��", Me.Caption, "cmd_Update_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub dg_ORT01W_Click()
Dim i As Double, strStorerkey As String
With dg_ORT01W
     intGridRow = .Row
     '��ܫȤ�Ȧs�ɤ��Ȥ���
     Call Clear_ORT01W_ConsigneeData
     .Col = 2: strStorerkey = Trim(.Text) '�f�D�s��
     .Col = 3   '�Ȥ�s��
     str_SQL = "Select * From ORT01W Where ConsigneeKey = '" & Trim(.Text) & "' and storerkey = '" & strStorerkey & "' "

     Dim rsTmp As New ADODB.Recordset
     rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
     If Not rsTmp.EOF Then
        Display_ORT01W_ConsigneeData rsTmp
     End If
     rsTmp.Close
     
     Call Clear_TRP01M_ConsigneeData
     .Col = 1   '���@���O
     If .Text = "��" Then
        .Col = 3
        '���@���O�G���ʡA�����w���ɤ��Ȥ���
        str_SQL = "Select * From TRP01M Where ConsigneeKey = '" & Trim(.Text) & "' and storerkey = '" & strStorerkey & "' "
        Call Confirm_Recordset_Closed(tmp_Rs)
        Call DB_CheckConnectStatus
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If Not tmp_Rs.EOF Then
           Display_TRP01M_ConsigneeData tmp_Rs
        End If
        tmp_Rs.Close
     End If
     
End With
'�ϥտ���Ӧ��ơG������b�̫�
With dg_ORT01W
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�h�f�q����J�ΫȤᲧ�ʺ��@"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475

Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'��ܩҦ��ݽT�{���Ȥ���
Call Display_ORT01W

'��ܥثe��m���Ȥ���
dg_ORT01W.Row = 1
Call dg_ORT01W_Click

'���o �l���ϸ�
cmb_ZIP.Clear: cmb_Zip_New.Clear: intloop = 0
ReDim arZip(1) As String
str_SQL = "SELECT RTRIM(ZIP) AS �l���ϸ�,RTRIM(Isnull(Description,'')) AS ���� " & _
          "From TRP02M Order by ZIP"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_ZIP.AddItem tmp_Rs.Fields("�l���ϸ�").Value & "  " & tmp_Rs.Fields("����").Value
   cmb_Zip_New.AddItem tmp_Rs.Fields("�l���ϸ�").Value & "  " & tmp_Rs.Fields("����").Value
   intloop = intloop + 1
   If UBound(arZip) < intloop Then
      ReDim Preserve arZip(intloop) As String
   End If
   arZip(intloop - 1) = tmp_Rs.Fields("�l���ϸ�").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_ZIP.ListIndex = -1

'���o ���إN�X
cmb_VehicleType.Clear: intloop = 0
ReDim arVehicleType(1) As String
str_SQL = "SELECT RTRIM(Vehicle_Type) AS �N�X, RTRIM(Description) AS �������� " & _
          "From TRP15M Order by Vehicle_Type"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_VehicleType.AddItem tmp_Rs.Fields("�N�X").Value & "  " & tmp_Rs.Fields("��������").Value
   intloop = intloop + 1
   If UBound(arVehicleType) < intloop Then
      ReDim Preserve arVehicleType(intloop) As String
   End If
   arVehicleType(intloop - 1) = tmp_Rs.Fields("�N�X").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_VehicleType.ListIndex = -1

'���o �S��ݨD
cmb_ExtraDemand1.Clear: cmb_ExtraDemand2.Clear: intloop = 0
ReDim arExtraDemand(1) As String
str_SQL = "SELECT RTRIM(Extra_Demand_Code) AS �N�X, RTRIM(Description) AS �S��ݨD " & _
          "From TRP04M Order by Extra_Demand_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_ExtraDemand1.AddItem tmp_Rs.Fields("�N�X").Value & "  " & tmp_Rs.Fields("�S��ݨD").Value
   cmb_ExtraDemand2.AddItem tmp_Rs.Fields("�N�X").Value & "  " & tmp_Rs.Fields("�S��ݨD").Value
   intloop = intloop + 1
   If UBound(arExtraDemand) < intloop Then
      ReDim Preserve arExtraDemand(intloop) As String
   End If
   arExtraDemand(intloop - 1) = tmp_Rs.Fields("�N�X").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_ExtraDemand1.ListIndex = -1: cmb_ExtraDemand2.ListIndex = -1

'���o �h�B�u��
cmb_PickTool.Clear: intloop = 0
ReDim arPickTool(1) As String
str_SQL = "SELECT RTRIM(Code) AS �N�X, RTRIM(Description) AS �h�B�u�� " & _
          "From CodeLKUP Where ListName = 'MOVETOOL'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_PickTool.AddItem tmp_Rs.Fields("�N�X").Value & "  " & tmp_Rs.Fields("�h�B�u��").Value
   intloop = intloop + 1
   If UBound(arPickTool) < intloop Then
      ReDim Preserve arPickTool(intloop) As String
   End If
   arPickTool(intloop - 1) = tmp_Rs.Fields("�N�X").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_PickTool.ListIndex = -1

End Sub

Private Sub Form_Resize()
'�����j�p�ܰ�
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '�ܤp
   fam_Command.Left = fam_Command.Left - (dbsrcFormWidth - Me.ScaleWidth)
   fam_ConsignHead.Left = fam_ConsignHead.Left - (dbsrcFormWidth - Me.ScaleWidth)
   fam_ConsignDetail.Left = fam_ConsignDetail.Left - (dbsrcFormWidth - Me.ScaleWidth)
   
   dg_ORT01W.Width = dg_ORT01W.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_ORT01W.Height = dg_ORT01W.Height - (dbsrcFormHeight - Me.ScaleHeight)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   fam_Command.Left = fam_Command.Left + (Me.ScaleWidth - dbsrcFormWidth)
   fam_ConsignHead.Left = fam_ConsignHead.Left + (Me.ScaleWidth - dbsrcFormWidth)
   fam_ConsignDetail.Left = fam_ConsignDetail.Left + (Me.ScaleWidth - dbsrcFormWidth)
   
   dg_ORT01W.Width = dg_ORT01W.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_ORT01W.Height = dg_ORT01W.Height + (Me.ScaleHeight - dbsrcFormHeight)
   
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
Set frm_OP_Other_OrderImport = Nothing
End Sub

Private Sub SetGrid_Format_ORT01W()
'�g�q����J�ˮ֧P�_�A�ݥ� USER �T�{���Ȥ���
Dim sub_var1 As Integer, sub_var2 As Integer
dg_ORT01W.Visible = False
With dg_ORT01W
     .Rows = 2
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
     .ColWidth(2) = 800
     .ColWidth(3) = 2000
     .ColWidth(4) = 2500
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "No"
     .Col = 1: .Text = "��"
     .Col = 2: .Text = "�f�D"
     .Col = 3: .Text = "�Ȥ�s��"
     .Col = 4: .Text = "�Ȥ�W��"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignCenterCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_ORT01W.Visible = True
End Sub

Private Sub Display_ORT01W()
'��� ORT01W �Ȥ��ƼȦs��

Call SetGrid_Format_ORT01W
Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

str_SQL = "SELECT Rtrim(StorerKey) as �f�D , Rtrim(ConsigneeKey) as �Ȥ�s�� , Case Transaction_Status When '1' Then '�s' else '��' End as �������O , isnull(Rtrim(Full_Name),'') as �Ȥ�W�� " & _
         "FROM ORT01W order by TRANSACTION_STATUS desc,CONSIGNEEKEY"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   Set tmp_Rs = Nothing
   Exit Sub
Else
   Do While Not tmp_Rs.EOF
      With dg_ORT01W
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0    '�Ǹ�
        .Text = .Row
        .Col = 1    '���@�@�~�O
        .Text = tmp_Rs.Fields("�������O").Value
        .Col = 2    '�f�D
        .Text = tmp_Rs.Fields("�f�D").Value
        .Col = 3    '�Ȥ�s��
        .Text = tmp_Rs.Fields("�Ȥ�s��").Value
        .Col = 4    '�Ȥ�W��
        .Text = tmp_Rs.Fields("�Ȥ�W��").Value
      End With
      tmp_Rs.MoveNext
   Loop
   tmp_Rs.Close
   Set tmp_Rs = Nothing
End If
End Sub
Private Sub Clear_ORT01W_ConsigneeData()
'�M���Ȥ������GORT01W �ݨϥΪ̽T�{���Ȥ�Ȧs���
txt_Storer_New.Text = ""
txt_ConsigneeKey_New.Text = ""
txt_Class_New.Text = ""
cmb_Zip_New.ListIndex = -1
txt_AreaCode_New.Text = ""
txt_FullName_New.Text = ""
txt_Address_New.Text = ""
txt_Contact_New.Text = ""
txt_Phone_New.Text = ""
txt_ShortName.Text = ""
cmb_VehicleType.ListIndex = -1
cmb_ExtraDemand1.ListIndex = -1
cmb_ExtraDemand2.ListIndex = -1
txt_ChannelType.Text = ""
chk_MultiCustomer.Value = vbUnchecked
txt_UnLoad.Text = ""
End Sub
Private Sub Display_ORT01W_ConsigneeData(ByRef in_rs As ADODB.Recordset)
'��� �ݽT�{���Ȥ��� [ORT01W]
Dim i As Double
txt_Storer_New.Text = Trim(in_rs.Fields("StorerKey").Value)
txt_ConsigneeKey_New.Text = Trim(in_rs.Fields("ConsigneeKey").Value)

If IsNull(in_rs.Fields("Class").Value) Then
   txt_Class_New.Text = ""
Else
   txt_Class_New.Text = in_rs.Fields("Class").Value
End If

If IsNull(in_rs.Fields("ZIP").Value) Then
   cmb_Zip_New.ListIndex = -1
Else
   For i = 0 To cmb_Zip_New.ListCount - 1
       If arZip(i) = Trim(in_rs.Fields("ZIP").Value) Then
          ZipQueryAreaCode = False
          cmb_Zip_New.ListIndex = i
          ZipQueryAreaCode = True
          Exit For
       End If
   Next i
End If

If IsNull(in_rs.Fields("Area_Code").Value) Then
    txt_AreaCode_New.Text = ""
Else
   txt_AreaCode_New.Text = Trim(in_rs.Fields("Area_Code").Value)
End If

If IsNull(in_rs.Fields("Full_Name").Value) Then
   txt_FullName_New.Text = ""
Else
   txt_FullName_New.Text = Trim(in_rs.Fields("Full_Name").Value)
End If

If IsNull(in_rs.Fields("Address").Value) Then
   txt_Address_New.Text = ""
Else
   txt_Address_New.Text = Trim(in_rs.Fields("Address").Value)
End If

If IsNull(in_rs.Fields("Contact").Value) Then
   txt_Contact_New.Text = ""
Else
   txt_Contact_New.Text = Trim(in_rs.Fields("Contact").Value)
End If

If IsNull(in_rs.Fields("Phone").Value) Then
   txt_Phone_New.Text = ""
Else
   txt_Phone_New.Text = Trim(in_rs.Fields("Phone").Value)
End If

If IsNull(in_rs.Fields("Short_Name").Value) Then
   txt_ShortName.Text = ""
Else
   txt_ShortName.Text = Trim(in_rs.Fields("Short_Name").Value)
End If

If IsNull(in_rs.Fields("Vehicle_Type").Value) Then
   cmb_VehicleType.ListIndex = -1
Else
   For i = 0 To cmb_VehicleType.ListCount - 1
       If arVehicleType(i) = Trim(in_rs.Fields("Vehicle_Type").Value) Then
          cmb_VehicleType.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(in_rs.Fields("Extra_Demand_Code").Value) Then
   cmb_ExtraDemand1.ListIndex = -1
Else
   For i = 0 To cmb_ExtraDemand1.ListCount - 1
       If arExtraDemand(i) = Trim(in_rs.Fields("Extra_Demand_Code").Value) Then
          cmb_ExtraDemand1.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(in_rs.Fields("Extra_Demand_Code2").Value) Then
   cmb_ExtraDemand2.ListIndex = -1
Else
   For i = 0 To cmb_ExtraDemand2.ListCount - 1
       If arExtraDemand(i) = Trim(in_rs.Fields("Extra_Demand_Code2").Value) Then
          cmb_ExtraDemand2.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(in_rs.Fields("Channel_Type").Value) Then
   txt_ChannelType.Text = ""
Else
   txt_ChannelType.Text = Trim(in_rs.Fields("Channel_Type").Value)
End If

If IsNull(in_rs.Fields("Multi_Customer").Value) Then
   chk_MultiCustomer.Value = vbUnchecked
Else
   If Trim(in_rs.Fields("Multi_Customer").Value) = "N" Then
      chk_MultiCustomer.Value = vbUnchecked
   Else
      chk_MultiCustomer.Value = vbChecked
   End If
End If

If IsNull(in_rs.Fields("Unload_Type").Value) Then
   txt_UnLoad.Text = ""
Else
   txt_UnLoad.Text = Trim(in_rs.Fields("Unload_Type").Value)
End If

End Sub
Private Sub Clear_TRP01M_ConsigneeData()
'�M������ơG�w���ɫȤ������
txt_Storer.Text = ""
txt_ConsigneeKey.Text = ""
txt_Class.Text = ""
cmb_ZIP.ListIndex = -1
txt_AreaCode.Text = ""
txt_FullName.Text = ""
txt_Address.Text = ""
txt_Contact.Text = ""
txt_Phone.Text = ""
txt_ShortName.Text = ""
cmb_VehicleType.ListIndex = -1
cmb_ExtraDemand1.ListIndex = -1
cmb_ExtraDemand2.ListIndex = -1
txt_ChannelType.Text = ""
chk_MultiCustomer.Value = vbUnchecked
txt_UnLoad.Text = ""
End Sub
Private Sub Display_TRP01M_ConsigneeData(ByRef in_rs As ADODB.Recordset)
'��� �w���ɤ��Ȥ��� [TRP01M]
Dim i As Double
txt_Storer.Text = Trim(in_rs.Fields("StorerKey").Value)
txt_ConsigneeKey.Text = Trim(in_rs.Fields("ConsigneeKey").Value)
If IsNull(in_rs.Fields("Class").Value) Then
   txt_Class.Text = ""
Else
   txt_Class.Text = in_rs.Fields("Class").Value
End If
If IsNull(in_rs.Fields("ZIP").Value) Then
   cmb_ZIP.ListIndex = -1
Else
   For i = 0 To cmb_ZIP.ListCount - 1
       If arZip(i) = Trim(in_rs.Fields("ZIP").Value) Then
          ZipQueryAreaCode = False
          cmb_ZIP.ListIndex = i
          ZipQueryAreaCode = True
          Exit For
       End If
   Next i
End If
If IsNull(in_rs.Fields("Area_Code").Value) Then
    txt_AreaCode.Text = ""
Else
   txt_AreaCode.Text = Trim(in_rs.Fields("Area_Code").Value)
End If
If IsNull(in_rs.Fields("Full_Name").Value) Then
   txt_FullName.Text = ""
Else
   txt_FullName.Text = Trim(in_rs.Fields("Full_Name").Value)
End If
If IsNull(in_rs.Fields("Address").Value) Then
   txt_Address.Text = ""
Else
   txt_Address.Text = Trim(in_rs.Fields("Address").Value)
End If
If IsNull(in_rs.Fields("Contact").Value) Then
   txt_Contact.Text = ""
Else
   txt_Contact.Text = Trim(in_rs.Fields("Contact").Value)
End If
If IsNull(in_rs.Fields("Phone").Value) Then
   txt_Phone.Text = ""
Else
   txt_Phone.Text = Trim(in_rs.Fields("Phone").Value)
End If
If IsNull(in_rs.Fields("Short_Name").Value) Then
   txt_ShortName.Text = ""
Else
   txt_ShortName.Text = Trim(in_rs.Fields("Short_Name").Value)
End If
If IsNull(in_rs.Fields("Grid_Code").Value) Then
   txt_GridCode.Text = ""
Else
   txt_GridCode.Text = Trim(in_rs.Fields("Grid_Code").Value)
End If
If IsNull(in_rs.Fields("Vehicle_Type").Value) Then
   cmb_VehicleType.ListIndex = -1
Else
   For i = 0 To cmb_VehicleType.ListCount - 1
       If arVehicleType(i) = Trim(in_rs.Fields("Vehicle_Type").Value) Then
          cmb_VehicleType.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(in_rs.Fields("Extra_Demand_Code").Value) Then
   cmb_ExtraDemand1.ListIndex = -1
Else
   For i = 0 To cmb_ExtraDemand1.ListCount - 1
       If arExtraDemand(i) = Trim(in_rs.Fields("Extra_Demand_Code").Value) Then
          cmb_ExtraDemand1.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(in_rs.Fields("Extra_Demand_Code2").Value) Then
   cmb_ExtraDemand2.ListIndex = -1
Else
   For i = 0 To cmb_ExtraDemand2.ListCount - 1
       If arExtraDemand(i) = Trim(in_rs.Fields("Extra_Demand_Code2").Value) Then
          cmb_ExtraDemand2.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(in_rs.Fields("Channel_Type").Value) Then
   txt_ChannelType.Text = ""
Else
   txt_ChannelType.Text = Trim(in_rs.Fields("Channel_Type").Value)
End If

If IsNull(in_rs.Fields("Multi_Customer").Value) Then
   chk_MultiCustomer.Value = vbUnchecked
Else
   If Trim(in_rs.Fields("Multi_Customer").Value) = "N" Then
      chk_MultiCustomer.Value = vbUnchecked
   Else
      chk_MultiCustomer.Value = vbChecked
   End If
End If

If IsNull(in_rs.Fields("Unload_Type").Value) Then
   txt_UnLoad.Text = ""
Else
   txt_UnLoad.Text = Trim(in_rs.Fields("Unload_Type").Value)
End If
If IsNull(in_rs.Fields("Pick_Tool").Value) Then
   cmb_PickTool.ListIndex = -1
Else
   For i = 0 To cmb_PickTool.ListCount - 1
       If arPickTool(i) = Trim(in_rs.Fields("Pick_Tool").Value) Then
          cmb_PickTool.ListIndex = i
          Exit For
       End If
   Next i
End If
End Sub

Private Sub txt_Address_KeyPress(KeyAscii As Integer)
'TRP01M �B�e�a�} ���i�s��
KeyAscii = 0
End Sub

Private Sub txt_AreaCode_KeyPress(KeyAscii As Integer)
'TRP01M �B�e�ϽX ���i�s��
KeyAscii = 0
End Sub



Private Sub txt_AreaCode_New_LostFocus()    'daniel-20041001
    If Len(txt_AreaCode.Text) = 0 Then Exit Sub
    str_SQL = "select * from dbo.TRP03M Where AREA_CODE = '" & txt_AreaCode.Text & "'"
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧹B�e�ϽX"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
    tmp_Rs.Close
End Sub

Private Sub txt_Class_KeyPress(KeyAscii As Integer)
'TRP01M �Ȥᵥ�� ���i�s��
KeyAscii = 0
End Sub

Private Sub txt_ConsigneeKey_KeyPress(KeyAscii As Integer)
'TRP01M �Ȥ�s�� ���i�s��
KeyAscii = 0
End Sub

Private Sub txt_Contact_KeyPress(KeyAscii As Integer)
'TRP01M �s���H ���i�s��
KeyAscii = 0
End Sub

Private Sub txt_FullName_KeyPress(KeyAscii As Integer)
'TRP01M �Ȥ�W�� ���i�s��
KeyAscii = 0
End Sub

Private Sub txt_ChannelType_KeyPress(KeyAscii As Integer)
'�q�����A
Select Case KeyAscii
     Case 97 To 122     '�ഫ�j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          If Trim(txt_ChannelType.Text) <> "KA" And Trim(txt_ChannelType.Text) <> "GT" Then
             msg_text = "�q�����A��ƿ��~�G�u�i��J KA �� GT "
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_ChannelType.SelStart = 0: txt_ChannelType.SelLength = Len(txt_ChannelType.Text)
             txt_ChannelType.SetFocus
          End If
End Select
End Sub

Private Sub txt_Phone_KeyPress(KeyAscii As Integer)
'TRP01M �q�� ���i�s��
KeyAscii = 0
End Sub

Private Sub txt_Storer_KeyPress(KeyAscii As Integer)
'TRP01M �f�D ���i�s��
KeyAscii = 0
End Sub

Private Function CheckOP_ComsigneeData() As Boolean
'�Ȥ��� [�T�{�s��] �ˮ�
CheckOP_ComsigneeData = False
msg_text = ""
If Len(Trim(txt_Storer_New.Text)) = 0 And Len(Trim(txt_Storer.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J[�f�D]"
   Else
      msg_text = msg_text & vbCrLf & "����J[�f�D]"
   End If
End If
If Len(Trim(txt_ConsigneeKey_New.Text)) = 0 And Len(Trim(txt_ConsigneeKey.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J[�Ȥ�s��]"
   Else
      msg_text = msg_text & vbCrLf & "����J[�Ȥ�s��]"
   End If
End If

If Len(RTrim(cmb_Zip_New)) = 0 Then
   If msg_text = "" Then
      msg_text = "����J[�l���ϸ�]"
   Else
      msg_text = msg_text & vbCrLf & "����J[�l���ϸ�]"
   End If
End If

If msg_text = "" Then
   CheckOP_ComsigneeData = True
Else
   msg_text = "�Ȥ��Ʋ��`�A�Эץ���A���� [�T�{�s��]�G" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

End Function

