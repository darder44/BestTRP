VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_Other_OrderImport 
   Caption         =   "其它訂單轉入及客戶異動維護"
   ClientHeight    =   7140
   ClientLeft      =   270
   ClientTop       =   990
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11280
   WindowState     =   2  '最大化
   Begin VB.Frame fam_Command 
      Height          =   720
      Left            =   4260
      TabIndex        =   53
      Top             =   -75
      Width           =   7155
      Begin VB.CommandButton cmd_OrderImport 
         BackColor       =   &H00FF8080&
         Caption         =   "訂單及客戶資料轉入"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         Style           =   1  '圖片外觀
         TabIndex        =   31
         Top             =   120
         Width           =   2250
      End
      Begin VB.CommandButton cmd_Update 
         BackColor       =   &H8000000B&
         Caption         =   "確認存檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2790
         Style           =   1  '圖片外觀
         TabIndex        =   32
         Top             =   135
         Width           =   1860
      End
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "離  開"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Style           =   1  '圖片外觀
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
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   780
         Width           =   1995
      End
      Begin VB.ComboBox cmb_Zip 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         Style           =   2  '單純下拉式
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
         BackStyle       =   0  '透明
         Caption         =   "貨        主"
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   43
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶編號"
         Height          =   180
         Index           =   1
         Left            =   1950
         TabIndex        =   42
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "郵遞區號"
         Height          =   180
         Index           =   2
         Left            =   255
         TabIndex        =   41
         Top             =   870
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "運送區碼"
         Height          =   180
         Index           =   3
         Left            =   4335
         TabIndex        =   40
         Top             =   855
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶名稱"
         Height          =   180
         Index           =   4
         Left            =   255
         TabIndex        =   39
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "運送地址"
         Height          =   180
         Index           =   5
         Left            =   255
         TabIndex        =   38
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "樓層"
         Height          =   180
         Index           =   7
         Left            =   4335
         TabIndex        =   37
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "聯絡人"
         Height          =   180
         Index           =   8
         Left            =   435
         TabIndex        =   36
         Top             =   2790
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "電話"
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
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4755
         Style           =   2  '單純下拉式
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
         Style           =   2  '單純下拉式
         TabIndex        =   25
         Top             =   570
         Width           =   5445
      End
      Begin VB.ComboBox cmb_ExtraDemand1 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   1290
         Style           =   2  '單純下拉式
         TabIndex        =   26
         Top             =   915
         Width           =   5445
      End
      Begin VB.ComboBox cmb_ExtraDemand2 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   1290
         Style           =   2  '單純下拉式
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
         Caption         =   "指送客戶"
         Height          =   180
         Left            =   465
         TabIndex        =   45
         Top             =   2070
         Width           =   1260
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  '平面
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2445
         Width           =   705
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  '平面
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
         BackStyle       =   0  '透明
         Caption         =   "搬運工具"
         Height          =   180
         Index           =   19
         Left            =   3975
         TabIndex        =   54
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶簡稱"
         Height          =   180
         Index           =   6
         Left            =   495
         TabIndex        =   52
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "車種代碼"
         Height          =   180
         Index           =   10
         Left            =   495
         TabIndex        =   51
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "特殊需求 1"
         Height          =   180
         Index           =   11
         Left            =   360
         TabIndex        =   50
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "特殊需求 2"
         Height          =   180
         Index           =   12
         Left            =   360
         TabIndex        =   49
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "通路型態"
         Height          =   180
         Index           =   13
         Left            =   495
         TabIndex        =   48
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "卸貨難易度"
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
         Caption         =   "  待使用者確認之客戶資料  "
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
         Caption         =   "  已建檔之之客戶資料         "
         Height          =   180
         Index           =   17
         Left            =   1245
         TabIndex        =   2
         Top             =   2790
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "矩陣圖碼"
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
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬

Private intloop As Double
Private ZipQueryAreaCode As Boolean
Private intGridRow As Double

Private arZip() As String
Private arVehicleType() As String
Private arExtraDemand() As String
Private arPickTool() As String        '搬運工具

Private rs_ORT01W As ADODB.Recordset

Private Sub cmb_Zip_New_Change()
'取回 郵遞區號 所屬之 運送區域代碼
If ZipQueryAreaCode = False Then Exit Sub
If cmb_Zip_New.ListIndex = -1 Then Exit Sub

'取得選取之 Company 之所有 Branch
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
'取回 郵遞區號 所屬之 運送區域代碼
'If ZipQueryAreaCode = False Then Exit Sub
If cmb_Zip_New.ListIndex = -1 Then Exit Sub

'取得選取之 Company 之所有 Branch
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
'離開
Unload Me
End Sub

Private Sub cmd_OrderImport_Click()
'訂單及客戶資料轉入

On Error GoTo err_Handle

    Tran_Level = cn.BeginTrans
'退貨資料轉入WMS，要排除利豐的，利豐退貨寫入ASN Table
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
            "where o.B_PHONE2 is null and o.priority in ('R','RC','A2B') and o.storerkey not in ('LLFA01','LMBO01','LPSI01','LCHF01','LKYF01', 'LNCE01') and o.type <> '刪單' " & _
            "group by od.storerkey ,o.orderkey , o.externorderkey , o.priority ,od.storerkey ,o.orderkey , od.externorderkey , o.priority , o.consigneekey , o.c_company ,od.sku, s.descr,cast(o.notes as varchar(300)),s.packkey " & _
            "order by od.storerkey , o.orderkey "

Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = 3
Dim rsKeycount As New ADODB.Recordset
rsTmp.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
    
'沒有退貨單資料
If rsTmp.EOF Then GoTo LFMBO

    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    Dim strKeycount As String, strOrderkey As String, intLineNumber As Integer
    
    '讀取ini參數，其他排車是否入WMS系統
    Dim objIni As New vbIniFile, strOtherOrder2WMS As String
    objIni.FileName = App.Path & "/" & App.title & ".ini"
    
    strOtherOrder2WMS = objIni.ReadData("OPTION", "OtherOrder2WMS", "YES")
    Set objIni = Nothing
    
    If UCase(strOtherOrder2WMS) = "YES" Then 'WMS新增採購單
    
        rsTmp.MoveFirst
        rsTmp.Filter = "Priority = 'R' Or Priority = 'RC'"
        Do While Not rsTmp.EOF
        '亞培的訂單排除不寫入Exceed
        If RTrim(rsTmp("StorerKey")) = "LABT01" Then GoTo NextRow
        
        '寫入WMS
        If Trim(rsTmp("orderkey")) <> strOrderkey Then
    
            '取系統訂單單號
            rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
            '單號+1
            cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
            strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
            rsKeycount.Close

            '寫入表頭
            str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,BuyersReference , sellername,selleraddress1,externpokey,potype,notes) " & _
                      "values( '" & strKeycount & "','" & rsTmp("StorerKey") & "','" & RTrim(GetWord(Trim(rsTmp("ExternOrderKey")), 1, 18)) & "','" & rsTmp("consigneekey") & "','" & rsTmp("C_company") & "','" & rsTmp("OrderKey") & "','" & rsTmp("priority") & "','" & rsTmp("notes") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
            intLineNumber = 1
            strOrderkey = Trim(rsTmp("orderkey"))
    
        End If
    
            '寫入表身
            str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,Skudescription,StorerKey,QtyOrdered,packkey) " & _
                    "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & rsTmp("OrderLineNumber") & "','" & rsTmp("SKU") & "','" & rsTmp("descr") & "','" & rsTmp("StorerKey") & "','" & rsTmp("openqty") & "','" & rsTmp("packkey") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
            intLineNumber = intLineNumber + 1
NextRow:
        rsTmp.MoveNext

        Loop
        rsTmp.Close: Set rsTmp = Nothing
        
LFMBO:
    '寫入利豐的
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
                "where  o.priority in ('R','RC','A2B') and o.storerkey in ('LLFA01','LMBO01','LPSI01','LCHF01','LKYF01', 'LNCE01') and o.B_PHONE2 is null and o.type <> '刪單' " & _
                "group by od.storerkey ,o.orderkey , o.externorderkey , o.priority ,od.storerkey ,o.orderkey , od.externorderkey , o.priority , rtrim(isnull(o.externordertype,'')),o.consigneekey , o.c_company ,od.sku, s.descr,cast(o.notes as varchar(300)),s.packkey ,orderlinenumber,o.b_company,o.buyerpo,o.CustomerOrderkey,convert(char(8),o.orderdate,112),convert(char(8),o.deliverydate,112),od.externlineno,od.otheruom,od.lottable05,od.retailsku,od.lottable06,od.lottable03 " & _
                "order by o.orderkey ,orderlinenumber "

            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
            
    If rsTmp.EOF Then GoTo final  ': MsgBox "查詢結果：沒有需維護之客戶資料傳回，請繼續進行 [排車作業]", vbOKOnly, Me.Caption:
        
    rsTmp.MoveFirst
    rsTmp.Filter = "Priority = 'R' Or Priority = 'RC'"
    strOrderkey = ""
    '寫入ASN的資料
    Do While Not rsTmp.EOF
            '寫入WMS
            If Trim(rsTmp("orderkey")) <> strOrderkey Then
                '檢查是否貨主單號重複，重複則不寫入
                str_SQL = "select externasnkey from " & strWMSDB & "..asn where asntype = 'R' and storerkey = '" & Trim(rsTmp("storerkey")) & "' and externasnkey = '" & Trim(rsTmp("externorderkey")) & "'"
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.CursorLocation = 3
                tmp_Rs.Open str_SQL, cn
                    If tmp_Rs.EOF Then
                        tmp_Rs.Close
                        '取系統訂單單號
                        rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='ASN' ", cn
                        '單號+1
                        cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'ASN'", RowsAffect, adExecuteNoRecords
                        strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
                        rsKeycount.Close
                        
                        '寫入表頭
                        If Trim(rsTmp("StorerKey")) = "LLFA01" Then
                            str_SQL = "insert into " & strWMSDB & "..asn (asnKey,StorerKey,externasnkey , sellersreference,asntype,notes,SellersReference2,OtherReference,ASNDate,VesselDate) " & _
                                      "values( '" & strKeycount & "','" & rsTmp("StorerKey") & "','" & RTrim(GetWord(Trim(rsTmp("ExternOrderKey")), 1, 18)) & "','" & rsTmp("consigneekey") & "','" & rsTmp("priority") & "','" & rsTmp("notes") & "','" & _
                                      rsTmp("b_company") & "','" & rsTmp("customerorderkey") & "','" & rsTmp("orderdate") & "','" & rsTmp("deliverydate") & "') "
                                                              
                        Else
                            '毛寶
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
                '寫入表身
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
tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "ORTD11_IMPORTN"
Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'顯示 [執行中] 訊息
Load frm_WaitWindows
frm_WaitWindows.Tag = Me.Name
frm_WaitWindows.ZOrder
frm_WaitWindows.Refresh
DoEvents: DoEvents

cn.CommitTrans: Tran_Level = 0

'非同步執行
'On Error GoTo err_Handle
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
Loop

Me.WindowState = 2
 
If tmp_Rs.EOF Then
   'Release [執行中] 訊息視窗
   Unload frm_WaitWindows
   Set frm_WaitWindows = Nothing
   tmp_Rs.Close
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：沒有需維護之客戶資料傳回，請繼續進行 [排車作業]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_OrderImport.Enabled = True
   cmd_Update.Enabled = True
   Exit Sub
End If

Do While Not tmp_Rs.EOF
   With dg_ORT01W
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0    '序號
        .Text = .Row
        .Col = 1    '維護作業別
        .Text = tmp_Rs.Fields("異動類別").Value
        .Col = 2    '貨主
        .Text = tmp_Rs.Fields("貨主").Value
        .Col = 3    '客戶編號
        .Text = tmp_Rs.Fields("客戶編號").Value
        .Col = 4    '客戶名稱
        .Text = tmp_Rs.Fields("客戶名稱").Value
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close

'計算轉入訂單數量
str_SQL = "Select Count(*) as RecCount From ORT02W"
Set tmp_Rs = Nothing

'顯示目前位置之客戶資料
dg_ORT01W.Row = 1
Call dg_ORT01W_Click

'Release [執行中] 訊息視窗
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
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單及客戶資料轉入", Me.Caption, "cmd_OrderImport_Click", tmpString & str_SQL
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_OrderImport.Enabled = True
   cmd_Update.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Update_Click()
'確認存檔

'清除特殊字元
Call myFormExCharFilter(Me)

On Error GoTo err_Handle

'存檔資料檢核
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
tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_ConsigneeData_Other_ImportUpdate"
'貨主
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Storer_New.Text) > 0 Then
   tmp_Cmd.Parameters("StorerKey").Value = Trim(txt_Storer_New.Text)
Else
   tmp_Cmd.Parameters("StorerKey").Value = Trim(txt_Storer.Text)
End If

'客戶編號
Set tmp_para = tmp_Cmd.CreateParameter("ConsigneeKey", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_ConsigneeKey_New.Text) > 0 Then
   tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_ConsigneeKey_New.Text)
Else
   tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_ConsigneeKey.Text)
End If

'郵遞區號
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

'運送區碼檢查 daniel
str_SQL = "select * from dbo.TRP03M Where AREA_CODE = '" & Trim(txt_AreaCode_New.Text) & "'"
Call DB_CheckConnectStatus
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    msg_text = "查詢結果：無符合搜尋條件之運送區碼"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    txt_AreaCode_New.SetFocus: Screen.MousePointer = 0
    Exit Sub
End If
tmp_Rs.Close

'運送區碼
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

'運送地址
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

'聯絡人
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

'電話
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

'客戶等級
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

'特殊需求 1
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_ExtraDemand1.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = arExtraDemand(cmb_ExtraDemand1.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = arExtraDemand(0)
End If

'特殊需求 2
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code2", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_ExtraDemand2.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = arExtraDemand(cmb_ExtraDemand2.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = arExtraDemand(0)
End If

'客戶名稱
Set tmp_para = tmp_Cmd.CreateParameter("Full_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_FullName_New.Text) > 0 Then
   tmp_Cmd.Parameters("Full_Name").Value = Trim(txt_FullName_New.Text)
Else
   tmp_Cmd.Parameters("Full_Name").Value = Trim(txt_FullName.Text)
End If

'客戶簡稱
Set tmp_para = tmp_Cmd.CreateParameter("Short_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_ShortName.Text) > 0 Then
   tmp_Cmd.Parameters("Short_Name").Value = Trim(txt_ShortName.Text)
Else
   tmp_Cmd.Parameters("Short_Name").Value = ""
End If

'通路型態
Set tmp_para = tmp_Cmd.CreateParameter("Channel_Type", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_ChannelType.Text) > 0 Then
   tmp_Cmd.Parameters("Channel_Type").Value = Trim(txt_ChannelType.Text)
Else
   tmp_Cmd.Parameters("Channel_Type").Value = Null
End If

'拆櫃難易度
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
'指送客戶
Set tmp_para = tmp_Cmd.CreateParameter("Multi_Customer", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
If chk_MultiCustomer.Value = vbChecked Then
   tmp_Cmd.Parameters("Multi_Customer").Value = "Y"
Else
   tmp_Cmd.Parameters("Multi_Customer").Value = "N"
End If

'Grid_Code 矩陣圖碼
Set tmp_para = tmp_Cmd.CreateParameter("Grid_Code", adVarChar, adParamInput, 5)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_UnLoad.Text) > 0 Then
   tmp_Cmd.Parameters("Grid_Code").Value = Trim(txt_GridCode.Text)
Else
   '未輸入矩陣圖碼：以郵遞區號碼加一個字元 + [1]
   If cmb_ZIP.ListIndex <> -1 Then
      tmp_Cmd.Parameters("Grid_Code").Value = arZip(cmb_ZIP.ListIndex) & "1"
   Else
      tmp_Cmd.Parameters("Grid_Code").Value = Null
   End If
End If
'車種代碼
Set tmp_para = tmp_Cmd.CreateParameter("Vehicle_Type", adVarChar, adParamInput, 2)
tmp_Cmd.Parameters.Append tmp_para
If cmb_VehicleType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Vehicle_Type").Value = arVehicleType(cmb_VehicleType.ListIndex)
Else
   tmp_Cmd.Parameters("Vehicle_Type").Value = Null
End If
'搬運工具
Set tmp_para = tmp_Cmd.CreateParameter("PICK_TOOL", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_PickTool.ListIndex <> -1 Then
   tmp_Cmd.Parameters("PICK_TOOL").Value = arPickTool(cmb_PickTool.ListIndex)
Else
   tmp_Cmd.Parameters("PICK_TOOL").Value = Null
End If

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'非同步執行
cmd_Update.Enabled = False
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   'Debug.Print tmp_cmd.State
   DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
Loop
cmd_Update.Enabled = True

'待維護客戶資料 >> 已存檔之資料行刪除
If intGridRow = 0 Then Exit Sub
dg_ORT01W.Visible = False

Dim i As Integer, j As Integer

'1. 將刪除列資料由下一列資料取代
'   而後的資料列往上移一列
With dg_ORT01W
     For i = intGridRow To .Rows - 2   '會有多一行空白列
         .Row = i
         For j = 0 To .Cols - 1
             .Col = j
             .Text = .TextArray((.Row + 1) * .Cols + .Col)
         Next j
         DoEvents
         '防止最後第一列往上移給最後第二列時，會是弄白資料列，[序號] 欄位不能有值
         '有資料的列，[序號] 必須重新編號
         .Col = 0
         If Val(.Text) = 0 Then .Text = "" Else .Text = .Row
     Next i
'2. Grid 總列數 - 1
     .Rows = .Rows - 1
     .Row = 1
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
'3. Reset 變數
intGridRow = 0
dg_ORT01W.Visible = True

'4. 顯示目前資料行之客戶資料
Call dg_ORT01W_Click

'顯示所有待確認之客戶資料
Call Display_ORT01W
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單及客戶資料轉入-確認存檔", Me.Caption, "cmd_Update_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub dg_ORT01W_Click()
Dim i As Double, strStorerkey As String
With dg_ORT01W
     intGridRow = .Row
     '顯示客戶暫存檔之客戶資料
     Call Clear_ORT01W_ConsigneeData
     .Col = 2: strStorerkey = Trim(.Text) '貨主編號
     .Col = 3   '客戶編號
     str_SQL = "Select * From ORT01W Where ConsigneeKey = '" & Trim(.Text) & "' and storerkey = '" & strStorerkey & "' "

     Dim rsTmp As New ADODB.Recordset
     rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
     If Not rsTmp.EOF Then
        Display_ORT01W_ConsigneeData rsTmp
     End If
     rsTmp.Close
     
     Call Clear_TRP01M_ConsigneeData
     .Col = 1   '維護類別
     If .Text = "異" Then
        .Col = 3
        '維護類別：異動，取為已建檔之客戶資料
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
'反白選取該行資料：必須放在最後
With dg_ORT01W
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "退貨訂單轉入及客戶異動維護"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475

Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'顯示所有待確認之客戶資料
Call Display_ORT01W

'顯示目前位置之客戶資料
dg_ORT01W.Row = 1
Call dg_ORT01W_Click

'取得 郵遞區號
cmb_ZIP.Clear: cmb_Zip_New.Clear: intloop = 0
ReDim arZip(1) As String
str_SQL = "SELECT RTRIM(ZIP) AS 郵遞區號,RTRIM(Isnull(Description,'')) AS 說明 " & _
          "From TRP02M Order by ZIP"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_ZIP.AddItem tmp_Rs.Fields("郵遞區號").Value & "  " & tmp_Rs.Fields("說明").Value
   cmb_Zip_New.AddItem tmp_Rs.Fields("郵遞區號").Value & "  " & tmp_Rs.Fields("說明").Value
   intloop = intloop + 1
   If UBound(arZip) < intloop Then
      ReDim Preserve arZip(intloop) As String
   End If
   arZip(intloop - 1) = tmp_Rs.Fields("郵遞區號").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_ZIP.ListIndex = -1

'取得 車種代碼
cmb_VehicleType.Clear: intloop = 0
ReDim arVehicleType(1) As String
str_SQL = "SELECT RTRIM(Vehicle_Type) AS 代碼, RTRIM(Description) AS 車輛種類 " & _
          "From TRP15M Order by Vehicle_Type"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_VehicleType.AddItem tmp_Rs.Fields("代碼").Value & "  " & tmp_Rs.Fields("車輛種類").Value
   intloop = intloop + 1
   If UBound(arVehicleType) < intloop Then
      ReDim Preserve arVehicleType(intloop) As String
   End If
   arVehicleType(intloop - 1) = tmp_Rs.Fields("代碼").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_VehicleType.ListIndex = -1

'取得 特殊需求
cmb_ExtraDemand1.Clear: cmb_ExtraDemand2.Clear: intloop = 0
ReDim arExtraDemand(1) As String
str_SQL = "SELECT RTRIM(Extra_Demand_Code) AS 代碼, RTRIM(Description) AS 特殊需求 " & _
          "From TRP04M Order by Extra_Demand_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_ExtraDemand1.AddItem tmp_Rs.Fields("代碼").Value & "  " & tmp_Rs.Fields("特殊需求").Value
   cmb_ExtraDemand2.AddItem tmp_Rs.Fields("代碼").Value & "  " & tmp_Rs.Fields("特殊需求").Value
   intloop = intloop + 1
   If UBound(arExtraDemand) < intloop Then
      ReDim Preserve arExtraDemand(intloop) As String
   End If
   arExtraDemand(intloop - 1) = tmp_Rs.Fields("代碼").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_ExtraDemand1.ListIndex = -1: cmb_ExtraDemand2.ListIndex = -1

'取得 搬運工具
cmb_PickTool.Clear: intloop = 0
ReDim arPickTool(1) As String
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 搬運工具 " & _
          "From CodeLKUP Where ListName = 'MOVETOOL'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_PickTool.AddItem tmp_Rs.Fields("代碼").Value & "  " & tmp_Rs.Fields("搬運工具").Value
   intloop = intloop + 1
   If UBound(arPickTool) < intloop Then
      ReDim Preserve arPickTool(intloop) As String
   End If
   arPickTool(intloop - 1) = tmp_Rs.Fields("代碼").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_PickTool.ListIndex = -1

End Sub

Private Sub Form_Resize()
'視窗大小變動
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '變小
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
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_OP_Other_OrderImport = Nothing
End Sub

Private Sub SetGrid_Format_ORT01W()
'經訂單轉入檢核判斷，需由 USER 確認之客戶資料
Dim sub_var1 As Integer, sub_var2 As Integer
dg_ORT01W.Visible = False
With dg_ORT01W
     .Rows = 2
     .FixedRows = 1
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 300
     .ColWidth(1) = 400
     .ColWidth(2) = 800
     .ColWidth(3) = 2000
     .ColWidth(4) = 2500
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "No"
     .Col = 1: .Text = "※"
     .Col = 2: .Text = "貨主"
     .Col = 3: .Text = "客戶編號"
     .Col = 4: .Text = "客戶名稱"
     '設定列表之文字對齊
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
'顯示 ORT01W 客戶資料暫存檔

Call SetGrid_Format_ORT01W
Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

str_SQL = "SELECT Rtrim(StorerKey) as 貨主 , Rtrim(ConsigneeKey) as 客戶編號 , Case Transaction_Status When '1' Then '新' else '異' End as 異動類別 , isnull(Rtrim(Full_Name),'') as 客戶名稱 " & _
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
        .Col = 0    '序號
        .Text = .Row
        .Col = 1    '維護作業別
        .Text = tmp_Rs.Fields("異動類別").Value
        .Col = 2    '貨主
        .Text = tmp_Rs.Fields("貨主").Value
        .Col = 3    '客戶編號
        .Text = tmp_Rs.Fields("客戶編號").Value
        .Col = 4    '客戶名稱
        .Text = tmp_Rs.Fields("客戶名稱").Value
      End With
      tmp_Rs.MoveNext
   Loop
   tmp_Rs.Close
   Set tmp_Rs = Nothing
End If
End Sub
Private Sub Clear_ORT01W_ConsigneeData()
'清除客戶資料欄位：ORT01W 待使用者確認之客戶暫存資料
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
'顯示 待確認之客戶資料 [ORT01W]
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
'清除欄位資料：已建檔客戶資料欄位
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
'顯示 已建檔之客戶資料 [TRP01M]
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
'TRP01M 運送地址 不可編輯
KeyAscii = 0
End Sub

Private Sub txt_AreaCode_KeyPress(KeyAscii As Integer)
'TRP01M 運送區碼 不可編輯
KeyAscii = 0
End Sub



Private Sub txt_AreaCode_New_LostFocus()    'daniel-20041001
    If Len(txt_AreaCode.Text) = 0 Then Exit Sub
    str_SQL = "select * from dbo.TRP03M Where AREA_CODE = '" & txt_AreaCode.Text & "'"
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        msg_text = "查詢結果：無符合搜尋條件之運送區碼"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
    tmp_Rs.Close
End Sub

Private Sub txt_Class_KeyPress(KeyAscii As Integer)
'TRP01M 客戶等級 不可編輯
KeyAscii = 0
End Sub

Private Sub txt_ConsigneeKey_KeyPress(KeyAscii As Integer)
'TRP01M 客戶編號 不可編輯
KeyAscii = 0
End Sub

Private Sub txt_Contact_KeyPress(KeyAscii As Integer)
'TRP01M 連絡人 不可編輯
KeyAscii = 0
End Sub

Private Sub txt_FullName_KeyPress(KeyAscii As Integer)
'TRP01M 客戶名稱 不可編輯
KeyAscii = 0
End Sub

Private Sub txt_ChannelType_KeyPress(KeyAscii As Integer)
'通路型態
Select Case KeyAscii
     Case 97 To 122     '轉換大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          If Trim(txt_ChannelType.Text) <> "KA" And Trim(txt_ChannelType.Text) <> "GT" Then
             msg_text = "通路型態資料錯誤：只可輸入 KA 或 GT "
             MsgBox msg_text, vbOKOnly + vbInformation, msg_title
             txt_ChannelType.SelStart = 0: txt_ChannelType.SelLength = Len(txt_ChannelType.Text)
             txt_ChannelType.SetFocus
          End If
End Select
End Sub

Private Sub txt_Phone_KeyPress(KeyAscii As Integer)
'TRP01M 電話 不可編輯
KeyAscii = 0
End Sub

Private Sub txt_Storer_KeyPress(KeyAscii As Integer)
'TRP01M 貨主 不可編輯
KeyAscii = 0
End Sub

Private Function CheckOP_ComsigneeData() As Boolean
'客戶資料 [確認存檔] 檢核
CheckOP_ComsigneeData = False
msg_text = ""
If Len(Trim(txt_Storer_New.Text)) = 0 And Len(Trim(txt_Storer.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入[貨主]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入[貨主]"
   End If
End If
If Len(Trim(txt_ConsigneeKey_New.Text)) = 0 And Len(Trim(txt_ConsigneeKey.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入[客戶編號]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入[客戶編號]"
   End If
End If

If Len(RTrim(cmb_Zip_New)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入[郵遞區號]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入[郵遞區號]"
   End If
End If

If msg_text = "" Then
   CheckOP_ComsigneeData = True
Else
   msg_text = "客戶資料異常，請修正後再執行 [確認存檔]：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

End Function

