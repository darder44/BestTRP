VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_Pallet_CSVehicle_id_no 
   Caption         =   "���n�Ϩ����פJ"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   10230
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "�t�e���"
      TabPicture(0)   =   "frm_Pallet_CSVehicle_id_no.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgMain"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "���"
      TabPicture(1)   =   "frm_Pallet_CSVehicle_id_no.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgMainCT"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761087
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainCT 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8454016
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
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
   Begin VB.Frame Frame20 
      Caption         =   "���n�Ϩ����פJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9840
      Begin VB.ComboBox Cbx_Area 
         Height          =   300
         ItemData        =   "frm_Pallet_CSVehicle_id_no.frx":0038
         Left            =   2520
         List            =   "frm_Pallet_CSVehicle_id_no.frx":003A
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Cmd_Openfiles 
         BackColor       =   &H0080FFFF&
         Caption         =   "�}���ɮ�"
         Height          =   375
         Left            =   3480
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cboSheet 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   4365
      End
      Begin VB.FileListBox filLocalFile 
         Height          =   1530
         Left            =   4560
         Pattern         =   "*.xls"
         TabIndex        =   4
         ToolTipText     =   "����� ""*.xls"" �ɮ�"
         Top             =   1200
         Width           =   5190
      End
      Begin VB.DirListBox dirLocalDir 
         Height          =   1560
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Local Directory"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.DriveListBox drvLocalDrive 
         Height          =   300
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Local Drive List"
         Top             =   750
         Width           =   2040
      End
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H0080FFFF&
         Caption         =   "�}�l�פJ"
         Height          =   375
         Left            =   2400
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "�Х���ܰϰ�A�i��פJ:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�u�@��"
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
         Index           =   17
         Left            =   4560
         TabIndex        =   6
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm_Pallet_CSVehicle_id_no"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsMain As ADODB.Recordset
Private rsMainCT As ADODB.Recordset
Private rsMainItrn As ADODB.Recordset

Private Sub cboSheet_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFile.Path, 1) = "\" Then
    strFilePath = filLocalFile.Path
Else
    strFilePath = filLocalFile.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = ""

If Right(filLocalFile.Path, 1) <> "\" Then
    strFilePath = filLocalFile.Path & "\"
Else
    strFilePath = filLocalFile.Path
End If

Set rsMain = New ADODB.Recordset

Call Excel2Recordset(strFilePath & filLocalFile.FileName, cboSheet, strFieldName, rsMain)

Set dgMain.DataSource = rsMain

If rsMain Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMain
    MsgBox "���u�@��@ " & rsMain.RecordCount & "�����ӡA�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If


Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub



Private Sub Cbx_Area_Click()

If RTrim(Cbx_Area.Text) <> "" Then
    Call EnableMenu   '�}�ҩҦ��\��
    If Left(RTrim(Cbx_Area.Text), 2) = "CB" Then
        filLocalFile.Pattern = "*���ϰt�e��ƨ����פJ.xls"
    Else
        filLocalFile.Pattern = "*�n�ϰt�e��ƨ����פJ.xls"
    End If
Else
    Call DisableMenu   '�����Ҧ��\��
    
End If

End Sub

Private Sub Cmd_Openfiles_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFile.Path, 1) = "\" Then
    strFilePath = filLocalFile.Path
Else
    strFilePath = filLocalFile.Path & "\"
End If

'�إ����W�ٰ}�C
strFieldName = ""

If Right(filLocalFile.Path, 1) <> "\" Then
    strFilePath = filLocalFile.Path & "\"
Else
    strFilePath = filLocalFile.Path
End If

Set rsMain = New ADODB.Recordset

Call Excel2Recordset(strFilePath & filLocalFile.FileName, "�t�e���", strFieldName, rsMain)

'�إ����W�ٰ}�C
strFieldName = ""

Call Excel2Recordset(strFilePath & filLocalFile.FileName, "���", strFieldName, rsMainCT)

Set dgMain.DataSource = rsMain
Set dgMainCT.DataSource = rsMainCT

If rsMain Is Nothing And rsMainCT Is Nothing Then

    MsgBox "�d�L���!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMain
    SetDataGridColWidth Me.Caption, dgMainCT
    rsMain.Sort = "���u�s��"
    rsMainCT.Sort = "���u�s��"
    MsgBox "�t�e��Ʀ@: " & rsMain.RecordCount & " ���A����Ʀ@: " & rsMainCT.RecordCount & " ���A�нT�{���ƻP���e�O�_�P��l�ɮ׬۲�!!", 64, "Excel2Recordset"
    
End If


Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdImport_Click()

Dim Str_RouteNo As String, Str_CarNo As String, str_MaxRouteNo As String, str_TMSOrders As String, str_CTOrders As String
Dim str_TMSALLOrders As String, str_TRPType As String, intDriveTimes As Integer
Dim Long_CTcs As Long '���c��
Dim rsTmp As New ADODB.Recordset


On Error GoTo err_Handle

str_TMSALLOrders = ""
If Cbx_Area.Text = "CB���Ϩ����פJ" Then str_TRPType = "CB"
If Cbx_Area.Text = "SB�n�Ϩ����פJ" Then str_TRPType = "SB"

Tran_Level = 0
'===============================================�ˬd==================================================================

If (rsMain.RecordCount = 0 Or rsMain Is Nothing) And (rsMainCT.RecordCount = 0 Or rsMainCT Is Nothing) Then Exit Sub

'�t�e���
If rsMain.RecordCount = 0 Or rsMain Is Nothing Then
Else
    '�t�e��Ʀ���ơA�ˬd�t�e��Ʀ��L���`
    Str_RouteNo = "": Str_CarNo = "": str_MaxRouteNo = "": str_TMSOrders = ""
    rsMain.MoveFirst
    Do While Not rsMain.EOF
            '�@�Ӹ��u�s���u�঳�@�Ө���
            If Str_RouteNo <> Trim(rsMain("���u�s��")) Then
                '�������P�h���������M���u�s��
                Str_RouteNo = Trim(rsMain("���u�s��"))
                Str_CarNo = Trim(rsMain("����"))
            Else
                '���u�s���ۦP�A�h��������O�_�ۦP
                If Trim(rsMain("����")) <> Str_CarNo Then
                    Screen.MousePointer = vbDefault
                    MsgBox "���u�s��:" & Str_RouteNo & "  �X�{��إH�W����:" & Str_CarNo & " ; " & Trim(rsMain("����")) & "�A�нT�{�����C�פJ����", vbOKOnly + vbCritical, "�����ˬd"
                    Exit Sub
                End If
            End If
            
            '�����ˬd
            str_SQL = "select vehicle_id_no from trp09m(nolock) where vehicle_id_no = '" & RTrim(rsMain.Fields("����")) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then  '���s���ǭn��
                MsgBox "���y�D�ɤ��A�䤣�즹:" & RTrim(rsMain.Fields("����")) & " �����A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": Screen.MousePointer = 0
                Exit Sub
            End If
        rsMain.MoveNext
    Loop
    rsMain.MoveFirst
End If

'��泡���ˬd���ƶq&��������
If rsMainCT.RecordCount = 0 Or rsMainCT Is Nothing Then
Else
    '����Ʀ���ơA�ˬd����Ʀ��L���`
    Str_RouteNo = "": Str_CarNo = "": str_MaxRouteNo = "": str_TMSOrders = "": Long_CTcs = 0
    rsMainCT.MoveFirst
    Do While Not rsMainCT.EOF
            '�@�Ӹ��u�s���u�঳�@�Ө���
            If Str_RouteNo <> Trim(rsMainCT("���u�s��")) Then
                '�������P�h���������M���u�s��
                Str_RouteNo = Trim(rsMainCT("���u�s��"))
                Str_CarNo = Trim(rsMainCT("����"))
            Else
                '���u�s���ۦP�A�h��������O�_�ۦP
                If Trim(rsMainCT("����")) <> Str_CarNo Then
                    Screen.MousePointer = vbDefault
                    MsgBox "���u�s��:" & Str_RouteNo & "  �X�{��إH�W����:" & Str_CarNo & " ; " & Trim(rsMainCT("����")) & "�A�нT�{�����C�פJ����", vbOKOnly + vbCritical, "�����ˬd"
                    Exit Sub
                End If
            End If
            
            '���c�Ƥ��i�H>=���c��
            If Val(rsMainCT.Fields("���c��")) >= Val(rsMainCT.Fields("�q��c��")) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "TMS�渹:" & Val(rsMainCT.Fields("TMS�渹")) & "  �����c��:" & Val(rsMainCT.Fields("���c��")) & "�j�󵥩�q��c��: " & Val(rsMainCT.Fields("�q��c��")) & " �A�нT�{�����C�פJ����", vbOKOnly + vbCritical, "�����ˬd"
                    Exit Sub
            End If
            
            '�����ˬd
            str_SQL = "select vehicle_id_no from trp09m(nolock) where vehicle_id_no = '" & RTrim(rsMainCT.Fields("����")) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then  '���s���ǭn��
                MsgBox "���y�D�ɤ��A�䤣�즹:" & RTrim(rsMainCT.Fields("����")) & " �����A�Х���ӫ~�D�ɷs�ذӫ~��ơA�q����J�פ�!!": Screen.MousePointer = 0
                Exit Sub
            End If
            
        rsMainCT.MoveNext
    Loop
    rsMainCT.MoveFirst
End If

'===============================================��s==================================================================
Tran_Level = cn.BeginTrans
DoEvents: DoEvents

'�}�l��s�t�e����`����
If rsMain.RecordCount = 0 Or rsMain Is Nothing Then
Else
    Str_RouteNo = "": Str_CarNo = ""
    dgMain.Enabled = False
    rsMain.MoveFirst
    '�}�l��ssdn02t����

    Do While Not rsMain.EOF
        If RTrim(rsMain.Fields("����")) = "000-31" Then '�D�~������s
        Else
            str_TMSALLOrders = str_TMSALLOrders & "'" & Format(Trim(rsMain("TMS�渹")), "0000000000") & "',"
                If Str_RouteNo <> Trim(rsMain("���u�s��")) Then
                    '�������P�h���������M���u�s��
                    Str_RouteNo = Trim(rsMain("���u�s��"))
                    Str_CarNo = Trim(rsMain("����"))
                    '���o�̤j���{�s���i��s�W
                        str_SQL = "select MaxRouteNO = right(max(c_route_no),3)+1,�̤j���u�s��=max(c_route_no) from sdn01t where left(c_route_no,1) = 'N' and convert(char(8),delivery_date,112) = '" & RTrim(rsMain.Fields("��f���")) & "'"
                        Call Confirm_Recordset_Closed(tmp_Rs)
                        Call ReDim_Recordset(tmp_Rs)
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        str_MaxRouteNo = "N" & Right(RTrim(rsMain.Fields("��f���")), 6) & Format(tmp_Rs.Fields("MaxRouteNO"), "000")
                        tmp_Rs.Close
                        
                        '���ͨ���
                        str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                                  "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & RTrim(rsMain.Fields("��f���")) & "' and Vehicle_ID_No = '" & RTrim(rsMain.Fields("����")) & "'"
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
                        tmp_Rs.Close
                        
                        '�s�W�@�Ӹ��s��ƨ�sdn01t
                        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,Receiver,SDNStatus,AddUser,Drive_Times) " & _
                        "select " & _
                        "'" & RTrim(rsMain.Fields("��f���")) & "' " & _
                        ",'" & str_MaxRouteNo & "' " & _
                        ",'" & RTrim(rsMain.Fields("����")) & "' " & _
                        ",�r�p = rtrim(isnull(t9.driver,'')) " & _
                        ",�дڤH = rtrim(isnull(t9.receiver,'')) " & _
                        ",sdnstatus = 0 " & _
                        ",adduser = 'Vehicle_Update' " & _
                        ",'" & intDriveTimes & "' " & _
                        "from trp09m t9 " & _
                        "where vehicle_id_no = '" & RTrim(rsMain.Fields("����")) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                End If
                
                'sdn02t,sdn03t ��s�@�����u�s���Ψ����A
                '�s�W�����A������s�e�������M���u�s��
                '�٨S��TRP_TYPE = "CB"
'                str_SQL = "update s2 " & _
'                        "Set s2.RouteNo_old = s2.Route_no " & _
'                        ",s2.CarNo_old = s2.vehicle_id_no " & _
'                        ",s2.Route_no = '" & str_MaxRouteNo & "' " & _
'                        ",s2.vehicle_id_no = '" & Trim(rsMain("����")) & "' " & _
'                        ",s2.trp_type = '" & str_TRPType & "' " & _
'                        "from sdn02t s2 " & _
'                        "where receipt_no = '" & Format(Trim(rsMain("TMS�渹")), "0000000000") & "'"
                str_SQL = "update s2 " & _
                        "Set s2.Route_no = '" & str_MaxRouteNo & "' " & _
                        ",s2.vehicle_id_no = '" & Trim(rsMain("����")) & "' " & _
                        ",s2.trp_type = '" & str_TRPType & "' " & _
                        "from sdn02t s2 " & _
                        "where receipt_no = '" & Format(Trim(rsMain("TMS�渹")), "0000000000") & "'"
                 cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                 cn.Execute "update s3 set s3.Route_no = '" & str_MaxRouteNo & "' from sdn03t s3 where receipt_no = '" & Format(Trim(rsMain("TMS�渹")), "0000000000") & "'", RowsAffect, adExecuteNoRecords
        End If
        rsMain.MoveNext
    Loop
End If


'��泡���Aby�q��ƧǡA���s�W����檺�q��A�A�}�l��s����

If rsMainCT.RecordCount = 0 Or rsMainCT Is Nothing Then
Else
rsMainCT.MoveFirst
rsMainCT.Sort = "TMS�渹"
Str_RouteNo = "": Str_CarNo = "": str_MaxRouteNo = "": Long_CTcs = 0: str_CTOrders = "": str_TMSOrders = ""
Do While Not rsMainCT.EOF
    If RTrim(rsMainCT.Fields("���c��")) > 0 And RTrim(rsMainCT.Fields("TMS�渹")) <> str_TMSOrders Then
        str_TMSOrders = RTrim(rsMainCT.Fields("TMS�渹"))
        '���o�̤j���渹�i��s�W
            str_SQL = "select AvailNo = cast(code as integer) from codelkup where listname = 'cutordersno'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_CTOrders = "CT" & Format(tmp_Rs.Fields("AvailNo"), "00000000")
            cn.Execute "update codelkup set code = '" & Val(tmp_Rs.Fields("AvailNo")) + 1 & "' where listname = 'cutordersno'", RowsAffect, adExecuteNoRecords
            tmp_Rs.Close
            
        '�s�W�@��sdn02t��檺�q���ơC
            str_SQL = "insert sdn02t(C_ROUTE_NO, ROUTE_NO, STORERKEY, EXTERN, RECEIPT_DATE, ARRIVE_DATE, CUST_NAME, SHIP_CS, SHIP_CBM, SHIP_WT, CAR_NOTES, SDNStatus, SDN_NOTE, C_Route_Time, C_Route_Total, RECEIPT_NO, OnTimeDelivery, PODOnTime, RejectOrder, DESCRIPTION, CONFIRM_DATE, CONSIGNEEKEY, CONFIRM_USERID, CUSTSIGNDATE, RBCCode, RSCCode, CONFIRM_Notes, PRIORITY, SCHEDULEDATE, CustomerOrderkey1, Scan, SDNSendDate, CUST_Handle, TRP_Handle, Advance, INV_Handle, TRP_Cost, Sorting_Cost, Total_Cost, VEHICLE_ID_NO, ExpectReceiptOK, SdnFeedBack, InvBack, C_RECEIPT_NO, SDNBack, OTQty, OTConfirmUser, Facility, BConsigneekey, ReturnStatus) " & _
                    "select s2.C_ROUTE_NO, s2.ROUTE_NO, s2.STORERKEY, s2.EXTERN, s2.RECEIPT_DATE, s2.ARRIVE_DATE, s2.CUST_NAME, s2.SHIP_CS, s2.SHIP_CBM, s2.SHIP_WT, s2.CAR_NOTES, s2.SDNStatus, s2.SDN_NOTE, s2.C_Route_Time, s2.C_Route_Total, " & _
                    "'" & str_CTOrders & "', s2.OnTimeDelivery, s2.PODOnTime, s2.RejectOrder, s2.DESCRIPTION, s2.CONFIRM_DATE, s2.CONSIGNEEKEY, s2.CONFIRM_USERID, s2.CUSTSIGNDATE, s2.RBCCode, s2.RSCCode, s2.CONFIRM_Notes, s2.PRIORITY, s2.SCHEDULEDATE, " & _
                    "s2.CustomerOrderkey1, s2.Scan, s2.SDNSendDate, s2.CUST_Handle, s2.TRP_Handle, s2.Advance, s2.INV_Handle, s2.TRP_Cost, s2.Sorting_Cost, s2.Total_Cost, s2.VEHICLE_ID_NO, s2.ExpectReceiptOK, s2.SdnFeedBack, s2.InvBack, s2.C_RECEIPT_NO, s2.SDNBack, " & _
                    "s2.OTQty , s2.OTConfirmUser, s2.Facility, s2.BConsigneekey, s2.ReturnStatus " & _
                    "from sdn02t s2 " & _
                    "where s2.receipt_no = '" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
         '����쥻��SDN03T���ӡA���@����~���C�i����C
            str_SQL = "select S3.*,sp.casecnt " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no " & _
                    "join sdn01t s1 on s1.c_route_no = s2.c_route_no " & _
                    "join gv_skuxpack sp on sp.storerkey = s2.storerkey and s3.product_no = sp.sku " & _
                    "where s2.receipt_no = '" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.CursorLocation = adUseClient
            tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
            If tmp_Rs.RecordCount <> 0 Then
                tmp_Rs.MoveFirst
                Do While Not tmp_Rs.EOF
                    Long_CTcs = Val(Trim(rsMainCT("���c��")))
                    If Long_CTcs > 0 Then
                        If Val(tmp_Rs.Fields("ship_qty")) / Val(tmp_Rs.Fields("casecnt")) >= Val(Trim(rsMainCT("���c��"))) Then
                            '����A�hinsert�@���s������
                            str_SQL = "insert sdn03t(C_ROUTE_NO, ROUTE_NO, STORERKEY, RECEIPT_NO, SEQ_NO, SubSeq_No, EXTERN, PRODUCT_NO, SHIP_UNIT, SHIP_QTY, SIGN_QTY, WEIGHT, VOLUMN_WEIGHT, RSC_CODE, RBC_CODE, CONFIRM_DATE, DESCRIPTION, ORDER_QTY, SHIP_TIME, Responsible) " & _
                            "values('" & RTrim(tmp_Rs.Fields("C_ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("STORERKEY")) & "','" & str_CTOrders & "','" & RTrim(tmp_Rs.Fields("SEQ_NO")) & "','" & _
                             RTrim(tmp_Rs.Fields("SubSeq_No")) & "','" & RTrim(tmp_Rs.Fields("EXTERN")) & "','" & RTrim(tmp_Rs.Fields("PRODUCT_NO")) & "','" & RTrim(tmp_Rs.Fields("SHIP_UNIT")) & "','" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("���c��"))) & "','0" & _
                            "','0','0','" & RTrim(tmp_Rs.Fields("RSC_CODE")) & "','" & RTrim(tmp_Rs.Fields("RBC_CODE")) & "','" & RTrim(tmp_Rs.Fields("CONFIRM_DATE")) & "','" & RTrim(tmp_Rs.Fields("DESCRIPTION")) & "','" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("���c��"))) & "','" & RTrim(tmp_Rs.Fields("SHIP_TIME")) & "','" & RTrim(tmp_Rs.Fields("Responsible")) & "')"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                            Long_CTcs = Long_CTcs - Val(Trim(rsMainCT("���c��")))
                            '��ssdn03t���ƶq
                            str_SQL = "update sdn03t set ship_qty = ship_qty - '" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("���c��"))) & "',order_qty = order_qty - '" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("���c��"))) & "' from sdn03t  where receipt_no = '" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "' and seq_no = '" & RTrim(tmp_Rs.Fields("seq_no")) & "'"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        Else
                            '������A�����Nsdn03t��line insert �L�h
                            str_SQL = "insert sdn03t(C_ROUTE_NO, ROUTE_NO, STORERKEY, RECEIPT_NO, SEQ_NO, SubSeq_No, EXTERN, PRODUCT_NO, SHIP_UNIT, SHIP_QTY, SIGN_QTY, WEIGHT, VOLUMN_WEIGHT, RSC_CODE, RBC_CODE, CONFIRM_DATE, DESCRIPTION, ORDER_QTY, SHIP_TIME, Responsible) " & _
                            "values('" & RTrim(tmp_Rs.Fields("C_ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("STORERKEY")) & "', '" & str_CTOrders & "', '" & RTrim(tmp_Rs.Fields("SEQ_NO")) & "', '" & RTrim(tmp_Rs.Fields("SubSeq_No")) & "', '" & RTrim(tmp_Rs.Fields("EXTERN")) & _
                            "', '" & RTrim(tmp_Rs.Fields("PRODUCT_NO")) & "', '" & RTrim(tmp_Rs.Fields("SHIP_UNIT")) & "', '" & RTrim(tmp_Rs.Fields("SHIP_QTY")) & "', '" & RTrim(tmp_Rs.Fields("SIGN_QTY")) & "', '" & RTrim(tmp_Rs.Fields("Weight")) & "', '" & RTrim(tmp_Rs.Fields("VOLUMN_WEIGHT")) & "', '" & RTrim(tmp_Rs.Fields("RSC_CODE")) & "', '" & RTrim(tmp_Rs.Fields("RBC_CODE")) & _
                            "', '" & RTrim(tmp_Rs.Fields("CONFIRM_DATE")) & "', '" & RTrim(tmp_Rs.Fields("Description")) & "', '" & Val(RTrim(tmp_Rs.Fields("ORDER_QTY"))) & "', '" & RTrim(tmp_Rs.Fields("SHIP_TIME")) & "', '" & RTrim(tmp_Rs.Fields("Responsible")) & "')"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                            Long_CTcs = Long_CTcs - Val(tmp_Rs.Fields("casecnt")) * Val(tmp_Rs.Fields("ship_qty"))
                        '��ssdn03t���ƶq
                            str_SQL = "update sdn03t set ship_qty = 0,order_qty = 0 from sdn03t where receipt_no = '" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "' and seq_no = '" & RTrim(tmp_Rs.Fields("seq_no")) & "'"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        End If
                    End If
                    tmp_Rs.MoveNext
                Loop
                tmp_Rs.Close
            End If
          '�M��sdn03t���A�X�f�q=0��line
          str_SQL = "delete s3 " & _
                    "from sdn03t s3 " & _
                    "where s3.ship_qty = 0 and s3.receipt_no in ('" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "','" & str_CTOrders & "') "
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
          '�̫᭫�s��s��ӳ渹��sdn02t,sdn03t�~�n���q�A�c�Ƹ��
          str_SQL = "Update s2 " & _
                    "set ship_CS = (select sum(sdn03t.ship_qty/sp.casecnt) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku where sdn03t.receipt_no = s2.receipt_no), " & _
                    "ship_CBM =  (select sum(sdn03t.ship_qty*sp.stdcube) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku  where sdn03t.receipt_no = s2.receipt_no), " & _
                    "ship_WT = (select sum( sdn03t.ship_qty*sp.stdgrosswgt)from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku  where sdn03t.receipt_no = s2.receipt_no) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no " & _
                    "where s2.receipt_no in ('" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "','" & str_CTOrders & "') "
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
          str_SQL = "Update s3 " & _
                    "set s3.weight = (select sum(case when isnull(sp.casecnt,0) = 0 then 0 else sdn03t.ship_qty*sp.stdgrosswgt end) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku where sdn03t.receipt_no = s2.receipt_no), " & _
                    "s3.volumn_weight =(select sum(case when isnull(sp.casecnt,0) = 0 then 0 else sdn03t.ship_qty*sp.stdcube end) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku where sdn03t.receipt_no = s2.receipt_no) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no " & _
                    "where s2.receipt_no in ('" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "','" & str_CTOrders & "') "
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            '��srecordset���q�渹�X=���᪺�q��C
            rsMainCT.Fields("TMS�渹") = str_CTOrders
    End If
    rsMainCT.MoveNext
Loop
    '�}�l��s����
    rsMainCT.MoveFirst
    rsMainCT.Sort = "���u�s��"
    Str_RouteNo = "": Str_CarNo = ""
    dgMainCT.Enabled = False
    rsMainCT.MoveFirst
    '�}�l��ssdn02t����
    Do While Not rsMainCT.EOF
        If RTrim(rsMainCT.Fields("����")) = "000-31" Then '�D�~������s
        Else
        str_TMSALLOrders = str_TMSALLOrders & "'" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "',"
                If Str_RouteNo <> Trim(rsMainCT("���u�s��")) Then
                    '�������P�h���������M���u�s��
                    Str_RouteNo = Trim(rsMainCT("���u�s��"))
                    Str_CarNo = Trim(rsMainCT("����"))
                    '���o�̤j���{�s���i��s�W
                        str_SQL = "select MaxRouteNO = right(max(c_route_no),3)+1,�̤j���u�s��=max(c_route_no) from sdn01t where left(c_route_no,1) = 'N' and convert(char(8),delivery_date,112) = '" & RTrim(rsMainCT.Fields("��f���")) & "'"
                        Call Confirm_Recordset_Closed(tmp_Rs)
                        Call ReDim_Recordset(tmp_Rs)
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        str_MaxRouteNo = "N" & Right(RTrim(rsMainCT.Fields("��f���")), 6) & Format(tmp_Rs.Fields("MaxRouteNO"), "000")
                        tmp_Rs.Close
                        
                        '���ͨ���
                        str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                                  "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & RTrim(rsMainCT.Fields("��f���")) & "' and Vehicle_ID_No = '" & RTrim(rsMainCT.Fields("����")) & "'"
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
                        tmp_Rs.Close
                        
                        '�s�W�@�Ӹ��s��ƨ�sdn01t
                        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,Receiver,SDNStatus,AddUser,Drive_Times) " & _
                        "select " & _
                        "'" & RTrim(rsMainCT.Fields("��f���")) & "' " & _
                        ",'" & str_MaxRouteNo & "' " & _
                        ",'" & RTrim(rsMainCT.Fields("����")) & "' " & _
                        ",�r�p = rtrim(isnull(t9.driver,'')) " & _
                        ",�дڤH = rtrim(isnull(t9.receiver,'')) " & _
                        ",sdnstatus = 0 " & _
                        ",adduser = 'Vehicle_Update' " & _
                        ",'" & intDriveTimes & "' " & _
                        "from trp09m t9 " & _
                        "where vehicle_id_no = '" & RTrim(rsMainCT.Fields("����")) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                End If
                
                'sdn02t,sdn03t ��s�@�����u�s���Ψ����A
                '�s�W�����A������s�e�������M���u�s��
                '�٨S��TRP_TYPE = "CB"
                str_SQL = "update s2 " & _
                        "Set s2.Route_no = '" & str_MaxRouteNo & "' " & _
                        ",s2.vehicle_id_no = '" & Trim(rsMainCT("����")) & "' " & _
                        ",s2.trp_type = '" & str_TRPType & "'" & _
                        "from sdn02t s2 " & _
                        "where receipt_no = '" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "'"
                 cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                 cn.Execute "update s3 set s3.Route_no = '" & str_MaxRouteNo & "' from sdn03t s3 where receipt_no = '" & Format(Trim(rsMainCT("TMS�渹")), "0000000000") & "'", RowsAffect, adExecuteNoRecords
                 

        End If

        rsMainCT.MoveNext
    Loop
End If

    str_TMSALLOrders = Mid(str_TMSALLOrders, 1, Len(str_TMSALLOrders) - 1)
    cn.CommitTrans: Tran_Level = 0
    '���X�o����s��Sdn���
    str_SQL = "select " & _
            "�G�����u�s�� = RTrim(s2.c_route_no) " & _
            ",���u�s��=rtrim(s2.route_no) " & _
            ",����=rtrim(s2.vehicle_id_no) " & _
            ",TMS�渹=rtrim(s2.receipt_no) " & _
            ",�q�渹�X=rtrim(s2.extern) " & _
            ",�Ȥ�W�� = rtrim(s2.cust_name) " & _
            ",�~�� = rtrim(s3.product_no) " & _
            ",�c�� = rtrim(s2.ship_cs) " & _
            ",���n = rtrim(s2.ship_CBM) " & _
            ",���q =  rtrim(s2.ship_WT) " & _
            ",�Х� = rtrim(s2.trp_type) " & _
            ",�q��q = sum(s3.order_qty) " & _
            ",�X�f�q = sum(s3.ship_qty) " & _
            "from sdn02t s2 (nolock) join sdn03t s3 (nolock) on s2.receipt_no = s3.receipt_no " & _
            "where s2.receipt_no in (" & str_TMSALLOrders & ") " & _
            "group by  RTrim(s2.c_route_no),rtrim(s2.route_no) ,rtrim(s2.vehicle_id_no) ,rtrim(s2.routeno_old) ,rtrim(s2.carno_old) ,rtrim(s2.receipt_no),rtrim(s2.extern) ,rtrim(s2.cust_name) ,rtrim(s3.product_no),rtrim(s2.ship_cs), rtrim(s2.ship_CBM), rtrim(s2.ship_WT),rtrim(s2.trp_type)  " & _
            "order by RTrim(s2.c_route_no) "

    Call Confirm_Recordset_Closed(tmp_Rs)
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
    Recordset2Excel "���n�ܨ�����s", tmp_Rs
    '��X���EXCEL
    Set MyXlsApp = Nothing
    tmp_Rs.Close
    dgMain.Enabled = True
    msg_text = "���n�Ϩ����פJ����"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    
    
    Set dgMain.DataSource = Nothing
    Set dgMainCT.DataSource = Nothing
    
    Cbx_Area.Text = Cbx_Area.List(0)
    Call DisableMenu
    
    Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
    Set dgMain.DataSource = Nothing
    Set dgMainCT.DataSource = Nothing
End Sub





Private Sub dirLocalDir_Change()
    filLocalFile.Path = dirLocalDir.Path
End Sub

Private Sub drvLocalDrive_Change()
On Error GoTo DriveError
dirLocalDir.Path = drvLocalDrive.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub

Private Sub filLocalFile_Click()

On Error GoTo err_Handle
Set rsMain = Nothing: Set dgMain.DataSource = rsMain
Set rsMainCT = Nothing: Set dgMainCT.DataSource = rsMainCT
Dim strFilePath As String

'�T�{���|�O�_�a"\"
If Right(filLocalFile.Path, 1) = "\" Then
    strFilePath = filLocalFile.Path
Else
    strFilePath = filLocalFile.Path & "\"
End If

If Dir(strFilePath & filLocalFile.FileName) = "" Then: filLocalFile.Refresh: Exit Sub

cboSheet.Clear

If UCase(mySplit(filLocalFile.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFile.FileName)
    MyXlsApp.DisplayAlerts = False

    '�C�X�Ҧ��u�@��
    blDo = False
    cboSheet.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        
        cboSheet.AddItem MyXlsApp.Sheets(i).Name
  
        '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheet.ListIndex = -1

    '���ǮɭԨϥ�Microsoft.Jet.OLEDB.4.0��Ū��XLS�ASheet�������s�R�W�s�ɤ~�ॿ�T���
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheet.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "���n�Ϩ����פJ")
End Sub


Private Sub Form_Load()
Me.Height = 8000: Me.Width = 10500
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

Call DisableMenu   '�����Ҧ��\��
Cbx_Area.AddItem "    "
Cbx_Area.AddItem "CB���Ϩ����פJ"
Cbx_Area.AddItem "SB�n�Ϩ����פJ"


End Sub

Public Function EnableMenu()
'���}�Ҧ�Menu
drvLocalDrive.Enabled = True
dirLocalDir.Enabled = True
cmdImport.Enabled = True
Cmd_Openfiles.Enabled = True
SSTab1.Enabled = True
End Function

Public Function DisableMenu()
'�����Ҧ�Menu
drvLocalDrive.Enabled = False
dirLocalDir.Enabled = False
cmdImport.Enabled = False
Cmd_Openfiles.Enabled = False
SSTab1.Enabled = False
End Function

