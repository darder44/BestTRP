VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Pallet_Import 
   Caption         =   "�̪O��ƶפJ"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10680
   Begin VB.Frame fraLocalFiles 
      Caption         =   "�ɮ��`��"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10080
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
         Height          =   450
         Index           =   2
         Left            =   3840
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
         Top             =   2400
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Tab0_Import 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��  �J"
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
         Left            =   2640
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         Top             =   2400
         Width           =   1200
      End
      Begin VB.ComboBox cmb_Tab0_Storer 
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
         ItemData        =   "frm_Pallet_Import.frx":0000
         Left            =   960
         List            =   "frm_Pallet_Import.frx":000A
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   5
         Top             =   2520
         Width           =   1605
      End
      Begin VB.DriveListBox drvLocalDrive 
         Height          =   300
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Local Drive List"
         Top             =   270
         Width           =   2040
      End
      Begin VB.DirListBox dirLocalDir 
         Height          =   1560
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Local Directory"
         Top             =   720
         Width           =   4335
      End
      Begin VB.FileListBox filLocalFile 
         Height          =   2070
         Left            =   4560
         Pattern         =   "*.xls"
         TabIndex        =   1
         ToolTipText     =   "Local Files"
         Top             =   240
         Width           =   5190
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���O�G"
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
         Index           =   11
         Left            =   240
         TabIndex        =   8
         Top             =   2580
         Width           =   630
      End
   End
   Begin MSDataGridLib.DataGrid dg_Tab0_Pallet 
      Height          =   3255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5741
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
Attribute VB_Name = "frm_Pallet_Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_Excel As ADODB.Recordset         'excel
Private arStorer() As String                '�f�D
Private dbsrcFormHeight As Double           'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double            'Form �]�p�ɴ����e
Private str_SDN_Date, str_PalletNo, str_CarNo, str_Type, str_AreaStart, str_AreaEnd, str_uom, str_Cost, str_QTy, str_in, str_out, str_customer As String

Private Sub cmd_Exit_Click(Index As Integer)
    Set rs_Excel = Nothing
    '���}
    Unload Me
End Sub

Private Sub cmd_Tab0_Import_Click()
    If Len(Trim(cmb_Tab0_Storer.Text)) > 0 Then
        Select Case Trim(cmb_Tab0_Storer.Text)
            Case "B&Q�n��"
                 Call ImportB
            Case "B&Q����"
                 Call ImportC
            Case Else
                 msg_text = "�L���f�D���פJ�{��"
                 MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                 Exit Sub
        End Select
    Else
        msg_text = "�Х��I��f�D�A�פJ"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
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

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�W�U��"
End Sub

Private Sub Form_Load()
    Me.Height = 7600: Me.Width = 11000
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200
    
    'dirLocalDir.Path = "C:"
    RecievingSize = False
End Sub

Private Sub ImportB()

    '�}�l�פJ�ɮ�
    strExcelFileName = filLocalFile.Path & "\" & filLocalFile.FileName
    If Len(Trim(filLocalFile.FileName)) = 0 Then
        Exit Sub
    End If
    
    If strExcelFileName = "" Then
        '�L����ɮ�
        Exit Sub
    End If
    If FileLen(strExcelFileName) = 0 Then
        msg_text = "�ɮפj�p=0,�ɦW:" & str_file
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    On Error GoTo err_Handle
    str_file = Trim(filLocalFile.FileName)
    '�ˬd�O�_��������
'    Call Confirm_Recordset_Closed(tmp_rs)
'    str_SQL = "select * from bestroute where import_file='" & str_file & "'"
'    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_rs.EOF Then
'        msg_text = "�o���ɮ׸�Ƥw�פJ"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        Exit Sub
'    End If
    
    '�إ� Excel �����Ʈw�s��
    strExcel = "Provider=MSDASQL.1;Persist Security Info=False;Driver={Microsoft Excel Driver (*.xls)};DBQ= " & strExcelFileName
    Set cnExcel = New ADODB.Connection
    cnExcel.ConnectionString = strExcel
    cnExcel.Open
    Call ReDim_Recordset(rs_Excel)
    
    rs_Excel.CursorLocation = 3
    str_SQL = "select * from [����޲z$]"
    'rs_Excel.Open str_SQL, cnExcel, adOpenForwardOnly, adLockReadOnly      '�L�k���� Set dg_Tab0_Import.DataSource = rs_Excel
    rs_Excel.Open str_SQL, cnExcel, adOpenStatic, adLockOptimistic
    
    If rs_Excel.EOF Then
        rs_Excel.Close
        msg_text = "�d�ߵ��G�Gexcel�L���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
         cnExcel.Close
         Exit Sub
    Else
        rs_Excel.Sort = "�渹 desc"
'        Call OffLineRecordset(tmp_rs, rs_Excel)
        Set dg_Tab0_Pallet.DataSource = rs_Excel
        rs_Excel.MoveFirst
        
        Do While Not rs_Excel.EOF
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            If str_AreaStart = str_AreaEnd Then
                msg_text = "���~�T���G�_�I " & str_AreaStart & "�P���I:" & str_AreaEnd & "�ۦP"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If str_AreaStart <> "�n����B��" And str_AreaEnd <> "�n����B��" Then
                msg_text = "���~�T���G�_�I���I�n���n����B��"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If Len(Trim(rs_Excel("���"))) < 8 Or IsNull(rs_Excel("���")) Then MsgBox "�����즳�~!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("�渹"))) = 0 Or IsNull(rs_Excel("�渹")) Then MsgBox "�渹����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("���O"))) = 0 Or IsNull(rs_Excel("���O")) Then MsgBox "���O����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("�_�I"))) = 0 Or IsNull(rs_Excel("�_�I")) Then MsgBox "�_�I����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("���I"))) = 0 Or IsNull(rs_Excel("���I")) Then MsgBox "���I����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If IsNull(rs_Excel("�ƶq")) Then MsgBox "�ƶq���ର�s!", 64, "�̪O��ƶפJ": Exit Sub
            If Val(rs_Excel("�ƶq")) = 0 Then MsgBox "�ƶq���ର�s!", 64, "�̪O��ƶפJ": Exit Sub
            
            Call ReDim_Recordset(tmp_rs)
            tmp_rs.Open "select * from pallet_cst where rtrim(checkno) = '" & RTrim(rs_Excel("�渹")) & "' ", cn
            If Not tmp_rs.EOF Then MsgBox "�渹���ơA��J�פ�!", 64, "�פJ": Exit Sub
            
            rs_Excel.MoveNext
        Loop
        
        rs_Excel.MoveFirst
        int_order = 0: intLine = 0
        Tran_Level = 0
        Tran_Level = cn.BeginTrans
        
        Do While Not rs_Excel.EOF
            DoEvents: DoEvents
'            If IsNull(rs_Excel.Fields(0).Value) Then GoTo exitloop
            If str_PalletNo <> Trim(rs_Excel.Fields(1).Value) Then '�������--�P�_�q��s���w�q�O�_�n�b [������] ���s�W�@��
                str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            
            '�g�J���Y���
            str_SQL = "insert into pallet_cds(checkno,storer,carno,usertype,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & Trim(rs_Excel("�渹")) & "','BEST','" & UCase(Trim(rs_Excel("����"))) & "','','" & Trim(rs_Excel("���")) & "','�n����B��','" & User_id & "','" & Trim(rs_Excel("���")) & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
            intLine = 1
            
            End If
            
            'excel:���,�渹,����,���O,�_�I,���I,���,���,�ƶq
            '�s�W������
            str_SDN_Date = Format(rs_Excel.Fields(0).Value, "YYYYMMDD")
            str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            str_CarNo = Trim(rs_Excel.Fields(2).Value)
            str_Type = Trim(rs_Excel.Fields(3).Value)
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            str_uom = Trim(rs_Excel.Fields(6).Value)
            str_Cost = Trim(rs_Excel.Fields(7).Value)
            str_QTy = Trim(rs_Excel.Fields(8).Value)
            
            If str_AreaStart = "�n����B��" Then
                str_in = str_QTy
                str_out = 0
                str_customer = str_AreaEnd
            Else
                str_in = 0
                str_out = str_QTy
                str_customer = str_AreaStart
            End If
             
            'checkno,linenumber,storer,carno,usertype,customer,customernoSheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,keyinDate,Editdate,checkDate,AddUser,EditUser,CheckUser,KeyID
            str_SQL = "INSERT Pallet_Cst (checkno,linenumber,storer,carno,usertype,customer,chargedate,adddate,qtyin,qtyout,sortingqty,AddUser,keyindate)" & _
                     "VALUES ('" & str_PalletNo & "','" & intLine & "','Best','" & str_CarNo & "','" & str_Type & "', " & _
                     "'" & str_customer & "','" & str_SDN_Date & "','" & str_SDN_Date & "','" & str_in & "','" & str_out & "','0', " & _
                     "'�n����B��',getdate())"
                      
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_order = int_order + 1
            intLine = intLine + 1
            rs_Excel.MoveNext
        Loop
exitloop:
        cn.CommitTrans
        Tran_Level = 0
        msg_text = "�פJ����:" & int_order
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
    Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
    End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
    CreateErrorLog Me.Name & "�פJ�����-�פJ", Me.Caption, "Import_other", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault

End Sub

Private Sub ImportC()

    '�}�l�פJ�ɮ�
    strExcelFileName = filLocalFile.Path & "\" & filLocalFile.FileName
    If Len(Trim(filLocalFile.FileName)) = 0 Then
        Exit Sub
    End If
    
    If strExcelFileName = "" Then
        '�L����ɮ�
        Exit Sub
    End If
    If FileLen(strExcelFileName) = 0 Then
        msg_text = "�ɮפj�p=0,�ɦW:" & str_file
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    On Error GoTo err_Handle
    str_file = Trim(filLocalFile.FileName)
    '�ˬd�O�_��������
'    Call Confirm_Recordset_Closed(tmp_rs)
'    str_SQL = "select * from bestroute where import_file='" & str_file & "'"
'    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_rs.EOF Then
'        msg_text = "�o���ɮ׸�Ƥw�פJ"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        Exit Sub
'    End If
    
    '�إ� Excel �����Ʈw�s��
    strExcel = "Provider=MSDASQL.1;Persist Security Info=False;Driver={Microsoft Excel Driver (*.xls)};DBQ= " & strExcelFileName
    Set cnExcel = New ADODB.Connection
    cnExcel.ConnectionString = strExcel
    cnExcel.Open
    Call ReDim_Recordset(rs_Excel)
    
    rs_Excel.CursorLocation = 3
    str_SQL = "select * from [����޲z$]"
    'rs_Excel.Open str_SQL, cnExcel, adOpenForwardOnly, adLockReadOnly      '�L�k���� Set dg_Tab0_Import.DataSource = rs_Excel
    rs_Excel.Open str_SQL, cnExcel, adOpenStatic, adLockOptimistic
    
    If rs_Excel.EOF Then
        rs_Excel.Close
        msg_text = "�d�ߵ��G�Gexcel�L���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
         cnExcel.Close
         Exit Sub
    Else
        rs_Excel.Sort = "�渹 desc"
'        Call OffLineRecordset(tmp_rs, rs_Excel)
        Set dg_Tab0_Pallet.DataSource = rs_Excel
        rs_Excel.MoveFirst
        
        Do While Not rs_Excel.EOF
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            If str_AreaStart = str_AreaEnd Then
                msg_text = "���~�T���G�_�I " & str_AreaStart & "�P���I:" & str_AreaEnd & "�ۦP"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If str_AreaStart <> "������B��" And str_AreaEnd <> "������B��" Then
                msg_text = "���~�T���G�_�I���I�n��������B��"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If Len(Trim(rs_Excel("���"))) < 8 Or IsNull(rs_Excel("���")) Then MsgBox "�����즳�~!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("�渹"))) = 0 Or IsNull(rs_Excel("�渹")) Then MsgBox "�渹����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("���O"))) = 0 Or IsNull(rs_Excel("���O")) Then MsgBox "���O����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("�_�I"))) = 0 Or IsNull(rs_Excel("�_�I")) Then MsgBox "�_�I����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If Len(Trim(rs_Excel("���I"))) = 0 Or IsNull(rs_Excel("���I")) Then MsgBox "���I����ť�!", 64, "�̪O��ƶפJ": Exit Sub
            If IsNull(rs_Excel("�ƶq")) Then MsgBox "�ƶq���ର�s!", 64, "�̪O��ƶפJ": Exit Sub
            If Val(rs_Excel("�ƶq")) = 0 Then MsgBox "�ƶq���ର�s!", 64, "�̪O��ƶפJ": Exit Sub
            
            Call ReDim_Recordset(tmp_rs)
            tmp_rs.Open "select * from pallet_cst where rtrim(checkno) = '" & RTrim(rs_Excel("�渹")) & "' ", cn
            If Not tmp_rs.EOF Then MsgBox "�渹���ơA��J�פ�!", 64, "�פJ": Exit Sub
            
            rs_Excel.MoveNext
        Loop
        
        rs_Excel.MoveFirst
        int_order = 0: intLine = 0
        Tran_Level = 0
        Tran_Level = cn.BeginTrans
        
        Do While Not rs_Excel.EOF
            DoEvents: DoEvents
'            If IsNull(rs_Excel.Fields(0).Value) Then GoTo exitloop
            If str_PalletNo <> Trim(rs_Excel.Fields(1).Value) Then '�������--�P�_�q��s���w�q�O�_�n�b [������] ���s�W�@��
                str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            
            '�g�J���Y���
            str_SQL = "insert into pallet_cds(checkno,storer,carno,usertype,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & Trim(rs_Excel("�渹")) & "','BEST','" & UCase(Trim(rs_Excel("����"))) & "','','" & Trim(rs_Excel("���")) & "','������B��','" & User_id & "','" & Trim(rs_Excel("���")) & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
            intLine = 1
            
            End If
            
            'excel:���,�渹,����,���O,�_�I,���I,���,���,�ƶq
            '�s�W������
            str_SDN_Date = Format(rs_Excel.Fields(0).Value, "YYYYMMDD")
            str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            str_CarNo = Trim(rs_Excel.Fields(2).Value)
            str_Type = Trim(rs_Excel.Fields(3).Value)
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            str_uom = Trim(rs_Excel.Fields(6).Value)
            str_Cost = Trim(rs_Excel.Fields(7).Value)
            str_QTy = Trim(rs_Excel.Fields(8).Value)
            
            If str_AreaStart = "������B��" Then
                str_in = str_QTy
                str_out = 0
                str_customer = str_AreaEnd
            Else
                str_in = 0
                str_out = str_QTy
                str_customer = str_AreaStart
            End If
             
            'checkno,linenumber,storer,carno,usertype,customer,customernoSheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,keyinDate,Editdate,checkDate,AddUser,EditUser,CheckUser,KeyID
            str_SQL = "INSERT Pallet_Cst (checkno,linenumber,storer,carno,usertype,customer,chargedate,adddate,qtyin,qtyout,sortingqty,AddUser,keyindate)" & _
                     "VALUES ('" & str_PalletNo & "','" & intLine & "','Best','" & str_CarNo & "','" & str_Type & "', " & _
                     "'" & str_customer & "','" & str_SDN_Date & "','" & str_SDN_Date & "','" & str_in & "','" & str_out & "','0', " & _
                     "'������B��',getdate())"
                      
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_order = int_order + 1
            intLine = intLine + 1
            rs_Excel.MoveNext
        Loop
exitloop:
        cn.CommitTrans
        Tran_Level = 0
        msg_text = "�פJ����:" & int_order
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
    Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
    End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
    CreateErrorLog Me.Name & "�פJ�����-�פJ", Me.Caption, "Import_other", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub



