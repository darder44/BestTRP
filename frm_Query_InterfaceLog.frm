VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Query_InterfaceLog 
   Caption         =   "Interface Log"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11955
   WindowState     =   2  '�̤j��
   Begin VB.Frame Frame7 
      Appearance      =   0  '����
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11940
      Begin VB.CommandButton cmd_Tab7SaveToExcel 
         BackColor       =   &H00C0C0FF&
         Caption         =   "��Excel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   8700
         Picture         =   "frm_Query_InterfaceLog.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         Top             =   210
         Width           =   1020
      End
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "��  �}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Index           =   6
         Left            =   9720
         Picture         =   "frm_Query_InterfaceLog.frx":08CA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   210
         Width           =   1020
      End
      Begin VB.CommandButton cmd_Query7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�d��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   7680
         Picture         =   "frm_Query_InterfaceLog.frx":0D0C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   210
         Width           =   1020
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   240
      End
      Begin VB.Label Label2 
         Caption         =   "�{���N�b"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  '�a�k���
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "���۰ʧ�s"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "����i����I��d�ߧ�s"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "�|�����B�z�����~"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin TabDlg.SSTab SSTab4 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "���B�z"
      TabPicture(0)   =   "frm_Query_InterfaceLog.frx":114E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dg_QueryResult7_E"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�w�B�z"
      TabPicture(1)   =   "frm_Query_InterfaceLog.frx":116A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "dg_QueryResult7_N"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "famTab1Filter"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame famTab1Filter 
         Height          =   510
         Left            =   8400
         TabIndex        =   12
         Top             =   300
         Width           =   3375
         Begin VB.CommandButton cmdTab1Filter 
            Caption         =   "�z��"
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
            Left            =   2520
            TabIndex        =   15
            Top             =   120
            Width           =   735
         End
         Begin VB.ComboBox cboTranType 
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
            ItemData        =   "frm_Query_InterfaceLog.frx":1186
            Left            =   1200
            List            =   "frm_Query_InterfaceLog.frx":1188
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   13
            Top             =   130
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "������O�G"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Left            =   150
            TabIndex        =   14
            Top             =   200
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid dg_QueryResult7_E 
         Height          =   6195
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   10927
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8421631
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   12
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
      Begin MSDataGridLib.DataGrid dg_QueryResult7_N 
         Height          =   5715
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   10081
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   12
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '������U��
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   8145
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   12356
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   9703
            MinWidth        =   9703
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm_Query_InterfaceLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn_Self As ADODB.Connection
Private rs_disp7_E As ADODB.Recordset
Private rs_disp7_N As ADODB.Recordset
Private blInterfaceNEventEnable As Boolean  '�w�B�zInterfaceGrid Event Ĳ�o���ı���
Private blInterfaceEEventEnable As Boolean  '�w�B�zInterfaceGrid Event Ĳ�o���ı���
Private strOrderN, strOrderE As String

Private Sub cmd_Exit_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmd_Query7_Click()
On Error GoTo err_Handle

    Label5.Caption = 300     '�˼Ƭ��
    SSTab4.Tab = 0
    
    '���`���A�u��24�p�ɤw����
    str_SQL = "select " & _
                "�f�D = RTrim(Storerkey) " & _
                ",�s�W�� = convert(char,adddate,120) " & _
                ",���A = case when status = 0 then '���B�z' else '�w�B�z' end " & _
                ",����渹 = rtrim(tranNo) " & _
                ",������O = rtrim(trantype) " & _
                ",�ɮצW�� = rtrim(filename) " & _
                ",�q��i�� = isnull(ordercount,0) " & _
                ",�������O = rtrim(logtype) " & _
                ",�����T�� = rtrim(cast(logmsg as char)) " & _
                "From interfacelog " & _
                "where status = '9' and datediff(HH,convert(char,adddate,120),convert(char,getdate(),120)) < 24 " & _
                "order by �s�W�� desc"
                
    blInterfaceNEventEnable = False
    blInterfaceEEventEnable = False
    
    Call Confirm_Recordset_Closed(tmp_rs)
    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_rs.EOF Then
       Screen.MousePointer = vbDefault
       Set dg_QueryResult7_N.DataSource = Nothing
       GoTo Error
    End If
    
    Call ReDim_Recordset(rs_disp7_N)
    Call Replication_Recordset(tmp_rs, rs_disp7_N)
    tmp_rs.Close
         
    With dg_QueryResult7_N
         .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
         .HeadLines = 2                  '��ܦb DataGrid �������Ʀ�歺������r��ơC
         .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
         .RowHeight = 300                '�]�wDataGrid ������Ҧ���ƦC����
    End With
    rs_disp7_N.MoveFirst
    Set dg_QueryResult7_N.DataSource = rs_disp7_N
    With dg_QueryResult7_N
        .RowHeight = 250
        .Columns(0).Width = 500        '�Ǹ�
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 900        '
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1800        '
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 900        '
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1100        '
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 600       '
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 2500       '
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 600       '1
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 1200        '
        .Columns(8).Alignment = dbgLeft
        .Columns(9).Width = 3400        '
        .Columns(9).Alignment = dbgLeft

    End With
    
    blInterfaceNEventEnable = True
    
    'Distinct ������O
    Dim colDistinct As Collection
    Set colDistinct = New Collection
    Do Until rs_disp7_N.EOF
        AddToCollection colDistinct, rs_disp7_N.Fields("������O").Value
        rs_disp7_N.MoveNext
    Loop
    
    '�]�w������O�U�Կ��
    cboTranType.Clear
    cboTranType.AddItem ""
    
    Dim varTemp As Variant
    For Each varTemp In colDistinct
        cboTranType.AddItem varTemp
        'Debug.Print varTemp
    Next
    Set colDistinct = Nothing
    
    rs_disp7_N.MoveFirst
    SSTab4.Tab = 1
    
Error:
     str_SQL = "select " & _
                "�f�D = RTrim(Storerkey) " & _
                ",�s�W�� = convert(char,adddate,120) " & _
                ",���A = case when status = 0 then '���B�z' else '�w�B�z' end " & _
                ",����渹 = rtrim(tranNo) " & _
                ",������O = rtrim(trantype) " & _
                ",�ɮצW�� = rtrim(filename) " & _
                ",�q��i�� = isnull(ordercount,0) " & _
                ",�������O = rtrim(logtype) " & _
                ",�����T�� = logmsg " & _
                "From interfacelog " & _
                "where status = '0' " & _
                "order by �s�W�� desc"
                
    'Whitney Edit logmsg�r�Ʒ|������A�ഫ�����w���ת�char�A���e�|�Q�I�_ ",�����T�� = rtrim(cast(logmsg as char)) "
    
    Call Confirm_Recordset_Closed(tmp_rs)
    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If tmp_rs.EOF Then
       Label7.Visible = False
       Screen.MousePointer = vbDefault
       Set dg_QueryResult7_E.DataSource = Nothing
       Exit Sub
    End If
    
    Label7.Visible = True
    Call ReDim_Recordset(rs_disp7_E)
    Call Replication_Recordset(tmp_rs, rs_disp7_E)
    tmp_rs.Close
    
    With dg_QueryResult7_E
         .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
         .HeadLines = 2                  '��ܦb DataGrid �������Ʀ�歺������r��ơC
         .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
         .RowHeight = 300                '�]�wDataGrid ������Ҧ���ƦC����
    End With
    rs_disp7_E.MoveFirst
    Set dg_QueryResult7_E.DataSource = rs_disp7_E
    With dg_QueryResult7_E
        .RowHeight = 250
        .Columns(0).Width = 500        '�Ǹ�
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 900        '
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1800        '
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 900        '
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1100        '
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 600       '
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 2500       '
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 600      '1
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 600        '
        .Columns(8).Alignment = dbgLeft
        .Columns(9).Width = 10000
        .Columns(9).Alignment = dbgLeft

    End With
    
    blInterfaceEEventEnable = True
    SSTab4.Tab = 0
        
    Screen.MousePointer = vbDefault
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
    
End Sub

Private Sub cmd_Tab7SaveToExcel_Click()
'       '�d�ߵ��G>> �� EXCEL
'    If rs_disp7 Is Nothing Then Exit Sub
'    If rs_disp7.RecordCount = 0 Then Exit Sub
'    Dim ExcelTitle As String
'    Call DocStoreDirectory(strDocPath)
'
'    Dim strTranFileName As String           'Excel �ɮצW��
'    CmnDialog.DialogTitle = "��s Excel ��"
'    CmnDialog.InitDir = "c:\my documents"
'    CmnDialog.FileName = "AJ_" & Format(Now, "YYYYMMDDHHNNSS")
'    CmnDialog.Filter = "Excel�ɮ�(*.xls)|*.xls"
'    CmnDialog.FilterIndex = 1
'    CmnDialog.CancelError = True
'    'On Error Resume Next
'    CmnDialog.Flags = cdlOFNHideReadOnly    '���ð�Ū�֨����
'    CmnDialog.ShowOpen
'    If Err.Number = cdlCancel Then          '�� [�}������] ��ܤ�����A���U [����] �s
'       msg_text = "��� [����] ���s�A������ Excel ���ۦ�s��"
'       MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
'       strTranFileName = ""
'    Else
'       strTranFileName = CmnDialog.FileName
'       If Dir(strTranFileName) <> "" Then
'          Kill strTranFileName
'       End If
'    End If
'
'    On Error GoTo err_Handle
'    Screen.MousePointer = vbHourglass
'    If SaveTo_ExcelFile(strTranFileName, rs_disp7, "�վ��^��") = 1 Then
'       Screen.MousePointer = vbDefault
'       MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
'    Else
'       Screen.MousePointer = vbDefault
'       If Len(strTranFileName) > 0 Then
'          msg_text = "��s�@�~�����A�ɮצs���m�G" & strTranFileName
'          MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'       End If
'    End If
'    rs_disp7.MoveFirst
'    Exit Sub
'
'err_Handle:
'   Dim tmpString As String
'   Screen.MousePointer = vbDefault
'   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'   CreateErrorLog Me.Name & "--�վ�^�� EXCEL", Me.Caption, "cmd_Tab7SavetoExcel_Click", tmpString
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub
 
 
Private Sub AddToCollection(objCollection As Collection, Value As String)
    On Error GoTo ErrorHandler
        objCollection.Add Value, Value
    Exit Sub
ErrorHandler:
    If Err.Number <> 457 Then '457 = key already in collection
        'something else is wrong
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub cmdTab1Filter_Click()
If rs_disp7_N Is Nothing Then Exit Sub
If rs_disp7_N.RecordCount = 0 Then Exit Sub

If Not rs_disp7_N.EOF Then
    If cboTranType.Text = "" Then
        rs_disp7_N.Filter = adFilterNone
        rs_disp7_N.Sort = "�Ǹ� ASC"
    Else
        rs_disp7_N.Filter = adFilterNone
        rs_disp7_N.Filter = "������O='" & cboTranType.Text & "'"
        rs_disp7_N.Sort = "�Ǹ� ASC"
    End If
End If

StatusBar1.Panels(1).Text = "�@ " & rs_disp7_N.RecordCount & " ����ƦC�@"

Dim i As Integer

Dim intFinishedSum As Integer
intFinishedSum = 0

Dim intUpdatedSum As Integer
intUpdatedSum = 0

Do While Not rs_disp7_N.EOF
    If rs_disp7_N.Fields(9).Value = "�������" Or rs_disp7_N.Fields(9).Value = "���ɧ���" Then
        intFinishedSum = intFinishedSum + rs_disp7_N.Fields(7).Value
    End If

    If Left(rs_disp7_N.Fields(9).Value, 4) = "��s����" Then
        intUpdatedSum = intUpdatedSum + rs_disp7_N.Fields(7).Value
    End If

    rs_disp7_N.MoveNext
Loop

rs_disp7_N.MoveFirst

If (cboTranType.Text = "") Then
    StatusBar1.Panels(2).Text = ""
ElseIf (cboTranType.Text = "PC" Or cboTranType.Text = "SC" Or cboTranType.Text = "CI" Or cboTranType.Text = "CF" Or cboTranType.Text = "GR" Or cboTranType.Text = "RN") Then
    StatusBar1.Panels(2).Text = "�i�@ " & intFinishedSum & " �i " & cboTranType.Text & " �w��X�j"
Else
    StatusBar1.Panels(2).Text = "�i�@ " & intFinishedSum & " �i " & cboTranType.Text & " �w��J�j------ ��s " & intUpdatedSum & " ��"
End If

End Sub

Private Sub dg_QueryResult7_E_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�������D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_disp7_E) <> "Nothing" Then
        '�קK���� [���] ���ʧ@
        blInterfaceEEventEnable = False
        OrderFieldName = "[" & dg_QueryResult7_E.Columns(ColIndex).Caption & "]"
        If strOrderE = "ASC" Then
            strOrderE = "DESC"
            rs_disp7_E.Sort = OrderFieldName & " DESC "
        Else
            strOrderE = "ASC"
            rs_disp7_E.Sort = OrderFieldName & " ASC "
        End If
        blInterfaceEEventEnable = True
    End If
End Sub

Private Sub dg_QueryResult7_N_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�������D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_disp7_N) <> "Nothing" Then
        '�קK���� [���] ���ʧ@
        blInterfaceNEventEnable = False
        OrderFieldName = "[" & dg_QueryResult7_N.Columns(ColIndex).Caption & "]"
        If strOrderN = "ASC" Then
            strOrderN = "DESC"
            rs_disp7_N.Sort = OrderFieldName & " DESC "
        Else
            strOrderN = "ASC"
            rs_disp7_N.Sort = OrderFieldName & " ASC "
        End If
        blInterfaceNEventEnable = True
    End If
End Sub

Private Sub Form_Load()
cmd_Query7_Click
End Sub

Private Sub Form_Resize()
Const MARGIN As Single = 160
Dim wid As Single
Dim hgt As Single

' Don't bother if we're minimized.
If WindowState = vbMinimized Then Exit Sub

On Error Resume Next 'Resize Error�ɲ��L

' Add the code to resize the controls:
SSTab4.Move 0 * ScaleWidth, SSTab4.Top, 1 * ScaleWidth, 1 * ScaleHeight - SSTab4.Top - StatusBar1.Height

hgt = ScaleHeight - MARGIN - dg_QueryResult7_E.Top - SSTab4.Top - StatusBar1.Height
wid = ScaleWidth - 2 * MARGIN: If wid < 120 Then wid = 120
dg_QueryResult7_E.Move dg_QueryResult7_E.Left, dg_QueryResult7_E.Top, wid, hgt
hgt = ScaleHeight - MARGIN - dg_QueryResult7_E.Top - SSTab4.Top - famTab1Filter.Height - StatusBar1.Height
dg_QueryResult7_N.Move dg_QueryResult7_N.Left, dg_QueryResult7_N.Top, wid, hgt

'�a�k���
famTab1Filter.Move wid + dg_QueryResult7_N.Left - famTab1Filter.Width, famTab1Filter.Top, famTab1Filter.Width, famTab1Filter.Height

StatusBar1.Width = SSTab4.Width
 
End Sub

Private Sub Form_Terminate()
    '��s Menu [����]��[�w�}�����M��]
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
    Set frm_Query_InterfaceLog = Nothing
End Sub

Private Sub SSTab4_Click(PreviousTab As Integer)
If SSTab4.Tab = 1 Then
    famTab1Filter.Visible = True
    StatusBar1.Visible = True
Else
    famTab1Filter.Visible = False
    StatusBar1.Visible = False
End If
End Sub

Private Sub Timer1_Timer()

frm_Query_InterfaceLog.Caption = "Interface Log �ثe�ɶ��G" & Now()

Dim intTime As Integer
intTime = Val(Label5.Caption) - 1
    
Label5.Caption = Val(Label5.Caption) - 1

If Label5.Caption = "0" Then
    cmd_Query7_Click
Else
    If Label7.Visible = True And (intTime Mod 2) = 0 Then
        Label7.Caption = " �|�����B�z�����~"
    ElseIf Label7.Visible = True And (intTime Mod 2) = 1 Then
        Label7.Caption = "�|�����B�z�����~"
    End If
End If


End Sub
