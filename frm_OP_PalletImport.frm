VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_OP_PalletImport 
   Caption         =   "�����ƶפJ"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   4410
      _ExtentX        =   7779
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
      StartOfWeek     =   61407233
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
      TabIndex        =   4
      Top             =   2160
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   7
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
      TabIndex        =   5
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�פJ"
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
         Picture         =   "frm_OP_PalletImport.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   1
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
         Left            =   7080
         Picture         =   "frm_OP_PalletImport.frx":212A
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
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
         Picture         =   "frm_OP_PalletImport.frx":2BD3C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�}��"
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
         Picture         =   "frm_OP_PalletImport.frx":2C04E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   0
         Top             =   240
         Width           =   1065
      End
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   2760
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   8
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
            Object.Width           =   8149
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
Attribute VB_Name = "frm_OP_PalletImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmdImport_Click()
If rsMain Is Nothing Then MsgBox "�L��ƥi�ѶפJ�I", vbOKOnly + vbInformation, "�פJ": Exit Sub
If rsMain.RecordCount = 0 Then MsgBox "�L��ƥi�ѶפJ�I", vbOKOnly + vbInformation, "�פJ": Exit Sub
On Error GoTo err_Handle
Dim i As Long, k As Integer, strCheckNo As String

With rsMain
    .Filter = "��� = 'v'"
    If .EOF Then MsgBox "�L������פJ�����!": .Filter = "": Exit Sub
    .Sort = "CheckNo"
    .MoveFirst
    
    strCheckNo = Trim(rsMain("checkno"))
    
    Screen.MousePointer = 11: cmdImport.Enabled = False: dgMain.Enabled = False
    Tran_Level = cn.BeginTrans
    k = 1
    Do While Not .EOF
    
        '�s�W���Ӹ��
        str_SQL = "insert Pallet_Cst (CheckNo,LineNumber,Storer,CarNo,UserType,Customer,Customersheetno,chargedate,QtyIn,QtyOut,Notes,AddDate,keyindate,editdate,checkdate,AddUser,edituser,checkuser,keyid) " & _
                      "Values ('" & rsMain("CheckNo") & "','" & rsMain("LineNumber") & "','" & rsMain("Storer") & "','" & rsMain("CarNo") & "','" & rsMain("UserType") & "'," & _
                      "'" & rsMain("Customer") & "','" & rsMain("Customersheetno") & "','" & rsMain("chargedate") & "','" & rsMain("QtyIn") & "','" & rsMain("QtyOut") & "','" & rsMain("Notes") & "','" & Format(rsMain("AddDate"), "yyyy/mm/dd hh:mm:ss") & "','" & Format(rsMain("keyindate"), "yyyy/mm/dd hh:mm:ss") & "','" & Format(rsMain("editdate"), "yyyy/mm/dd hh:mm:ss") & "','" & Format(rsMain("checkdate"), "yyyy/mm/dd hh:mm:ss") & "','" & rsMain("AddUser") & "','" & rsMain("edituser") & "','" & rsMain("checkuser") & "','" & rsMain("keyid") & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
     
        i = i + 1 '���ӭp��+1
    
        If strCheckNo <> Trim(rsMain("checkno")) Then
            k = k + 1 '�̪O��p�� + 1
            strCheckNo = Trim(rsMain("checkno"))
        End If
        rsMain("���") = "x"
    
    .MoveNext
    Loop
    '�s�W���Y���
    str_SQL = "insert into pallet_cds " & _
    "select checkno " & _
    ", storer " & _
    ", carno " & _
    ", usertype " & _
    ", qtyin = sum(qtyin) " & _
    ", qtyout = sum(qtyout) " & _
    ", adddate " & _
    ", keyindate " & _
    ", editdate " & _
    ", checkdate " & _
    ", adduser " & _
    ", edituser " & _
    ", checkuser " & _
    "From pallet_cst " & _
    "where ltrim(rtrim(checkno)) not in (select ltrim(rtrim(checkno)) from pallet_cds) " & _
    "group by checkno , storer, carno, usertype, adddate, keyindate, editdate, checkdate, adduser, edituser, checkuser"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    cn.CommitTrans: Tran_Level = 0
    .Filter = ""

End With
MsgBox "�@��J " & k & "���̪O��A" & i & "�����Ӹ��!", vbOKOnly, Me.Caption

Screen.MousePointer = 0: cmdImport.Enabled = True: dgMain.Enabled = True
Exit Sub

err_Handle:
Screen.MousePointer = 0: cmdImport.Enabled = True: dgMain.Enabled = True
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "�� " & k & "���̪O��A" & i & "�����Ӹ�Ʀ��~�I")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With dgMain
If .DataSource Is Nothing Or rsMain.EOF Then Exit Sub
On Error GoTo err_Handle
If LastRow = Empty And .Col <> 1 Then Exit Sub

'�L��Ʃ��I��䥦���L�@�����}
If .Row = -1 Or .Col <> 1 Then Exit Sub
If rsMain("���") = "x" Then Exit Sub

If rsMain("���") <> "v" Then
    rsMain("���") = "v"
Else
    rsMain("���") = ""

End If

.Col = 2

End With

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
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
dgMain.Visible = False
With rsMain
    .MoveLast
    
    Do While Not .BOF
        If rsMain("���") <> "x" Then rsMain("���") = " "
        .MovePrevious
    Loop

End With
dgMain.Visible = True
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

StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdOpen_Click()
On Error GoTo err_Handle
Dim strLineTmp As String, i As Long, j As Long, k As Long, arrTmp, strCheckNo As String, intNO As Integer, intDouble As Integer
Screen.MousePointer = 11

With CmnDialog
    .FileName = ""
    .DialogTitle = "��ƶפJ"
    .CancelError = False
    .InitDir = "C:\"
    'ToDo: �]�w�q�ι�ܤ��������X�Ф��ݩ�
    .Filter = "��r�ɮ� (*.CSV)|*.csv"
    .ShowOpen
    
    If Len(.FileName) = 0 Then Screen.MousePointer = 0: Exit Sub
    
    '�}���ɮ�
    Open .FileName For Input As #1
 
End With

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.CursorLocation = adUseClient
tmp_rs.Open "select ��� = ' ' ,* from pallet_cst where 1 = 2 ", cn, adOpenKeyset, adLockPessimistic
Call Replication_Recordset(tmp_rs, rsMain)
tmp_rs.Close: Set tmp_rs = Nothing

With rsMain

'���X�Ҧ��渹
Dim rsTmp As New ADODB.Recordset
Call Confirm_Recordset_Closed(tmp_rs)
str_SQL = "select CheckNo = rtrim(CheckNo) from Pallet_cds "
tmp_rs.Open str_SQL, cn
Call Replication_Recordset(tmp_rs, rsTmp)
tmp_rs.Close: Set tmp_rs = Nothing
rsTmp.MoveFirst

'intNO = 1

'�פJ�ɮ�
Do While Not EOF(1)
    
    Line Input #1, strLineTmp '����Ʀ�
    If k = 0 Then '���L���D�C
        k = k + 1 '���ӭp��+1
    Else
        If Len(RTrim(strLineTmp)) > 2 Then
            arrTmp = Split(strLineTmp, ",") '������
            k = k + 1 '���ӭp��+1
            
            .AddNew
            j = 0
            For i = 0 To .Fields.Count - 1 '���L�����
            
            If i = 1 Then i = i + 1
                .Fields(i) = Trim(arrTmp(j))
                j = j + 1
            Next
            .Update
            
            If strCheckNo <> Trim(rsMain("checkno")) Then '�O�_���P�渹
               '�渹�O�_����
               rsTmp.MoveFirst
                rsTmp.Find "Checkno = '" & RTrim(rsMain("Checkno")) & "'"
                If Not rsTmp.EOF Then rsMain("���") = "x": intDouble = intDouble + 1 '�渹���ƭp��+1
                intNO = intNO + 1 '�̪O��p��+1
            Else '�P�渹
                If Not rsTmp.EOF Then rsMain("���") = "x": intDouble = intDouble + 1 '�渹���ƭp��+1
            End If
            strCheckNo = Trim(rsMain("checkno"))
        End If
    End If

Loop
    Close #1

.Sort = "�s��"

.MoveFirst

End With

With dgMain
Set dgMain.DataSource = rsMain

    .ColumnHeaders = True        '���D�����
    .RowHeight = 300
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Alignment = dbgCenter
    .Columns(3).Alignment = dbgCenter
    .Columns(10).Alignment = dbgRight
    .Columns(11).Alignment = dbgRight

End With

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True
MsgBox "�@�}�� " & intNO & " ���̪O��A" & k - 1 & " �����Ӹ�ơA" & intDouble & " �����ӳ渹����!(�����Gx)", vbOKOnly, Me.Caption

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, k & " �C " & i + 1 & " ���Ʀ��~�I")
Close #1
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub
