VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_TKARList 
   Caption         =   "�����b�ک��Ӫ�"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
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
   ScaleHeight     =   10950
   ScaleWidth      =   20250
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
      StartOfWeek     =   135593985
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
      Width           =   8295
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   17
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   16
         Top             =   960
         Width           =   1485
      End
      Begin VB.CommandButton cmdSaveToText 
         BackColor       =   &H00C0E0FF&
         Caption         =   "���r��"
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
         Picture         =   "frm_Report_TKARList.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   8
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   8
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
         Left            =   5880
         Picture         =   "frm_Report_TKARList.frx":030A
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
         Left            =   7080
         Picture         =   "frm_Report_TKARList.frx":1604
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
         Left            =   7080
         Picture         =   "frm_Report_TKARList.frx":2B216
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
         Left            =   4680
         Picture         =   "frm_Report_TKARList.frx":2B528
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   240
         Width           =   1065
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
         Left            =   2640
         TabIndex        =   19
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
         TabIndex        =   18
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
         Caption         =   "������"
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
      Top             =   10680
      Width           =   20250
      _ExtentX        =   35719
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
            Object.Width           =   29078
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
Attribute VB_Name = "frm_Report_TKARList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long, strDeliveryDateS As String, strDeliveryDateE As String

Private Sub cmd2Excel_Click()

Call WriteOut_RunLog("1/16.��X�p�O���Ӹ��")
Recordset2Excel "LTKK01�����b�ک��Ӫ�", rsMain
If rsMain Is Nothing Then Call Unload_RunLogForm: Exit Sub

'..�b���s��EXCEL
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String

With MyXlsApp: .Visible = False

If RTrim(Combo1) = "LTKK01" Then
    cn.Execute "if object_id ('tempdb..##LTKK01ARList') is not null drop table ##LTKK01ARList exec gs_LTKK01ARList '" & strDeliveryDateS & "' , '" & strDeliveryDateE & "' ", RowsAffect, adExecuteNoRecords
    
    Dim rsTmp As New ADODB.Recordset

'�����

    '�M��u�@��
    strSheet = "�����"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�WDATA�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
        
'    .Sheets.Add: .ActiveSheet.Name = "�|�p�ХI�ڸ��"
    str_SQL = "select * from gv_" & Combo1.Text & "Charge where 1 = 1 " & "and ���f��� between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' order by �д����O,�Ǹ�,���f���,���� "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("2/16.��X�������")
    Call OffLineRecordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp

    rsTmp.Close
    
'�M���βz�f
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�M���βz�f"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "SELECT �U���� = cast(RECEIPT_DATE as datetime),���f��� = cast(ARRIVE_DATE as datetime) " & _
            ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",���� = C_VEHICLE_ID_NO ,�W�I = areaend ,�Ȥ�渹 = orderkey,���a�N�X = SHIPTO " & _
            ",�Ȥ�W�� = FULL_NAME ,�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
            ",�γ~�O = NOTES1  " & _
            ",�ϧO = NOTES2 " & _
            ",�X�f�c�� = ship_cs " & _
            ",�ƶq = chargeqty " & _
            ",��� = uom " & _
            ",��������� = FULL_KG " & _
            ",�t�e�O��� = receivable " & _
            ",�t�e�O�`�� = sumreceivable " & _
            ",�z�f�O��� = SortingAR " & _
            ",�z�f�O�`�� = SUMSortingAR " & _
            ",���u�s�� = route_no " & _
            ",�q���O = channel " & _
            ",�a�}�O���� = short_name ,�Ƶ� = note " & _
            "from ##LTKK01ARList " & _
            "where priority <> 'R' " & _
            "and costkind <> '�쨮�h�^' and note like ('�M��%') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("3/16.��X�M���B�O���")
Call Replication_Recordset(tmp_Rs, rsTmp)

'�g�J���D�C
k = 65: j = 1: strCol = ""
For i = 0 To rsTmp.Fields.Count - 1
    l = i Mod 26
    .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
    '���W�L26
    If Chr(65 + l) = "Z" Then
        If strCol = "" Then
            strCol = "A"
        Else
            strCol = Chr(Asc(strCol) + 1)
        End If
    End If
Next i

.Range("A2").CopyFromRecordset rsTmp

rsTmp.Close

'�M���B�O���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�M���B�O���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�M���B�O',���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�t�e�O�`�� = sumreceivable " & _
                ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and note like ('�M��%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("4/16.��X�M���B�O���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'�M���z�f���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�M���z�f���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�M���z�f�O',���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�z�f�O�`�� = SUMSortingAR " & _
                ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and note like ('�M��%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("5/16.��X�M���z�f���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close

'�~�q�B�O
Screen.MousePointer = 11
'�M��u�@��
strSheet = "�~�q�B�O"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
Next

'�䤣��s�W�u�@��
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "SELECT �U���� = cast(RECEIPT_DATE as datetime),���f��� = cast(ARRIVE_DATE as datetime) " & _
            ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",���� = C_VEHICLE_ID_NO ,�W�I = areaend ,�Ȥ�渹 = orderkey,���a�N�X = SHIPTO " & _
            ",�Ȥ�W�� = FULL_NAME ,�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
            ",�γ~�O = NOTES1  " & _
            ",�ϧO = NOTES2 " & _
            ",�X�f�c�� = ship_cs " & _
            ",�ƶq = chargeqty " & _
            ",��� = uom " & _
            ",��������� = FULL_KG " & _
            ",�t�e�O�`�� = sumreceivable " & _
            ",���u�s�� = route_no " & _
            ",�q���O = channel " & _
            ",�a�}�O���� = short_name " & _
            "from ##LTKK01ARList " & _
            "where priority <> 'R' " & _
            "and costkind <> '�쨮�h�^' and rtrim(costcode) in ('000-67','002-09','002-43') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("6/16.��X�~�q�B�O���")
Call Replication_Recordset(tmp_Rs, rsTmp)

'�g�J���D�C
k = 65: j = 1: strCol = ""
For i = 0 To rsTmp.Fields.Count - 1
    l = i Mod 26
    .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
    '���W�L26
    If Chr(65 + l) = "Z" Then
        If strCol = "" Then
            strCol = "A"
        Else
            strCol = Chr(Asc(strCol) + 1)
        End If
    End If
Next i

.Range("A2").CopyFromRecordset rsTmp

rsTmp.Close

'�~�q�B�O���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�~�q�B�O���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�~�q�B�O',���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�t�e�O�`�� = sumreceivable " & _
                ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and rtrim(costcode) in ('000-67','002-09','002-43') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("7/16.��X�~�q�B�O���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'�h�f�βz�f
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�h�f�βz�f"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
'    .Sheets.Add: .ActiveSheet.Name = "�h�f�βz�f"

    str_SQL = "SELECT �U���� = cast(RECEIPT_DATE as datetime),���f��� = cast(ARRIVE_DATE as datetime) " & _
            ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",���� = C_VEHICLE_ID_NO,�_�I = areastart,�Ȥ�渹 = orderkey,���a�N�X = SHIPTO,�Ȥ�W�� = FULL_NAME,�~�� = reason " & _
            ",��] = case when priority = 'R' then '�q�����h�^' else rtrim(costkind) end,���~�O = SUSR1,�~�P�O = SUSR3,�γ~�O = NOTES1 " & _
            ",�ϧO = NOTES2,�X�f�c�� = ship_cs,�ƶq = chargeqty,��� = uom,�t�e�O��� = receivable,�t�e�O�`�� = sumreceivable,�z�f�O��� = SortingAR " & _
            ",�z�f�O�`�� = SUMSortingAR,���u�s�� = route_no,�q���O = channel,�a�}�O���� = short_name ,�Ƶ� = note " & _
            "from ##LTKK01ARList " & _
            "where (priority = 'R' or costkind = '�쨮�h�^') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("8/16.��X�h�f�βz�f���")
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'�h�f�B�O���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�h�f�B�O���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�h�f�B�O',���f��� = cast(ARRIVE_DATE as datetime) " & _
            ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",�~�� = reason ,���~�O = SUSR1,�~�P�O = SUSR3,�γ~�O = NOTES1 ,�ϧO = NOTES2,�t�e�O�`�� = sumreceivable " & _
            ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
            ",�q���O = channel,�a�}�O���� = short_name " & _
            "from ##LTKK01ARList " & _
            "where (priority = 'R' or costkind = '�쨮�h�^') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("9/16.��X�h�f�B�O���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'�h�f�z�f���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�h�f�z�f���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�h�f�z�f',���f��� = cast(ARRIVE_DATE as datetime) " & _
            ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",�~�� = reason ,���~�O = SUSR1,�~�P�O = SUSR3,�γ~�O = NOTES1 ,�ϧO = NOTES2,�z�f�O�`�� = SUMSortingAR " & _
            ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
            ",�q���O = channel,�a�}�O���� = short_name " & _
            "from ##LTKK01ARList " & _
            "where (priority = 'R' or costkind = '�쨮�h�^') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("10/16.��X�h�f�z�f���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'�t�e�βz�f
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�t�e�βz�f"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �U���� = cast(RECEIPT_DATE as datetime),���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",���� = C_VEHICLE_ID_NO ,�W�I = areaend ,�Ȥ�渹 = orderkey,���a�N�X = SHIPTO " & _
                ",�Ȥ�W�� = FULL_NAME ,�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�X�f�c�� = ship_cs " & _
                ",�ƶq = chargeqty " & _
                ",��� = uom " & _
                ",��������� = FULL_KG " & _
                ",�t�e�O��� = receivable " & _
                ",�t�e�O�`�� = sumreceivable " & _
                ",�z�f�O��� = SortingAR " & _
                ",�z�f�O�`�� = SUMSortingAR " & _
                ",���u�s�� = route_no " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name ,�Ƶ� = note " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and rtrim(costcode) not in ('000-67','002-09','002-43','Bonded') and note not like ('�M��%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("11/16.��X�t�e�βz�f���")
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'�t�e���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�t�e���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�t�e�O',���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�t�e�O�`�� = sumreceivable " & _
                ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and rtrim(costcode) not in ('000-67','002-09','002-43','Bonded') and note not like ('�M��%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("12/16.��X�t�e���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close

    '�O�|�βz�f
    Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�O�|�βz�f"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �U���� = cast(RECEIPT_DATE as datetime),���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",���� = C_VEHICLE_ID_NO ,�W�I = areaend ,�Ȥ�渹 = orderkey,���a�N�X = SHIPTO " & _
                ",�Ȥ�W�� = FULL_NAME ,�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�X�f�c�� = ship_cs " & _
                ",�ƶq = chargeqty " & _
                ",��� = uom " & _
                ",��������� = FULL_KG " & _
                ",�t�e�O��� = receivable " & _
                ",�t�e�O�`�� = sumreceivable " & _
                ",�z�f�O��� = SortingAR " & _
                ",�z�f�O�`�� = SUMSortingAR " & _
                ",���u�s�� = route_no " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name ,�Ƶ� = note " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and rtrim(costcode) = 'Bonded' and note not like ('�M��%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "

    Call Confirm_Recordset_Closed(tmp_Rs)

    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("13/16.��X�O�|�βz�f���")
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp

    rsTmp.Close

'�t�e���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�O�|���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next

    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�t�e�O',���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�t�e�O�`�� = sumreceivable " & _
                ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and costcode = 'Bonded' and note not like ('�M��%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "

    Call Confirm_Recordset_Closed(tmp_Rs)

    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("14/16.��X�O�|���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)

    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp

    rsTmp.Close
    
'�z�f���R
Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�z�f���R"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT �д����O = '�z�f�O',���f��� = cast(ARRIVE_DATE as datetime) " & _
                ",ñ���� = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",�~�� = reason ,���~�O = SUSR1 ,�~�P�O = SUSR3 " & _
                ",�γ~�O = NOTES1  " & _
                ",�ϧO = NOTES2 " & _
                ",�z�f�O�`�� = SUMSortingAR " & _
                ",�Ȥ�W�� = FULL_NAME ,���a�N�X = SHIPTO " & _
                ",�q���O = channel " & _
                ",�a�}�O���� = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '�쨮�h�^' and rtrim(costcode) not in ('000-67','002-09','002-43','Bonded') and note not like ('�M��%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("15/16.��X�z�f���R���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close


'�����I
    Screen.MousePointer = 11
    '�M��u�@��
    strSheet = "�����I"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '��w�u�@��
    Next
    
    '�䤣��s�W�u�@��
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec es_LTKK01ARP '" & txtDeliveryDateS & "','" & txtDeliveryDateE & "'"
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("16/16.��X�����I���")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '�g�J���D�C
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '���W�L26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    cn.Execute "if object_id ('tempdb..##LTKK01ARList') is not null drop table ##LTKK01ARList ", RowsAffect, adExecuteNoRecords
End If
.Visible = True: End With

Call Unload_RunLogForm
Set MyXlsApp = Nothing
Screen.MousePointer = 0

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
If Len(txtDeliveryDateS.Text) = 0 Or Len(txtDeliveryDateE.Text) = 0 Then MsgBox "�п�J�_�W����϶��I", vbOKOnly, Me.Caption: Exit Sub
strDeliveryDateS = txtDeliveryDateS.Text: strDeliveryDateE = txtDeliveryDateE.Text
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Dim chc_Orderdate As String, chc_DeliveryDate As String
    
'�q����
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and YMD between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and YMD = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and YMD = '" & txtOrderDateE.Text & "' "
End If

'��f���
chc_DeliveryDate = "and ��f�� between '" & strDeliveryDateS & "' and '" & strDeliveryDateE & "' "

str_SQL = "select * from gv_sdn05tdetail where 1 = 1 " & chc_Orderdate & chc_DeliveryDate

'�f�D
If Len(RTrim(Combo1.Text)) > 0 Then str_SQL = str_SQL & "and �f�D = '" & RTrim(Combo1.Text) & "' "

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
rsMain.Sort = "��f��,���u�s��,�f�D�渹"

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdSaveToText_Click()
Exit Sub

err_Handle:
    Close
    Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
    
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
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
Combo1.ListIndex = 0

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
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        Combo1.AddItem tmp_Rs("storerkey")
        tmp_Rs.MoveNext
    Next
    tmp_Rs.Close: Set tmp_Rs = Nothing
'    Combo1.ListIndex = 0
    Combo1.Text = "LTKK01"

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
Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

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
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
