VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Query_PalletDetail 
   Caption         =   "�̪O���Ӭd��"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
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
   ScaleWidth      =   11190
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3240
      TabIndex        =   12
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
      StartOfWeek     =   230883329
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
      TabIndex        =   10
      Top             =   2160
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   13
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
      TabIndex        =   11
      Top             =   0
      Width           =   10455
      Begin VB.ListBox lstType 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   4440
         Sorted          =   -1  'True
         Style           =   1  '���إ]�t�֨����
         TabIndex        =   26
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  '�a�k���
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
         Height          =   330
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1680
         Width           =   765
      End
      Begin VB.TextBox txtQtyout 
         Alignment       =   1  '�a�k���
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
         Height          =   330
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1680
         Width           =   765
      End
      Begin VB.TextBox txtQtyin 
         Alignment       =   1  '�a�k���
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
         Height          =   330
         Left            =   720
         MaxLength       =   8
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1680
         Width           =   765
      End
      Begin VB.ComboBox cboCustomer 
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
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   3285
      End
      Begin VB.ComboBox cboCarno 
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
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1485
      End
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
         Left            =   720
         MaxLength       =   8
         TabIndex        =   3
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
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   4
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
         Left            =   8040
         Picture         =   "frm_Query_PalletDetail.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ComboBox cboUserType 
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
         Left            =   2880
         TabIndex        =   0
         Top             =   240
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
         Left            =   8040
         Picture         =   "frm_Query_PalletDetail.frx":030A
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
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
         Left            =   9240
         Picture         =   "frm_Query_PalletDetail.frx":1604
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
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
         Left            =   9240
         Picture         =   "frm_Query_PalletDetail.frx":2B216
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
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
         Left            =   6840
         Picture         =   "frm_Query_PalletDetail.frx":2B528
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���l"
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
         Index           =   7
         Left            =   3000
         TabIndex        =   25
         Top             =   1725
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�٤J"
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
         Index           =   6
         Left            =   1560
         TabIndex        =   23
         Top             =   1725
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ɥX"
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
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   1725
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�"
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
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "����"
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
         TabIndex        =   18
         Top             =   300
         Width           =   480
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
         Left            =   2160
         TabIndex        =   17
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���"
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
         TabIndex        =   16
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���O"
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
         Left            =   2280
         TabIndex        =   15
         Top             =   300
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   6030
      Width           =   11190
      _ExtentX        =   19738
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
            Object.Width           =   13124
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
Attribute VB_Name = "frm_Query_PalletDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel Me.Caption, rsMain

'..�b���s��EXCEL
With MyXlsApp

  
End With
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC": txtQtyin = 0: txtQtyout = 0: txtQty = 0
Dim chc_Orderdate As String, chc_DeliveryDate As String, chc_Storerkey As String, strTmp As String, i As Integer

str_SQL = "select Convert(char(8),adddate,112) as ���, CheckNo as �渹,CarNo as ����, usertype as ���O,Customer as �Ȥ�,QtyIn as '�ɥX',QtyOut as '�٤J',�Ȥ�渹 =customersheetno, " & _
          "Notes as �Ƶ� ,edituser as ���� ,editdate as ���ʤ�� from Pallet_Cst where 1 = 1 "
    
'���
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   str_SQL = str_SQL & "and convert(Char(8),adddate,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   str_SQL = str_SQL & "and convert(Char(8),adddate,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   str_SQL = str_SQL & "and convert(Char(8),adddate,112) = '" & txtDeliveryDateE.Text & "' "
End If

'���O
'If Len(RTrim(cboUserType.Text)) > 0 Then str_SQL = str_SQL & " and usertype ='" & cboUserType.Text & "' "

'���O�����
strTmp = ""
For i = 0 To lstType.ListCount - 1
    If lstType.Selected(i) Then
            strTmp = strTmp & "'" & lstType.List(i) & "',"
    End If
Next
If Len(strTmp) > 0 Then str_SQL = str_SQL & "and usertype in (" & Left(strTmp, Len(strTmp) - 1) & ") "

'�Ȥ�
If Len(RTrim(cboCustomer.Text)) > 0 Then str_SQL = str_SQL & " and Customer ='" & cboCustomer.Text & "' "
'����
If Len(RTrim(cboCarno.Text)) > 0 Then str_SQL = str_SQL & " and Carno ='" & cboCarno.Text & "' "

str_SQL = str_SQL & "order by Convert(char(8),adddate,112),CheckNo,CarNo,���O,�Ȥ� "

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "���,�渹,����,���O,�Ȥ�"

rsMain.MoveFirst
txtQtyin = 0: txtQtyout = 0
Do While Not rsMain.EOF

    txtQtyin = Val(txtQtyin) + rsMain("�ɥX")
    txtQtyout = Val(txtQtyout) + rsMain("�٤J")
    
    rsMain.MoveNext
Loop

txtQty = Val(txtQtyin) - Val(txtQtyout)

rsMain.MoveFirst

Set dgMain.DataSource = rsMain: dgMain.Visible = False

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdSaveToText_Click()

If rsMain Is Nothing Then Exit Sub
On Error GoTo err_Handle
Screen.MousePointer = 11

Dim i As Integer, strFileName As String, strFileName1 As String, strCheck As String

'���r��
If Dir("C:\LTKK01\Ship2TKK", vbDirectory) = "" Then MkDirs "C:\LTKK01\Ship2TKK"
If Dir("C:\LTKK01\Ship2TKK\Backup", vbDirectory) = "" Then MkDirs "C:\LTKK01\Ship2TKK\Backup"
strFileName = "�X�f�^��-�t��" & Format(Now, "yyyymmddhhMMss") & ".csv"
strFileName1 = "�X�f�^��-��}" & Format(Now, "yyyymmddhhMMss") & ".csv"

Open "C:\LTKK01\Ship2TKK\" & strFileName For Output As #1
Open "C:\LTKK01\Ship2TKK\" & strFileName1 For Output As #2

'����}�l
Tran_Level = cn.BeginTrans

'��}��g�J�Ĥ@�����
Print #2, "��f�渹"; ","; "��f��"; ","; "��f��"; ","; "�q�渹�X"; ","; "����"; ","; "�ƶq"; ","; "�s�y��"; ","; "B"; ","; "�a�}�O"; ","; "�Ȥ�W��"; ","; "�Ƹ�"
Dim strA As String, strB As String, strC As String, strD As String, strE As String, intF As Integer, strG As String, strH As String, strI As String, strJ As String, strK As String, strL As String, strM As String

rsMain.MoveFirst
strA = RTrim(rsMain("��f�渹"))
strB = RTrim(rsMain("��f��"))
strC = RTrim(rsMain("��f��"))
strD = RTrim(rsMain("�q�渹�X"))
strE = RTrim(rsMain("����"))
strH = RTrim(rsMain("B"))
strI = RTrim(rsMain("�a�}�O"))
strJ = RTrim(rsMain("�Ȥ�W��"))
strK = RTrim(rsMain("�Ƹ�"))
strL = RTrim(rsMain("�q��ӷ�"))
strM = RTrim(rsMain("WMS�渹"))
strCheck = RTrim(rsMain("�q�渹�X")) & RTrim(rsMain("����"))

Do While Not rsMain.EOF

    If strCheck = RTrim(rsMain("�q�渹�X")) & RTrim(rsMain("����")) Then
        '�P�渹�~���ƶq�ۥ[
        intF = intF + RTrim(rsMain("�ƶq")): strG = strG & RTrim(rsMain("�s�y��")) & ";"
    Else
        '���P�渹�~��
        '�ˬd�O�_�t�γ�
        If Len(strL) > 0 Then
            '�t�γ�
            Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH
            
        Else
            '��}��
            Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK
        
        End If
        
    '��s���w�^��
    cn.Execute "update " & strWMSDB & "..orders set yfystatus = '2' ,TrafficCop = null where orderkey = '" & strM & "' ", RowsAffect, adExecuteNoRecords
    
    '�k�s
    strA = RTrim(rsMain("��f�渹"))
    strB = RTrim(rsMain("��f��"))
    strC = RTrim(rsMain("��f��"))
    strD = RTrim(rsMain("�q�渹�X"))
    strE = RTrim(rsMain("����"))
    intF = RTrim(rsMain("�ƶq"))
    strG = RTrim(rsMain("�s�y��")) & ";"
    strH = RTrim(rsMain("B"))
    strI = RTrim(rsMain("�a�}�O"))
    strJ = RTrim(rsMain("�Ȥ�W��"))
    strK = RTrim(rsMain("�Ƹ�"))
    strL = RTrim(rsMain("�q��ӷ�"))
    strM = RTrim(rsMain("WMS�渹"))
    strCheck = RTrim(rsMain("�q�渹�X")) & RTrim(rsMain("����"))
    End If
    rsMain.MoveNext
Loop

'�g�J�̫���
'�ˬd�O�_�t�γ�
If Len(strL) > 0 Then
    '�t�γ�
    Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH
    
Else
    '��}��
    Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK

End If

'��s���w�^��
cn.Execute "update " & strWMSDB & "..orders set yfystatus = '2' ,TrafficCop = null where orderkey = '" & strM & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'�����ɮ�
Close

'�ƥ��ɮ�
FileCopy "C:\LTKK01\Ship2TKK\" & strFileName, "C:\LTKK01\Ship2TKK\Backup\" & strFileName
FileCopy "C:\LTKK01\Ship2TKK\" & strFileName1, "C:\LTKK01\Ship2TKK\Backup\" & strFileName1

Set rsMain = Nothing: Set dgMain.DataSource = Nothing
Screen.MousePointer = 0
MsgBox "�X�f�����X����!!" & vbCrLf & "C:\LTKK01\Ship2TKK\Backup\" & strFileName & vbCrLf & "C:\LTKK01\Ship2TKK\Backup\" & strFileName1, vbOKOnly, Me.Caption
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

'���O
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(usertype) from pallet_cst order by usertype", cn, adOpenKeyset, adLockPessimistic

tmp_Rs.MoveFirst

For i = 0 To tmp_Rs.RecordCount - 1
    lstType.AddItem RTrim(tmp_Rs("usertype"))
    cboUserType.AddItem RTrim(tmp_Rs("usertype"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close
'cboUserType.ListIndex = -1

'�Ȥ�
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct (customer) from pallet_cst order by customer ", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.MoveFirst
For i = 0 To tmp_Rs.RecordCount - 1
    cboCustomer.AddItem RTrim(tmp_Rs("customer"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close
cboCustomer.ListIndex = -1

'����
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(carno) from pallet_cst order by carno ", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.MoveFirst
For i = 0 To tmp_Rs.RecordCount - 1
    cboCarno.AddItem RTrim(tmp_Rs("carno"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close
cboCarno.ListIndex = 0

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

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
