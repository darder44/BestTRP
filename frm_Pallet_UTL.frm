VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_Pallet_UTL 
   Caption         =   "�̪O�޲z"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   10980
   WindowState     =   2  '�̤j��
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3360
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4440
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
      StartOfWeek     =   61603841
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.CommandButton cmd_ClearDetail 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�M�ũ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   12
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Clean 
      BackColor       =   &H00C0FFFF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   4440
      Picture         =   "frm_Pallet_UTL.frx":0000
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   10
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmd_DelDetail 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�R������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5880
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   11
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "�s�W�Ȥ�W��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   3015
      Begin VB.ComboBox cboKind 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   16
         Top             =   840
         Width           =   1515
      End
      Begin VB.CommandButton cmd_CustAdd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�s�W"
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
         Height          =   945
         Left            =   1800
         Picture         =   "frm_Pallet_UTL.frx":0312
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   17
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txt_CustAdd 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmd_Edit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�ק�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1200
      Picture         =   "frm_Pallet_UTL.frx":2184
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   7
      Top             =   2880
      Width           =   1035
   End
   Begin VB.ComboBox lst_Cust 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Query 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�d��"
      DownPicture     =   "frm_Pallet_UTL.frx":2BD96
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   2280
      Picture         =   "frm_Pallet_UTL.frx":2D518
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   8
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Save 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�s�W"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      Picture         =   "frm_Pallet_UTL.frx":2D95A
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   6
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Del 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�R��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   3360
      Picture         =   "frm_Pallet_UTL.frx":2F7CC
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   9
      Top             =   2880
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '�m�����
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   5415
      Begin VB.TextBox txtSortingPL 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   1920
         Width           =   1155
      End
      Begin VB.TextBox txtSorting 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3480
         TabIndex        =   31
         Top             =   1920
         Width           =   1035
      End
      Begin VB.ComboBox cboUserType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   2
         Text            =   "cboUserType"
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txt_CheckNo 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1440
         Width           =   2835
      End
      Begin VB.TextBox txt_CDSOut 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   1035
      End
      Begin VB.TextBox txt_CDSIn 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txt_CarNo 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox txt_Date 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "½�O��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "�z�f���q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "�渹"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  '�m�����
         Caption         =   "AddUser"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "�̪O���O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "�٤J"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "�ɥX"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  '�m�����
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_PalletDetail 
      Height          =   2640
      Left            =   5640
      TabIndex        =   13
      Top             =   300
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   4657
      _Version        =   393216
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_PalletHead 
      Height          =   2880
      Left            =   3240
      TabIndex        =   14
      Top             =   4080
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5080
      _Version        =   393216
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   1
   End
End
Attribute VB_Name = "frm_Pallet_UTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private i, j, k, m, n As Integer
Private rs_Pallet_CDS As ADODB.Recordset
Private rs_Pallet_Cst As ADODB.Recordset
Private objMvdateTarget As Object

Private Sub cmd_Clean_Click()
    Me.txt_CDSIn.Text = "0"
    Me.txt_CDSOut.Text = "0"
    Me.txt_CheckNo.Text = ""
    Me.txt_CarNo.Text = ""
    Me.txtSortingPL.Text = "0"
    Me.txtSorting.Text = "0"
    txt_Date.Text = ""
    Call clear_PalletDetail
End Sub

Private Sub cmd_ClearDetail_Click()
Dim x As Integer
For x = 2 To dg_PalletDetail.Rows
    dg_PalletDetail.Row = dg_PalletDetail.Rows - 1
    Call cmd_DelDetail_Click
Next x
    
End Sub

Private Sub cmd_CustAdd_Click()
    On Error GoTo TextError
    If Len(Trim(Me.txt_CustAdd.Text)) = 0 Then MsgBox "�п�J�Ȥ�W��!!", vbOKOnly, Me.Caption: Exit Sub
    str_SQL = "select code from CodeLkup where listname='Cust_CDS' and code='" & Replace(Trim(Me.txt_CustAdd.Text), "-", "") & cboKind.Text & "'"
    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_rs.EOF Then
        cn.BeginTrans
            str_SQL = "insert CodeLkup (LISTNAME, Code, Addwho, Editwho) " & _
            "Values ('Cust_CDS','" & Replace(Trim(Me.txt_CustAdd.Text), "-", "") & cboKind.Text & "','" & User_id & "','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
        cn.CommitTrans
        Me.lst_Cust.AddItem Replace(Trim(Me.txt_CustAdd.Text), "-", "") & cboKind.Text
        msg_text = "�s�W�Ȥ�W��OK"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Else
        msg_text = "�Ȥ�W�٭���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CustAdd.SetFocus
    End If
    tmp_rs.Close
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Sub cmd_Del_Click()
On Error GoTo TextError
    If Len(Trim(Me.txt_CheckNo.Text)) = 0 Then
        msg_text = "�L�渹"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    DelRecord = MsgBox("�T�w�R����ƶ�?", vbQuestion + vbYesNo, "�R�����")
    If DelRecord = vbYes Then
    
    '�ˬd�O�_�w�T�{
    str_SQL = "select * from pallet_cst where len(checkuser) > 0 and CheckNo= '" & Trim(Me.txt_CheckNo.Text) & "'"
    Call Confirm_Recordset_Closed(rs_Pallet_CDS)
    rs_Pallet_CDS.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If rs_Pallet_CDS.EOF = False Then MsgBox "�t�w�T�{���ӡA�L�k�R���I", vbOKOnly, Me.Caption: Exit Sub
    '�}�l�R�����
        cn.BeginTrans
            str_SQL = "delete Pallet_Cst where rtrim(checkno) = '" & Trim(Me.txt_CheckNo.Text) & "' "
            str_SQL = str_SQL & "delete Pallet_CDS where rtrim(checkno) = '" & Trim(Me.txt_CheckNo.Text) & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
        cn.CommitTrans
        Call clear_PalletDetail
        Call clear_text
        '�M��dg_PalletHead
        dg_PalletHead.Rows = 2
        dg_PalletHead.Row = 1
        For i = 0 To dg_PalletHead.Cols - 1
            dg_PalletHead.Col = i
            dg_PalletHead.Text = ""
        Next i
    End If
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Sub cmd_DelDetail_Click()
dg_PalletDetail.Col = 6
If Len(dg_PalletDetail.Text) > 0 Then MsgBox "�w�T�{��ƵL�k�R���I", vbOKOnly, Me.Caption: txtUser.SetFocus: Exit Sub
    If dg_PalletDetail.Rows > 2 Then
        dg_PalletDetail.Rows = dg_PalletDetail.Rows - 1
    End If
    
End Sub

Private Sub cmd_Edit_Click()
On Error GoTo TextError
    If Len(Trim(Me.txt_CheckNo.Text)) = 0 Then
        msg_text = "�п�J�渹!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CheckNo.SetFocus
        Exit Sub
    Else
        Me.txt_CheckNo.Text = UCase(Me.txt_CheckNo.Text)
    End If
    If myIsDate(txt_Date.Text) = False Then Exit Sub

    If Len(Trim(Me.txt_CarNo.Text)) = 0 Then
        msg_text = "�п�J����!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CarNo.SetFocus
        Exit Sub
    Else
        Me.txt_CarNo.Text = UCase(Me.txt_CarNo.Text)
    End If
    If Len(Trim(cboUserType.Text)) = 0 Then
        msg_text = "�п�J�ܮw�O!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.cboUserType.SetFocus
        Exit Sub
    End If
    If QtyCheck = False Then
        msg_text = "���Y�P�����ƶq���~�Τ���!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If CustCheck = False Then
        msg_text = "���Ӥ��Ȥ����Ƥ���!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    DelRecord = MsgBox("�T�w�n�ק��ƶ�?", vbQuestion + vbYesNo, "�ק���")
    If DelRecord = vbYes Then
    
    str_SQL = "select * from Pallet_CDS where CheckNo='" & Trim(Me.txt_CheckNo.Text) & "'"
    Call Confirm_Recordset_Closed(tmp_rs)
    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "�䤣��ӵ��渹!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CheckNo.SetFocus
        Exit Sub
    End If
    
    tmp_rs.Close
        cn.BeginTrans
            
            '��s���Y���
            str_SQL = "Update Pallet_CDS " & _
            "set CarNo = '" & Trim(Me.txt_CarNo.Text) & "' " & _
            ",Storer = 'IDS' " & _
            ",UserType = '" & Trim(Me.cboUserType.Text) & "' " & _
            ",CheckNo = '" & Trim(Me.txt_CheckNo.Text) & "' " & _
            ",QtyIn = '" & Trim(Me.txt_CDSIn.Text) & "' " & _
            ",QtyOut = '" & Trim(Me.txt_CDSOut.Text) & "' " & _
            ",Adddate = '" & Trim(Me.txt_Date.Text) & "' " & _
            ",EditDate = getdate() " & _
            ",EditUser = '" & Trim(User_id) & "' " & _
            ",SortingPL = " & Trim(txtSortingPL.Text) & _
            " ,Sorting = " & Trim(txtSorting.Text) & _
            " where CheckNo = '" & Trim(txt_CheckNo.Text) & "'"

            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '��s����
            str_SQL = "delete Pallet_Cst where checkno = '" & Trim(Me.txt_CheckNo.Text) & "' and len(rtrim(isnull(checkuser,''))) = 0"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            Dim strNow As String: strNow = Format(Now(), "yyyy/mm/dd hh:MM:ss")
            num = 1
            For i = 1 To dg_PalletDetail.Rows - 1
                dg_PalletDetail.Row = i
                dg_PalletDetail.Col = 0: str_chargedate = Trim(dg_PalletDetail.Text): If Len(Trim(dg_PalletDetail.Text)) = 0 Then str_chargedate = Trim(txt_Date.Text)
                dg_PalletDetail.Col = 1: str_qtyin = Trim(dg_PalletDetail.Text)
                dg_PalletDetail.Col = 2: str_qtyout = Trim(dg_PalletDetail.Text)
                dg_PalletDetail.Col = 3: str_cstno = Trim(dg_PalletDetail.Text)
                dg_PalletDetail.Col = 4: str_cstnum = Trim(dg_PalletDetail.Text)
                dg_PalletDetail.Col = 5: str_Note = Trim(dg_PalletDetail.Text)
                dg_PalletDetail.Col = 6: str_checkuser = Trim(dg_PalletDetail.Text)
                
                If Len(str_checkuser) > 0 Then
                    str_SQL = "update Pallet_Cst set LineNumber = '" & num & "' where checkno = '" & Trim(Me.txt_CheckNo.Text) & "' and linenumber = ( " & _
                    "select top 1 linenumber from pallet_cst where rtrim(checkno) = '" & Trim(Me.txt_CheckNo.Text) & "' and rtrim(storer) = 'IDS' and rtrim(carno) ='" & UCase(Trim(Me.txt_CarNo.Text)) & "' and rtrim(usertype) = '" & Trim(Me.cboUserType.Text) & "' and rtrim(customer) = '" & str_cstno & "' and rtrim(customersheetno) = '" & str_cstnum & "' and chargedate = '" & str_chargedate & "' and qtyin = '" & str_qtyin & "' and qtyout = '" & str_qtyout & "' and rtrim(notes) = '" & str_Note & "' and rtrim(checkuser) = '" & str_checkuser & "' )"
                Else
                    str_SQL = "insert Pallet_Cst (chargedate , CarNo,Storer,UserType,CheckNo,Customer,QtyIn,QtyOut,Notes,AddDate,AddUser,Editdate,EditUser,LineNumber,Customersheetno,checkuser) " & _
                      "Values ('" & str_chargedate & "','" & UCase(Trim(Me.txt_CarNo.Text)) & "','IDS','" & Trim(Me.cboUserType.Text) & "','" & Trim(Me.txt_CheckNo.Text) & "','" & Trim(str_cstno) & "'," & _
                      "'" & Trim(str_qtyin) & "','" & Trim(str_qtyout) & "','" & Trim(str_Note) & "','" & Trim(Me.txt_Date.Text) & "','" & Trim(User_id) & "','" & strNow & "' , '" & Trim(User_id) & "','" & num & "','" & str_cstnum & "','" & str_checkuser & "' )"
                End If
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                num = num + 1
            Next i
                '��skeyindate
                str_SQL = "update pallet_cst set keyindate = (select top 1 keyindate from pallet_cds where ltrim(rtrim(checkno)) = '" & Trim(Me.txt_CheckNo.Text) & "') where ltrim(rtrim(checkno)) = '" & Trim(Me.txt_CheckNo.Text) & "' "
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        cn.CommitTrans
        
        '�M��dg_PalletHead
        dg_PalletHead.Rows = 2
        dg_PalletHead.Row = 1
        For i = 0 To dg_PalletHead.Cols - 1
            dg_PalletHead.Col = i
            dg_PalletHead.Text = ""
        Next i
        dg_PalletHead.Row = 1
        '�[�Jdg_PalletHead
        dg_PalletHead.Col = 0
        dg_PalletHead.Text = Trim(Me.txt_CheckNo.Text)
        dg_PalletHead.Col = 1
        dg_PalletHead.Text = Trim(Me.txt_CarNo.Text)
        dg_PalletHead.Col = 2
        dg_PalletHead.Text = Trim(Me.cboUserType.Text)
        dg_PalletHead.Col = 3
        dg_PalletHead.Text = Trim(User_id): dg_PalletHead.ColAlignment(3) = flexAlignLeft
        dg_PalletHead.Col = 4
        dg_PalletHead.Text = Trim(Me.txt_CDSIn.Text)
        dg_PalletHead.Col = 5
        dg_PalletHead.Text = Trim(Me.txt_CDSOut.Text)
        dg_PalletHead.Col = 6
        dg_PalletHead.Text = Trim(Me.txt_Date.Text)
        dg_PalletHead.Col = 7
        dg_PalletHead.Text = "����"
        Me.txt_CDSIn.Text = "0"
        Me.txt_CDSOut.Text = "0"
        Me.txt_CheckNo.Text = ""
        Me.txt_CarNo.Text = ""
        Me.txtSortingPL.Text = "0"
        Me.txtSorting.Text = "0"
        txt_Date.Text = ""
        Call clear_PalletDetail
        msg_text = "�ק��Ƨ���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CarNo.SetFocus
    End If
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmd_Query_Click()
    
    On Error GoTo err_Handle
    If Len(Trim(Me.txt_CheckNo.Text)) = 0 Then
        msg_text = "�п�J�渹"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CheckNo.SetFocus
        Exit Sub
    End If
    str_SQL = "select CarNo,isnull(UserType,''),isnull(CheckNo,''),isnull(AddUser,''),isnull(QtyIn,''), " & _
              "isnull(QtyOut,''),Convert(char(8),AddDate,112),sortingpl,sorting from Pallet_CDS where CheckNo= '" & Trim(Me.txt_CheckNo.Text) & "'"
    Call Confirm_Recordset_Closed(rs_Pallet_CDS)
    rs_Pallet_CDS.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If rs_Pallet_CDS.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "�d�ߵ��G�G�L�w�s���"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Call clear_PalletDetail
       Exit Sub
    End If
    Me.txt_CarNo.Text = Trim(rs_Pallet_CDS.Fields(0))
    Me.cboUserType.Text = Trim(rs_Pallet_CDS.Fields(1))
    Me.txt_CheckNo.Text = Trim(rs_Pallet_CDS.Fields(2))
    txtUser.Text = Trim(rs_Pallet_CDS.Fields(3))
    Me.txt_CDSIn.Text = Trim(rs_Pallet_CDS.Fields(4))
    Me.txt_CDSOut.Text = Trim(rs_Pallet_CDS.Fields(5))
    Me.txt_Date.Text = Trim(rs_Pallet_CDS.Fields(6)) & ""
    Me.txtSortingPL = Trim(rs_Pallet_CDS.Fields("sortingpl")) & ""
    Me.txtSorting = Trim(rs_Pallet_CDS.Fields("sorting")) & ""
    rs_Pallet_CDS.Close
    '���Ӹ��
    str_SQL = "select convert(char(8),ChargeDate,112),QtyIn,QtyOut,isnull(Customer,''),isnull(Customersheetno,''),isnull(Notes,''),isnull(checkuser,'') from  " & _
              "Pallet_Cst where CheckNo='" & Trim(Me.txt_CheckNo.Text) & "' order by checkuser desc"
    Call Confirm_Recordset_Closed(rs_Pallet_Cst)
    rs_Pallet_Cst.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    rs_Pallet_Cst.MoveFirst
    dg_PalletDetail.Rows = 2
    j = 1
    Do While Not rs_Pallet_Cst.EOF
        i = 0
        dg_PalletDetail.Row = j
        For i = 0 To rs_Pallet_Cst.Fields.Count - 1
            dg_PalletDetail.Col = i
            dg_PalletDetail.Text = Trim(rs_Pallet_Cst.Fields(i))
        Next
        j = j + 1
        If j > 1 Then
            With dg_PalletDetail
                .Rows = .Rows + 1
            End With
        End If
    rs_Pallet_Cst.MoveNext
    Loop
    rs_Pallet_Cst.Close
    With dg_PalletDetail
        .Rows = .Rows - 1
    End With
'    'Set tmp_cmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmd_Save_Click()
On Error GoTo TextError
    If Len(Trim(Me.txt_CheckNo.Text)) = 0 Then
        msg_text = "�п�J�渹!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CheckNo.SetFocus
        Exit Sub
    Else
        Me.txt_CheckNo.Text = UCase(Me.txt_CheckNo.Text)
    End If
    If myIsDate(Trim(Me.txt_Date.Text)) = False Then Exit Sub

    If Len(Trim(Me.txt_CarNo.Text)) = 0 Then
        msg_text = "�п�J����!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CarNo.SetFocus
        Exit Sub
    Else
        Me.txt_CarNo.Text = UCase(Me.txt_CarNo.Text)
    End If
    If Len(Trim(Me.cboUserType.Text)) = 0 Then
        msg_text = "�п�J�ܮw�O!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.cboUserType.SetFocus
        Exit Sub
    End If
    If QtyCheck = False Then
        msg_text = "���Y�P�����ƶq���~�Τ���!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If CustCheck = False Then
        msg_text = "���Ӥ��Ȥ����Ƥ���!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Call Confirm_Recordset_Closed(tmp_rs)
    str_SQL = "select * from Pallet_CDS where CheckNo='" & Trim(Me.txt_CheckNo.Text) & "'"
    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "�渹����!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Me.txt_CheckNo.SetFocus
        Exit Sub
    End If
    tmp_rs.Close
    cn.BeginTrans
    str_SQL = "insert Pallet_CDS (CarNo,Storer,UserType,CheckNo,QtyIn,QtyOut,AddDate,AddUser,SortingPL,Sorting) " & _
              "Values ('" & Trim(Me.txt_CarNo.Text) & "','IDS','" & Trim(Me.cboUserType.Text) & "','" & Trim(Me.txt_CheckNo.Text) & "','" & Trim(Me.txt_CDSIn.Text) & "', " & _
              "'" & Trim(Me.txt_CDSOut.Text) & "','" & Trim(Me.txt_Date.Text) & "','" & Trim(User_id) & "','" & txtSortingPL & "','" & txtSorting & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�s�W���Ӹ��
    num = 1
    For i = 1 To dg_PalletDetail.Rows - 1
        dg_PalletDetail.Row = i
        dg_PalletDetail.Col = 0: str_chargedate = Trim(dg_PalletDetail.Text): If Len(Trim(dg_PalletDetail.Text)) = 0 Then str_chargedate = Trim(txt_Date.Text)
        dg_PalletDetail.Col = 1: str_qtyin = Trim(dg_PalletDetail.Text)
        dg_PalletDetail.Col = 2: str_qtyout = Trim(dg_PalletDetail.Text)
        dg_PalletDetail.Col = 3: str_customer = Trim(dg_PalletDetail.Text)
        dg_PalletDetail.Col = 4: str_cstnum = Trim(dg_PalletDetail.Text)
        dg_PalletDetail.Col = 5: str_Note = Trim(dg_PalletDetail.Text)
        
        str_SQL = "insert Pallet_Cst (CarNo,Storer,UserType,CheckNo,Customer,QtyIn,QtyOut,Notes,AddDate,AddUser,LineNumber,Customersheetno,chargedate) " & _
                  "Values ('" & UCase(Trim(Me.txt_CarNo.Text)) & "','IDS','" & Trim(Me.cboUserType.Text) & "','" & Trim(Me.txt_CheckNo.Text) & "','" & Trim(str_customer) & "'," & _
                  "'" & Trim(str_qtyin) & "','" & Trim(str_qtyout) & "','" & Trim(str_Note) & "','" & Trim(Me.txt_Date.Text) & "','" & User_id & "','" & num & "','" & str_cstnum & "','" & str_chargedate & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        num = num + 1
    Next i
    cn.CommitTrans
    
    'dg_PalletHead
    j = dg_PalletHead.Rows
    dg_PalletHead.Row = j - 1
    dg_PalletHead.Col = 0
    If Len(Trim(dg_PalletHead.Text)) > 0 Then    '�s�W�@row
        j = dg_PalletHead.Rows + 1
        dg_PalletHead.Rows = j
        dg_PalletHead.Row = j - 1
    End If
    dg_PalletHead.Row = j - 1
    ',CheckNo,CarNo,UserType,CheckUser,QtyIn,QtyOut,CheckDate
    dg_PalletHead.Col = 0
    dg_PalletHead.Text = Trim(Me.txt_CheckNo.Text)
    dg_PalletHead.Col = 1
    dg_PalletHead.Text = Trim(Me.txt_CarNo.Text)
    dg_PalletHead.Col = 2
    dg_PalletHead.Text = Trim(Me.cboUserType.Text)
    dg_PalletHead.Col = 3
    dg_PalletHead.Text = Trim(User_id): dg_PalletHead.ColAlignment(3) = flexAlignLeft
    dg_PalletHead.Col = 4
    dg_PalletHead.Text = Trim(Me.txt_CDSIn.Text)
    dg_PalletHead.Col = 5
    dg_PalletHead.Text = Trim(Me.txt_CDSOut.Text)
    dg_PalletHead.Col = 6
    dg_PalletHead.Text = Trim(Me.txt_Date.Text)
    dg_PalletHead.Col = 7
    dg_PalletHead.Text = "����"
    Me.txt_CDSIn.Text = ""
    Me.txt_CDSOut.Text = ""
    Me.txt_CheckNo.Text = ""
    Me.txt_CarNo.Text = ""
    Call clear_PalletDetail
    msg_text = "�s�ɧ���"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Me.txt_CarNo.SetFocus
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Sub dg_PalletDetail_Click()
Dim x As Integer
x = dg_PalletDetail.Col
'�ˬd�O�_�w�T�{
dg_PalletDetail.Col = 6
If Len(dg_PalletDetail.Text) > 0 Then MsgBox "�w�T�{��ƵL�k�ק�I", vbOKOnly, Me.Caption: txtUser.SetFocus: Exit Sub
If x = 6 Then x = x - 1
dg_PalletDetail.Col = x

Text1.Visible = False

    If dg_PalletDetail.Col = 3 Then NextPosition1 dg_PalletDetail.Row, dg_PalletDetail.Col: Exit Sub '���ʿ����
    If dg_PalletDetail.Col = 0 Then '��ܤ�������
        mvDate.Top = dg_PalletDetail.Top '+ dg_PalletDetail.RowPos(r)     '�W��
        mvDate.Left = dg_PalletDetail.Left + dg_PalletDetail.ColWidth(0) + 30 '�k��
        mvDate.Visible = True
        mvDate.Value = Now
        Set objMvdateTarget = dg_PalletDetail
    End If
    
    NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col       '���ʤ�r���Text1
    
End Sub

Private Sub dg_PalletDetail_Scroll()
    On Error GoTo TextError
        Text1.Visible = False
        lst_Cust.Visible = False
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Sub dg_PalletHead_Click()
    If dg_PalletHead.Col = 7 Then
        dg_PalletHead.Col = 0
        i = dg_PalletHead.Row
        str_checkno = Trim(dg_PalletHead.Text)
        
        '���Y
        str_SQL = "select CarNo,isnull(UserType,''),isnull(CheckNo,''),isnull(AddUser,''),isnull(QtyIn,''), " & _
              "isnull(QtyOut,''),Convert(char(8),addDate,112) from Pallet_CDS where CheckNo= '" & str_checkno & "'"
        Call Confirm_Recordset_Closed(rs_Pallet_CDS)
        rs_Pallet_CDS.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If rs_Pallet_CDS.EOF Then
           Screen.MousePointer = vbDefault
           msg_text = "�d�ߵ��G�G�L�w�s���"
           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
           Exit Sub
        End If
        Me.txt_CarNo.Text = Trim(rs_Pallet_CDS.Fields(0))
        Me.cboUserType.Text = Trim(rs_Pallet_CDS.Fields(1))
        Me.txt_CheckNo.Text = Trim(rs_Pallet_CDS.Fields(2))
        Me.txtUser.Text = Trim(rs_Pallet_CDS.Fields(3))
        Me.txt_CDSIn.Text = Trim(rs_Pallet_CDS.Fields(4))
        Me.txt_CDSOut.Text = Trim(rs_Pallet_CDS.Fields(5))
        Me.txt_Date.Text = Trim(rs_Pallet_CDS.Fields(6))
        rs_Pallet_CDS.Close
        
        '����
        str_SQL = "select convert(char(8),chargedate,112),QtyIn,QtyOut,isnull(Customer,''),isnull(Customersheetno,''),isnull(Notes,''),isnull(checkuser,'') from  " & _
                  "Pallet_Cst where CheckNo='" & str_checkno & "' order by checkuser desc"
        Call Confirm_Recordset_Closed(rs_Pallet_Cst)
        rs_Pallet_Cst.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        rs_Pallet_Cst.MoveFirst
        dg_PalletDetail.Rows = 2
        j = 1
        Do While Not rs_Pallet_Cst.EOF
            i = 0
            dg_PalletDetail.Row = j
            For i = 0 To rs_Pallet_Cst.Fields.Count - 1
                dg_PalletDetail.Col = i
                dg_PalletDetail.Text = Trim(rs_Pallet_Cst.Fields(i))
            Next
            
            j = j + 1
            If j > 1 Then
                With dg_PalletDetail
                    .Rows = .Rows + 1
                End With
            End If
            
        rs_Pallet_Cst.MoveNext
        Loop
        
        rs_Pallet_Cst.Close
        With dg_PalletDetail
            .Rows = .Rows - 1
        End With
        Screen.MousePointer = vbDefault
    End If
        
End Sub

Private Sub Form_Activate()
    '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
    Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
    msg_title = "�̪O�޲z"
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''�d�I��Ӫ����L����ƥ�
''�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������������
'If KeyCode = vbKeyEscape Then
'   mvDate.Visible = False
'End If
'End Sub

Private Sub Form_Load()
    Me.Height = 7600: Me.Width = 11500
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200
    mvDate.Visible = False
    
'�T�{�H
    txtUser.Text = ""
    Frame1.Caption = "�ϥΪ̡J" & User_id
    If blAdmin = True Then cmd_CustAdd.Enabled = True: cboUserType.Enabled = True

'�g�P�N�e
    cboKind.AddItem ""
    cboKind.AddItem "-�g�P"
    cboKind.AddItem "-�N�e"
'    cboKind.AddItem "-����"
'    cboKind.AddItem "-���q"
'    cboKind.AddItem "-�Q��"
'    cboKind.AddItem "-�j�O"
'    cboKind.AddItem "-���a"
    cboKind.ListIndex = 0
    
'�ܮw�O
    '���Ѽ�
    Dim objIni As vbIniFile, arrTmp
    Set objIni = New vbIniFile
    objIni.FileName = striniFileName_FullPath
    
    arrTmp = Split(objIni.ReadData("OPTION", "WAREHOUSE", "0"), ";")
    
    For i = 0 To UBound(arrTmp)
        cboUserType.AddItem arrTmp(i)
    Next
    cboUserType.ListIndex = 0
    
'�Ȥ�W��
    Call Confirm_Recordset_Closed(tmp_rs)
    str_SQL = "select code from CodeLkup where listname='Cust_CDS'"
    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_rs.EOF Then
       Do While Not tmp_rs.EOF
          lst_Cust.AddItem Trim(tmp_rs.Fields("code"))
          tmp_rs.MoveNext
       Loop
    End If
    tmp_rs.Close
    With dg_PalletDetail
         .FixedRows = 1
         '�]�w���\��C���
         .AllowBigSelection = True
         '�]�w�C�����e��
         .ColWidth(0) = 1200  '�p�����
         .ColWidth(1) = 700   '�ɥX
         .ColWidth(2) = 700   '�٤J
         .ColWidth(3) = 2200  '�Ȥ�
         .ColWidth(4) = 1200  '�渹
         .ColWidth(5) = 1800  '�Ƶ�
         .ColWidth(6) = 1500  '�T�{
         '�]�w�C�����D
         .Row = 0
         .Col = 0: .Text = "�������"
         .Col = 1: .Text = "�ɥX"
         .Col = 2: .Text = "�٤J"
         .Col = 3: .Text = "�Ȥ�"
         .Col = 4: .Text = "�渹"
         .Col = 5: .Text = "�Ƶ�"
         .Col = 6: .Text = "�T�{"
         '�]�w�C����r���
         
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1
             .CellAlignment = flexAlignCenterCenter
         Next sub_var1
    End With
    'CheckNo,CarNo,UserType,CheckUser,QtyIn,QtyOut,CheckDate
    With dg_PalletHead
         .FixedRows = 1
         '�]�w���\��C���
         .AllowBigSelection = True
         '�]�w�C�����e��
         .ColWidth(0) = 1500
         .ColWidth(1) = 1000
         .ColWidth(2) = 1000
         .ColWidth(3) = 1500
         .ColWidth(4) = 600
         .ColWidth(5) = 600
         .ColWidth(6) = 1200
         .ColWidth(7) = 600
         '�]�w�C�����D
         .Row = 0
         .Col = 0: .Text = "�渹"
         .Col = 1: .Text = "����"
         .Col = 2: .Text = "�ܮw�O"
         .Col = 3: .Text = "AddUser"
         .Col = 4: .Text = "�ɥX"
         .Col = 5: .Text = "�٤J"
         .Col = 6: .Text = "���"
         .Col = 7: .Text = ""
         
         '�]�w�C����r���
'         .ColAlignment(0) = 4
'         .ColAlignment(1) = flexAlignLeft
'         .ColAlignment(2) = flexAlignLeft
'         .ColAlignment(3) = flexAlignLeft
'         .ColAlignment(4) = flexAlignRight
'         .ColAlignment(5) = flexAlignRight
'         .ColAlignment(6) = flexAlignLeft
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1
             .CellAlignment = flexAlignCenterCenter
         Next sub_var1
    End With
End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleWidth < dg_PalletDetail.Width + dg_PalletDetail.Left Then

Exit Sub

Else
dg_PalletDetail.Width = Me.ScaleWidth - dg_PalletDetail.Left
dg_PalletHead.Width = Me.ScaleWidth - dg_PalletHead.Left
End If
End Sub

Private Sub lst_Cust_Change()
On Error GoTo TextError
    dg_PalletDetail.Text = Me.lst_Cust.Text    '�N��r������e�g�ܹ����x�s��
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Sub lst_Cust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyHome Then
        dg_PalletDetail.Col = dg_PalletDetail.Col - 1
        NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
    End If
End Sub

Private Sub lst_Cust_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'        dg_PalletDetail.Col = dg_PalletDetail.Col + 1
'        KeyAscii = 0
'        NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
    If dg_PalletDetail.Col = 3 Then Call AddRow

End If
End Sub

Private Sub lst_Cust_LostFocus()
    Me.lst_Cust.Visible = False
End Sub

Private Sub lst_Cust_Click()
    If dg_PalletDetail.Col = 3 Then
        dg_PalletDetail.Text = Me.lst_Cust.Text
    End If
End Sub

Private Sub lst_User_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyHome Then
        Me.cboUserType.SetFocus
    End If
End Sub

Private Sub lst_User_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.txt_CDSIn.SetFocus
    End If
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

    objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
    mvDate.Visible = False

End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt_CarNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyHome Then
        cboUserType.SetFocus
    End If
End Sub

Private Sub txt_CarNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.cboUserType.SetFocus
    End If
End Sub

Private Sub txt_CDSIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txt_CDSIn.Text)) > 0 Then
            If IsNumeric(Trim(Me.txt_CDSIn.Text)) = False Then
                msg_text = "�ƶq����,�Э��s��J"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                Me.txt_CDSIn.SetFocus
                Exit Sub
            End If
        End If
        Me.txt_CDSOut.SetFocus
    End If
End Sub


Private Sub txt_CDSIn_LostFocus()
    If Len(Trim(Me.txt_CDSIn.Text)) > 0 Then
        If IsNumeric(Trim(Me.txt_CDSIn.Text)) = False Then
            msg_text = "�ƶq����,�Э��s��J"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Me.txt_CDSIn.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txt_CDSOut_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyHome Then
        Me.txt_CDSIn.SetFocus
    End If
End Sub

Private Sub txt_CDSOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txt_CDSOut.Text)) > 0 Then
            If IsNumeric(Trim(Me.txt_CDSOut.Text)) = False Then
                msg_text = "�ƶq����,�Э��s��J"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                Me.txt_CDSOut.SetFocus
                Exit Sub
            End If
        End If
        Me.txt_CheckNo.SetFocus
    End If
End Sub

Private Sub txt_CDSOut_LostFocus()
    If Len(Trim(Me.txt_CDSOut.Text)) > 0 Then
        If IsNumeric(Trim(Me.txt_CDSOut.Text)) = False Then
            msg_text = "�ƶq����,�Э��s��J"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Me.txt_CDSOut.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txt_CheckNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyHome Then
        Me.txt_CDSOut.SetFocus
    End If
End Sub

Private Sub txt_CheckNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        dg_PalletDetail.Row = 1
        dg_PalletDetail.Col = 0
        KeyAscii = 0
        NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col    '���U�Ӥ���m
    End If
End Sub

Private Sub txt_Date_Click()
    mvDate.Top = Frame1.Top ' + txt_Date.Top + txt_Date.Height
    mvDate.Left = Frame1.Left + txt_Date.Left + txt_Date.Width
    mvDate.Visible = True
    mvDate.Value = Now
    Set objMvdateTarget = txt_Date
End Sub

Private Sub txt_Date_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub

'Private Sub txt_UserTypeCode_Change()
'    If Len(Trim(Me.txt_UserTypeCode.Text)) = "0" Then Exit Sub
'    If Trim(Me.txt_UserTypeCode.Text) = "1" Then
'        Me.txt_UserType.Text = "���Z"
'    ElseIf Trim(Me.txt_UserTypeCode.Text) = "2" Then
'        Me.txt_UserType.Text = "���Z"
'    ElseIf Trim(Me.txt_UserTypeCode.Text) = "3" Then
'        Me.txt_UserType.Text = "���Z"
'    Else
'        Me.txt_UserType.Text = ""
'        Me.txt_UserTypeCode.Text = ""
'    End If
'End Sub

Private Sub cboUserType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyHome Then
        Me.txt_CarNo.SetFocus
    End If
End Sub

Private Sub cboUserType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt_CDSIn.SetFocus
    End If
End Sub

Public Sub NextPosition(ByVal r As Integer, ByVal C As Integer)     '���ʤ�r���
    On Error GoTo NextError
    Text1.Width = dg_PalletDetail.CellWidth                     '�e��
    Text1.Height = dg_PalletDetail.CellHeight                   '����
    Text1.Left = dg_PalletDetail.Left + dg_PalletDetail.ColPos(C) + 30 '����
    Text1.Top = dg_PalletDetail.Top + dg_PalletDetail.RowPos(r)     '�W��
    Text1.Text = dg_PalletDetail.Text       '�NMSFlexGrid�ثe�@���x�s�椺�e��m���r���
    Text1.Visible = True                '�N��r�����ܩ�e���W
    Text1.SetFocus                      '�N��в��ܤ�r���
    Exit Sub
NextError:
    MsgBox Err.Description
End Sub

Public Sub NextPosition1(ByVal r As Integer, ByVal C As Integer)     '���ʤ�r���
    On Error GoTo NextError
    lst_Cust.Width = dg_PalletDetail.CellWidth                     '�e��
    'lst_Cust.Height = dg_PalletDetail.CellHeight                   '����
    lst_Cust.Left = dg_PalletDetail.Left + dg_PalletDetail.ColPos(C) + 30 '����
    lst_Cust.Top = dg_PalletDetail.Top + dg_PalletDetail.RowPos(r)     '�W��
    lst_Cust.Text = dg_PalletDetail.Text       '�NMSFlexGrid�ثe�@���x�s�椺�e��m���r���
    lst_Cust.Visible = True                '�N��r�����ܩ�e���W
    lst_Cust.SetFocus                      '�N��в��ܤ�r����W
    Exit Sub
NextError:
    MsgBox Err.Description
End Sub

Private Sub Text1_LostFocus()

    On Error GoTo TextError
        Text1.Visible = False
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Sub Text1_Change()  '�N��r������e�g�ܹ����x�s��
    On Error GoTo TextError
    If dg_PalletDetail.Col = 0 Or dg_PalletDetail.Col = 1 Then
        If Len(Me.Text1.Text) > 0 Then
            If IsNumeric(Me.Text1.Text) = False Then
                msg_text = "�ƶq����,�Э��s��J"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            End If
        End If
    End If
    dg_PalletDetail.Text = Text1.Text   '�N��r������e�g�ܹ����x�s��
    Exit Sub
 
TextError:
    MsgBox Err.Description
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyHome Then
        If dg_PalletDetail.Col = 3 Then
            dg_PalletDetail.Col = dg_PalletDetail.Col - 1
            KeyAscii = 0
            NextPosition1 dg_PalletDetail.Row, dg_PalletDetail.Col
            Exit Sub
        End If
        If dg_PalletDetail.Col = 1 Then
            dg_PalletDetail.Col = dg_PalletDetail.Col - 1
            KeyAscii = 0
            NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
            Exit Sub
        End If
        If dg_PalletDetail.Col = 0 Then
            If dg_PalletDetail.Row = 1 Then
                Me.txt_CheckNo.SetFocus
            Else
                dg_PalletDetail.Row = dg_PalletDetail.Row - 1
                dg_PalletDetail.Col = 3
                NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
            End If
            Exit Sub
        End If
    End If
End Sub
Private Sub AddRow()
dg_PalletDetail.Col = 1
            If Len(Trim(Me.dg_PalletDetail.Text)) = 0 Then
                k = 0
            Else
                k = Trim(Me.dg_PalletDetail.Text)
        End If
            dg_PalletDetail.Col = 2
        If Len(Trim(Me.dg_PalletDetail.Text)) = 0 Then
                m = 0
        Else
                m = Trim(Me.dg_PalletDetail.Text)
        End If

        If k + m = 0 Then
                msg_text = "�ƶq����,�Э��s��J"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dg_PalletDetail.Col = 1
                NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
                Exit Sub
        End If
            dg_PalletDetail.Col = 3
        If Len(Trim(Me.dg_PalletDetail.Text)) = 0 Then
                msg_text = "�п�J�Ȥ�W��"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dg_PalletDetail.Col = 3
                Text1.Visible = False
                Exit Sub
        End If

        If dg_PalletDetail.Rows = dg_PalletDetail.Row + 1 Then
        '�O�_�ɥX�٦^������
        If k <> 0 And m <> 0 Then
        'Ū������
        Dim arrTemp(6)
            For i = 0 To dg_PalletDetail.Cols - 1
            dg_PalletDetail.Col = i
                arrTemp(i) = dg_PalletDetail.Text
                If i = 2 Then dg_PalletDetail.Text = 0 '�٤J�אּ0
            Next
            
        '�s�Wrow1
        dg_PalletDetail.Rows = dg_PalletDetail.Rows + 1
        dg_PalletDetail.Row = dg_PalletDetail.Rows - 1
                                    
        For i = 0 To dg_PalletDetail.Cols - 1
            dg_PalletDetail.Col = i
            dg_PalletDetail.Text = arrTemp(i)
            If i = 1 Then dg_PalletDetail.Text = 0
        Next
        
        End If
        
                If dg_PalletDetail.Rows > 1 Then    '�s�W�@row
                    j = dg_PalletDetail.Rows + 1
                    dg_PalletDetail.Rows = j
                    dg_PalletDetail.Row = j - 1
                End If
        Else
                dg_PalletDetail.Row = dg_PalletDetail.Rows - 1
        End If

            dg_PalletDetail.Col = 1
            KeyAscii = 0
            Call dg_PalletDetail_Click
'            NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
            Exit Sub
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo TextError
    Dim i As Integer
    
    '���Uesc���ä�����
    If KeyAscii = 27 Then mvDate.Visible = False

    If KeyAscii = 19 Then Call cmd_Save_Click

    If KeyAscii = vbKeyReturn Then                '�b���UEnter�ɡA�M�w�U��grid����m
        If dg_PalletDetail.Col = 5 Then
            Call AddRow
            Exit Sub
        End If
        If dg_PalletDetail.Col = 0 Then
            dg_PalletDetail.Col = dg_PalletDetail.Col + 1
            KeyAscii = 0
            NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
            Exit Sub
        End If
        If dg_PalletDetail.Col = 1 Then
            dg_PalletDetail.Col = dg_PalletDetail.Col + 1
            KeyAscii = 0
            NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
            Exit Sub
        End If
        If dg_PalletDetail.Col = 2 Then
            dg_PalletDetail.Col = dg_PalletDetail.Col + 1
            KeyAscii = 0
            NextPosition1 dg_PalletDetail.Row, dg_PalletDetail.Col
            Exit Sub
        End If
        If dg_PalletDetail.Col = 4 Then
            dg_PalletDetail.Col = dg_PalletDetail.Col + 1
            KeyAscii = 0
            NextPosition dg_PalletDetail.Row, dg_PalletDetail.Col
            Exit Sub
        End If
    End If
    Exit Sub
TextError:
    MsgBox Err.Description
End Sub

Private Function QtyCheck() As Boolean  '�ˬd�ƶq�O�_���T
    QtyCheck = True
'    If Val(Trim(txt_CDSIn.Text)) + Val(Trim(txt_CDSOut.Text)) = 0 Then QtyCheck = False: Exit Function
    sumin = 0
    sumout = 0
    m = 0
    n = 0
    j = dg_PalletDetail.Row
    For i = 1 To dg_PalletDetail.Rows - 1
         dg_PalletDetail.Row = i
         dg_PalletDetail.Col = 1
         If Len(Trim(dg_PalletDetail.Text)) > 0 Then
            sumin = sumin + Val(Trim(dg_PalletDetail.Text))
            m = Val(Trim(dg_PalletDetail.Text))
         Else
            m = 0
         End If
         dg_PalletDetail.Col = 2
         If Len(Trim(dg_PalletDetail.Text)) > 0 Then
            sumout = sumout + Val(Trim(dg_PalletDetail.Text))
            n = Val(Trim(dg_PalletDetail.Text))
         Else
            n = 0
         End If
         'If m + n = 0 Then QtyCheck = False
    Next i
    If Len(Trim(txt_CDSIn.Text)) = 0 Then
        m = 0
    Else
        m = Val(Trim(txt_CDSIn.Text))
    End If
    If sumin <> m Then QtyCheck = False
    If Len(Trim(Me.txt_CDSOut.Text)) = 0 Then
        n = 0
    Else
        n = Val(Trim(Me.txt_CDSOut.Text))
    End If
    If sumout <> n Then QtyCheck = False
End Function

Private Function CustCheck() As Boolean  '�ˬd�Ȥ���O�_��J
    CustCheck = True
    For i = 1 To dg_PalletDetail.Rows - 1
         dg_PalletDetail.Row = i
         dg_PalletDetail.Col = 3
         If Len(Trim(dg_PalletDetail.Text)) = 0 Then
             CustCheck = False
         End If
    Next i
End Function

Public Sub clear_PalletDetail()     '
    On Error GoTo NextError
    dg_PalletDetail.Rows = 2
    dg_PalletDetail.Row = 1
    For i = 0 To dg_PalletDetail.Cols - 1
        dg_PalletDetail.Col = i
        dg_PalletDetail.Text = ""
    Next i
    Exit Sub
NextError:
    MsgBox Err.Description
End Sub

Public Sub clear_PalletHead()     '
    On Error GoTo NextError
    dg_PalletHead.Rows = 2
    dg_PalletHead.Row = 1
    For i = 0 To dg_PalletHead.Cols - 1
        dg_PalletHead.Col = i
        dg_PalletHead.Text = ""
    Next i
    Exit Sub
NextError:
    MsgBox Err.Description
End Sub
Public Sub clear_text()     '
    On Error GoTo NextError
    Me.txt_CDSIn.Text = ""
    Me.txt_CDSOut.Text = ""
    Me.txt_CheckNo.Text = ""
    Me.txt_CarNo.Text = ""
    Me.txt_Date.Text = ""
    Me.cboUserType.ListIndex = 0
    txtSortingPL = 0
    txtSorting = 0
    txtUser.Text = ""
    Exit Sub
NextError:
    MsgBox Err.Description
End Sub
