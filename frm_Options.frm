VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frm_Options 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "選項"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm_Options.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "安全"
      TabPicture(1)   =   "frm_Options.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "排車"
      TabPicture(2)   =   "frm_Options.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "系統版本"
         Height          =   1455
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   6375
         Begin VB.CommandButton Cmd_updateversion 
            Caption         =   "更新"
            Height          =   375
            Left            =   3840
            TabIndex        =   22
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_Version 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   720
            Width           =   1815
         End
         Begin VB.PictureBox Picture4 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  '沒有框線
            Height          =   240
            Left            =   120
            Picture         =   "frm_Options.frx":0054
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   18
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "目前版本"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   21
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "系統版本如果不是最新，則要求User手動更新TMS，但是可以使用Bestold執行檔來執行更新前的TMS。先更新版本，再打包執行檔"
            Height          =   375
            Left            =   840
            TabIndex        =   19
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "簽單維護期限"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   12
         Top             =   780
         Width           =   6375
         Begin VB.TextBox txtDay 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   270
            Left            =   720
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "0"
            Top             =   1020
            Width           =   615
         End
         Begin VB.PictureBox Picture3 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  '沒有框線
            Height          =   240
            Left            =   120
            Picture         =   "frm_Options.frx":29C66
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   13
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "天。(系統管理員權限才可修改)"
            Height          =   255
            Left            =   1440
            TabIndex        =   16
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label3 
            Caption         =   "        限制超過多少天以前的簽單將無法再異動與維護，含運費資料維護。(需重新登入)"
            Height          =   375
            Left            =   840
            TabIndex        =   14
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "排車資料權限控管"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   8
         Top             =   780
         Width           =   6375
         Begin VB.CheckBox chkRouteModify 
            Caption         =   "啟用權限控管"
            Height          =   255
            Left            =   840
            TabIndex        =   10
            Top             =   1080
            Width           =   1935
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  '沒有框線
            Height          =   240
            Left            =   120
            Picture         =   "frm_Options.frx":53878
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   9
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   $"frm_Options.frx":5A0CA
            Height          =   735
            Left            =   840
            TabIndex        =   11
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "登入控管"
         Height          =   1455
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   6375
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  '沒有框線
            Height          =   240
            Left            =   120
            Picture         =   "frm_Options.frx":5A172
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   7
            Top             =   240
            Width           =   240
         End
         Begin VB.CheckBox chkLoginControl 
            Caption         =   "啟用登入控管 (資訊部權限才可修改)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   6
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   $"frm_Options.frx":83D84
            Height          =   735
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   5415
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "套用"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_updateversion_Click()

cn.Execute "insert into versioncheck(adddate,version,project) values(getdate(),'" & RTrim(txt_Version.Text) & "','BestTms')", RowsAffect, adExecuteNoRecords
MsgBox "版本更新成功!", vbOKOnly + vbSystemModal, "版本更新"
End Sub

Private Sub cmdApply_Click()

'簽單維護期限
cn.Execute "update codelkup set description = '" & Val(txtDay) & "',editwho = '" & User_id & "' ,editdate = getdate() where listname = 'Options' and code = 'DueDate'", RowsAffect, adExecuteNoRecords
If RowsAffect = 0 Then cn.Execute "insert into codelkup(listname,code,description,adddate,addwho,editdate,editwho) values('OPTIONS','DueDate'," & Val(txtDay) & ",getdate(),'" & User_id & "',getdate(),'" & User_id & "') ", RowsAffect, adExecuteNoRecords

'登入控管
cn.Execute "update codelkup set description = '" & chkLoginControl & "',editwho = '" & User_id & "' ,editdate = getdate() where listname = 'Options' and code = 'logincontrol'", RowsAffect, adExecuteNoRecords

'路線編號資料修正
cn.Execute "update codelkup set description = '" & chkRouteModify & "',editwho = '" & User_id & "' ,editdate = getdate() where listname = 'Options' and code = 'RouteModifyControl'", RowsAffect, adExecuteNoRecords
If RowsAffect = 0 Then cn.Execute "insert into codelkup(listname,code,description,adddate,addwho,editdate,editwho) values('OPTIONS','RouteModifyControl'," & chkRouteModify & ",getdate(),'" & User_id & "',getdate(),'" & User_id & "') ", RowsAffect, adExecuteNoRecords
blRouteModifyControl = chkRouteModify

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Call cmdApply_Click
Call cmdCancel_Click
End Sub


Private Sub Form_Load()

chkLoginControl = 0
If blRouteModifyControl Then chkRouteModify = 1 '路線編號資料修正
    
Dim rsOptions As New ADODB.Recordset
str_SQL = "Select * From codelkup where ListName = 'OPTIONS' "

rsOptions.Open str_SQL, cn, adOpenKeyset, adLockPessimistic

rsOptions.MoveFirst
Do While Not rsOptions.EOF
    If UCase(RTrim(rsOptions("code"))) = "LOGINCONTROL" Then chkLoginControl = Val(rsOptions("Description"))     '登入控管
    If UCase(RTrim(rsOptions("code"))) = "DUEDATE" Then txtDay = Val(rsOptions("Description"))     '簽單維護期限

rsOptions.MoveNext
Loop

rsOptions.Close: Set rsOptions = Nothing

txt_Version = RTrim(App.Major & "." & App.Minor & "." & App.Revision)

SSTab1.Tab = 0

'If blAdmin Then chkLoginControl.Enabled = True
If UCase(User_id) = "ADMINISTRATOR" Or UCase(strComputerName) = "BESTDB" Or UCase(strComputerName) = "GEMINI_NB" Then chkLoginControl.Enabled = True: Cmd_updateversion.Visible = True '資訊部帳號才可修改
If blAdmin Then chkRouteModify.Enabled = True   '管理員群組可以修改
If blAdmin Then txtDay.Enabled = True

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub
