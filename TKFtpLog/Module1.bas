Attribute VB_Name = "Module1"
Option Explicit

Public rsOffline As ADODB.Recordset
Dim strMsg As String

Sub Main()
'On Error Resume Next
On Error GoTo errHandle

'檢查是否重複執行
If App.PrevInstance Then Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "程式重複執行"): End

Dim strSql As String, strServerName As String, strDataBase As String, strID As String, strPassWord As String, strCN As String
Dim strFileName, strLastFileName As String, strFullFileName As String, strSourcePath As String, strToPath As String, strBackupPath1 As String, strBackupPath2 As String, strFileDateTime As String, strCurrentFrom As String, strCurrentTo As String
Dim i As Integer
Dim arrTmp
Dim fso As New Scripting.FileSystemObject

'讀取ini參數
Dim objIni As New vbIniFile
Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String

objIni.FileName = App.Path & "/" & App.Title & ".ini"

strServerName = objIni.ReadData("System", "ServerName", ".")
strDataBase = objIni.ReadData("System", "DataBase", "")
strID = objIni.ReadData("System", "ID", "")
strPassWord = objIni.ReadData("System", "PassWord", "sa")
strSourcePath = objIni.ReadData("Path", "SourcePath", "")
strToPath = objIni.ReadData("Path", "Topath", "")
strBackupPath1 = objIni.ReadData("Path", "BackupPath1", "")
strBackupPath2 = objIni.ReadData("Path", "BackupPath2", "")
strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
strSubject = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Subject", "")
strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")

'設定與建立目錄
If strSourcePath = "" Then strSourcePath = App.Path
If (Right(strSourcePath, 1) <> "/" And Right(strSourcePath, 1) <> "\") Then strSourcePath = strSourcePath & "\"
If (Right(strToPath, 1) <> "/" And Right(strToPath, 1) <> "\") Then strToPath = strToPath & "\"

'如果是網路路徑則不檢查目的資料夾是否存在
If Replace(strToPath, "\\", "") = strToPath Then If Dir(strToPath) = "" Then MkDirs (strToPath)
If (Right(strBackupPath1, 1) <> "/" And Right(strBackupPath1, 1) <> "\") Then strBackupPath1 = strBackupPath1 & "\"
If Dir(strBackupPath1) = "" Then MkDirs (strBackupPath1)
If (Right(strBackupPath2, 1) <> "/" And Right(strBackupPath2, 1) <> "\") Then strBackupPath2 = strBackupPath2 & "\"
If Dir(strBackupPath2) = "" Then MkDirs (strBackupPath2)

'***********
'*資料庫連接
'***********
strCN = "PROVIDER=MSDASQL;driver={SQL Server};server=" & strServerName & ";uid=" & strID & ";pwd=" & strPassWord & ";database=" & strDataBase & ";"
Dim rsMain As New ADODB.Recordset

'*************
'*開始檔案處理
'*************
arrTmp = Array("Orders\", "Receive\")

For i = 0 To UBound(arrTmp)
    strFileName = "": strLastFileName = ""
    strCurrentFrom = strSourcePath & arrTmp(i)
    strCurrentTo = strToPath & arrTmp(i)
    If Dir(strCurrentFrom, vbDirectory) = "" Then MkDirs (strCurrentFrom)
    If Dir(strCurrentTo, vbDirectory) = "" Then MkDirs (strCurrentTo)
    
    strFileName = Dir(strCurrentFrom & "*.*")
    
LoadFile:
    
    '找不到檔案則換下一個目錄
    If strFileName = "" Then GoTo NextPath
    
    '如果檔名重覆則結束現階段作業
    If strLastFileName = strFileName Then Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "檔名重覆結束現階段作業!"): GoTo NextPath
    
    strLastFileName = strFileName
    
    strFullFileName = strCurrentFrom & strFileName
    strFileDateTime = FileDateTime(strFullFileName)
    
    '檔案開啟中則跳過下一個檔案
    If iscompleted(strFullFileName) = False Then
        Call ErrorMsgbox(App.Title, Err.Number, Err.Description, strFullFileName & "檔案開啟中稍後處裡!")
        GoTo NextFile
    End If
    
'    If FileLen(strFullFileName) = 0 Then Call ErrorMsgbox(App.Title, Err.Number, "檔案長度為 0 ", "檔案名稱： " & strFileName): strFileName = Dir: GoTo LoadFile
        
    '*********
    '*檔案分配
    '*********
    If Len(strBackupPath1) > 1 Then FileCopy strFullFileName, strBackupPath1 & strFileName
    If Len(strBackupPath2) > 1 Then FileCopy strFullFileName, strBackupPath2 & strFileName
    
    FileCopy strFullFileName, strCurrentTo & strFileName
    
    '檢查是否複製成功
    If fso.FileExists(strCurrentTo & strFileName) = True Then
    
        strTextbody = strFileName & " 檔案時間：" & strFileDateTime & " 檔案大小：" & FileLen(strFullFileName)

        '寫入資料庫
        strSql = "insert into gt_filelog(storerkey,filename,filedate,filelen) values('LTKK01','" & strFileName & "','" & strFileDateTime & "','" & FileLen(strFullFileName) & "')"
        rsMain.Open strSql, strCN

        Kill strFullFileName
        Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "檔案複製完成，檔案名稱：" & strTextbody)

        If Len(RTrim(strTextbody)) > 0 And Len(RTrim(strFrom)) > 0 Then

           '傳送郵件
            strSubject = "FTP接檔通知(" & strFileName & ")"
            Dim objEmail As Object
            Set objEmail = CreateObject("CDO.Message")

            objEmail.From = strFrom
            objEmail.To = strTo
            objEmail.CC = strCC   ' 副本
            objEmail.BCC = strBCC ' 密件副本
            objEmail.Subject = strSubject
            objEmail.TextBody = strTextbody
            objEmail.AddAttachment strAddAttachment

            objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.bestlog.com.tw"
            objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            'SMTP 伺服器需要驗證時
            If Len(RTrim(strEmailID)) > 0 Then
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
            End If
            objEmail.Configuration.Fields.Update
            objEmail.Send

            Set objEmail = Nothing

            Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "Email作業結束")

        End If
End If

NextFile:
    strFileName = Dir: GoTo LoadFile

NextPath:

Next i

Exit Sub

errHandle:
Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "檔案名稱： " & strFileName)
End Sub

Public Sub OffLineRecordset(ByRef rsBefore As ADODB.Recordset, ByRef rsAfter As ADODB.Recordset)
Dim i As Integer, lngStep As Long
Set rsAfter = New ADODB.Recordset
On Error GoTo errHandle

'建立虛擬Recordset
'Call ReDim_Recordset(rsAfter)
rsAfter.Fields.Append "項次", adInteger
'rsAfter.Fields.Append "選取", adChar, 1
For i = 0 To rsBefore.Fields.Count - 1
    rsAfter.Fields.Append rsBefore.Fields(i).Name, rsBefore.Fields(i).Type, rsBefore.Fields(i).DefinedSize
Next i

With rsAfter
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
i = 0
Do While Not rsBefore.EOF
   lngStep = lngStep + 1
   rsAfter.AddNew
'   rsAfter.Fields(0).Value = " "
    rsAfter.Fields(0).Value = lngStep
   For i = 0 To rsBefore.Fields.Count - 1
       rsAfter.Fields(i + 1).Value = "" & rsBefore.Fields(i).Value
   Next i
   rsAfter.Update
   rsBefore.MoveNext
Loop

Screen.MousePointer = 0
Exit Sub
errHandle:
Call ErrorMsgbox("OffineRecordset", Err.Number, Err.Description, "")
End Sub

Public Sub ErrorMsgbox(ByVal strMeCaption, strErrNo, strErrDescr As String, strNote As String)
    Dim FileNumber
    FileNumber = FreeFile
    strMsg = strMeCaption & Chr(9) & strErrNo & Chr(9) & strErrDescr & Chr(9) & strNote
    
'產生程式執行記錄
Open App.Path & "\" & App.Title & ".log" For Append As #FileNumber
    
'寫入狀態值
Print #FileNumber, Format(Now, "yyyy-mm-dd ttttt ") & strMsg
Close #FileNumber

'MsgBox "錯誤!!" & vbCrLf & vbCrLf & "Number：" & strErrNo & vbCrLf & "Description：" & strErrDescr & vbCrLf & "備註：" & strNote, vbOKOnly + vbInformation, App.Title & "_" & strMeCaption
Screen.MousePointer = vbDefault
    
End Sub
Public Function MkDirs(ByVal PathIn As String) As Boolean
    Dim nPos As Long
    MkDirs = True  '先假設成功
    If Right$(PathIn, 1) <> "\" Then PathIn = PathIn + "\"
    nPos = InStr(1, PathIn, "\")
    Do While nPos > 0
        If Dir$(Left$(PathIn, nPos), vbDirectory) = "" Then
            On Error GoTo Failed
            MkDir Left$(PathIn, nPos)
            On Error GoTo 0
        End If
        nPos = InStr(nPos + 1, PathIn, "\")
    Loop
    Exit Function
Failed:
    MkDirs = False
End Function

Function iscompleted(FileName As String) As Boolean
'判斷檔案是否獨佔模式

'測試結果，檔案被開啟或是傳輸中都可以判斷出
'
'strTranFileName = "\\192.168.1.203\d$\FTP\best01-tw-ftp-test\outbox\TestWMS_20140721.zip"
'
'If iscompleted(strTranFileName) = turn Then
'
'    msg_text = "檔案已被開啟"
'
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'
'    Exit Sub
'
'Else
'
'    msg_text = "可抓取"
'
'       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'
'    Exit Sub
'
'End If

Dim f As Integer

On Error GoTo openfailed

    f = FreeFile()

    Open FileName For Binary Lock Read Write As #f

    Close #f

    iscompleted = True

    Exit Function

openfailed:

    iscompleted = False

End Function
