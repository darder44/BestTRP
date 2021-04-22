Attribute VB_Name = "Module1"
Option Explicit

Public rsOffline As ADODB.Recordset
Dim strMsg As String

Sub Main()
'On Error Resume Next
On Error GoTo errHandle

'�ˬd�O�_���ư���
If App.PrevInstance Then Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "�{�����ư���"): End

Dim strSql As String, strServerName As String, strDataBase As String, strID As String, strPassWord As String, strCN As String
Dim strFileName, strLastFileName As String, strFullFileName As String, strSourcePath As String, strToPath As String, strBackupPath1 As String, strBackupPath2 As String, strFileDateTime As String, strCurrentFrom As String, strCurrentTo As String
Dim i As Integer
Dim arrTmp
Dim fso As New Scripting.FileSystemObject

'Ū��ini�Ѽ�
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

'�]�w�P�إߥؿ�
If strSourcePath = "" Then strSourcePath = App.Path
If (Right(strSourcePath, 1) <> "/" And Right(strSourcePath, 1) <> "\") Then strSourcePath = strSourcePath & "\"
If (Right(strToPath, 1) <> "/" And Right(strToPath, 1) <> "\") Then strToPath = strToPath & "\"

'�p�G�O�������|�h���ˬd�ت���Ƨ��O�_�s�b
If Replace(strToPath, "\\", "") = strToPath Then If Dir(strToPath) = "" Then MkDirs (strToPath)
If (Right(strBackupPath1, 1) <> "/" And Right(strBackupPath1, 1) <> "\") Then strBackupPath1 = strBackupPath1 & "\"
If Dir(strBackupPath1) = "" Then MkDirs (strBackupPath1)
If (Right(strBackupPath2, 1) <> "/" And Right(strBackupPath2, 1) <> "\") Then strBackupPath2 = strBackupPath2 & "\"
If Dir(strBackupPath2) = "" Then MkDirs (strBackupPath2)

'***********
'*��Ʈw�s��
'***********
strCN = "PROVIDER=MSDASQL;driver={SQL Server};server=" & strServerName & ";uid=" & strID & ";pwd=" & strPassWord & ";database=" & strDataBase & ";"
Dim rsMain As New ADODB.Recordset

'*************
'*�}�l�ɮ׳B�z
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
    
    '�䤣���ɮ׫h���U�@�ӥؿ�
    If strFileName = "" Then GoTo NextPath
    
    '�p�G�ɦW���Ыh�����{���q�@�~
    If strLastFileName = strFileName Then Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "�ɦW���е����{���q�@�~!"): GoTo NextPath
    
    strLastFileName = strFileName
    
    strFullFileName = strCurrentFrom & strFileName
    strFileDateTime = FileDateTime(strFullFileName)
    
    '�ɮ׶}�Ҥ��h���L�U�@���ɮ�
    If iscompleted(strFullFileName) = False Then
        Call ErrorMsgbox(App.Title, Err.Number, Err.Description, strFullFileName & "�ɮ׶}�Ҥ��y��B��!")
        GoTo NextFile
    End If
    
'    If FileLen(strFullFileName) = 0 Then Call ErrorMsgbox(App.Title, Err.Number, "�ɮת��׬� 0 ", "�ɮצW�١G " & strFileName): strFileName = Dir: GoTo LoadFile
        
    '*********
    '*�ɮפ��t
    '*********
    If Len(strBackupPath1) > 1 Then FileCopy strFullFileName, strBackupPath1 & strFileName
    If Len(strBackupPath2) > 1 Then FileCopy strFullFileName, strBackupPath2 & strFileName
    
    FileCopy strFullFileName, strCurrentTo & strFileName
    
    '�ˬd�O�_�ƻs���\
    If fso.FileExists(strCurrentTo & strFileName) = True Then
    
        strTextbody = strFileName & " �ɮ׮ɶ��G" & strFileDateTime & " �ɮפj�p�G" & FileLen(strFullFileName)

        '�g�J��Ʈw
        strSql = "insert into gt_filelog(storerkey,filename,filedate,filelen) values('LTKK01','" & strFileName & "','" & strFileDateTime & "','" & FileLen(strFullFileName) & "')"
        rsMain.Open strSql, strCN

        Kill strFullFileName
        Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "�ɮ׽ƻs�����A�ɮצW�١G" & strTextbody)

        If Len(RTrim(strTextbody)) > 0 And Len(RTrim(strFrom)) > 0 Then

           '�ǰe�l��
            strSubject = "FTP���ɳq��(" & strFileName & ")"
            Dim objEmail As Object
            Set objEmail = CreateObject("CDO.Message")

            objEmail.From = strFrom
            objEmail.To = strTo
            objEmail.CC = strCC   ' �ƥ�
            objEmail.BCC = strBCC ' �K��ƥ�
            objEmail.Subject = strSubject
            objEmail.TextBody = strTextbody
            objEmail.AddAttachment strAddAttachment

            objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.bestlog.com.tw"
            objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            'SMTP ���A���ݭn���Ү�
            If Len(RTrim(strEmailID)) > 0 Then
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
            End If
            objEmail.Configuration.Fields.Update
            objEmail.Send

            Set objEmail = Nothing

            Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "Email�@�~����")

        End If
End If

NextFile:
    strFileName = Dir: GoTo LoadFile

NextPath:

Next i

Exit Sub

errHandle:
Call ErrorMsgbox(App.Title, Err.Number, Err.Description, "�ɮצW�١G " & strFileName)
End Sub

Public Sub OffLineRecordset(ByRef rsBefore As ADODB.Recordset, ByRef rsAfter As ADODB.Recordset)
Dim i As Integer, lngStep As Long
Set rsAfter = New ADODB.Recordset
On Error GoTo errHandle

'�إߵ���Recordset
'Call ReDim_Recordset(rsAfter)
rsAfter.Fields.Append "����", adInteger
'rsAfter.Fields.Append "���", adChar, 1
For i = 0 To rsBefore.Fields.Count - 1
    rsAfter.Fields.Append rsBefore.Fields(i).Name, rsBefore.Fields(i).Type, rsBefore.Fields(i).DefinedSize
Next i

With rsAfter
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '���ݳs������
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
    
'���͵{������O��
Open App.Path & "\" & App.Title & ".log" For Append As #FileNumber
    
'�g�J���A��
Print #FileNumber, Format(Now, "yyyy-mm-dd ttttt ") & strMsg
Close #FileNumber

'MsgBox "���~!!" & vbCrLf & vbCrLf & "Number�G" & strErrNo & vbCrLf & "Description�G" & strErrDescr & vbCrLf & "�Ƶ��G" & strNote, vbOKOnly + vbInformation, App.Title & "_" & strMeCaption
Screen.MousePointer = vbDefault
    
End Sub
Public Function MkDirs(ByVal PathIn As String) As Boolean
    Dim nPos As Long
    MkDirs = True  '�����]���\
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
'�P�_�ɮ׬O�_�W���Ҧ�

'���յ��G�A�ɮ׳Q�}�ҩάO�ǿ餤���i�H�P�_�X
'
'strTranFileName = "\\192.168.1.203\d$\FTP\best01-tw-ftp-test\outbox\TestWMS_20140721.zip"
'
'If iscompleted(strTranFileName) = turn Then
'
'    msg_text = "�ɮפw�Q�}��"
'
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'
'    Exit Sub
'
'Else
'
'    msg_text = "�i���"
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
