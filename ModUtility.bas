Attribute VB_Name = "ModUtility"
Public cnSource           As New Connection
Public cnTarget           As New Connection
Public sSource            As String
Public sTarget            As String

Public sDiff              As String
Public sDB                As String

Public sPath As String
Public sTabel As String
Public sColDate As String
Public sColTime As String
Public sColLogID As String



Public Function GetSetting() As String
    Dim fSetting As FileSystemObject
    Set fSetting = New FileSystemObject
    Dim sFile       As String
    
    sFile = ""
    
    sFile = App.Path & "\setting.ini"
    If Not fSetting.FileExists(App.Path & "\setting.ini") Then
       MsgBox "Sorry..!!! setting.ini file not found in. Call technical support to get setting.ini created", vbInformation
       'ReadPath = "Error: Setting File not found.."
       Exit Function
    End If
    
    GetSetting = sFile
    
End Function

Public Function ReadPath() As String
    On Error GoTo ReadPathErr
    
    Dim fSetting As FileSystemObject
    Set fSetting = New FileSystemObject
    
    Dim sFile       As String
    
    sFile = GetSetting
        
    sSource = ReadIni(sFile, "source", "Path")
    sSource = sSource & sDB
    If IsNull(sSource) Then
        ReadPath = "Error: Source Database not found..."
        Exit Function
    End If
    
    sTarget = ReadIni(sFile, "target", "Path")
    If IsNull(sTarget) Then
        ReadPath = "Error: Target Database not found..."
        Exit Function
    End If
    
    ReadPath = "Success"
    
    sDiff = ReadIni(sFile, "TimeDiff", "Min")
    
    Exit Function
ReadPathErr:
    MsgBox err.Description, vbInformation, "Time Tracker"
    frmDownload.lblFinal.Caption = "Failure....!!!"
    
End Function
Public Sub StoreInTextFile(ByVal strEmpIdtxt As String, ByVal strDatetxt As String, ByVal strtimetxt As String, ByVal stat As Integer, ByVal Prod_Name As String)
    Dim fname As FileSystemObject
    Set fname = New FileSystemObject
    Dim s As Object
    If Not fname.FolderExists(App.Path & "\" & Prod_Name) Then fname.CreateFolder (App.Path & "\" & Prod_Name)
    'End If
    
  '  If Not fname.FileExists(App.Path & "\" & Prod_Name & "\" & Prod_Name & Format(Date, "MMM-yy") & "Import.sur") Then
  '     fname.CreateTextFile (App.Path & "\" & Prod_Name & "\" & Prod_Name & Format(Date, "MMM-yy") & "Import.sur")
  '  End If
    Set s = fname.OpenTextFile(App.Path & "\" & Prod_Name & "\" & Prod_Name & Format(Date, "MMM-yy") & "import.sur", ForAppending, True)
    s.WriteLine strEmpIdtxt & "," & strDatetxt & "," & strtimetxt & "," & stat
    s.Close
End Sub

Public Function OpenTargetDB() As String
    On Error GoTo err

    If cnTarget.State Then Exit Function
    
    cnTarget.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sTarget & ";Persist Security Info=False;Jet OLEDB:Database Password=rAMAcHANDRA"
    
    OpenTargetDB = "Connected Target DB"
    
    Exit Function
    
err:
    OpenTargetDB = "Error Connecting to Target DB."
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical, "Time Tracker"
End Function


Public Function OpenSourceDB() As String
    On Error GoTo err

    If cnSource.State Then Exit Function
    
    Select Case sDB
        Case "ATT2000.mdb"
            cnSource.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sSource & ";Persist Security Info=False"
        Case "Warden.mdb"
            cnSource.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sSource & ";Persist Security Info=False;Jet OLEDB:Database Password=RamaChandra"
    '.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Ostpl.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=RamaChandra"
         Case "File.mdb"
            cnSource.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Left(sSource, Len(sSource) - Len(sDB)) & ";Persist Security Info=False;Jet OLEDB:Database"
        Case Else
            GoTo err
    End Select
    
    OpenSourceDB = "Connected Source DB"
    
    Exit Function
    
err:
    OpenSourceDB = "Error Connecting to Source DB."
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical, "Time Tracker"
End Function
Public Sub CloseDatabase()
    If cnSource.State Then cnSource.Close
    If cnTarget.State Then cnTarget.Close
End Sub

Public Sub AttDownloadData()
    On Error GoTo AttDownloadDataErr
    
    Dim rsDest As New ADODB.Recordset, rsSource As New ADODB.Recordset
    
    Dim cnt As Long, n As Long, rsCount As Long
    Dim MyTime As Date, mydate As Date, MySQl As String, flag As Boolean
    
    If Trim(sDiff) = "" Then
        sDiff = 2
    End If
    
    frmDownload.lblTimeDiff.Caption = "Time difference between each swip  > " & sDiff & " Min"
    
    Set rsSource = cnSource.Execute("SELECT count(userid) from Checkinout")
    If IsNull(rsSource(0)) Then
        rsCount = 0
    Else
        rsCount = rsSource(0)
    End If

    Set rsSource = cnSource.Execute("SELECT Checkinout.userid, badgenumber, checktime from Checkinout inner join USERINFO on USERINFO.Userid = Checkinout.Userid order by checktime")
    If rsSource.EOF Then
        frmDownload.lblStatus.Caption = "No Data to move..."
        Exit Sub
    Else
        rsSource.MoveFirst
    End If
    
    cnt = 1
    Do While Not rsSource.EOF
        MyTime = Format(rsSource("Checktime"), "HH:mm:ss")
        mydate = Format(rsSource("Checktime"), "dd/MMM/yyyy")
        
        DoEvents
        
        StoreInTextFile Trim(rsSource("badgenumber")), mydate, MyTime, 1, "att2000"
          
        frmDownload.lblStatus.Caption = "Importing  : " & rsSource("badgenumber") & " on " & mydate & vbCrLf & cnt & "/" & rsCount
          
        'check for duplicate
        If checkDuplidate(Trim(rsSource("badgenumber")), mydate & " " & MyTime, "1", sDiff) Then
            flag = True
        End If
        
        If flag = False Then
            MySQl = ""
            MySQl = "Insert into tbl_trn_log (logid,logdate,logtime)"
            MySQl = MySQl & " values('" & Trim(rsSource("badgenumber")) & "',"
            MySQl = MySQl & "#" & mydate & "#,#" & MyTime & "# )" '
            cnTarget.Execute (MySQl)
        End If
        rsSource.MoveNext
        cnt = cnt + 1
        flag = False
    Loop
    
    If cnt > 1 Then
        MsgBox "Data move successful..", vbInformation, "Time Traker"
    Else
        MsgBox "Oh.. not data to move", vbInformation, "Time Traker"
    End If
    'cnSource.Execute ("Delete * from Checkinout")
    Exit Sub

AttDownloadDataErr:
    MsgBox "AttDownloadDataErr " & err.Description, vbInformation, "Time Tracker"
    frmDownload.lblFinal.Caption = "Failure....!!!"
 
End Sub

Public Sub WardenDownloadData()
    On Error GoTo WardenDownloadDataErr
    
    Dim rsDest As New ADODB.Recordset, rsSource As New ADODB.Recordset
    
    Dim cnt As Long, n As Long, rsCount As Long
    Dim MyTime As Date, mydate As Date, MySQl As String, flag As Boolean
    
    If Trim(sDiff) = "" Then
        sDiff = 2
    End If
    
    frmDownload.lblTimeDiff.Caption = "Time difference between each swip  > " & sDiff & " Min"
    
    Set rsSource = cnSource.Execute("SELECT count(logid) from log where logid <> NULL and  type <> '" & "Y" & "'")
    If IsNull(rsSource(0)) Then
        rsCount = 0
    Else
        rsCount = rsSource(0)
    End If

    Set rsSource = cnSource.Execute("SELECT logid, Logdate, deviceid from log where logid <> NULL and  type <> '" & "Y" & "' order by Logdate")
    If rsSource.EOF Then
        frmDownload.lblStatus.Caption = "No Data to move..."
        Exit Sub
    Else
        rsSource.MoveFirst
    End If
    
    cnt = 1
    Do While Not rsSource.EOF
        If IsNull(rsSource("logid")) Or Trim(rsSource("logid")) = "" Then GoTo Nxt
        
        MyTime = Format(rsSource("Logdate"), "HH:mm:ss")
        mydate = Format(rsSource("Logdate"), "dd/MMM/yyyy")
        
        DoEvents
        
        StoreInTextFile Trim(rsSource("userid")), mydate, MyTime, 1, "Warden"
          
        frmDownload.lblStatus.Caption = "Importing  : " & rsSource("logid") & " on " & mydate & vbCrLf & cnt & "/" & rsCount
          
        'check for duplicate
        If checkDuplidate(Trim(rsSource("logid")), mydate & " " & MyTime, "1", sDiff) Then
            flag = True
        End If
        
        If flag = False Then
            MySQl = ""
            MySQl = "Insert into tbl_trn_log (logid,logdate,logtime)"
            MySQl = MySQl & " values('" & Trim(rsSource("logid")) & "',"
            MySQl = MySQl & "#" & mydate & "#,#" & MyTime & "# )" '
            cnTarget.Execute (MySQl)
        End If
        rsSource.MoveNext
        cnt = cnt + 1
        flag = False
        cnSource.Execute ("update log set type = '" & "Y" & "' where logid = '" & rsSource("logid") & "' and logdate = #" & rsSource("logdate") & "#"), n
Nxt:
        rsSource.MoveNext
    Loop
    
    If cnt > 1 Then
        MsgBox "Data move successful..", vbInformation, "Time Traker"
    Else
        MsgBox "Oh.. no data to move", vbInformation, "Time Traker"
    End If
    
    Exit Sub

WardenDownloadDataErr:
    MsgBox "WardenDownloadDataErr " & err.Description, vbInformation, "Time Tracker"
    frmDownload.lblFinal.Caption = "Failure....!!!"
 
End Sub


Public Function checkDuplidate(ByVal EMPID As String, ByVal TDate As Date, ByVal strMCNo As String, ByVal txtDiff As String) As Boolean
    'On Error Resume Next
    Dim sql As String
    Dim rsDup As New ADODB.Recordset
    Dim TmpTime As Date
    Dim gblTimeDiff As Long
    gblTimeDiff = Val(txtDiff)
    TmpTime = Format(TDate, "HH:mm")
    sql = "select * from tbl_trn_log where logid= '" & EMPID & "' and logdate = #" & Format(TDate, "dd/MMM/yyyy") & "#"
    sql = sql & " and Machineid = '" & strMCNo & "'"
    sql = sql & " and (logtime between #" & DateAdd("n", -gblTimeDiff, TmpTime) & "# and #" & DateAdd("n", gblTimeDiff, TmpTime) & "#)"
'    Debug.Print sql
    Set rsDup = cnTarget.Execute(sql)
    If Not rsDup.EOF And Not rsDup.BOF Then
        checkDuplidate = True
    Else
        checkDuplidate = False
    End If
End Function



Public Sub FileLoad()
    On Error GoTo FileLoad
    
    Dim rsDest As New ADODB.Recordset, rsSource As New ADODB.Recordset
    
    Dim cnt As Long, n As Long, rsCount As Long
    Dim MyTime As Date, mydate As Date, MySQl As String, flag As Boolean
    
    
    If Trim(sDiff) = "" Then
        sDiff = 2
    End If
    
    frmDownload.lblTimeDiff.Caption = "Time difference between each swip  > " & sDiff & " Min"
    
    MySQl = "SELECT count(" + sColLogID + ") from " + sTabel + " where " + sColLogID + " <> NULL"
    
    ''Set rsSource = cnSource.Execute("SELECT count(logid) from log where logid <> NULL and  type <> '" & "Y" & "'")
    Set rsSource = cnSource.Execute(MySQl)
    
    If IsNull(rsSource(0)) Then
        rsCount = 0
    Else
        rsCount = rsSource(0)
    End If

    MySQl = "SELECT " + sColLogID + " as logid," + sColDate + " as logdate  from " + sTabel + " where " + sColLogID + " <> NULL order by " + sColDate

    'Set rsSource = cnSource.Execute("SELECT logid, Logdate, deviceid from log where logid <> NULL and  type <> '" & "Y" & "' order by Logdate")
    Set rsSource = cnSource.Execute(MySQl)
    If rsSource.EOF Then
        frmDownload.lblStatus.Caption = "No Data to move..."
        Exit Sub
    Else
        rsSource.MoveFirst
    End If
    
    cnt = 1
    Do While Not rsSource.EOF
        If IsNull(rsSource("logid")) Or Trim(rsSource("logid")) = "" Then GoTo Nxt
        
        MyTime = Format(rsSource("Logdate"), "HH:mm:ss")
        mydate = Format(rsSource("Logdate"), "dd/MMM/yyyy")
        
        DoEvents
        
        StoreInTextFile Trim(rsSource("logid")), mydate, MyTime, 1, "SourceDB"
          
        frmDownload.lblStatus.Caption = "Importing  : " & rsSource("logid") & " on " & mydate & vbCrLf & cnt & "/" & rsCount
          
        'check for duplicate
        If checkDuplidate(Trim(rsSource("logid")), mydate & " " & MyTime, "1", sDiff) Then
            flag = True
        End If
        
        If flag = False Then
            MySQl = ""
            MySQl = "Insert into tbl_trn_log (logid,logdate,logtime)"
            MySQl = MySQl & " values('" & Trim(rsSource("logid")) & "',"
            MySQl = MySQl & "#" & mydate & "#,#" & MyTime & "# )" '
            cnTarget.Execute (MySQl)
        End If
        rsSource.MoveNext
        cnt = cnt + 1
        flag = False
        'cnSource.Execute ("update log set type = '" & "Y" & "' where logid = '" & rsSource("logid") & "' and logdate = #" & rsSource("logdate") & "#"), n
Nxt:
        rsSource.MoveNext
    Loop
    
    If cnt > 1 Then
        If MsgBox("Data move successful. Do you want to delete from source database?", vbInformation + vbYesNo, "Time Traker") = vbYes Then
            MySQl = "(Delete * from " + sTabel + ")"
            cnSource.Execute MySQl
        End If
    Else
        MsgBox "Oh.. no data to move", vbInformation, "Time Traker"
    End If
    
    Exit Sub

FileLoad:
    MsgBox "FileLoad " & err.Description, vbInformation, "Time Tracker"
    frmDownload.lblFinal.Caption = "Failure....!!!"
 
End Sub
