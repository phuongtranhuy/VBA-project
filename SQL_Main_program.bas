Attribute VB_Name = "SQL_Main_program"
Public ObjWshNw As Object
Public Const sqltemplate_path = "C:\temp\SQL table template.XLSX"



Sub SQL_download()
     ' On Error GoTo errHandling
     ' FOR THIS CODE TO WORK
     ' In VBE you need to go Tools References and check Microsoft Active X Data Objects 2.x library
    application.ScreenUpdating = False
    

    
    '------------------------------------********** SET UP CONNECTION TO MS SQL SERVER *************------------------------------------------
    Dim Cn As New ADODB.Connection
    Dim rs1 As New ADODB.RecordSet
    Dim rs2 As New ADODB.RecordSet
    Dim iCols As Integer
    
    Set Cn = New ADODB.Connection
    
    Dim Server_Name As String
    Dim Database_Name As String
    Dim user_id As String
    Dim Password As String
    Dim SQLStr As String
    
    Set ObjWshNw = CreateObject("WScript.Network")
    USERID_HSCODEDATA = ObjWshNw.UserName
    
    Server_Name = "tcp:SGPVSQL58.apac.bosch.com,1433" ' Enter your server name here
    Database_Name = "DB_CTXFC1_SQL" ' Enter your database name here
    user_id = "BOSCH\WOM.C_TXF-C1-INT" ' enter your user ID here
    Password = "" ' Enter your password here
    

    
    Set Cn = New ADODB.Connection
    Cn.ConnectionString = "Driver={ODBC Driver 17 for SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & user_id & ";Pwd=" & Password & _
    ";Trusted_Connection=yes;Encrypt=yes;TrustServerCertificate=yes;DataTypeCompatibility=80;MultipleActiveResultSets=True;"
    
    Cn.ConnectionTimeout = 100
    Cn.Open
    '------------------------------------************************************************************-----------------------------------------

       SQLStr_GET_part1 = "SELECT  [Priority]" & _
                                ",[Logical System Group],[Product Number]" & _
                                ",[Product Short Text],[Documentation Status] + ' - ' + SUBSTRING([system id documentation status is found],4,3) as [Documentation Status]" & _
                                ",[Product Type],[Configurable Material],[Term code],[Termcode text]" & _
                                ",[Prod#hierarchy],[Hierachy Definition]" & _
                                ",[Where used Part Number],[Where Used Description],[Direct Component],[Direct Usage]" & _
                                ",[Classification Criteria],[Weight],[Hardness Grade on MO table],"
  
    
       SQLStr_GET_part2 = "[Product Group],[HS code After clean up]" & _
                         ",[Reason],CAST([Comment] AS text) as [Comment],"
                        

                        
                         If Task = "2nd Classification" Then '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                             
                            Task_name = "2nd_classify"
                            task_label = "2nd"
                             
                            download_time_header = "[" & "2nd classify down time" & "]"
                            
                            where_condition = "FROM [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                " WHERE " & _
                                                    "[2nd PIC] = 'Block by " & PIC & "'" & " AND " & "[Logical System Group] = " & "'" & Logicalgroup & "' ORDER BY [Priority] ASC, [Term code]  desc, [Product Number] asc"
                            
                            SQLStr_GET = SQLStr_GET_part1 & SQLStr_GET_part2 & "[2nd PIC],trim([AA Projects] + ' ' + [ONDEMAND PROJECT]),[Comment of Requestor],[Reference Number from feedback (AA-AS)],[Data Storage Folder] " & where_condition
                            

                            SQLStr_LOCKCOUNT = "SELECT COUNT([2nd PIC]) FROM [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                        " WHERE [2nd PIC] = 'Block by " & PIC & "' AND (([Need to Check] = '' AND [Auto-classify] = '') OR ([Need to Check] <> '' AND [Auto-classify] <> '' )) AND [Logical System Group] = " & "'" & Logicalgroup & "'"
                                            
                                            Set rs1 = New ADODB.RecordSet
                                            rs1.Open SQLStr_LOCKCOUNT, Cn, adOpenForwardOnly, adLockReadOnly
                                            count_locking = rs1(0)
                            '<><><><><><><><><><><> MODIFY UPDATE QUERY STRING FOR EACH LOGICAL SYSTEM GROUP <><><><><><><><><><><>

                        
                        ElseIf Task = "QC 2nd Classification" Then '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                            
                            Task_name = "QC_2nd_classify"
                            task_label = "QC"
                            
                            download_time_header = "[" & "QC down time" & "]"
                            
                            
                            SQLStr_GET_part1 = "SELECT  [Priority]" & _
                                                ",[Logical System Group],[Product Number]" & _
                                                ",[Product Short Text],[Documentation Status] + ' - ' + SUBSTRING([system id documentation status is found],4,3) as [Documentation Status]" & _
                                                ",[Product Type],[Configurable Material],[Term code],[Termcode text]" & _
                                                ",[Prod#hierarchy],[Hierachy Definition]" & _
                                                ",[Where used Part Number],[Where Used Description],[Direct Component],[Direct Usage]" & _
                                                ",[Classification Criteria],[Weight]" & _
                                                ",CAST([Comment by BOT] AS text) as [Comment by BOT],[Hardness Grade on MO table],"
                            
                            
                            where_condition = "FROM [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                " WHERE " & _
                                                    "[QC check by] = 'Block by " & PIC & "'" & " AND " & "[Logical System Group] = " & "'" & Logicalgroup & "' ORDER BY [Priority] ASC,  [Term code]  desc, [Product Number] asc"
                            
                            SQLStr_GET = SQLStr_GET_part1 & SQLStr_GET_part2 & "CAST([Comment by QC] AS text) as [Comment by QC],[2nd PIC],[QC check by],[Product Group before QC check],trim([AA Projects] + ' ' + [ONDEMAND PROJECT]),[Product Group In GTS],[Comment of Requestor],[Reference Number from feedback (AA-AS)],[Data Storage Folder] " & where_condition
                                        

                            SQLStr_LOCKCOUNT = "SELECT COUNT([QC check by]) FROM [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                    " WHERE [QC check by] = 'Block by " & PIC & "'" & _
                                                            "  AND [Logical System Group] = " & "'" & Logicalgroup & "'"

                                            Set rs1 = New ADODB.RecordSet
                                            rs1.Open SQLStr_LOCKCOUNT, Cn, adOpenForwardOnly, adLockReadOnly
                                            count_locking = rs1(0)
                                                      


                            
                            
                            
                        ElseIf Task = "Create new Product Group" Then '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                            
                                                                        Call update_new_group_SQL_WL(USERID_HSCODEDATA) '***********************
                            
                            SQLStr_GET_part1 = "SELECT [Priority]" & _
                                                            ",[Logical System Group],[Product Number]" & _
                                                            ",[Product Short Text],[Documentation Status] + ' - ' + SUBSTRING([system id documentation status is found],4,3) as [Documentation Status]" & _
                                                            ",[Product Type],[Configurable Material],[Term code],[Termcode text]" & _
                                                            ",[Prod#hierarchy],[Hierachy Definition]" & _
                                                            ",[Where used Part Number],[Where Used Description],[Direct Component],[Direct Usage]" & _
                                                            ",[Classification Criteria],[Weight]" & _
                                                            ",CAST([Comment by BOT] AS text) as [Comment by BOT],[Hardness grade on MO Table],"
                            
                            Task_name = "NewGroup_classify"
                            task_label = "NewGroup"
                            
                            download_time_header = "[" & "Create new group down time" & "]"
                            
                            where_condition = "FROM [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                " WHERE " & _
                                                    "[PIC New product group] = 'Block by " & PIC & "'" & " AND " & "[Logical System Group] = " & "'" & Logicalgroup & "' ORDER BY [Priority_termcode] ASC, [Count_termcode] desc, [Term code], [Product Number] asc"
                            
                            
                            SQLStr_GET = SQLStr_GET_part1 & SQLStr_GET_part2 & "CAST([Comment by QC] AS text) as [Comment by QC],[2nd PIC],[QC check by],[PIC New product group]" & _
                                                                               ",CAST([Change history] AS text) as [Change history],[New Product Group],trim([AA Projects] + ' ' + [ONDEMAND PROJECT]),[Product Group In GTS],[Comment of Requestor],[Reference Number from feedback (AA-AS)],[Data Storage Folder] " & where_condition
                            
                            
                            count_locking = 0
                            SQLStr_UPDATE_DOWN = "UPDATE [DB_CTXFC1_SQL].[dbo].[P94_worklist]" & _
                                                            " SET " & _
                                                                "[PIC New product group] = 'Block by " & PIC & "'," & download_time_header & "= GETDATE() " & _
                                                                   "WHERE (([New Product Group] LIKE '%yes%' or [New Product Group] LIKE '%WAIT%') " & _
                                                                        "AND [Logical System Group] = " & "'" & Logicalgroup & "'" & " AND UPPER([P94 Upload status]) not like '%DONE%' ) " & _
                                                                            "AND [QC check by] <> '' AND [QC check by] not like '%Block%'  ORDER BY [Priority_termcode] ASC, [Count_termcode] desc, [Term code], [Product Number] asc"
                        
                        ElseIf Task = "Check Rules Based" Then '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                        

                        
                            Task_name = "Check_Rules_Based"
                            task_label = "Check_Rules_Based"


                         SQLStr_GET = "SELECT [Logical System Group]" & _
                                                                                ",[Product Number]" & _
                                                                                ",[Product Short Text],[Term code text]" & _
                                                                                ",[Term code],[HS code],[Hardness grade]" & _
                                                                                ",[Rule Name],[Number of part similar]" & _
                                                                                ",[Risk point],[Reference Number]" & _
                                                                                ",[Reference Group]" & _
                                                                                ",[Predicted reference group],CAST([Predicted reference number] AS text) AS [Predicted reference number]" & _
                                                                                ",[Where Used Description],[Hierachy Definition],[Hardness grade on MO Table],[Product Group]" & _
                                                                                ",[HS code After clean up],CAST([Comment] AS text) as [Comment],[Risk_checked_by],[New Product Group],[PIC New product group]" & _
                                                                                ",[Ref Part number used],[P94 Upload status],[Ref Part number (Logical System Group)],[Upload time to P94] FROM [DB_CTXFC1_SQL].[dbo].[MO_TABLE_DATA] " & _
                                                                                     " WHERE [Risk_checked_by] = 'Block by " & PIC & "'" & _
                                                                                        " AND [Logical System Group] = " & "'" & Logicalgroup & "'" & " ORDER BY [Term code] asc, [Product Number] asc"


                        Result = MsgBox("Press 'YES' to if you have been assigned for a specific task" & vbNewLine & "Press 'NO' if not ", vbYesNoCancel + vbDefaultButton2)
                        
                        If Result = vbYes Then 'ALREADY BLOCKED ON SYSTEM FOR SPECIFC PART NUMBERS

                                                                                        
                                        file_name = PIC & "_" & Task_name & "_PEP_WEEK_" & Logicalgroup & ".xlsx"
                                        Call copy_worksheet(sqltemplate_path, file_name, task_label) '*************************
                                                    
                                                    
                                        windowfilename = Dir("C:\temp\" & file_name)
                                        Windows(windowfilename).Activate
                                    Workbooks(file_name).Sheets("Sheet1").Activate
                                                                                        
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                                ' MsgBox SQLStr_GET
                                                 rs.Open SQLStr_GET, Cn, adOpenForwardOnly, adLockReadOnly
                                                 
                                                 
                                                         With ActiveWorkbook.Sheets("Sheet1")
                                                         
                                                               .Range("A2" & ":" & "AB1000").ClearContents
                                                               .Range("A2").CopyFromRecordset rs
                                                                MsgBox rs.RecordCount & " parts " & "from task: (" & Task_name & ") has been downloaded"
                                                                
                                                                rs.Close
                                                                Cn.Close
                                                                application.ScreenUpdating = True
                                                                hscodefinder1.Hide
                                                                
                                                                Exit Sub
                                                         End With
                                                                                        
                                                                                        
                        ElseIf Result = vbNo Then 'DAILY TASK
                        Call update_priorityforrulebase_SQL_WL
                        
                SQLStr_LOCKCOUNT = "SELECT COUNT(*) FROM [DB_CTXFC1_SQL].[dbo].[MO_TABLE_DATA] " & _
                                                        " WHERE [Risk_checked_by] = 'Block by " & PIC & "'" & " AND [Logical System Group] = " & "'" & Logicalgroup & "'"

                                                                    Set rs1 = New ADODB.RecordSet
                                                                    rs1.Open SQLStr_LOCKCOUNT, Cn, adOpenForwardOnly, adLockReadOnly
                                                                    count_locking = rs1(0)
                                                                    
                If (200 - count_locking) < 100 Then
            
                    MsgBox "You are locking: " & count_locking & " parts" & vbNewLine & UCase(Task_name) & " task"
                        
                        
                        Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                       'MsgBox SQLStr_GET
                        rs.Open SQLStr_GET, Cn, adOpenForwardOnly, adLockReadOnly
                        
                        file_name = PIC & "_" & Task_name & "_PEP_WEEK_" & Logicalgroup & ".xlsx"
                        Call copy_worksheet(sqltemplate_path, file_name, task_label) '*************************
                                                    
                                                    
                                    windowfilename = Dir("C:\temp\" & file_name)
                                    Windows(windowfilename).Activate
                                    Workbooks(file_name).Sheets("Sheet1").Activate
                        
                                With ActiveWorkbook.Sheets("Sheet1")
                                
                                     .Range("A2" & ":" & "AB1000").ClearContents
                                      .Range("A2").CopyFromRecordset rs
                                       MsgBox rs.RecordCount & " parts " & "from task: (" & Task_name & ") has been downloaded"
                                End With

            
                                 
                        
                    application.ScreenUpdating = True
                    'MsgBox SQLStr_GET
            End If
                         
                        rs.Close
                        Cn.Close
                        application.ScreenUpdating = True
                        hscodefinder1.Hide
                        Exit Sub


End Sub

Sub SQL_upload()
 
'------------clear all filter and unhide all cells-----------
  On Error Resume Next
    Sheet1.ShowAllData
    Sheet1.Columns.EntireColumn.Hidden = False
    Sheet1.Rows.EntireRow.Hidden = False
  On Error GoTo 0
'------------clear all filter and unhide all cells-----------

'-------------BACK UP FILE BEFORE SAVING TO SERVER--------------------
filelink_backup = "C:\TEMP" & "\" & ActiveWorkbook.name

If Dir("C:\backupPEPfile_before_save\", vbDirectory) = vbNullString Then
VBA.MkDir "C:\backupPEPfile_before_save\"
End If

Set fs1 = CreateObject("Scripting.FileSystemObject")


fs1.COPYFILE filelink_backup, "C:\backupPEPfile_before_save\"
'Set file_backup = GetObject(filelink_backup)
'file_backup.ChangeFileAccess Mode:=xlReadOnly

filelink_delete = "C:\backupPEPfile_before_save\" & ActiveWorkbook.name
'-----------------------------------------------------------------------


application.ScreenUpdating = False
                                 'MsgBox ActiveWorkbook.Path
                                If InStr(ActiveWorkbook.path, "backupPEPfile") > 0 Then
                                      MsgBox "You are not allowed to upload data from any file in 'backupPEPfile' folder"
                                      Exit Sub
                                ElseIf InStr(ActiveWorkbook.name, "Y_") = 0 And InStr(ActiveWorkbook.name, "ECCN") = 0 Then
                                      MsgBox "Please stay at PEP WEEK file before saving"
                                      Exit Sub

                                End If
                                
    Set ObjWshNw = CreateObject("WScript.Network")
    
    
    USERID_HSCODEDATA = ObjWshNw.UserName
    PIC = PIC_(USERID_HSCODEDATA)
    

    
    '------------------------------------********** SET UP CONNECTION TO MS SQL SERVER *************------------------------------------------
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    
    Dim Server_Name As String
    Dim Database_Name As String
    Dim user_id As String
    Dim Password As String
    
    Dim rs As ADODB.RecordSet
    Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
    

    
    Server_Name = "tcp:SGPVSQL58.apac.bosch.com,1433" ' Enter your server name here
    Database_Name = "DB_CTXFC1_SQL" ' Enter your database name here
    user_id = "BOSCH\WOM.C_TXF-C1-INT" ' enter your user ID here
    Password = "" ' Enter your password here


    Set Cn = New ADODB.Connection
    Cn.ConnectionString = "Driver={ODBC Driver 17 for SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & user_id & ";Pwd=" & Password & _
    ";Trusted_Connection=yes;Encrypt=yes;TrustServerCertificate=yes;DataTypeCompatibility=80;MultipleActiveResultSets=True;"
    
    Cn.ConnectionTimeout = 100
    Cn.Open
    '----------------------------------------------************************************--------------------------------------------------------
    PRODUCT_COL = 0: partnum_col = 0: hs_col = 0: comment_col = 0: hardness_col = 0: qc_comment_col = 0: qc_check_col = 0: newgroup_check_col = 0: needtocheck_col = 0: NEW_PRODUCT_GROUP_STATUS = 0
            
            For j = 1 To 50
               If Trim(Sheets("Sheet1").Cells(1, j)) = "Product Group" Then
                   PRODUCT_COL = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Product Number" Then
                   partnum_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "ECCN" Then
                   eccn_col = j
               
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "HS code After clean up" Then
                   hs_col = j
               
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Reason" Then
                   reason_col = j
               
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Comment" Then
                   comment_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "ECCN Comment" Then 'for task QC ECCN only
                   eccn_comment_col = j
                   
               ElseIf UCase(Trim(Sheets("Sheet1").Cells(1, j))) = "HARDNESS GRADE ON MO TABLE" Then
                   hardness_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Comment by QC" Then
                   qc_comment_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "QC check by" Then
                   qc_check_col = j
               
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "QC ECCN" Then 'for task QC ECCN only
                   qc_eccn_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "ECCN PIC" Then 'for task ECCN only
                   eccn_pic_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "PIC New product group" Then
                   newgroup_check_col = j
                
                ElseIf UCase(Trim(Sheets("Sheet1").Cells(1, j))) = "NEW PRODUCT GROUP" Then
                   newgroup_STATUS_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Need to Check" Then
                   needtocheck_col = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Product Group In GTS" Then
                   group_in_GTS_col = j
               
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Check by" Then 'for task spliting group
                   split_group_check = j
                   
               ElseIf Trim(Sheets("Sheet1").Cells(1, j)) = "Risk_checked_by" Then 'for task check rules based only [Product Group In GTS]
                   risk_check_col = j
               End If
            Next j
            
                If PRODUCT_COL = 0 Or partnum_col = 0 Then
                    MsgBox "Invalid template!!!" & Chr(13) & "Missing column 'Part Number' or 'Product Group'"
                    Exit Sub
                End If
                        
                        col = Mid(Sheets("Sheet1").Cells(1, PRODUCT_COL).Address(ColumnAbsolute:=True), 2, 1)
                        sort_range = col & ":" & col
                                            
                                            ActiveWorkbook.Activate
                                    
                         ActiveWorkbook.Sheets("Sheet1").AutoFilter.Sort.SortFields.Clear
                         ActiveWorkbook.Sheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
                            (sort_range), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                            xlSortNormal
                            With ActiveWorkbook.Sheets("Sheet1").AutoFilter.Sort
                                 .Header = xlYes
                                 .MatchCase = False
                                 .Orientation = xlTopToBottom
                                 .SortMethod = xlPinYin
                                 .Apply
                            End With

   '----------------------------------------------*******--------------------------------------------------------
            
            
  With Sheets("Sheet1")
            
        lastrow_pep = .Cells(.Cells.Rows.Count, partnum_col).End(xlUp).Row
                                            
                        If InStr(ActiveWorkbook.name, "2nd_classify") > 0 And InStr(ActiveWorkbook.name, "QC_2nd_classify") = 0 Then '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                             
                             count_done = 0
                             For i = 2 To lastrow_pep
                                
                                If .Cells(i, PRODUCT_COL) <> "" And .Cells(i, partnum_col) <> "" Then
                                    
                                     If InStr(UCase(Trim(.Cells(i, PRODUCT_COL))), "UNCLASS") > 0 Then
                                        .Cells(i, PRODUCT_COL) = "UNCLASSIFY"
                                        application.Wait (Now + TimeValue("0:00:01"))
                                     End If
                                     
                                    '----------------------------------------------------------
                                    'checking for comment weight
            
                                        If InStr(.Cells(i, PRODUCT_COL), "KG") <> 0 And .Cells(i, comment_col) = "" Then
                                            MsgBox "Line: " & i & " comment not found. As this group requires to check product's weight, please fill the weight in comment COL O."
                                            .Cells(i, comment_col).Select
                                            Exit Sub
                                        End If
                                    '---------------------------------------------------------
                                    
                                    '----------------------------------------------------------
                                    'remove "'" out of comment
            
                                        If InStr(.Cells(i, comment_col), "'") <> 0 Then
                                            .Cells(i, comment_col) = Replace(.Cells(i, comment_col), "'", "")
                                        End If

                                    '---------------------------------------------------------
                                    
                                    
                                    '----------------------------------------------------------
                                    'remove ";" "|" "#" out of product group
            
                                        If InStr(.Cells(i, PRODUCT_COL), ";") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), ";", "")
                                        End If
                                        
                                        If InStr(.Cells(i, PRODUCT_COL), "|") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), "|", "")
                                        End If
                                        
                                       If InStr(.Cells(i, PRODUCT_COL), "#") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), "#", "")
                                        End If
                                        
                                    '---------------------------------------------------------
                                     
                                                                         
                                     If .Cells(i, PRODUCT_COL) = "UNCLASSIFY" And .Cells(i, reason_col) = "" Then
                                        MsgBox "Please specify reason for group 'UNCLASSIFY'" & " at row: " & i
                                        Exit Sub
                                     ElseIf .Cells(i, PRODUCT_COL) = "UNCLASSIFY" And .Cells(i, hs_col) <> "" Then
                                        .Cells(i, hs_col) = ""
                                     End If
                                     
                                     
                                     If .Cells(i, PRODUCT_COL) <> "UNCLASSIFY" And .Cells(i, hs_col) = "" Then
                                        MsgBox "Please fill hs code for group " & "'" & .Cells(i, PRODUCT_COL) & "'" & " at row: " & i
                                        Exit Sub
                                     End If
                                     
                                    '----------------------------------------------------------
                                        'Loc - checking for screw group

                                        If Cells(i, hs_col) = "731814" Or Cells(i, hs_col) = "731815" Or Cells(i, hs_col) = "731816" Then
                                        'IF HS CODE IS SCREW

                                        For K = 2 To LastRow_hscodedata
                                        'GO THROUGH HSCODEDATA

                                            If HSCODEDATA_ref(K, 3) <> "" Then 'IF THERE R SUB GROUPS

                                                If InStr(HSCODEDATA_ref(K, 3), Cells(i, PRODUCT_COL)) <> 0 Then 'IF USING SUB GROUP TO CLASSIFY

                                                    Cells(i, PRODUCT_COL) = HSCODEDATA_ref(K, 0) 'CHANGE BACK TO MAIN GORUP

                                                    Cells(i, comment_col) = Cells(i, comment_col) & " ,THIS GROUP HAS BEEN AUTOMATICALLY CHANGED TO ITS MAIN GROUP"
                                                    'ADDING CMT
                                                    Exit For
                                                End If
                                            End If
                                        Next K
                                        End If
                                     '---------------------------------------------------------
                                End If
                                
                             Next i
                            '#########################################################################################
                            'progress_form.Show
                            For i = 2 To lastrow_pep
                            
                                If .Cells(i, PRODUCT_COL) <> "" And .Cells(i, partnum_col) <> "" Then
                                
                                    history_changed = PIC & "|" & .Cells(i, PRODUCT_COL) & "|" & "2nd Classification" & "|" & Format(Now(), "DD-MM-YYYY hh mm AMPM") & "|" & "" & ";"
                                    
                                    
                                    SQLStr = "UPDATE [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                "SET " & _
                                                    "[Product Group] = " & "'" & Trim(.Cells(i, PRODUCT_COL)) & "'" & "," & _
                                                    "[HS code After clean up] = " & "'" & .Cells(i, hs_col) & "'" & "," & _
                                                    "[Reason] = " & "'" & .Cells(i, reason_col) & "'" & "," & _
                                                    "[Comment] = [Comment] + " & "'" & .Cells(i, comment_col) & ";'" & "," & _
                                                    "[2nd PIC] = " & "'" & PIC & "'" & "," & _
                                                    "[Hardness Grade on MO table] = " & "'" & Trim(UCase(.Cells(i, hardness_col))) & "'" & "," & _
                                                    "[User ID of PIC 2nd Classification] = " & "'" & UCase(USERID_HSCODEDATA) & "'" & "," & _
                                                    "[Change history] = [Change history] + " & "'" & history_changed & "'" & "," & _
                                                    "[HS Code before QC check] = " & "'" & .Cells(i, hs_col) & "'" & "," & _
                                                    "[Product Group before QC check] = " & "'" & Trim(.Cells(i, PRODUCT_COL)) & "'" & "," & _
                                                    "[2nd classify upload time] = " & "GETDATE() " & _
                                                        " WHERE " & _
                                                              "[Product Number] = " & "'" & Trim(.Cells(i, partnum_col)) & "'" & " AND " & "[2nd PIC]  like '%Block%'"
                                      

                                                    Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                                    rs.Open SQLStr, Cn, adOpenForwardOnly, adLockReadOnly
                                                    count_done = count_done + 1
                                       


                                End If
                                                            progress_form.progress_bar_1.Visible = True
                                                            progress_status = (i * 258) / lastrow_pep
                                                            progress_percentage = (i / lastrow_pep) * 100
                                                            progress_form.progress_bar_1.Width = Round(progress_status)
                                                            progress_form.progress_bar_1.caption = Round(progress_percentage) & "%"

                                                            DoEvents
                             Next i
                                            
                                            Call SQL_update_auto_QC(USERID_HSCODEDATA) ' ///////////////////////////
                                            
                                            
                        ElseIf InStr(ActiveWorkbook.name, "QC_2nd_classify") > 0 Then '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                             
                             count_done = 0
                             For i = 2 To lastrow_pep
                                If .Cells(i, PRODUCT_COL) <> "" And .Cells(i, partnum_col) <> "" Then
                                     
                                     If InStr(UCase(Trim(.Cells(i, PRODUCT_COL))), "UNCLASS") > 0 Then
                                        .Cells(i, PRODUCT_COL) = "UNCLASSIFY"
                                        application.Wait (Now + TimeValue("0:00:01"))
                                     End If
                                                                     
                                    'remove "'" out of comment
            
                                        If InStr(.Cells(i, comment_col), "'") <> 0 Then
                                            .Cells(i, comment_col) = Replace(.Cells(i, comment_col), "'", "")
                                        End If
                                        
                                        If InStr(.Cells(i, qc_comment_col), "'") <> 0 Then
                                            .Cells(i, qc_comment_col) = Replace(.Cells(i, qc_comment_col), "'", "")
                                        End If
                                    '---------------------------------------------------------
                                    
                                    '----------------------------------------------------------
                                    'remove ";" "|" "#" out of product group
            
                                        If InStr(.Cells(i, PRODUCT_COL), ";") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), ";", "")
                                        End If
                                        
                                        If InStr(.Cells(i, PRODUCT_COL), "|") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), "|", "")
                                        End If
                                        
                                       If InStr(.Cells(i, PRODUCT_COL), "#") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), "#", "")
                                        End If
                                        
                                    '---------------------------------------------------------
                                     
                                                                         
                                     If .Cells(i, PRODUCT_COL) = "UNCLASSIFY" And .Cells(i, reason_col) = "" Then
                                        MsgBox "Please specify reason for group 'UNCLASSIFY'" & " at row: " & i
                                        Exit Sub
                                     ElseIf .Cells(i, PRODUCT_COL) = "UNCLASSIFY" And .Cells(i, hs_col) <> "" Then
                                        .Cells(i, hs_col) = ""
                                     End If
                                     
                                     
                                     If .Cells(i, PRODUCT_COL) <> "UNCLASSIFY" And .Cells(i, hs_col) = "" Then
                                        MsgBox "Please fill hs code for group " & "'" & .Cells(i, PRODUCT_COL) & "'" & " at row: " & i
                                        Exit Sub
                                     End If
                                     
                                    '----------------------------------------------------------
                                        'Loc - checking for screw group

                                        If Cells(i, hs_col) = "731814" Or Cells(i, hs_col) = "731815" Or Cells(i, hs_col) = "731816" Then
                                        'IF HS CODE IS SCREW

                                        For K = 2 To LastRow_hscodedata
                                        'GO THROUGH HSCODEDATA

                                            If HSCODEDATA_ref(K, 3) <> "" Then 'IF THERE R SUB GROUPS

                                                If InStr(HSCODEDATA_ref(K, 3), Cells(i, PRODUCT_COL)) <> 0 Then 'IF USING SUB GROUP TO CLASSIFY


                                                    Cells(i, PRODUCT_COL) = HSCODEDATA_ref(K, 0) 'CHANGE BACK TO MAIN GORUP

                                                    Cells(i, qc_comment_col) = Cells(i, qc_comment_col) & " ,THIS GROUP HAS BEEN AUTOMATICALLY CHANGED TO ITS MAIN GROUP"
                                                    'ADDING CMT
                                                    Exit For
                                                End If
                                            End If
                                        Next K
                                        End If
                                     '---------------------------------------------------------
                                     
                                End If
                                
                             Next i
                             
                            '#########################################################################################
                            
                            
                            For i = 2 To lastrow_pep
                                
                                If .Cells(i, PRODUCT_COL) <> "" And .Cells(i, partnum_col) <> "" Then
                                       
                                       If InStr(.Cells(i, qc_check_col), "Block by") = 0 And .Cells(i, qc_check_col) <> "" Then

                                                If Trim(UCase(.Cells(i, hardness_col))) = "" And .Cells(i, PRODUCT_COL) <> "" And UCase(.Cells(i, PRODUCT_COL)) <> "UNCLASSIFY" Then
                                                    hardness = "990"
                                                Else
                                                    hardness = Trim(UCase(.Cells(i, hardness_col)))
                                                End If
                                            
                                            history_changed = PIC & "|" & .Cells(i, PRODUCT_COL) & "|" & "QC 2nd Classification" & "|" & Format(Now(), "DD-MM-YYYY hh mm AMPM") & "|" & "" & ";"
                                            
                                            SQLStr = "UPDATE [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                        "SET " & _
                                                            "[Product Group] = " & "'" & Trim(.Cells(i, PRODUCT_COL)) & "'" & "," & _
                                                            "[HS code After clean up] = " & "'" & .Cells(i, hs_col) & "'" & "," & _
                                                            "[Reason] = " & "'" & .Cells(i, reason_col) & "'" & "," & _
                                                            "[Hardness Grade on MO table] = " & "'" & hardness & "'" & "," & _
                                                            "[User ID of PIC 2nd Classification] = " & "'" & hardness & "'" & "," & _
                                                            "[Comment by QC] = [Comment by QC] + " & "'" & .Cells(i, qc_comment_col) & ";'" & "," & _
                                                            "[QC check by] = " & "'" & .Cells(i, qc_check_col) & "'" & "," & _
                                                            "[Change history] = [Change history] + " & "'" & history_changed & "'" & "," & _
                                                            "[Product Group before New Group] = " & "'" & Trim(.Cells(i, PRODUCT_COL)) & "'" & "," & _
                                                            "[Product Group before AI vs QC check] = " & "'" & Trim(.Cells(i, PRODUCT_COL)) & "'" & "," & _
                                                            "[QC upload time] = " & "GETDATE() ," & _
                                                            "[Product Group In GTS] = " & "'" & Trim(.Cells(i, group_in_GTS_col)) & "'" & _
                                                                " WHERE " & _
                                                                      "[Product Number] = " & "'" & Trim(.Cells(i, partnum_col)) & "'" & " AND " & "[QC check by]  like '%Block%'"
                                               'MsgBox SQLStr
                                                                             Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                                                             rs.Open SQLStr, Cn, adOpenForwardOnly, adLockReadOnly
                                                                             count_done = count_done + 1
                                                                             


                                        End If
                                  
                                  End If
                                                            progress_form.progress_bar_1.Visible = True
                                                            progress_status = (i * 258) / lastrow_pep
                                                            progress_percentage = (i / lastrow_pep) * 100
                                                            progress_form.progress_bar_1.Width = Round(progress_status)
                                                            progress_form.progress_bar_1.caption = Round(progress_percentage) & "%"

                                                            DoEvents
                                Next i
                                
                        
                        Call SQL_update_auto_QC(USERID_HSCODEDATA)
                        
                        ElseIf InStr(ActiveWorkbook.name, "NewGroup_classify") > 0 Then '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                             
                             count_done = 0
                             For i = 2 To lastrow_pep
                                If .Cells(i, PRODUCT_COL) <> "" And .Cells(i, partnum_col) <> "" Then
                                     
                                     If InStr(UCase(Trim(.Cells(i, PRODUCT_COL))), "UNCLASS") > 0 Then
                                        .Cells(i, PRODUCT_COL) = "UNCLASSIFY"
                                        application.Wait (Now + TimeValue("0:00:01"))
                                     End If
                                    
                                    '----------------------------------------------------------
                                    'remove "'" out of comment
            
                                        If InStr(.Cells(i, comment_col), "'") <> 0 Then
                                            .Cells(i, comment_col) = Replace(.Cells(i, comment_col), "'", "")
                                        End If
                                        
                                        If InStr(.Cells(i, qc_comment_col), "'") <> 0 Then
                                            .Cells(i, qc_comment_col) = Replace(.Cells(i, qc_comment_col), "'", "")
                                        End If
                                    '---------------------------------------------------------
                                    
                                    '----------------------------------------------------------
                                    'remove ";" "|" "#" out of product group
            
                                        If InStr(.Cells(i, PRODUCT_COL), ";") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), ";", "")
                                        End If
                                        
                                        If InStr(.Cells(i, PRODUCT_COL), "|") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), "|", "")
                                        End If
                                        
                                       If InStr(.Cells(i, PRODUCT_COL), "#") <> 0 Then
                                            .Cells(i, PRODUCT_COL) = Replace(.Cells(i, PRODUCT_COL), "#", "")
                                        End If
                                        
                                    '---------------------------------------------------------
                                     
                                                                         
                                     If .Cells(i, PRODUCT_COL) = "UNCLASSIFY" And .Cells(i, reason_col) = "" Then
                                        MsgBox "Please specify reason for group 'UNCLASSIFY'" & " at row: " & i
                                        Exit Sub
                                     ElseIf .Cells(i, PRODUCT_COL) = "UNCLASSIFY" And .Cells(i, hs_col) <> "" Then
                                        .Cells(i, hs_col) = ""
                                     End If
                                     
                                     
                                     If .Cells(i, PRODUCT_COL) <> "UNCLASSIFY" And .Cells(i, hs_col) = "" Then
                                        MsgBox "Please fill hs code for group " & "'" & .Cells(i, PRODUCT_COL) & "'" & " at row: " & i
                                        Exit Sub
                                     End If
                                     
                                End If
                                
                             Next i
                             '#########################################################################################
                             'progress_form.Show
                             For i = 2 To lastrow_pep
                            
                                If .Cells(i, PRODUCT_COL) <> "" And .Cells(i, partnum_col) <> "" Then
                                
                                     If InStr(.Cells(i, newgroup_check_col), "Block by") = 0 And .Cells(i, newgroup_check_col) <> "" Then
                                         
                                         
                                               If Trim(UCase(.Cells(i, hardness_col))) = "" And .Cells(i, PRODUCT_COL) <> "" And UCase(.Cells(i, PRODUCT_COL)) <> "UNCLASSIFY" Then
                                                    hardness = "990"
                                                Else
                                                    hardness = Trim(UCase(.Cells(i, hardness_col)))
                                                End If
                                                
                                                                                         
                                               If Trim(UCase(.Cells(i, newgroup_STATUS_col))) <> "WAIT" And Trim(UCase(.Cells(i, newgroup_STATUS_col))) <> "RECHECK" Then
                                                    NEW_PRODUCT_GROUP_STATUS = "Done"
                                                Else
                                                    NEW_PRODUCT_GROUP_STATUS = Trim(UCase(.Cells(i, newgroup_STATUS_col)))
                                                End If
                                                
                                        NEW_COMMENT = "CREATE NEW COMMENT: " & .Cells(i, comment_col)
                                        
                                        If NEW_PRODUCT_GROUP_STATUS = "Done" Then 'normal workflow
                                        
                                        
                                        history_changed = PIC & "|" & .Cells(i, PRODUCT_COL) & "|" & "Create new Product Group" & "|" & Format(Now(), "DD-MM-YYYY hh mm AMPM") & "|" & "" & ";"
                                            
                                         SQLStr = "UPDATE [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                     "SET " & _
                                                         "[Product Group] = " & "'" & Trim(.Cells(i, PRODUCT_COL)) & "'" & "," & _
                                                         "[HS code After clean up] = " & "'" & .Cells(i, hs_col) & "'" & "," & _
                                                         "[Reason] = " & "'" & .Cells(i, reason_col) & "'" & "," & _
                                                         "[Comment] = [Comment] + " & "'" & NEW_COMMENT & ";'" & "," & _
                                                         "[Hardness Grade on MO table] = " & "'" & hardness & "'" & "," & _
                                                        "[User ID of PIC 2nd Classification] = " & "'" & hardness & "'" & "," & _
                                                         "[Change history] = [Change history] + " & "'" & history_changed & "'" & "," & _
                                                         "[New Product Group] = " & "'" & NEW_PRODUCT_GROUP_STATUS & "'" & "," & _
                                                         "[PIC New product group] = " & "'" & .Cells(i, newgroup_check_col) & "'" & "," & _
                                                         "[Create new group up time] = " & "GETDATE(), " & _
                                                         "[Product Group In GTS] = " & "'" & Trim(.Cells(i, group_in_GTS_col)) & "'" & _
                                                             " WHERE " & _
                                                                   "[Product Number] = " & "'" & Trim(.Cells(i, partnum_col)) & "'" & " AND [PIC New product group]  like '%Block%'"
                                    'MsgBox SQLStr
                                                                        Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                                                        rs.Open SQLStr, Cn, adOpenForwardOnly, adLockReadOnly
                                                                        count_done = count_done + 1
                                                                



                                        ElseIf NEW_PRODUCT_GROUP_STATUS = "WAIT" Then

                                         SQLStr = "UPDATE [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                     "SET " & _
                                                         "[New Product Group] = " & "'" & NEW_PRODUCT_GROUP_STATUS & "'" & "," & _
                                                         "[Comment] = [Comment] + " & "'" & NEW_COMMENT & ";'" & "," & _
                                                             " WHERE " & _
                                                                   "[Product Number] = " & "'" & Trim(.Cells(i, partnum_col)) & "'" & " AND [PIC New product group]  like '%Block%'"
                                                                        
                                                                        Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                                                        rs.Open SQLStr, Cn, adOpenForwardOnly, adLockReadOnly
                                                                        count_done = count_done + 1
                                        
                                        ElseIf NEW_PRODUCT_GROUP_STATUS = "RECHECK" Then
                                        
                                        SQLStr = "UPDATE [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                                     "SET " & _
                                                         "[Product Group] = ''," & _
                                                         "[HS code After clean up] = ''," & _
                                                         "[Reason] = ''," & _
                                                         "[Comment] = [Comment] + " & "'" & NEW_COMMENT & ";'" & "," & _
                                                         "[2nd PIC] = ''," & _
                                                         "[QC check by] = ''," & _
                                                         "[New Product Group] = ''," & _
                                                         "[PIC New product group] = ''," & _
                                                         "[Hardness Grade on MO table] = ''," & _
                                                        "[User ID of PIC 2nd Classification] = ''," & _
                                                             " WHERE " & _
                                                                   "[Product Number] = " & "'" & Trim(.Cells(i, partnum_col)) & "'" & " AND [PIC New product group]  like '%Block%'"
                                                                   
                                                                   
                                                                        Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                                                        rs.Open SQLStr, Cn, adOpenForwardOnly, adLockReadOnly
                                                                        count_done = count_done + 1
                                        
                                        End If
                                                                                 
                                     
                                     End If
                                     
                                 End If
                                 
                                                            progress_form.progress_bar_1.Visible = True
                                                            progress_status = (i * 258) / lastrow_pep
                                                            progress_percentage = (i / lastrow_pep) * 100
                                                            progress_form.progress_bar_1.Width = Round(progress_status)
                                                            progress_form.progress_bar_1.caption = Round(progress_percentage) & "%"

                                                            DoEvents
                            Next i
                        
                             
                        Else
                                    
                                    MsgBox "Please stay at correct PEP file"
                                    Exit Sub
                            
                        End If
                        
    '----------------------*****-----------------------------
    
                                    If count_done = 0 Then
                                            MsgBox "Nothing to upload!!!"
                                                                        
                                            windowfilename = Dir("C:\temp\" & ActiveWorkbook.name)
                                            Windows(windowfilename).Activate
                                            ActiveWorkbook.Sheets("Sheet1").Activate
                                                                        
                                        
                                    Else
                                                MsgBox count_done & " parts number has been uploaded to SQL server"
                                                        
                                                        
                                                  If USERID_HSCODEDATA = "TNP2HC" Or USERID_HSCODEDATA = "MTE1HC" Or USERID_HSCODEDATA = "RON2HC" Or USERID_HSCODEDATA = "NLY2HC" Then
                                                        Call update_new_group_SQL_WL(USERID_HSCODEDATA)
                                                  End If
                                                        
                                                Set rs = Nothing
                                                Cn.Close
                                                
                                                application.ScreenUpdating = True
                                                ActiveWorkbook.Sheets("Sheet1").Activate
                                                    
                                                    If Dir("C:\backupPEPfile\", vbDirectory) = vbNullString Then
                                                    VBA.MkDir "C:\backupPEPfile\"
                                                    End If
                                                    
                                                    Dim backupfolder As String: backupfolder = "C:\backupPEPfile\"
                                                    ActiveWorkbook.Save
                                                    ActiveWorkbook.SaveCopyAs FileName:=backupfolder & "(" & Format(Now(), "DD-MMM-YYYY hh mm AMPM") & " )" & ActiveWorkbook.name
                                                    
                                                    linkofbackupfilereadonly = backupfolder & "(" & Format(Now(), "DD-MMM-YYYY hh mm AMPM") & " )" & ActiveWorkbook.name & ".xlsx"
                                                    application.DisplayAlerts = False
                                                    On Error Resume Next
                                                    VBA.SetAttr linkofbackupfilereadonly, vbReadOnly
                                                    application.DisplayAlerts = True
                                                    ActiveWorkbook.Close SaveChanges:=True
                                    End If
        
    End With
    
'--------------------DELETE FILE IN TEMP AND BACKUPBEFORESAVE----------------------
    If InStr(UCase(filelink_backup), "C:\TEMP") > 0 Then

    Set current_file = GetObject(filelink_backup)
    current_file.Save
    current_file.Close
    Set fs = CreateObject("Scripting.FileSystemObject")


    fs.COPYFILE filelink_backup, "C:\backupPEPfile\"

    Kill filelink_backup
    Kill filelink_delete

    MsgBox ("your temp file was deleted")
    End If
'--------------------DELETE FILE IN TEMP AND BACKUPBEFORESAVE----------------------

hscodefinder1.Hide



    

End Sub

Sub update_new_group_SQL_WL(ByVal USERID As String)

    '------------------------------------********** SET UP CONNECTION TO MS SQL SERVER *************------------------------------------------
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    
    Dim Server_Name As String
    Dim Database_Name As String
    Dim user_id As String
    Dim Password As String
    
    Dim rs As ADODB.RecordSet
    Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
    
    
    
    Server_Name = "tcp:SGPVSQL58.apac.bosch.com,1433" ' Enter your server name here
    Database_Name = "DB_CTXFC1_SQL" ' Enter your database name here
    user_id = "BOSCH\WOM.C_TXF-C1-INT" ' enter your user ID here
    Password = "" ' Enter your password here


    Set Cn = New ADODB.Connection
    Cn.ConnectionString = "Driver={ODBC Driver 17 for SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & user_id & ";Pwd=" & Password & _
    ";Trusted_Connection=yes;Encrypt=yes;TrustServerCertificate=yes;DataTypeCompatibility=80;MultipleActiveResultSets=True;"
    
    Cn.ConnectionTimeout = 1000
    Cn.Open
    '----------------------------------------------************************************--------------------------------------------------------
    
    
    SQL_update_product_group = "UPDATE [dbo].[P94_worklist] " & _
                                    " SET [Product Group] = [Auto-classify] " & _
                                        " WHERE [Auto-classify] <> '' AND [Need to Check] = '' AND [Product Group] = ''" & vbNewLine & vbNewLine & _
                               "UPDATE [dbo].[P94_worklist] " & _
                                    " SET [Product Group] = 'UNCLASSIFY' " & _
                                        " WHERE [Auto-classify] <> '' AND (UPPER([Need to Check]) like '%UNCLASS%' OR UPPER([Auto-classify]) like '%UNCLASS%') AND [Product Group] = ''"
    
    SQL_merge_ref_product_group = "MERGE [dbo].[P94_worklist] t " & _
                                        "USING " & _
                                            "(" & _
                                                   " SELECT * " & _
                                                   " FROM ( " & _
                                                      " SELECT *, " & _
                                                             " row_number() OVER (PARTITION BY [PRODUCT_GROUP] ORDER BY [PRODUCT_GROUP]) AS row_number " & _
                                                      " FROM [DB_CTXFC1_SQL].[dbo].[hscodedata] " & _
                                                      " ) AS ROWS " & _
                                        " WHERE row_number = 1 ) AS s" & _
                                    " ON (t.[Product Group] = s.[PRODUCT_GROUP] AND s.[PART NUMBER] <> '') " & _
                                    " WHEN MATCHED " & _
                                       " THEN UPDATE SET " & _
                                           " t.[Ref Part number used]= s.[PART NUMBER], " & _
                                           " t.[HS code After clean up] = s.[HS CODE], " & _
                                           " t.[New Product Group] = '' " & _
                                    " WHEN NOT MATCHED BY SOURCE AND (t.[Product Group] <> '' and t.[Product Group] <> 'UNCLASSIFY' AND t.[P94 Upload status] = '' and [QC check by] <> '' and [QC check by] not like '%Block%' and t.[New Product Group] <> 'Wait' and t.[New Product Group] <> 'Recheck' ) " & _
                                        " THEN UPDATE SET " & _
                                            " t.[Ref Part number used] = '', " & _
                                            " t.[New Product Group]= 'yes'; "
    
    SQL_update_partref_partnum = "UPDATE [dbo].[P94_worklist] " & _
                                    " SET [P94 Upload status] = 'done', [Ref Part number used] = [Product Number], [QC check by] = 'Ref part' " & _
                                        " WHERE [Product Number] = [Ref Part number used] "
    
    SQL_update_partref_partnum_1 = " WITH WL AS  " & _
                                   " (SELECT * FROM P94_worklist as b " & _
                                          "  WHERE Exists " & _
                                              " (SELECT * FROM hscodedata as a " & _
                                                "  WHERE b.[Product Number] = a.[PART NUMBER] and [P94 Upload status] = '')) " & _
                                 " UPDATE WL SET [P94 Upload status] = 'done', [QC check by] = 'Ref part' "
    
    SQL_finish_update = "UPDATE [dbo].[P94_worklist] " & _
                                    " SET [New Product Group] = '' WHERE [P94 Upload status] = 'done' " & vbNewLine & vbNewLine & _
                         "UPDATE [dbo].[P94_worklist] set [PIC New product group] = '' , [New Product Group] = '' " & _
                                   " WHERE [Auto-classify] = '' AND [New Product Group] = 'yes' AND ([QC check by] like '%Block%' OR [QC check by] = '') " & vbNewLine & vbNewLine & _
                         "UPDATE [dbo].[P94_worklist] " & _
                                    " SET [Ref Part number used] = '',[Ref Part number (Logical System Group)] = '' WHERE [Auto-classify] = '' " & _
                                                                                                           " and ([QC check by] like '%Block%'  or [QC check by] = '')"
                                                                                                           


    '////////////////////////////////////////////////////////////// FOR P94_worklist  ///////////////////////////////////////////////////////////////////////////////
    sql_logical_ref_part1 = "UPDATE  [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                         " SET [Ref Part number (Logical System Group)] = 'Y_UBK' WHERE UPPER([Product Group]) not like '%UNCLASS%' AND [Product Group] <> '' " & _
                                             " AND [Product Group] not like 'DC-%' AND [Product Group] not like 'BT-%' " & _
                                             " AND [Product Group] not like 'PT-%' AND [Product Group] not like 'TT-%' AND [Ref Part number used] <> ''  "
    
    
    sql_logical_ref_part2 = "UPDATE  [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                               " SET [Ref Part number (Logical System Group)] = 'Y_BR' WHERE UPPER([Product Group]) not like '%UNCLASS%' AND [Product Group] <> '' " & _
                                   " AND [Product Group] like 'DC-%' AND [Ref Part number used] <> '' "
    
    
    sql_logical_ref_part3 = "UPDATE  [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                " SET [Ref Part number (Logical System Group)] = 'Y_PT' WHERE UPPER([Product Group]) not like '%UNCLASS%' AND [Product Group] <> '' " & _
                                   " AND [Product Group] like 'PT-%' AND [Ref Part number used] <> '' "
                               
    sql_logical_ref_part4 = "UPDATE  [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                " SET [Ref Part number (Logical System Group)] = 'Y_TT01' WHERE UPPER([Product Group]) not like '%UNCLASS%' AND [Product Group] <> '' " & _
                                  " AND [Product Group] like 'TT-%' AND [Ref Part number used] <> ''  "
                                  
                                  
    sql_logical_ref_part5 = "UPDATE  [DB_CTXFC1_SQL].[dbo].[P94_worklist] " & _
                                " SET [Ref Part number (Logical System Group)] = 'Y_ST' WHERE UPPER([Product Group]) not like '%UNCLASS%' AND [Product Group] <> '' " & _
                                  " AND [Product Group] like 'BT-%' AND [Ref Part number used] <> ''  "
                               
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    SQL_compare_1stgroup_AutoGroup = "UPDATE [dbo].[P94_worklist] " & _
                                            " SET [QC check by] = 'AutoQC (1st)'" & _
                                                " WHERE [Product Group] <> '' and [Auto-classify] <> '' and [Need to check] <> '' " & _
                                                    "AND [Product Group] <> 'UNCLASSIFY' and [Product Group] = [Auto-classify] and [P94 Upload status] = '' and [New Product Group] = '' and [QC check by] = ''"
    
    SQL_compare_REFGroup_AutoGroup = "UPDATE [dbo].[P94_worklist] " & _
                                            " SET [Ref Part number used] = ''" & _
                                                " WHERE [Product Group] <> '' and [Auto-classify] <> '' and [Auto-classify] <> 'UNCLASSIFY' " & _
                                                    "AND [Product Group] <> 'UNCLASSIFY' and [Product Group] <> [Auto-classify] and ([QC check by] like '%Block%'  or [QC check by] = '')"
                               
   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                               
    SQL_update_logical_for_refpart = sql_logical_ref_part1 & vbNewLine & vbNewLine & sql_logical_ref_part2 & vbNewLine & vbNewLine & sql_logical_ref_part3 & vbNewLine & vbNewLine & sql_logical_ref_part4 & vbNewLine & vbNewLine & sql_logical_ref_part5

    'SQL_update_logical_for_refpart_MO = sql_logical_ref_part1_MO & vbNewLine & vbNewLine & sql_logical_ref_part2_MO & vbNewLine & vbNewLine & sql_logical_ref_part3_MO & vbNewLine & vbNewLine & sql_logical_ref_part4_MO & vbNewLine & vbNewLine & sql_logical_ref_part5_MO
    

    
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                            rs.Open SQL_update_product_group, Cn, adOpenForwardOnly, adLockReadOnly
                                            
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient 'auto fill QC check by due to [PRoduct Group]  = [Auto-classify]
                                            rs.Open SQL_compare_1stgroup_AutoGroup, Cn, adOpenForwardOnly, adLockReadOnly
                                            
                                            
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                            rs.Open SQL_merge_ref_product_group, Cn, adOpenForwardOnly, adLockReadOnly
                                            
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                            rs.Open SQL_update_partref_partnum, Cn, adOpenForwardOnly, adLockReadOnly
                                            
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                            rs.Open SQL_update_partref_partnum_1, Cn, adOpenForwardOnly, adLockReadOnly
                                            
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient 'delete REF part  due to [PRoduct Group] <> [Auto-classify]
                                            rs.Open SQL_compare_REFGroup_AutoGroup, Cn, adOpenForwardOnly, adLockReadOnly
                                            
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                            rs.Open SQL_finish_update, Cn, adOpenForwardOnly, adLockReadOnly
                                            'MsgBox SQL_finish_update
                                            Set rs = New ADODB.RecordSet: rs.CursorLocation = adUseClient
                                            rs.Open SQL_update_logical_for_refpart, Cn, adOpenForwardOnly, adLockReadOnly

   
End Sub



Sub SQL_interraction()

    application.ScreenUpdating = False
    '------------------------------------********** SET UP CONNECTION TO MS SQL SERVER *************------------------------------------------
    Dim Cn As New ADODB.Connection
    Dim rs1 As New ADODB.RecordSet
    Dim rs2 As New ADODB.RecordSet
    Dim iCols As Integer
    
    Set Cn = New ADODB.Connection
    
    Dim Server_Name As String
    Dim Database_Name As String
    Dim user_id As String
    Dim Password As String
    Dim SQLStr As String
    
    Set ObjWshNw = CreateObject("WScript.Network")
    USERID_HSCODEDATA = ObjWshNw.UserName
    
    Server_Name = "tcp:SGPVSQL58.apac.bosch.com,1433" ' Enter your server name here
    Database_Name = "DB_CTXFC1_SQL" ' Enter your database name here
    user_id = "BOSCH\WOM.C_TXF-C1-INT" ' enter your user ID here
    Password = "" ' Enter your password here
    

    
    Set Cn = New ADODB.Connection
    Cn.ConnectionString = "Driver={ODBC Driver 17 for SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & user_id & ";Pwd=" & Password & _
    ";Trusted_Connection=yes;Encrypt=yes;TrustServerCertificate=yes;DataTypeCompatibility=80;MultipleActiveResultSets=True;"
    
    Cn.ConnectionTimeout = 100
    Cn.Open
    '------------------------------------************************************************************-----------------------------------------



    SQL_get_column_name_p94 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " & _
                           " WHERE TABLE_NAME = 'REFERENCE FULL SCHEME' and [COLUMN_NAME] <> 'Product Number' and [COLUMN_NAME] <> 'PRODUCT GROUP'" & _
                              "  ORDER BY ORDINAL_POSITION "

    
    SQL_get_column_count = "SELECT COUNT(COLUMN_NAME) FROM INFORMATION_SCHEMA.COLUMNS " & _
                           " WHERE TABLE_NAME = 'REFERENCE FULL SCHEME' and [COLUMN_NAME] <> 'Product Number' and [COLUMN_NAME] <> 'PRODUCT GROUP'" & _
                              ""



    Set rs = New ADODB.RecordSet
    rs.Open SQL_get_column_name_p94, Cn, adOpenForwardOnly, adLockReadOnly
    
    
                    Set rs1 = New ADODB.RecordSet
                    rs1.Open SQL_get_column_count, Cn, adOpenForwardOnly, adLockReadOnly
                    
                    'MsgBox rs1(0)
                    
            ReDim SCHEME_LIST(0 To CInt(rs1(0))) As String
            
                    With rs
                        If Not .EOF Then
                            .MoveFirst
                            'hscodefinder1.Scheme.Clear
                            i = 0
                            Do While Not .EOF
                                SCHEME_LIST(i) = .Fields(0).value
                                'hscodefinder1.Scheme.AddItem .Fields(0).value
                                .MoveNext
                                i = i + 1
                            Loop
                        End If

                    End With
    
    hscodefinder1.Scheme.value = ""
  application.ScreenUpdating = True
End Sub













