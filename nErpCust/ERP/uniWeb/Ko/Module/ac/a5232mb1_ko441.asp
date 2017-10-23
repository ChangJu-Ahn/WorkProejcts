<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%
Response.Expires = -1                                                                '☜ : will expire the response immediately
Response.Buffer = True                                                               '☜ : The server does not send output to the client until all of the ASP scripts on the current page have been processed

%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adoQuery.vbs" -->
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear   

 
    Call HideStatusWnd 
    '---------------------------------------Common-----------------------------------------------------------                                                              '☜: Hide Processing message
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
   
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
             
    End Select
    
    Call SubCloseDB(lgObjConn)
    
    
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
End Sub


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()    
    Dim iLoopMax
    Dim iKey1
    Dim strWhere    
    Dim iDx
                    
    On Error Resume Next    
    Err.Clear                                                               '☜: Clear Error status
               
	 strWhere = FilterVar(lgKeyStream(0), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(1), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(2), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(3), "''", "S") & ", " & _ 
				FilterVar(lgKeyStream(4), "''", "S")   
            

    Call SubMakeSQLStatements("MR", strWhere, "X", C_LIKE)                              '☜ : Make sql statements   
    
    Call GetSum
                          
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKeyIndex = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else
    
       Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex)       
       
       lgStrData = ""
       iDx = 1

       Do While Not lgObjRs.EOF        
       
           lgStrData = lgStrData & Chr(11) & Trim(lgObjRs("acct_cd"))
           lgStrData = lgStrData & Chr(11) & Trim(lgObjRs("acct_nm"))
           lgStrData = lgStrData & Chr(11) & lgObjRs(2)
           lgStrData = lgStrData & Chr(11) & lgObjRs(3)   
           lgStrData = lgStrData & Chr(11) & lgObjRs(4)   
           lgStrData = lgStrData & Chr(11) & lgObjRs(5)   
           lgStrData = lgStrData & Chr(11) & lgObjRs(6)  
           lgStrData = lgStrData & Chr(11) & lgObjRs(7)  
           lgStrData = lgStrData & Chr(11) & lgObjRs(8)  
           lgStrData = lgStrData & Chr(11) & lgObjRs(9)  
           lgStrData = lgStrData & Chr(11) & lgObjRs(10)  
           lgStrData = lgStrData & Chr(11) & lgObjRs(11)  
           lgStrData = lgStrData & Chr(11) & lgObjRs(12)  
           lgStrData = lgStrData & Chr(11) & lgObjRs(13)
           lgStrData = lgStrData & Chr(11) & lgObjRs(14)    
           
           lgStrData = lgStrData & Chr(11) & lgLngMaxRow + iDx
           lgStrData = lgStrData & Chr(11) & Chr(12)                                            
           
           lgObjRs.MoveNext
          
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------                                                                                     
            iDx =  iDx + 1  
            If iDx > lgMaxCount Then               
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1 
               Exit Do
            End If                          
            
        Loop
        
    End If
            
      'lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1       
      If iDx <= lgMaxCount Then
         lgStrPrevKeyIndex = ""
      End If   
      
      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
         ObjectContext.SetAbort
      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"                    
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub      

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

	Dim iCnt, strFg, strType, strOrgId
    On Error Resume Next   
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 
 On Error Resume Next 
 Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
       
  
     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
   
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
     Dim iSelCount
     Dim strZeroFg1
    
    Select Case Mid(pDataType,1,1)
        Case "M"
                  
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           strZeroFg1 = lgKeyStream(5)
           

  
           Select Case Mid(pDataType,2,1)
           
               Case "R"     
					If strZeroFg1 = "Y" then

						lgStrSQL = " SELECT *  "
						lgStrSQL = lgStrSQL & " FROM ufn_a5232ma1_ko441(" & pCode & ") "
'						lgStrSQL = lgStrSQL & " ORDER BY acct_cd "
                    else
						lgStrSQL = " SELECT *  "
						lgStrSQL = lgStrSQL & " FROM ufn_a5232ma1_ko441(" & pCode & ") "
                                                lgStrSQL = lgStrSQL & " WHERE (M01 <> 0 or M02 <> 0 or M03 <> 0 or M04 <> 0 or M05 <> 0 or "
                                                lgStrSQL = lgStrSQL & "  M06 <> 0 or M07 <> 0 or M08 <> 0 or M09 <> 0 or"
                                                lgStrSQL = lgStrSQL & "  M10 <> 0 or M11 <> 0 or M12 <> 0) "
'						lgStrSQL = lgStrSQL & " ORDER BY acct_cd "

                   end if
                   
'						lgStrSQL = " SELECT * , "
'						lgStrSQL = lgStrSQL & " acct_nm = dbo.ufn_x_getcodename('a_acct',acct_cd,'') "
'						lgStrSQL = lgStrSQL & " FROM dbo.ufn_a5277ma1(" & pCode & ") "
'						lgStrSQL = lgStrSQL & " ORDER BY acct_cd "                           
                 

           End Select             
    End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
'call svrmsgbox(lgStrSQL, vbinformation, i_mkscript)
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Response.Write "<BR> Commit Event occur"
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Response.Write "<BR> Abort Event occur"
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MX"
                 Call DisplayMsgBox("800350", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MY"
                 Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MZ"
                 Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)     '이 기간에 대해 이미 입력된 기간근태사항이 있습니다
                 ObjectContext.SetAbort
                 Call SetErrorStatus
    End Select
End Sub


%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData   "<%=lgstrData%>"                
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                
                '--------
               ' .frm1.txtSSumAmt.value  = "<%=liPSum%>"
             '   .frm1.txtTDrAmt.value = "<%=liDrSum%>"
             '   .frm1.txtTCrAmt.value = "<%=liCrSum%>"
             '   .frm1.txtTSumAmt.value   = "<%=liSum%>"
                '------
                
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
