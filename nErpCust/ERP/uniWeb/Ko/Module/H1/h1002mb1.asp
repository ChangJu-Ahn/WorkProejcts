<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '��: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)    
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query        
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
   On Error Resume Next
   Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    
    On Error Resume Next    
    Err.Clear
      
    strWhere =  FilterVar(lgKeyStream(0), "''", "S")
    strWhere = strWhere & " AND DILIG_TYPE LIKE " & FilterVar(lgKeyStream(1),"'%'", "S")
    
    If lgKeyStream(2) = "1" Then
        strWhere = strWhere & " ORDER BY DILIG_CD ASC"
    ElseIf lgKeyStream(2) = "2" Then
        strWhere =  strWhere & " ORDER BY DILIG_NM ASC"
    ElseIf lgKeyStream(2) = "3" Then
        strWhere =  strWhere & " ORDER BY DILIG_SEQ ASC"
    End If

    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                              '�� : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKey = ""
    
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))                  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_TIME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_TIME_NM"))
                               
            If ConvSPChars(lgObjRs("BAS_MARGIR")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
                    
            If ConvSPChars(lgObjRs("WK_DAY")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
                    
            If ConvSPChars(lgObjRs("ATTEND_DAY")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
                    
            If ConvSPChars(lgObjRs("WEEK_CNT_APPLY")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
            
            If ConvSPChars(lgObjRs("HOLIDAY_APPLY")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
            
            If ConvSPChars(lgObjRs("SYS_FLAG")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
                               
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_SEQ"))

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
	   	    lgObjRs.MoveNext
          
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
                       
        Loop 
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
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
    Err.Clear 

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
	On Error Resume Next   
	Err.Clear
	dim i
    For i = 6 To 11
		IF arrColVal(i) = "YES" then
		    arrColVal(i) = "Y"
		ElseIF arrColVal(i) = "NO" then
		    arrColVal(i) = "N"
		End IF	       
	Next	 
	lgStrSQL = "INSERT INTO HCA010T("
	lgStrSQL = lgStrSQL & " DILIG_CD , DILIG_SEQ , DILIG_NM , DILIG_TYPE , DAY_TIME," 
	lgStrSQL = lgStrSQL & " BAS_MARGIR , WK_DAY , ATTEND_DAY , WEEK_CNT_APPLY ,"
	lgStrSQL = lgStrSQL & " HOLIDAY_APPLY , SYS_FLAG , ISRT_DT , ISRT_EMP_NO ," 
	lgStrSQL = lgStrSQL & " UPDT_DT , UPDT_EMP_NO )" 
	lgStrSQL = lgStrSQL & " VALUES(" 
	    
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(12))),"0","D")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")    & ","
	lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S")    & ","  
	    
	      
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        
	lgStrSQL = lgStrSQL & ")"
	    
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '��: Clear Error status
    For i = 6 To 11
		IF arrColVal(i) = "YES" then
		    arrColVal(i) = "Y"
		ElseIF arrColVal(i) = "NO" then
		    arrColVal(i) = "N"
		End IF	       
	Next      
  
    lgStrSQL = "UPDATE  HCA010T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " DILIG_NM	   = " & FilterVar(UCase(arrColVal(3)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " DILIG_TYPE     = " & FilterVar(UCase(arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " DAY_TIME       = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " BAS_MARGIR     = " & FilterVar(UCase(arrColVal(6)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " WK_DAY         = " & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " ATTEND_DAY     = " & FilterVar(UCase(arrColVal(8)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " WEEK_CNT_APPLY = " & FilterVar(UCase(arrColVal(9)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " HOLIDAY_APPLY  = " & FilterVar(UCase(arrColVal(10)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " SYS_FLAG       = " & FilterVar(UCase(arrColVal(11)), "''", "S") & ","    
    lgStrSQL = lgStrSQL & " DILIG_SEQ      = " & FilterVar(Trim(UCase(arrColVal(12))),"0","D")    
    lgStrSQL = lgStrSQL & " WHERE            "
    lgStrSQL = lgStrSQL & " DILIG_CD       = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  HCA010T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " DILIG_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)

	 Dim iSelCount
	 
    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

           Select Case Mid(pDataType,2,1)
               Case "R"               
                       lgStrSQL = "SELECT TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & "	DILIG_CD, 	DILIG_SEQ, 	DILIG_NM, DILIG_TYPE, "                         
                       lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0085", "''", "S") & ", DILIG_TYPE) DILIG_TYPE_NM, "                       
                       lgStrSQL = lgStrSQL & "	DAY_TIME, "
                       lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0086", "''", "S") & ", DAY_TIME) DAY_TIME_NM, "
                       lgStrSQL = lgStrSQL & "	BAS_MARGIR, " 
                       lgStrSQL = lgStrSQL & "	WK_DAY, ATTEND_DAY, WEEK_CNT_APPLY, "
                       lgStrSQL = lgStrSQL & "	HOLIDAY_APPLY, SYS_FLAG "   
                       lgStrSQL = lgStrSQL & "  FROM HCA010T "
                       lgStrSQL = lgStrSQL & " WHERE DILIG_CD " & pComp & pCode 
           End Select             
    End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"        
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with        
	         
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
