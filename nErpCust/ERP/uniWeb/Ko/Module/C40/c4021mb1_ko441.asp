<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 1000
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	    
    If lgKeyStream(0) = "" then
       strWhere =  " C.YYYYMM " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere =  " C.YYYYMM  LIKE " & FilterVar(lgKeyStream(0), "''", "S")
    End if 
    
    If lgKeyStream(1) = "" then
       strWhere = strWhere & " AND C.DEPT_CD LIKE " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND C.DEPT_CD LIKE " & FilterVar(lgKeyStream(1), "''", "S")
    End if

    If  lgKeyStream(2) = "" then
        strWhere = strWhere & " AND C.COST_CD LIKE  " & FilterVar(Trim(lgKeyStream(2)) & "%", "''", "S") & " " 
    else
        strWhere = strWhere & " AND C.COST_CD = " & FilterVar(lgKeyStream(2), "''", "S")
    end if

    If  lgKeyStream(3) = "" then
        strWhere = strWhere & " AND C.EMP_NO LIKE  " & FilterVar(Trim(lgKeyStream(3)) & "%", "''", "S") & " " 
    else
        strWhere = strWhere & " AND C.EMP_NO  = " & FilterVar(lgKeyStream(3), "''", "S")
    end if


    Call SubMakeSQLStatements("MR",strWhere,"X","")                                 '¡Ù : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
            		        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))            
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ORG_CHANGE_ID"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DIR_INDIR"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DIR_INDIR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COST_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COST_NM"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NM"))                       
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
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
'-----------
	Dim itxtSpread
	Dim itxtSpreadArr
	Dim itxtSpreadArrCount
	Dim iCUCount
	Dim iDCount
	Dim ii
	
	itxtSpread = ""
	             
	iCUCount = Request.Form("txtCUSpread").Count
	iDCount  = Request.Form("txtDSpread").Count

	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount + iDCount)
	             
	For ii = 1 To iDCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
	Next
	For ii = 1 To iCUCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next
 
   itxtSpread = Join(itxtSpreadArr,"")
'---------        
	arrRowVal = Split(itxtSpread, gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
                                                                   
    lgStrSQL = "INSERT INTO C_INDV_COSTCENTER_KO441 ("
    lgStrSQL = lgStrSQL & " Yyyymm," 
    lgStrSQL = lgStrSQL & " EMP_NO," 
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID," 
    lgStrSQL = lgStrSQL & " DEPT_CD,"
    lgStrSQL = lgStrSQL & " DIR_INDIR,"
    lgStrSQL = lgStrSQL & " COST_CD,"
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD,"                
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES (" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","    
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","         
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

                     	                                 	               
    lgStrSQL = "UPDATE  C_INDV_COSTCENTER_KO441"
    lgStrSQL = lgStrSQL & " SET "                               
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID 	= "     & FilterVar(UCase(arrColVal(4)), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " DEPT_CD  		= "     & FilterVar(UCase(arrColVal(5)), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " DIR_INDIR  		= "     & FilterVar(UCase(arrColVal(6)), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " COST_CD  		= "     & FilterVar(UCase(arrColVal(7)), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD  	= "     & FilterVar(UCase(arrColVal(8)), "''", "S")                
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "   Yyyymm  = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "     AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")    

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    lgStrSQL = "DELETE  C_INDV_COSTCENTER_KO441 "
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "   Yyyymm  = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")   

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
	               lgStrSQL = "Select TOP " & iSelCount  & "  C.EMP_NO,   H.NAME EMP_NAME, C.ORG_CHANGE_ID,  D.DEPT_CD,  D.DEPT_NM,  "
	               lgStrSQL = lgStrSQL & "      C.DIR_INDIR, M.MINOR_NM DIR_INDIR_NM, C.COST_CD,N.COST_NM, C.BIZ_AREA_CD, Z.BIZ_AREA_NM "
	               lgStrSQL = lgStrSQL & " FROM C_INDV_COSTCENTER_KO441 C "
	               lgStrSQL = lgStrSQL & " INNER  JOIN HAA010T  H  ON  C.EMP_NO = H.EMP_NO "
	               lgStrSQL = lgStrSQL & " left outer   JOIN B_ACCT_DEPT   D  ON  C.ORG_CHANGE_ID = D.ORG_CHANGE_ID AND  C.DEPT_CD  = D.DEPT_CD "
	               lgStrSQL = lgStrSQL & " left outer  JOIN  B_MINOR M ON   M.MAJOR_CD  ='H0071'  AND M.MINOR_CD = C.DIR_INDIR "
	               lgStrSQL = lgStrSQL & " left outer  JOIN  B_COST_CENTER  N  ON  C.COST_CD = N.COST_CD "         
	               lgStrSQL = lgStrSQL & " left outer  JOIN  B_BIZ_AREA Z  ON  C.BIZ_AREA_CD = Z.BIZ_AREA_CD "        	               	               	               	               		                                                    
	               lgStrSQL = lgStrSQL & " WHERE   "  & pCode
	               lgStrSQL = lgStrSQL & " ORDER BY C.EMP_NO "
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
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
