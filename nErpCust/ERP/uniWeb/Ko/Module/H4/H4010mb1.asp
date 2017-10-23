<%@ LANGUAGE=VBSCript%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	Dim lgStrPrevKey,lgStrPrevKey1
	Const C_SHEETMAXROWS_D = 100
    call LoadBasisGlobalInf()
	Call loadInfTB19029B( "Q", "H","NOCOOKIE","MB")
    
    lgSvrDateTime = GetSvrDateTime
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)    
	If lgCurrentSpd = "M" Then    
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	else
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)    
	end if
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
  
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    If lgCurrentSpd = "M" Then
        Call SubBizQueryMulti()
    Else
        Call SubBizQueryMulti1()
    End if
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strWhere, FirstDate
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If lgKeyStream(7) = "" then
       strWhere = "" & FilterVar("%", "''", "S") & ")"
    Else
       strWhere =  FilterVar(lgKeyStream(7), "''", "S") & ")"
    End if 
             
    strWhere = strWhere & " AND (B.DILIG_DT BETWEEN " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")
    strWhere = strWhere & "                     AND " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(1)),NULL),"NULL","S") & ")"

    If lgKeyStream(5) = "" then
       strWhere = strWhere & " AND ( C.DILIG_CD LIKE " & "" & FilterVar("%", "''", "S") & ")"
    Else
       strWhere = strWhere & " AND ( C.DILIG_CD LIKE " & FilterVar(lgKeyStream(5), "''", "S") & ")"
    End if
    
    If  lgKeyStream(3) = "" then
        strWhere = strWhere & " AND ( A.INTERNAL_CD LIKE  " & FilterVar(Trim(lgKeyStream(4)) & "%", "''", "S") & ") " 
    else
'        strWhere = strWhere & " AND  ( A.INTERNAL_CD  = "  & FilterVar(Trim(lgKeyStream(4)),"''","S") & ")"
' 2003.04.01 by LSM 등록된 근태일이 해당 부서조건을 충족시키는지 (내부부서코드 -> 부서코드로 수정)
 		strWhere = strWhere & " AND (dbo.ufn_H_get_dept_cd(A.emp_no, B.DILIG_DT) = " & FilterVar(lgKeyStream(3), "''", "S") & ") "
    end if
    
    If lgKeyStream(2) = "2" then
       strWhere = strWhere & " AND (A.RETIRE_DT IS NULL OR A.RETIRE_DT >=  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(1)),NULL),"NULL","S") & ")"
    Else
    End if
        
    If lgKeyStream(8) = "3" then
       strWhere = strWhere & "  AND  ( C.DILIG_TYPE LIKE " & "" & FilterVar("%", "''", "S") & "))"
    ElseIf lgKeyStream(8) = "2" then 
       strWhere = strWhere & "  AND  ( C.DILIG_TYPE LIKE " & "" & FilterVar("2", "''", "S") & "))"
    Else
       strWhere = strWhere & "  AND  ( C.DILIG_TYPE LIKE " & "" & FilterVar("1", "''", "S") & " ))"
    End if
       strWhere = strWhere & " ORDER BY A.DEPT_CD ASC, A.EMP_NO ASC, B.DILIG_DT ASC "
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			lgStrPrevKey = ""        
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found.   
			Call SetErrorStatus()
    Else
		Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
			lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("DILIG_DT"),"")
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_HH"), ggAmtOfMoney.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_MM"), ggAmtOfMoney.DecPoint,0)    
            
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
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti1()
    Dim iDx
    Dim iKey1
    Dim strWhere
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If lgKeyStream(7) = "" then
       strWhere = "" & FilterVar("%", "''", "S") & ")"
    Else
       strWhere =  FilterVar(lgKeyStream(7), "''", "S") & ")"
    End if 
         
    If lgKeyStream(0) = "" then
       strWhere = strWhere & " AND  ( A.DILIG_DT >=  '1900" & gComDateType & "01" & gComDateType & "01' )"
       
    Else
       strWhere = strWhere & " AND  ( A.DILIG_DT >= " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S") & ")"
    End if
    
    If lgKeyStream(1) = "" then
       strWhere = strWhere & " AND  ( A.DILIG_DT <=  '2500" & gComDateType & "12" & gComDateType & "31' )"
    Else
       strWhere = strWhere & " AND  ( A.DILIG_DT <=  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(1)),NULL),"NULL","S") & ")"
    End if

    If lgKeyStream(5) = "" then
       strWhere = strWhere & " AND ( A.DILIG_CD LIKE " & "" & FilterVar("%", "''", "S") & ")"
    Else
       strWhere = strWhere & " AND ( A.DILIG_CD LIKE " & FilterVar(lgKeyStream(5), "''", "S") & ")"
    End if
    
    If  lgKeyStream(3) = "" then
        strWhere = strWhere & " AND ( B.INTERNAL_CD LIKE  " & FilterVar(Trim(lgKeyStream(4)) & "%", "''", "S") & ") " 
    else
        strWhere = strWhere & " AND  ( B.INTERNAL_CD  = " & FilterVar(lgKeyStream(4), "''", "S") & ")"
    end if
    
    If lgKeyStream(2) = "2" then
       strWhere = strWhere & " AND (B.RETIRE_DT IS NULL OR b.RETIRE_DT >=  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(1)),NULL),"NULL","S") & ")"
    Else
    End if
        
    If lgKeyStream(8) = "3" then
       strWhere = strWhere & "  AND  ( C.DILIG_TYPE LIKE " & "" & FilterVar("%", "''", "S") & "))"
    ElseIf lgKeyStream(8) = "2" then 
       strWhere = strWhere & "  AND  ( C.DILIG_TYPE LIKE " & "" & FilterVar("2", "''", "S") & "))"
    Else
       strWhere = strWhere & "  AND  ( C.DILIG_TYPE LIKE " & "" & FilterVar("1", "''", "S") & " ))"
    End if
    strWhere = strWhere & " GROUP BY C.DILIG_CD, A.DILIG_CD  ,C.DILIG_NM"
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		lgStrPrevKey1 = ""        
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found.   
         Call SetErrorStatus()
    Else
		Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)
        lgstrData1 = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))
            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs(2), ggAmtOfMoney.DecPoint,0)
            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs(3), ggAmtOfMoney.DecPoint,0)
            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs(4), ggAmtOfMoney.DecPoint,0)
            
            lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
            lgstrData1 = lgstrData1 & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
            iDx =  iDx + 1		    
            If iDx > C_SHEETMAXROWS_D Then
		       lgStrPrevKey1 = lgStrPrevKey1 + 1
               Exit Do
            End If   
		    
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
		lgStrPrevKey1 = ""
    End If   

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx
    Dim iKey1
    Dim iKey2

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        iKey1     = FilterVar(lgKeyStream(0), "''", "S")
        iKey2     = FilterVar(arrColVal(2),""  ,"S")

        Select Case arrColVal(0)
            Case "C"
                    Call SubMakeSQLStatements("MC",iKey1,iKey2,C_EQ)
                    If FncOpenRs(arrColVal(0),lgObjRs,lgStrSQL,"X" ,"X" ) = True Then
                       If FncRsExists(lgObjRs) = True Then
                          Call ServerMesgBox("기존 데이타가 존재 합니다.!", vbCritical, I_MKSCRIPT)	
                          lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
                          ObjectContext.SetAbort
                          Call SetErrorStatus
                       End If
                    Else
                       Call SubBizSaveMultiCreate(lgObjRs,arrColVal)
                    End If
            Case "U"
                    Call SubMakeSQLStatements("MU",iKey1,iKey2,C_EQ)
                    If FncOpenRs(arrColVal(0),lgObjRs,lgStrSQL,"X" ,"X" ) = True Then
                       Call SubBizSaveMultiUpdate(lgObjRs,arrColVal)
                    Else
                       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If 
            Case "D"
                    Call SubMakeSQLStatements("MD",iKey1,iKey2,C_EQ)
                    If FncOpenRs(arrColVal(0),lgObjRs,lgStrSQL,"X" ,"X" ) = True Then
                       Call SubBizSaveMultiDelete(lgObjRs)
                    Else
                       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
        End Select
        
        Call SubCloseRs(lgObjRs)                                                             '☜: Release RecordSSet
        
        If lgErrorStatus    = "YES" Then
           Exit For
        End If
    Next
	
End Sub    

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(lgObjRs,arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO B_MAJOR("
    lgStrSQL = lgStrSQL & " MAJOR_CD     ," 
    lgStrSQL = lgStrSQL & " MINOR_CD     ," 
    lgStrSQL = lgStrSQL & " MINOR_NM     ," 
    lgStrSQL = lgStrSQL & " MINOR_TYPE   ," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "", "D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(lgSvrDateTime),NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(lgSvrDateTime),NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(lgObjRs,arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  B_MINOR"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MAJOR_NM   = " & FilterVar(UCase(arrColVal(3)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " MINOR_TYPE = " & FilterVar(Trim(UCase(arrColVal(4))), "", "D")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " MAJOR_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(lgObjRs)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  B_MINOR"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     MAJOR_CD   = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND MINOR_CD   = " & FilterVar(arrColVal(2),""  ,"S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Select Case Mid(pDataType,1,1)
        Case "M"
			If lgCurrentSpd = "M" Then
			   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
			else 
			   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1
			End If   
           
           Select Case Mid(pDataType,2,1)
               Case "C"

               Case "D"

               Case "R"
                       If lgCurrentSpd = "M" Then
                          lgStrSQL = "SELECT TOP " & iSelCount  & " A.EMP_NO, A.NAME, B.DILIG_DT, C.DILIG_CD, B.DILIG_HH, B.DILIG_MM, C.DILIG_NM, A.DEPT_CD, C.DILIG_TYPE, A.DEPT_NM "
                          lgStrSQL = lgStrSQL & "  FROM HAA010T A, HCA060T B, HCA010T C  "
                          lgStrSQL = lgStrSQL & " WHERE ( B.EMP_NO = A.EMP_NO ) AND  ( B.DILIG_CD = C.DILIG_CD ) AND ( ( B.EMP_NO " & pComp & " " & pCode
                       Else
                          lgStrSQL = " SELECT A.DILIG_CD, C.DILIG_NM, SUM(A.DILIG_CNT), SUM(A.DILIG_HH)  + FLOOR(SUM(A.DILIG_MM) / 60),  CAST(SUM(A.DILIG_MM) AS int) % 60 "
                          lgStrSQL = lgStrSQL & " FROM HCA060T A, HAA010T B, HCA010T C  "
                          lgStrSQL = lgStrSQL & "  WHERE ( B.EMP_NO = A.EMP_NO ) AND  ( A.DILIG_CD = C.DILIG_CD ) AND  ( ( A.EMP_NO  " & pComp & " " &  pCode
                       End If             
               Case "U"

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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
          
              With Parent
                If .lgCurrentSpd = "M" Then
                   .ggoSpread.Source     = .frm1.vspdData
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                   .ggoSpread.SSShowData "<%=lgstrData%>"
					if .topleftOK then
						.DBQueryOk
					else
						.lgCurrentSpd = "S"						
						.DBQuery
					end if
                Elseif .lgCurrentSpd = "S" then
                   .ggoSpread.Source     = .frm1.vspdData1
                   .lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
                   .ggoSpread.SSShowData "<%=lgstrData1%>"   
	               .DBQueryOk                                   
                End If  
      
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
