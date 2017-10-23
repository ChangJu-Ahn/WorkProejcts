<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  

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
    Dim strEmp_no
    Dim strGazet_start_dt
    Dim strGazet_end_dt
    Dim strGazet_cd
    Dim Rvalue
    Dim strWhere

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strEmp_no = FilterVar(lgKeyStream(0),"'%'", "S")        
    strGazet_cd = FilterVar(lgKeyStream(3),"'%'", "S")
    
    if lgKeyStream(1) = "" then
		strGazet_start_dt = " " & FilterVar(UNIConvDate("1900-01-01"), "''", "S") & ""
	else
		strGazet_start_dt = FilterVar(UNIConvDate(lgKeyStream(1)),"'" & UNIConvDate("1901-01-01") & "'", "S")	
	end if					
    if lgKeyStream(2) = "" then
		strGazet_end_dt = " " & FilterVar(UNIConvDate("2999-12-31"), "''", "S") & ""
	else
		strGazet_end_dt = FilterVar(UNIConvDate(lgKeyStream(2)),"'" & UNIConvDate("2999-12-31") & "'", "S")
	end if		
	
    strWhere = strEmp_no
    strWhere = strWhere & " AND (HBA010T.GAZET_DT >= " & strGazet_start_dt & " AND HBA010T.GAZET_DT <= " & strGazet_end_dt & ")  "
    strWhere = strWhere & " AND HBA010T.GAZET_CD LIKE " & strGazet_cd 
    strWhere = strWhere & " AND HBA010T.EMP_NO = HAA010T.EMP_NO "
    strWhere = strWhere & " AND HAA010T.INTERNAL_CD LIKE  " & FilterVar(lgKeyStream(4) & "%", "''", "S") & ""
    strWhere = strWhere & " ORDER BY HBA010T.GAZET_DT DESC,HBA010T.GAZET_CD DESC,HAA010T.EMP_NO ASC "
	

    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else


        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
                               
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("gazet_dt"),Null)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gazet_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd2"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("entr_dt"),Null)
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("retire_dt"),Null)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sch_ship_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("school_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("major_nm"))
            
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
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
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
                       lgStrSQL = lgStrSQL & " HBA010T.GAZET_DT,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0029", "''", "S") & ",HBA010T.gazet_cd) gazet_nm,"
                       lgStrSQL = lgStrSQL & " HAA010T.NAME,"
                       lgStrSQL = lgStrSQL & " HBA010T.EMP_NO,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetDeptName(HBA010T.dept_cd,HBA010T.gazet_dt) dept_nm  , "
					   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",HBA010T.roll_pstn) roll_pstn_nm  , " 
					   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",HBA010T.pay_grd1) pay_grd1_nm  , "                                               
                       lgStrSQL = lgStrSQL & " HAA010T.ENTR_DT,"
                       lgStrSQL = lgStrSQL & " HAA010T.RETIRE_DT,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0007", "''", "S") & ",HAA010T.sch_ship) sch_ship_nm  , "
                       lgStrSQL = lgStrSQL & " HBA010T.DEPT_CD,"
                       lgStrSQL = lgStrSQL & " HBA010T.ROLL_PSTN,"
                       lgStrSQL = lgStrSQL & " HBA010T.PAY_GRD1,"
                       lgStrSQL = lgStrSQL & " HBA010T.PAY_GRD2,"
                       lgStrSQL = lgStrSQL & " HAA010T.SCH_SHIP,"
                       lgStrSQL = lgStrSQL & " t.SCHOOL_NM,"
                       lgStrSQL = lgStrSQL & " t.MAJOR_NM, "
                       lgStrSQL = lgStrSQL & " HAA010T.INTERNAL_CD "
                       lgStrSQL = lgStrSQL & " FROM HBA010T, HAA010T  left outer join "
                       lgStrSQL = lgStrSQL & "	( select a.EMP_NO,a.SCH_SHIP ,SCHOOL_NM,MAJOR_NM"
                       lgStrSQL = lgStrSQL & "	  from HAA030T a join (select  EMP_NO,SCH_SHIP,max(ADMI_DT) ADMI_DT from HAA030T group by emp_no,SCH_SHIP ) b "
                       lgStrSQL = lgStrSQL & "		on a.EMP_NO=b.EMP_NO and a.SCH_SHIP=b.SCH_SHIP and a.ADMI_DT=b.ADMI_DT"
                       lgStrSQL = lgStrSQL & "  ) t on  t.EMP_NO = HAA010T.EMP_NO AND t.SCH_SHIP = HAA010T.SCH_SHIP "
                       lgStrSQL = lgStrSQL & " WHERE HBA010T.EMP_NO " & pComp & " " &  pCode
'Response.Write lgStrSQL     
                  
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
