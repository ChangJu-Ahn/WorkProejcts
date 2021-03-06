<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
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
    Call SubBizQueryMulti()
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strRoll_pstn
    Dim strPay_grd1
    dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	dim strNat_cd
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    
    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQ)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("family_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("rel_cd"))
            lgstrData = lgstrData & Chr(11) & ""
         
			call CommonQueryRs(" nat_cd "," HAA010T "," EMP_NO = " &  FilterVar(lgKeyStream(0), "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
			strNat_cd = Replace(lgF0, Chr(11), "")  ' 주민번호 check를 위해서 
			if strNat_cd="KR" then            
				lgstrData = lgstrData & Chr(11) & ConvSPChars(Mid(lgObjRs("res_no"), 1, 6) & "-" & Mid(lgObjRs("res_no"), 7, 7))
			else 
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_no"))
			end if
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sch_ship"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("occup_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("comp_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("supp_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            If ConvSPChars(lgObjRs("reside_type")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If

            If ConvSPChars(lgObjRs("mdcl_insur")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If

            If ConvSPChars(lgObjRs("allow_type")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "1"
            Else
                lgstrData = lgstrData & Chr(11) & "0"
            End If
            
            lgstrData = lgstrData & Chr(11) & C_SHEETMAXROWS_D + iDx
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
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To C_SHEETMAXROWS_D
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HAA020T( emp_no, family_nm, rel_cd, res_no,"
    lgStrSQL = lgStrSQL & " sch_ship, occup_nm, comp_nm, roll_pstn, supp_cd,"
    lgStrSQL = lgStrSQL & " reside_type, mdcl_insur, allow_type, Isrt_emp_no," 
    lgStrSQL = lgStrSQL & " Isrt_dt, Updt_emp_no, Updt_dt      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    ' 주민번호 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    ' 최종학력 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S") & ","
    ' 직업 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S") & ","
    ' 근무회사 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S") & ","
    ' 직위 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(9), "''", "S") & ","
    ' 부양 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S") & ","

    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(12)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(13)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgStrSQL = "UPDATE  HAA020T"
    lgStrSQL = lgStrSQL & " SET " 
    ' 주민번호 
    lgStrSQL = lgStrSQL & " res_no = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    ' 최종학력 
    lgStrSQL = lgStrSQL & " sch_ship = " & FilterVar(UCase(arrColVal(6)), "''", "S") & ","
    ' 직업 
    lgStrSQL = lgStrSQL & " occup_nm = " & FilterVar(arrColVal(7), "''", "S") & ","
    ' 근무회사 
    lgStrSQL = lgStrSQL & " comp_nm = " & FilterVar(arrColVal(8), "''", "S") & ","
    ' 직위 
    lgStrSQL = lgStrSQL & " roll_pstn = " & FilterVar(arrColVal(9), "''", "S") & ","
    ' 부양 
    lgStrSQL = lgStrSQL & " supp_cd = " & FilterVar(UCase(arrColVal(10)), "''", "S") & ","

    lgStrSQL = lgStrSQL & " reside_type = " & FilterVar(UCase(arrColVal(11)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " mdcl_insur = " & FilterVar(UCase(arrColVal(12)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " allow_type = " & FilterVar(UCase(arrColVal(13)), "''", "S") & ","

    lgStrSQL = lgStrSQL & " Updt_emp_no = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " Updt_dt = " & FilterVar(lgSvrDateTime, "''", "S")  

    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   family_nm = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   rel_cd = " & FilterVar(UCase(arrColVal(4)), "''", "S")

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

    lgStrSQL = "DELETE  HAA020T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "         emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   family_nm = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND   rel_cd = " & FilterVar(UCase(arrColVal(4)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
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
                       lgStrSQL = lgStrSQL & " family_nm, rel_cd, res_no, sch_ship, occup_nm, comp_nm,"
                       lgStrSQL = lgStrSQL & " roll_pstn, supp_cd, reside_type, mdcl_insur, allow_type"
                       lgStrSQL = lgStrSQL & " FROM  HAA020T "
                       lgStrSQL = lgStrSQL & " WHERE emp_no " & pComp & pCode
                       lgStrSQL = lgStrSQL & " ORDER BY rel_cd, res_no ASC"
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
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
