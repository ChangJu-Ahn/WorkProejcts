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
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = " WHERE a.emp_no = b.emp_no"
    if  lgKeyStream(0) <> "" then
        iKey1 = iKey1 & " AND a.chng_pay_grd1 = " & FilterVar(lgKeyStream(0), "''", "S")
    end if
    if  lgKeyStream(1) <> "" then
        iKey1 = iKey1 & " AND a.chng_pay_grd2 = " & FilterVar(lgKeyStream(1), "''", "S")
    end if
    if  lgKeyStream(2) <> "" then
        iKey1 = iKey1 & " AND a.chng_roll_pstn = " & FilterVar(lgKeyStream(2), "''", "S")
    end if
    if  lgKeyStream(3) <> "" then
        iKey1 = iKey1 & " AND a.chng_cd = " & FilterVar(lgKeyStream(3), "''", "S")
    end if
    if  lgKeyStream(5) <> "" then
        iKey1 = iKey1 & " AND a.chng_dept_cd = " & FilterVar(lgKeyStream(5), "''", "S")
    end if
    if  lgKeyStream(4) <> "" then
        iKey1 = iKey1 & " AND a.dept_cd = " & FilterVar(lgKeyStream(4), "''", "S")
    end if
    if  Trim(lgKeyStream(6)) <> "" then
        iKey1 = iKey1 & " AND a.promote_dt = " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(6)),NULL),"NULL","S")
    end if
    iKey1 = iKey1 & " AND b.internal_cd LIKE  " & FilterVar(lgKeyStream(7) & "%", "''", "S") & ""     ' 자료권한 추가 

    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQGT)                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ocpt_type_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("func_cd_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("role_cd_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("entr_dt"),"")
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("resent_promote_dt"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_pay_grd1_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_pay_grd2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_roll_pstn_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_ocpt_type_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_func_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_role_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("promote_dt"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("chng_nm"))

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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HBA080T"
    lgStrSQL = lgStrSQL & " WHERE EMP_NO     = " & FilterVar(UCase(arrColVal(2)), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND PROMOTE_DT = " & FilterVar(UNIConvDate(arrColVal(3)),NULL,"S") 
    lgStrSQL = lgStrSQL & "   AND CHNG_CD    = " & FilterVar(UCase(arrColVal(4)), "''", "S")    
    
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
                       lgStrSQL = "Select TOP " & iSelCount
                       lgStrSQL = lgStrSQL & " a.EMP_NO,"
                       lgStrSQL = lgStrSQL & " b.name,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetDeptName(a.dept_cd,a.PROMOTE_DT) dept_nm  , "
					   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",a.PAY_GRD1) PAY_GRD1_nm, "	                                                                   
                       lgStrSQL = lgStrSQL & " a.PAY_GRD2,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",a.ROLL_PSTN) ROLL_PSTN_nm, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ",a.OCPT_TYPE) OCPT_TYPE_nm, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0004", "''", "S") & ",a.FUNC_CD) FUNC_CD_nm, "                                              
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ",a.ROLE_CD) ROLE_CD_nm, "                        
                       lgStrSQL = lgStrSQL & " a.ENTR_DT,"
                       lgStrSQL = lgStrSQL & " a.RESENT_PROMOTE_DT,"                     
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetDeptName(a.CHNG_DEPT_CD,a.PROMOTE_DT) CHNG_DEPT_nm, "                                              

                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",a.CHNG_PAY_GRD1) CHNG_PAY_GRD1_nm, "                       
                       lgStrSQL = lgStrSQL & " a.CHNG_PAY_GRD2,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",a.CHNG_ROLL_PSTN) CHNG_ROLL_PSTN_nm, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ",a.CHNG_OCPT_TYPE) CHNG_OCPT_TYPE_nm, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0004", "''", "S") & ",a.CHNG_FUNC_CD) CHNG_FUNC_nm, "                                                                     
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ",a.CHNG_ROLE_CD) CHNG_ROLE_nm, "                       
                       lgStrSQL = lgStrSQL & " a.PROMOTE_DT,"
                       lgStrSQL = lgStrSQL & " a.CHNG_CD,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0029", "''", "S") & ",a.CHNG_CD) CHNG_nm "                       
                       lgStrSQL = lgStrSQL & " FROM HAA010T b, HBA080T a "
                       lgStrSQL = lgStrSQL & pCode
                  
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
