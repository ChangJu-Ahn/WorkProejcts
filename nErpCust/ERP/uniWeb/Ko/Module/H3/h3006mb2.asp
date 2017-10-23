<%@ LANGUAGE=VBSCript%>
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

    Dim strSQL
    Dim strCnhg_pay_grd2
    Dim IntRetCD
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    iKey1 = iKey1 & " WHERE (retire_dt IS NULL OR retire_dt > " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(4)),NULL),"NULL","S") & ")"
    iKey1 = iKey1 & " AND entr_dt <= " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")
    if  lgKeyStream(1) <> "" then
        iKey1 = iKey1 & " AND pay_grd1 = " & FilterVar(lgKeyStream(1), "''", "S")
    end if
    if  lgKeyStream(2) <> "" then
        iKey1 = iKey1 & " AND pay_grd2 = " & FilterVar(lgKeyStream(2), "''", "S")
    end if

    iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(3) & "%", "''", "S") & ""

    iKey1 = iKey1 & " AND (resent_promote_dt IS NULL OR resent_promote_dt <= " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S") & ")"
    iKey1 = iKey1 & " ORDER BY roll_pstn, pay_grd1, pay_grd2"

    Call SubMakeSQLStatements("MR",iKey1,"X","")                                 '☆ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ocpt_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ocpt_type_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("func_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("func_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("role_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("role_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("entr_dt"),"")
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("resent_promote_dt"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1"))

            strSQL = " pay_grd =  " & FilterVar(ConvSPChars(lgObjRs("pay_grd1")), "''", "S") & ""
            Call CommonQueryRs(" pay_grd, hobong_size, hobong_updn "," HDA320T ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            if  Replace(lgF2, Chr(11), "") = "1" then
                strCnhg_pay_grd2 = Right(("000" & Cstr((Cint(lgObjRs("pay_grd2"))+Cint(Replace(lgF1,Chr(11),""))))), len(lgObjRs("pay_grd2")))
            else
                strCnhg_pay_grd2 = Right(("000" & Cstr((Cint(lgObjRs("pay_grd2"))-Cint(Replace(lgF1,Chr(11),""))))), len(lgObjRs("pay_grd2")))
            end if
            strSQL = "     pay_grd1 =  " & FilterVar(ConvSPChars(lgObjRs("pay_grd1")), "''", "S") & ""
            strSQL = strSQL & " AND pay_grd2 =  " & FilterVar(strCnhg_pay_grd2 , "''", "S") & ""
            strSQL = strSQL & " AND apply_strt_dt = (SELECT MAX(apply_strt_dt) FROM hdf010t WHERE apply_strt_dt <= getdate()) "
            Call CommonQueryRs(" pay_grd2 "," HDF010T ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            if  Replace(lgF0, Chr(11), "") <>"X" then
                lgstrData = lgstrData & Chr(11) & Replace(lgF0, Chr(11), "")
            else
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd2"))
            end if
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ocpt_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("func_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("role_cd"))
            lgstrData = lgstrData & Chr(11) & ""    ' 승급예정일 
            lgstrData = lgstrData & Chr(11) & ""    ' 변동사유 
            lgstrData = lgstrData & Chr(11) & ""    ' 변동사유 

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
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
          Select Case Mid(pDataType,2,1)
               Case "R"
			           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
               
                       lgStrSQL =           "SELECT  emp_no, "
                       lgStrSQL = lgStrSQL & "       name, "
                       lgStrSQL = lgStrSQL & "       dept_cd, "
                       lgStrSQL = lgStrSQL & "       dept_nm, "
                       lgStrSQL = lgStrSQL & "       pay_grd1, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",pay_grd1) pay_grd1_nm, "
                       lgStrSQL = lgStrSQL & "       pay_grd2, "
                       lgStrSQL = lgStrSQL & "       roll_pstn, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm, "
                       lgStrSQL = lgStrSQL & "       ocpt_type, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ",ocpt_type) ocpt_type_nm, "
                       lgStrSQL = lgStrSQL & "       func_cd, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0004", "''", "S") & ",func_cd) func_nm, "
                       lgStrSQL = lgStrSQL & "       role_cd, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ",role_cd) role_nm, "
                       lgStrSQL = lgStrSQL & "       entr_dt, "
                       lgStrSQL = lgStrSQL & "       resent_promote_dt "
                       lgStrSQL = lgStrSQL & "  FROM HAA010T "
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
                Call .CancelRestoreToolBar()             
                .DBAutoQueryOk        
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
