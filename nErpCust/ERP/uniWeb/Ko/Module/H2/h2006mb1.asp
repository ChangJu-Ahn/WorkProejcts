<%@ LANGUAGE=VBSCript%>
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

    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H","NOCOOKIE","MB")    
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    iKey1 = iKey1 & " AND a.emp_no LIKE " & FilterVar(lgKeyStream(1),"'%'", "S")
    iKey1 = iKey1 & " AND b.internal_cd LIKE  " & FilterVar(lgKeyStream(2) & "%", "''", "S") & ""
    iKey1 = iKey1 & " AND (b.retire_dt >" & FilterVar(lgKeyStream(0), "''", "S") & " or b.retire_dt is null)"
   

    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQLT)                                 '¡Ù : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("entr_dt"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn_nm")) ' FuncCodeName(1,"H0002",lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("warnt1_name"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("warnt1_start"),"")
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("warnt1_end"),"")

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
                       lgStrSQL =           "SELECT  b.name, b.entr_dt, a.emp_no, b.dept_cd, b.dept_nm, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", b.roll_pstn) roll_pstn_nm, "
                       lgStrSQL = lgStrSQL & "       a.warnt1_name, a.warnt1_start, a.warnt1_end "
                       lgStrSQL = lgStrSQL & "  FROM HAA040T a, HAA010T b "
                       lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no "
                       lgStrSQL = lgStrSQL & "   AND warnt1_end " & pComp & pCode
                       lgStrSQL = lgStrSQL & "  Union "
                       lgStrSQL = lgStrSQL & "SELECT  b.name, b.entr_dt, a.emp_no, b.dept_cd, b.dept_nm, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", b.roll_pstn) roll_pstn_nm, "
                       lgStrSQL = lgStrSQL & "       a.warnt2_name, a.warnt2_start, a.warnt2_end "
                       lgStrSQL = lgStrSQL & " From  HAA040T a, HAA010T b "
                       lgStrSQL = lgStrSQL & "Where a.emp_no = b.emp_no "
                       lgStrSQL = lgStrSQL & "  And warnt2_end " & pComp & pCode
                       lgStrSQL = lgStrSQL & "  Union "
                       lgStrSQL = lgStrSQL & "SELECT  b.name, b.entr_dt, a.emp_no, b.dept_cd, b.dept_nm, "
                       lgStrSQL = lgStrSQL & "       dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", b.roll_pstn) roll_pstn_nm, "
                       lgStrSQL = lgStrSQL & "       a.warnt_INSUR_NM, a.warnt_start, a.warnt_end "
                       lgStrSQL = lgStrSQL & " From  HAA040T a, HAA010T b "
                       lgStrSQL = lgStrSQL & "Where a.emp_no = b.emp_no "
                       lgStrSQL = lgStrSQL & "  And warnt_end " & pComp & pCode
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
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
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
