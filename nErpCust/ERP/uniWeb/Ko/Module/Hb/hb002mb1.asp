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
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

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
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_nm"))
         
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("entr_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("retire_dt"))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("year_area_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("year_area_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("native_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("native_nm"))

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_MONEY"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(ReadPROV_TYPE(lgObjRs("PROV_TYPE")))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tel_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hand_tel_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("email_addr"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("addr"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("zip_cd"))
            
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

    lgStrSQL = "INSERT INTO HAA011T( EMP_NO, EMP_NM, RES_NO, ENTR_DT, RETIRE_DT, DEPT_CD, YEAR_AREA_CD, NATIVE_CD, DAY_MONEY, PROV_TYPE, TEL_NO, HAND_TEL_NO, EMAIL_ADDR,ADDR, ZIP_CD, " 
    lgStrSQL = lgStrSQL & " Isrt_emp_no, Isrt_dt, Updt_emp_no, Updt_dt      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "null", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S") & ","
    
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(10), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & FilterVar(PutPROV_TYPE(arrColVal(11)), "''", "S") & ","
    
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(12), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(13), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(14), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(15), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(16), "''", "S") & ","

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
    lgStrSQL = "UPDATE  HAA011T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " EMP_NM = "		& FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & " RES_NO = "		& FilterVar(UCase(arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " ENTR_DT = "		& FilterVar(UCase(arrColVal(5)), "NULL", "S") & ","
    lgStrSQL = lgStrSQL & " RETIRE_DT = "	& FilterVar(arrColVal(6), "NULL", "S") & ","
    lgStrSQL = lgStrSQL & " DEPT_CD = "		& FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " YEAR_AREA_CD =" & FilterVar(arrColVal(8), "''", "S") & ","
    lgStrSQL = lgStrSQL & " NATIVE_CD = "	& FilterVar(UCase(arrColVal(9)), "''", "S") & ","

    lgStrSQL = lgStrSQL & " DAY_MONEY = "	& FilterVar(UNICdbl(arrColVal(10), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & " PROV_TYPE = "	& FilterVar(PutPROV_TYPE(arrColVal(11)), "''", "S") & ","

    lgStrSQL = lgStrSQL & " TEL_NO = "		& FilterVar(UCase(arrColVal(12)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " HAND_TEL_NO = " & FilterVar(UCase(arrColVal(13)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " EMAIL_ADDR = "	& FilterVar(UCase(arrColVal(14)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " ADDR = "		& FilterVar(UCase(arrColVal(15)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " ZIP_CD = "		& FilterVar(UCase(arrColVal(16)), "''", "S") & ","

    lgStrSQL = lgStrSQL & " Updt_emp_no = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " Updt_dt = "		& FilterVar(lgSvrDateTime, "''", "S")  

    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")

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

    lgStrSQL = "DELETE  HAA011T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "         emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

Function ReadPROV_TYPE(Byval pData)
	If pData = "Y" Then
		ReadPROV_TYPE = "1"
	Else
		ReadPROV_TYPE = "0"
	End If
End Function

Function PutPROV_TYPE(Byval pData)
	If pData = "1" Then
		PutPROV_TYPE = "Y"
	Else
		PutPROV_TYPE = "N"
	End If
End Function

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
                       lgStrSQL = lgStrSQL & " a.EMP_NO, a.EMP_NM, a.RES_NO, a.ENTR_DT, a.RETIRE_DT, a.DEPT_CD, a.YEAR_AREA_CD, a.NATIVE_CD, A.DAY_MONEY, A.PROV_TYPE"
                       lgStrSQL = lgStrSQL & " , a.TEL_NO, a.HAND_TEL_NO,EMAIL_ADDR, a.ADDR, ZIP_CD "
                       lgStrSQL = lgStrSQL & " , b.dept_nm, c.year_area_nm, case a.NATIVE_CD when '1' then '내국인' when '2' then '외국인' end native_nm "
                       lgStrSQL = lgStrSQL & " FROM  HAA011T a"
                       lgStrSQL = lgStrSQL & "	left outer join  B_ACCT_DEPT b on a.dept_cd = b.dept_cd and b.org_change_dt = ( select max(org_change_dt) from B_ACCT_DEPT where org_change_dt <= case when a.RETIRE_DT is not null then a.RETIRE_DT else a.ENTR_DT end)"
                       lgStrSQL = lgStrSQL & "	left outer join  HFA100T c on a.YEAR_AREA_CD = c.YEAR_AREA_CD "

                       If Trim(pCode) <> "''" Then
							lgStrSQL = lgStrSQL & " WHERE A.emp_no = " & pCode 
					   End If
                       lgStrSQL = lgStrSQL & " ORDER BY A.emp_no, A.ENTR_DT ASC"
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
