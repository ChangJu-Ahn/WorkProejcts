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

    Call HideStatusWnd
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
                                                          
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

    lgStrSQL = "DELETE  HDA080T"
    lgStrSQL = lgStrSQL & " WHERE ALLOW_CD = " & FilterVar(Request("txtAllow_cd"), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

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

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    iKey1 = iKey1 & " ORDER BY dept_CD, duty_strt "

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
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))          
            lgstrData = lgstrData & Chr(11) & UNINUMClientFormat(lgObjRs("duty_strt"), 0, 0)
            lgstrData = lgstrData & Chr(11) & UNINUMClientFormat(lgObjRs("duty_end"), 0, 0)
            lgstrData = lgstrData & Chr(11) & UNINUMClientFormat(lgObjRs("allow"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("apply_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("apply_type_nm"))
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
	Dim Dept_cd
	Dim strWhere
	Dim IntRetCD
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	 '  중복기간 체크 
	Dept_CD = FilterVar(UCase(arrColVal(2)), "''", "S") 
	strWhere = " not (Duty_Strt >= " & UNIConvNum(arrColVal(5),0) & " OR Duty_End <= " & UNIConvNum(arrColVal(4),0) 
	strWhere = strWhere & " ) AND Dept_cd = " & Dept_cd & " AND ALLOW_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	IntRetCD = CommonQueryRs("Count(*) ", " HDA080T ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    if intRetCD = true then
		if Cint(replace(lgF0,chr(11),"")) > 0 then
		
			Call DisplayMsgBox("800496", vbInformation, "", "", I_MKSCRIPT)
			
%>
<script language=vbscript >		
			parent.frm1.vspdData.Row = <%=arrColVal(8)%>
			parent.frm1.vspdData.Col = 4  ' C_DUTY_Strt
			parent.frm1.vspdData.Action = 0 ' go to 
</script>
<%
			 ObjectContext.SetAbort
             Call SetErrorStatus
			 Call SubCloseDB(lgObjConn)
			response.end
		end if
    end if

    lgStrSQL = "INSERT INTO HDA080T( DEPT_CD  , ALLOW_CD  , DUTY_STRT ,DUTY_END , ALLOW  ,"
    lgStrSQL = lgStrSQL & " APPLY_TYPE , ISRT_EMP_NO, ISRT_DT  , UPDT_EMP_NO , UPDT_DT  )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(4),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    
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
	Dim Dept_cd
	Dim strWhere
	Dim IntRetCD
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    '  중복기간 체크 
    Dept_CD = FilterVar(UCase(arrColVal(2)), "''", "S") 
	strWhere = " not (Duty_Strt >= " & UNIConvNum(arrColVal(5),0) & " OR Duty_End <= " & UNIConvNum(arrColVal(4),0) 
	strWhere = strWhere & " ) AND Dept_cd = " & Dept_cd & " AND ALLOW_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	strWhere = strWhere & " AND Duty_Strt  <> " & UNIConvNum(arrColVal(4),0) 
	IntRetCD = CommonQueryRs("Count(*) ", " HDA080T ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    
    if intRetCD = true then
		if Cint(replace(lgF0,chr(11),"")) > 0 then
			Call DisplayMsgBox("800496", vbInformation, "", "", I_MKSCRIPT)			
%>
<script language=vbscript >		
			parent.frm1.vspdData.Row = <%=arrColVal(8)%>
			parent.frm1.vspdData.Col = 5  ' C_DUTY_END
			parent.frm1.vspdData.Action = 0 ' go to 
</script>
<%
			 ObjectContext.SetAbort
             Call SetErrorStatus
			 Call SubCloseDB(lgObjConn)
			response.end
		end if
    end if
    lgStrSQL = "UPDATE  HDA080T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & "       DUTY_END = " & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & "       ALLOW = " & UNIConvNum(arrColVal(6), 0) & ","
    lgStrSQL = lgStrSQL & "       APPLY_TYPE = " & FilterVar(arrColVal(7), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & "       UPDT_DT = " & FilterVar(GetSvrDateTime,"NULL","S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       DEPT_CD = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND ALLOW_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND DUTY_STRT = " & UNIConvNum(arrColVal(4), 0) 
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

    lgStrSQL = "DELETE  HDA080T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       DEPT_CD = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND ALLOW_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND DUTY_STRT = " & UNIConvNum(arrColVal(4), 0) 

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
                       lgStrSQL = "SELECT TOP " & iSelCount  & " dept_cd, dbo.ufn_H_GetCodeName(" & FilterVar("H_CURRENT_DEPT", "''", "S") & " ,DEPT_CD  ,'' ) as DEPT_nm, "
                       lgStrSQL = lgStrSQL & " duty_strt, duty_end, allow, apply_type, dbo.ufn_GetCodeName(" & FilterVar("H0091", "''", "S") & ",apply_type) apply_type_nm "  
                       lgStrSQL = lgStrSQL & " FROM  HDA080T "
                       lgStrSQL = lgStrSQL & " WHERE ALLOW_CD " & pComp & pCode
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
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	

