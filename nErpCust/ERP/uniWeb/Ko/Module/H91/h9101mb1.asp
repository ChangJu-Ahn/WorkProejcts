<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
                                                                          '¢Ð: Clear Error status
	dim lgGetSvrDateTime
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
   
    lgGetSvrDateTime = GetSvrDateTime
    
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
    dim strWhere
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TAX_STD_LOW"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TAX_STD_HIGH"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(CDbl(lgObjRs("TAX_RATE"))*100,0,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ACCUM_SUB_AMT"), ggAmtOfMoney.DecPoint,0)
            
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
	dim strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	strWhere = " not (TAX_STD_LOW >= " & UNIConvNum(arrColVal(3),0) & " OR TAX_STD_HIGH <= " & UNIConvNum(arrColVal(2),0) & " ) "
	
	IntRetCD = CommonQueryRs("Count(*) ", " HFA010T ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    
    if intRetCD = true then
		if Cint(replace(lgF0,chr(11),"")) > 0 then
			Call DisplayMsgBox("800497", vbInformation, "", "", I_MKSCRIPT)
			
%>
<script language=vbscript >		
			parent.frm1.vspdData.Row = <%=arrColVal(6)%>
			parent.frm1.vspdData.Col = 1  ' C_TAX_STD_LOW_AMT
			parent.frm1.vspdData.Action = 0 ' go to 
</script>
<%
			 ObjectContext.SetAbort
             Call SetErrorStatus
			 Call SubCloseDB(lgObjConn)
			response.end
		end if
    end if
    
    lgStrSQL = "INSERT INTO HFA010T("
    lgStrSQL = lgStrSQL & " TAX_STD_LOW,"
    lgStrSQL = lgStrSQL & " TAX_STD_HIGH,"
    lgStrSQL = lgStrSQL & " TAX_RATE," 
    lgStrSQL = lgStrSQL & " ACCUM_SUB_AMT,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(2),0)      & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(3),0)      & ","
    lgStrSQL = lgStrSQL & Cdbl(UNIConvNum(arrColVal(4),0))/100     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)      & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
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
	dim strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    strWhere = " not (TAX_STD_LOW >= " & UNIConvNum(arrColVal(3),0) & " OR TAX_STD_HIGH <= " & UNIConvNum(arrColVal(2),0) & " ) "
	strWhere = strWhere & " AND TAX_STD_LOW <> " & UNIConvNum(arrColVal(2),0) & " and TAX_STD_HIGH <> " & UNIConvNum(arrColVal(3),0)
	IntRetCD = CommonQueryRs("Count(*) ", " HFA010T ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    
    if intRetCD = true then
		if Cint(replace(lgF0,chr(11),"")) > 0 then
			Call DisplayMsgBox("800497", vbInformation, "", "", I_MKSCRIPT)
			
%>
<script language=vbscript >		
			parent.frm1.vspdData.Row = <%=arrColVal(6)%>
			parent.frm1.vspdData.Col = 1  ' C_TAX_STD_HIGH_AMT
			parent.frm1.vspdData.Action = 0 ' go to 
</script>
<%
			 ObjectContext.SetAbort
             Call SetErrorStatus
			 Call SubCloseDB(lgObjConn)
			response.end
		end if
    end if
    
    lgStrSQL = "UPDATE  HFA010T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " TAX_STD_HIGH   = " & UNIConvNum(arrColVal(3),0)  & "," 
    lgStrSQL = lgStrSQL & " TAX_RATE       = " & Cdbl(UNIconvnum(arrColVal(4),0))/100 & " ," 
    lgStrSQL = lgStrSQL & " ACCUM_SUB_AMT  = " & UNIConvNum(arrColVal(5),0)  & "," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO    = " & FilterVar(gUsrId, "''", "S")  & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT         = " & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " TAX_STD_LOW   = "    & UNIConvNum(arrColVal(2),0)

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

    lgStrSQL = "DELETE  HFA010T "
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " TAX_STD_LOW  = "        & UNIConvNum(arrColVal(2),0)

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
                       lgStrSQL = "Select TOP " & iSelCount  & "  TAX_STD_LOW , TAX_STD_HIGH ,TAX_RATE , ACCUM_SUB_AMT ,ISRT_EMP_NO , ISRT_DT , UPDT_EMP_NO , UPDT_DT  "
                       lgStrSQL = lgStrSQL & " FROM HFA010T "
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
