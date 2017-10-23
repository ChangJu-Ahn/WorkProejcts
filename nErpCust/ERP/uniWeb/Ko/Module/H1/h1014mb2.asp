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

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

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
    
    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQ)                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("strt_yy"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("strt_yy_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Mid(lgObjRs("strt_mmdd"),1,2))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Mid(lgObjRs("strt_mmdd"),3,2))
            lgstrData = lgstrData & Chr(11) & "~"
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("end_yy"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("end_yy_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Mid(lgObjRs("end_mmdd"),1,2))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Mid(lgObjRs("end_mmdd"),3,2))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("year_cnt10"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("year_cnt8"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("year_retr"))
            
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
	Dim strStrt, strEnd
	Dim IntRetCD
	Dim strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	strStrt = FilterVar((2+cint(arrColVal(2))) & arrColVal(3) & arrColVal(4), "''", "S")
	strEnd  = FilterVar((2+cint(arrColVal(5))) & arrColVal(6) & arrColVal(7), "''", "S")

	strWhere = "  not ( convert(varchar(1), 2 + strt_yy) + strt_mmdd >= " & strEnd & " or " 
	strWhere = strWhere & " convert(varchar(1), 2 + end_yy)  + end_mmdd  <= " & strStrt & " ) "
	
	IntRetCD = CommonQueryRs("Count(*) ", " hfb010t ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    if intRetCD = true then
		if Cint(replace(lgF0,chr(11),"")) > 0 then
			Call DisplayMsgBox("800496", vbInformation, "", "", I_MKSCRIPT)
			
%>
<script language=vbscript >		
			parent.frm1.vspdData.Row = <%=arrColVal(11)%>
			parent.frm1.vspdData.Col = 1  ' C_STRT_YY
			parent.frm1.vspdData.Action = 0 ' go to 
</script>
<%
			 ObjectContext.SetAbort
             Call SetErrorStatus
			 Call SubCloseDB(lgObjConn)
			response.end
		end if
    end if
    
    lgStrSQL = "INSERT INTO HFB010T( strt_yy,strt_mmdd, end_yy, end_mmdd,"
    lgStrSQL = lgStrSQL & " year_cnt10, year_cnt8, year_retr,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO, ISRT_DT , UPDT_EMP_NO , UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(2),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3) & arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6) & arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0) & ","
    
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
	
	Dim strStrt, strEnd
	Dim IntRetCD
	Dim strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strStrt = FilterVar((2+ cint(arrColVal(2))) & arrColVal(3) & arrColVal(4), "''", "S")
	strEnd  = FilterVar((2 + cint(arrColVal(5))) & arrColVal(6) & arrColVal(7), "''", "S")

	strWhere = "  not ( convert(varchar(1), 2 + strt_yy) + strt_mmdd >= " & strEnd & " or " 
	strWhere = strWhere & " convert(varchar(1), 2 + end_yy)  + end_mmdd  <= " & strStrt & " ) "
	strWhere = strWhere & " AND  NOT (strt_yy = " & arrColVal(2) & " AND strt_mmdd =  " & FilterVar(arrColVal(3) & arrColVal(4), "''", "S") & ")"

	IntRetCD = CommonQueryRs("Count(*) ", " hfb010t ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    if intRetCD = true then
		if Cint(replace(lgF0,chr(11),"")) > 0 then
			Call DisplayMsgBox("800496", vbInformation, "", "", I_MKSCRIPT)
			
%>
<script language=vbscript >		
			parent.frm1.vspdData.Row = <%=arrColVal(11)%>
			parent.frm1.vspdData.Col = 6  ' C_END_YY
			parent.frm1.vspdData.Action = 0 ' go to 
</script>
<%
			 ObjectContext.SetAbort
             Call SetErrorStatus
			 Call SubCloseDB(lgObjConn)
			response.end
		end if
    end if
    
    lgStrSQL = "UPDATE  HFB010T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & "       end_yy = " & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & "       end_mmdd = " & FilterVar(arrColVal(6) & arrColVal(7), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       year_cnt10 = " & UNIConvNum(arrColVal(8),0) & ","
    lgStrSQL = lgStrSQL & "       year_cnt8 = " & UNIConvNum(arrColVal(9),0) & ","
    lgStrSQL = lgStrSQL & "       year_retr = " & UNIConvNum(arrColVal(10),0) & ","
    lgStrSQL = lgStrSQL & "       UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & "       UPDT_DT = " & FilterVar(GetSvrDateTime,"NULL","S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       STRT_YY = " & UNIConvNum(arrColVal(2),0)
    lgStrSQL = lgStrSQL & "   AND STRT_MMDD = " & FilterVar(arrColVal(3) & arrColVal(4), "''", "S")

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

    lgStrSQL = "DELETE  HFB010T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       STRT_YY = " & UNIConvNum(arrColVal(2),0)
    lgStrSQL = lgStrSQL & "   AND STRT_MMDD = " & FilterVar(arrColVal(3) & arrColVal(4), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select 	strt_yy,  dbo.ufn_GetCodeName(" & FilterVar("H0098", "''", "S") & ",strt_yy) strt_yy_nm, "
                       lgStrSQL = lgStrSQL & " strt_mmdd,  end_yy, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0098", "''", "S") & ",end_yy) end_yy_nm, "
                       lgStrSQL = lgStrSQL & " end_mmdd, year_cnt10, year_cnt8, year_retr "                       
                       lgStrSQL = lgStrSQL & " From  HFB010T "
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

