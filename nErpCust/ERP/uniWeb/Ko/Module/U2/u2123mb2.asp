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
	Dim lgStrColorFlag
	
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""
    lgOpModeCRUD      = Request("txtMode")
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)

    Call SubOpenDB(lgObjConn)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    'Call SubBizQueryCond()
    'If lgErrorStatus <> "YES" Then
		Call SubBizQueryMulti()
	'End If
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
' Name : SubBizQueryCond
' Desc : Verify Condition Value from Db
'============================================================================================================
Sub SubBizQueryCond()

    On Error Resume Next
    Err.Clear
	
	If lgKeyStream(0) <> "" then 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CI")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(lgObjRs("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(1) <> "" then 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CB")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("SCM003", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """ & ConvSPChars(lgObjRs("BP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next
    Err.Clear

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
			
			If lgObjRs("GRP_FLG") = 1 Then
				lgstrData = lgstrData & Chr(11) & "[ÃÑ°è]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "1" & gRowSep
			Else
				lgstrData = lgstrData & Chr(11) & ""
			End If
			
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
			
			If ConvSPChars(lgObjRs("CONFIRM_YN")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONFIRM_YN"))
            Else
				lgstrData = lgstrData & Chr(11) & ""
            End If
            
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("CONFIRM_QTY"),ggQty.DecPoint,0)
            
            If ConvSPChars(lgObjRs("M_TYPE")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("M_TYPE"))
            Else
				lgstrData = lgstrData & Chr(11) & ""
            End If
            
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("INSPECT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPLIT_SEQ_NO"))
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
    Call SubCloseRs(lgObjRs)

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)

        Select Case arrColVal(0)
            Case "C"                            '¢Ð: Create
                    Call SubBizSaveMultiCreate(arrColVal)
            Case "U"                            '¢Ð: Update
					If Trim(lgErrorStatus) = "NO" Then
	                    Call SubBizSaveMultiUpdate(arrColVal)
					End If
            Case "D"							'¢Ð: Delete
					Call SubBizDelCheck(arrColVal)
					If Trim(lgErrorStatus) = "NO" Then
						Call SubBizSaveMultiDelete(arrColVal)
					End If
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    


'============================================================================================================
' Name : SubBizDelCheck
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDelCheck(arrColVal)
    On Error Resume Next
    Err.Clear

	lgStrSQL =            " SELECT isNULL(RCPT_QTY,0) RCPT_QTY "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO		= " & FilterVar(Trim(UCase(arrColVal(2))),"","S") 
	lgStrSQL = lgStrSQL & "    AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")
    lgStrSQL = lgStrSQL & "    AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgOrgQty = lgObjRs("RCPT_QTY")
    End If
    
	If Cdbl(lgOrgQty) > 0 Then
	    Call DisplayMsgBox("SCM005", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus
	End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
    
    lgStrSQL =            " UPDATE  M_SCM_FIRM_PUR_RCPT "
    lgStrSQL = lgStrSQL & "    SET " 
    lgStrSQL = lgStrSQL & " CONFIRM_YN			= " &  FilterVar(Trim(UCase(arrColVal(5))),"","S") & "," 
    lgStrSQL = lgStrSQL & " CONFIRM_QTY			= " &  FilterVar(UNIConvNum(arrColVal(6), 0),"","D") & ","
    lgStrSQL = lgStrSQL & " M_TYPE				= " &  FilterVar(Trim(UCase(arrColVal(7))),"","S") & ","  
    lgStrSQL = lgStrSQL & " UPDT_USER_ID		= " &  FilterVar(gUsrId,"","S")                    & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT				= " &  FilterVar(GetSvrDateTime,"''","S")
    lgStrSQL = lgStrSQL & " WHERE PO_NO			= " &  FilterVar(Trim(UCase(arrColVal(2))),"","S")
    lgStrSQL = lgStrSQL & "	AND PO_SEQ_NO		= " &  FilterVar(Trim(UCase(arrColVal(3))),"","D")	
    lgStrSQL = lgStrSQL & " AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")
        
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear

    lgStrSQL = "DELETE  M_SCM_FIRM_PUR_RCPT "
    lgStrSQL = lgStrSQL & " WHERE PO_NO			= " &  FilterVar(Trim(UCase(arrColVal(2))),"","S")
    lgStrSQL = lgStrSQL & "	AND PO_SEQ_NO		= " &  FilterVar(Trim(UCase(arrColVal(3))),"","D")	
    lgStrSQL = lgStrSQL & " AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
			iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
			Select Case Mid(pDataType,2,1)
				Case "R"
					
					lgStrSQL	= " SELECT	TOP " & iSelCount & " A.GRP_FLG, B.PO_NO, B.PO_SEQ_NO, B.SPLIT_SEQ_NO, B.PLAN_DVRY_DT, " _
								& "			A.PLAN_DVRY_QTY, B.CONFIRM_YN, B.M_TYPE, A.CONFIRM_QTY, A.INSPECT_QTY, B.RCPT_QTY " _
								& "   FROM	( " _
								& "          SELECT GROUPING(PO_NO) + GROUPING(PO_SEQ_NO) + GROUPING(SPLIT_SEQ_NO) GRP_FLG, " _
								& "					PO_NO, PO_SEQ_NO, SPLIT_SEQ_NO, " _
								& "					SUM(ISNULL(PLAN_DVRY_QTY,0)) PLAN_DVRY_QTY, " _
								& "					SUM(ISNULL(CONFIRM_QTY,0)) CONFIRM_QTY, " _
								& "					SUM(ISNULL(INSPECT_QTY,0)) INSPECT_QTY " _
								& "			   FROM	M_SCM_FIRM_PUR_RCPT (NOLOCK) " _
								& "			  WHERE PO_NO = " & FilterVar(lgKeyStream(0),"''", "S") _
								& "				AND PO_SEQ_NO = " & FilterVar(lgKeyStream(1),"''", "S") _
								& "				AND CONFIRM_QTY - RCPT_QTY > 0 " _
								& "				AND DLVY_NO IS NOT NULL " _
								& "			  GROUP BY PO_NO,PO_SEQ_NO,SPLIT_SEQ_NO WITH ROLLUP " _
								& "			 HAVING	SUM(ISNULL(PLAN_DVRY_QTY,0)) <> 0 OR SUM(ISNULL(CONFIRM_QTY,0)) <> 0) A " _
								& "   LEFT OUTER JOIN M_SCM_FIRM_PUR_RCPT B(NOLOCK) " _
								& " 	ON A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO AND A.SPLIT_SEQ_NO = B.SPLIT_SEQ_NO " _
								& "  WHERE A.GRP_FLG IN (0,1) "
					
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
	Dim lsMsg
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
                .ggoSpread.Source     = .frm1.vspdData2
                .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
                .lgStrColorFlag = "<%=lgStrColorFlag%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBDtlQueryOk()
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         