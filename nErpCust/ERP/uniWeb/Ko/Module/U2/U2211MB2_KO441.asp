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

	Dim sBpCd, sDlvyNo
	
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")

    lgErrorStatus    = "NO"
    lgErrorPos        = ""

    lgLngMaxRow			= Request("txtMaxRows")
    lgStrPrevKey			= UNICInt(Trim(Request("lgStrPrevKey")),0)

    lgOpModeCRUD		= Request("txtMode")
    sBpCd						= Request("sBpCd")
    sDlvyNo					= Request("sDlvyNo")

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

     'strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
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

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))	
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DLVY_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOCUMENT_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TITLE"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INS_USER"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOCUMENT_ABBR"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("INS_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOCUMENT_TEXT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INSRT_USER_ID"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("INSRT_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UPDT_USER_ID"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("UPDT_DT"))

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
					
					lgStrSQL	= " SELECT	TOP " & iSelCount & " 	BP_CD, DLVY_NO, 	DOCUMENT_NO, TITLE, INS_USER, DOCUMENT_ABBR, INS_DT, " _
								    & "				DOCUMENT_TEXT, INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT " _
									& "    FROM	M_SCM_DOCUMENT_HDR_KO441 " _
									& " WHERE	BP_CD = " & FilterVar(sBpCd,"''", "S") _
									& "       AND	DLVY_NO = " & FilterVar(sDlvyNo,"''", "S") _
									& "  ORDER BY BP_CD, DLVY_NO, DOCUMENT_NO "

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
      
</Script>	