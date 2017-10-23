<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->


<%

    call LoadBasisGlobalInf()        
	Call loadInfTB19029B( "I", "*","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Dim interface_batch_id
    Dim change_code
    Dim strComRegNo
    

    interface_batch_id = ""
    change_code = ""
    strComRegNo = ""
    

    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             'Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
'             Call SubBizDelete()                          
    End Select
    
    Call SubCloseDB(lgObjConn)


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"					
'                   Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    if interface_batch_id = "" then
                       Call SubBizBatchNo(arrColval)                      
                       Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
                    else
                       Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
                    end if
                    
            Case "D"
'                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
    Next

End Sub  


'============================================================================================================
' Name : SubBizbatchNo
' Desc : Batch
'============================================================================================================
Sub SubBizBatchNo(arrColval)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


        Call SubCreateCommandObject(lgObjComm)
        Call SubBizBatchMultiNo(arrColval)                            '☜: Run Batch
        Call SubCloseCommandObject(lgObjComm)


End Sub


'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMultiNo(arrColval)
    on error resume next

    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim bErrorRaised
    Dim strWhereFlag
    Dim batchid

             
    With lgObjComm
		.CommandTimeout = 0
        .CommandText = "dbo.usp_make_interface_batch_id"
        .CommandType = adCmdStoredProc
        .Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
                
        Set batchid = lgObjComm.CreateParameter("@interface_batch_id",adNumeric,adParamOutput)
        batchid.Precision = 15
        batchid.NumericScale = 0
        
        lgObjComm.Parameters.Append  batchid
                                
        lgObjComm.Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD <> 1 then
            'strMsg_cd = lgObjComm.Parameters("@return_msg_cd").Value            
            'if strMsg_Cd <> "" Then
			'	Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
			'END IF
			interface_batch_id = lgObjComm.Parameters("@interface_batch_id").Value	
        ELSE
'			plant_nm = lgObjComm.Parameters("@plant_nm").Value            
        end if
        
    Else    
        lgErrorStatus     = "YES"                                                         '☜: Set error status
         If lgObjComm.ActiveConnection.Errors.Count > 0 then
			strNativeErr = lgObjComm.ActiveConnection.Errors(0).NativeError
			bErrorRaised = True
		End If
		
		Select Case Trim(strNativeErr)
			Case "8115"																'%1!을(를) 데이터 형식 %2!(으)로 변환하는 중 산술 오버플로 오류가 발생했습니다.
'				Call DisplayMsgBox("121515", vbInformation, "", "", I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
		End Select
		If bErrorRaised = False Then
	        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	    End if    
   End if
End Sub


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next
    Err.Clear                   
                        
                   
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear      
    
    if  change_code = "" then
        change_code = Trim(arrColVal(4))
    end if
    
    if strComRegNo = "" then
        strComRegNo = Trim(arrColVal(5))
    end if

    
    lgStrSQL = "UPDATE  XXSB_DTI_MAIN"
    lgStrSQL = lgStrSQL & " SET "     
    lgStrSQL = lgStrSQL & " INTERFACE_BATCH_ID = '" & interface_batch_id     & "' "    
    lgStrSQL = lgStrSQL & " WHERE   "
    lgStrSQL = lgStrSQL & " CONVERSATION_ID   =   " & FilterVar(arrColVal(2), "''", "S")     
    
    
    lgStrSQL = lgStrSQL & " UPDATE  XXSB_DTI_STATUS"
    lgStrSQL = lgStrSQL & " SET "     
    'lgStrSQL = lgStrSQL & " DTI_STATUS = " & FilterVar(arrColVal(4), "''", "S")  & ","   
    lgStrSQL = lgStrSQL & " LAST_UPDATE_DATE  = getdate(), "
    lgStrSQL = lgStrSQL & " SBDESCRIPTION =  " & FilterVar(arrColVal(3), "''", "S")  & ""       
    lgStrSQL = lgStrSQL & " WHERE   "
    lgStrSQL = lgStrSQL & " DIRECTION = '1' AND SUPBUY_TYPE ='AR' "        
    lgStrSQL = lgStrSQL & " and CONVERSATION_ID   =   " & FilterVar(arrColVal(2), "''", "S")     
    

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
		
	    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.frm1.hbatchid.value = "<%=interface_batch_id%>"
             Parent.frm1.hchangecode.value = "<%=change_code%>"
             Parent.frm1.hcomregno.value = "<%=strComRegNo%>"
             Parent.ExeNumOk2
          Else
             'Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
             Parent.ExeNumNot
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	