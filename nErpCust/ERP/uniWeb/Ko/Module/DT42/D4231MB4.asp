<%@ LANGUAGE=VBSCript %>
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
                    Call SubBizBatch(arrColval)
            Case "U"
                    'Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
                    Call SubBizBatch2(arrColval)
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
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch(arrColval)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


        Call SubCreateCommandObject(lgObjComm)
        Call SubBizBatchMulti(arrColval)                            '☜: Run Batch
        Call SubCloseCommandObject(lgObjComm)


End Sub

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch2(arrColval)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


        Call SubCreateCommandObject(lgObjComm)
        Call SubBizBatchMulti2(arrColval)                            '☜: Run Batch
        Call SubCloseCommandObject(lgObjComm)


End Sub


'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti(arrColval)
    on error resume next

    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim bErrorRaised
    Dim strWhereFlag    
    Dim strTaxBillNo
    Dim strByrEmpName,  strByrDeptName,   strByrTelNum
    Dim strByrEmail,  strAmend, strRemark, strTaxType
    Dim strConverid


    strWhereFlag     =  Trim(arrColVal(2))
    strTaxBillNo     =  Trim(arrColVal(3))
    strByrEmpName    =  Trim(arrColVal(4))    
    strByrDeptName   =  Trim(arrColVal(5))    
    strByrTelNum     =  Trim(arrColVal(6))    
    strByrEmail      =  Trim(arrColVal(7))    
    strAmend         =  Trim(arrColVal(8))
    strRemark        =  Trim(arrColVal(9))
    strTaxType       =  Trim(arrColVal(10))
   
   if  strTaxType = "FI" then
      if strWhereFlag = "PMS" then
          strTaxType = "PMS"
      end if
   end if   

          
    With lgObjComm
		.CommandTimeout = 0
        .CommandText = "dbo.usp_dt_send_tax_sm"
        .CommandType = adCmdStoredProc
        .Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
                        
        .Parameters.Append lgObjComm.CreateParameter("@tax_type"      ,adVarChar, adParamInput,3, strTaxType)        
        .Parameters.Append lgObjComm.CreateParameter("@tax_no"        ,adVarChar, adParamInput,16, strTaxBillNo)
        .Parameters.Append lgObjComm.CreateParameter("@user_id"       ,adVarChar, adParamInput,13, gUsrID)                        
        .Parameters.Append lgObjComm.CreateParameter("@change_remark" ,adVarChar, adParamInput,150, strRemark)                 
        .Parameters.Append lgObjComm.CreateParameter("@byr_emp_name"  ,adVarChar, adParamInput,30, strByrEmpName) 
        .Parameters.Append lgObjComm.CreateParameter("@amend_code"    ,adVarChar, adParamInput,2, strAmend)         
        .Parameters.Append lgObjComm.CreateParameter("@byr_dept_name" ,adVarChar, adParamInput,40, strByrDeptName) 
        .Parameters.Append lgObjComm.CreateParameter("@byr_tel_num"   ,adVarChar, adParamInput,20, strByrTelNum)         
        .Parameters.Append lgObjComm.CreateParameter("@byr_email"     ,adVarChar, adParamInput,40, strByrEmail) 
        
        .Parameters.Append lgObjComm.CreateParameter("@make_conversation_id"   ,adVarChar ,adParamOutput, 35)
        .Parameters.Append lgObjComm.CreateParameter("@return_msg_cd"   ,adVarChar ,adParamOutput, 10)
                	        
'        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@in_item_cd"     ,adVarChar,adParamInput,Len(Trim(strItemCd)), strItemCd)

'        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_nm"     ,adVarChar,adParamOutput,40)
'        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@in_item_nm"     ,adVarChar,adParamOutput,40)

        lgObjComm.Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@return_msg_cd").Value            
            if strMsg_Cd <> "" Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
			END IF
			strConverid = lgObjComm.Parameters("@make_conversation_id").Value			
        ELSE
'			plant_nm = lgObjComm.Parameters("@plant_nm").Value
            'strConverid = lgObjComm.Parameters("@make_conversation_id").Value			
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
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti2(arrColval)
    on error resume next

    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim bErrorRaised
    Dim strWhereFlag
    Dim strTaxBillNo   
    Dim strConverid
    Dim strTaxType


    strWhereFlag   =  Trim(arrColVal(2))
    strConverid    =  Trim(arrColVal(3))      
    strTaxType     =  Trim(arrColVal(4))    

   
   if  strTaxType = "SD" then
      if strWhereFlag = "PMS" then
          strTaxType = "PMS"
      end if
   end if   

          
    With lgObjComm
		.CommandTimeout = 0
        .CommandText = "dbo.usp_save_cancel_conversation_id"
        .CommandType = adCmdStoredProc
        .Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
                        
        .Parameters.Append lgObjComm.CreateParameter("@@conversationid" ,adVarChar, adParamInput,35, strConverid)        
        
        'lgObjComm.Parameters.Append lgObjComm.CreateParameter("@make_conversation_id"   ,adVarChar ,adParamOutput, 35)
        'lgObjComm.Parameters.Append lgObjComm.CreateParameter("@return_msg_cd"   ,adVarChar ,adParamOutput, 10)
                	        
'        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@in_item_cd"     ,adVarChar,adParamInput,Len(Trim(strItemCd)), strItemCd)

'        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_nm"     ,adVarChar,adParamOutput,40)

        lgObjComm.Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD <> 1 then
            'strMsg_cd = lgObjComm.Parameters("@return_msg_cd").Value            
            'if strMsg_Cd <> "" Then
			'	Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
			'END IF
        ELSE
'			plant_nm = lgObjComm.Parameters("@plant_nm").Value
            'strConverid = lgObjComm.Parameters("@make_conversation_id").Value			
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
             Parent.DBSaveOk
          Else
             'Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	