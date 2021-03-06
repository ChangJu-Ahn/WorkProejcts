<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
	Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
	Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
	
	Dim lgErrorStatus,	lgErrorPos,	lgOpModeCRUD 
    Dim lgKeyStream,	lgLngMaxRow
    Dim lgObjConn,		lgObjComm
    Dim strNativeErr

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim plant_nm,item_nm
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

 
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
      Case CStr(UID_M0006)
          Call SubBizBatch()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data

        Call SubCreateCommandObject(lgObjComm)
        Call SubBizBatchMulti(arrColVal)                            '☜: Run Batch
        Call SubCloseCommandObject(lgObjComm)


        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Call DisplayMsgBox("800407", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
           Exit For
        End If

    Next
    IF lgErrorStatus = "NO"	Then
    		Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
	END IF
End Sub


'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti(arrColVal)
    on error resume next
    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim strPlantCd,strItemCd
    Dim bErrorRaised

	strPlantCd	=	Trim(lgKeyStream(0))
	strItemCd	=	Trim(lgKeyStream(1))
    
    With lgObjComm
		.CommandTimeout = 0
        .CommandText = arrColVal(0)
        .CommandType = adCmdStoredProc
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd"     ,adVarChar,adParamInput,Len(Trim(strPlantCd)), strPlantCd)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@in_item_cd"     ,adVarChar,adParamInput,Len(Trim(strItemCd)), strItemCd)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adVarChar,adParamInput,Len(Trim(gUsrID)), gUsrID)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_nm"     ,adVarChar,adParamOutput,40)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@in_item_nm"     ,adVarChar,adParamOutput,40)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"   ,adVarChar ,adParamOutput,6)

        lgObjComm.Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value            
            if strMsg_Cd <> "" Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
			END IF
            Response.end
        ELSE
			plant_nm = lgObjComm.Parameters("@plant_nm").Value
			item_nm = lgObjComm.Parameters("@in_item_nm").Value
			
        end if
        
    Else    
        lgErrorStatus     = "YES"                                                         '☜: Set error status
         If lgObjComm.ActiveConnection.Errors.Count > 0 then
			strNativeErr = lgObjComm.ActiveConnection.Errors(0).NativeError
			bErrorRaised = True
		End If
		
		Select Case Trim(strNativeErr)
			Case "8115"																'%1!을(를) 데이터 형식 %2!(으)로 변환하는 중 산술 오버플로 오류가 발생했습니다.
				Call DisplayMsgBox("121515", vbInformation, "", "", I_MKSCRIPT)
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

	If Trim("<%=lgErrorStatus%>") = "NO" Then
		Parent.frm1.txtPlantNm.value = "<%=ConvSPChars(plant_nm)%>"
		parent.frm1.txtItemNm.value = "<%=ConvSPChars(item_nm)%>"
    End If

</Script>
