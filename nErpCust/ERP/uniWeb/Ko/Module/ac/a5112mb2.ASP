<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<%

    On Error Resume Next
    Err.Clear

    Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("Q","A", "COOKIE", "QB")

    Dim lgFiscStart
	Dim lgStartDt
	Dim lgEndDt
	Dim txtClassType
	Dim txtBizArea
	Dim txtPrintOpt
	Dim strZeroFg
	Dim strHqBrchFg
	Dim strUserId

	Dim lgBalLamt
	Dim lgTotLamt
	Dim lgThisLamt
	Dim lgThisRamt
	Dim lgTotRamt
	Dim lgBalRamt

	Dim yyyy,mm,dd

	Dim lgSp_Id

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    lgFiscStart		= Trim(Request("lgFiscStart"))
	lgStartDt		= Trim(Request("lgStartDt"))
	lgEndDt			= Trim(Request("lgEndDt"))
	txtClassType	= Trim(Request("txtClassType"))
	txtBizArea		= Trim(Request("txtBizArea"))
	txtPrintOpt		= Trim(Request("txtPrintOpt"))
	strZeroFg		= Trim(Request("strZeroFg"))
	strHqBrchFg		= Trim(Request("strHqBrchFg"))
	strUserId		= Trim(Request("strUserId"))
	
	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
	lgInternalCd		= Trim(Request("lgInternalCd"))	
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    Call SubBizBatch()
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch()
    On Error Resume Next
    Err.Clear

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatchMulti()                            '☜: Run Batch
    Call SubCloseCommandObject(lgObjComm)


    If lgErrorStatus    = "YES" Then
  '     lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
    End If
    
    IF lgErrorStatus = "NO"	Then
    		'Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
	END IF
End Sub


'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti()
	 On Error Resume NEXT
	 Err.Clear

    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim strSp

    Dim  strNativeErr

    strSp = "usp_a_tb"
    lgstrData = ""
    strNativeErr = ""

	'권한 관리 추가 
	Dim BizAreaCd
	
	BizAreaCd = txtBizArea
	
	If BizAreaCd = "" Then	 	 	
		If lgAuthBizAreaCd <> "" Then			
			BizAreaCd = lgAuthBizAreaCd
		End If
	Else
		If lgAuthBizAreaCd <> "" Then
			If UCASE(lgAuthBizAreaCd) <> UCASE(BizAreaCd) Then
		        Call DisplayMsgBox("124200", vbInformation, "", "", I_MKSCRIPT)
				Response.end
			End If
		End If
	End If

'Response.Write BizAreaCd
'Response.end
	With lgObjComm
	   .CommandText = strSp
	   .CommandType = adCmdStoredProc

	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	 adInteger,	adParamReturnValue)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fisc_dt",		 adVarWChar,	adParamInput,		10, lgFiscStart)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@start_dt",		 adVarWChar,	adParamInput,		10, lgStartDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@end_dt",			 adVarWChar,	adParamInput,		10, lgEndDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@class_type",		 adVarWChar,	adParamInput,		20, txtClassType)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@biz_area_cd",	 adVarWChar,	adParamInput,		10, BizAreaCd)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@hq_brch_fg",		 adVarWChar,	adParamInput,		1, strHqBrchFg)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@zero_fg",		 adWChar,	adParamInput,		1, strZeroFg)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@print_opt",		 adWChar,	adParamInput,		1, txtPrintOpt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",		 adVarWChar,	adParamInput,		13, strUserId)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",		     adVarWChar,	adParamOutput,		6)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_id",		     adVarWChar,	adParamOutput,		13)	   		  

	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@BalLamt",		 adVarWChar,	adParamOutput,		20)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TotLamt",		 adVarWChar,	adParamOutput,		20)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ThisLamt",		 adVarWChar,	adParamOutput,		20)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ThisRamt",		 adVarWChar,	adParamOutput,		20)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TotRamt",		 adVarWChar,	adParamOutput,		20)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@BalRamt",		 adVarWChar,	adParamOutput,		20)

	   lgObjComm.Execute ,, adExecuteNoRecords
	End With

    If Err.number = 0 Then
       IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
'		Response.Write "IntRetCD=" & intRetCd
       If IntRetCD <> 1 then
          strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
          if strMsg_Cd <> "" Then
		       Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
		       Call SubCloseCommandObject(lgObjComm)
			 END IF
          Response.end

       Else

		lgSp_Id = lgObjComm.Parameters("@sp_id").Value

		lgBalLamt		= lgObjComm.Parameters("@BalLamt").Value
		lgTotLamt		= lgObjComm.Parameters("@TotLamt").Value
		lgThisLamt		= lgObjComm.Parameters("@ThisLamt").Value
		lgThisRamt		= lgObjComm.Parameters("@ThisRamt").Value
		lgTotRamt		= lgObjComm.Parameters("@TotRamt").Value
		lgBalRamt		= lgObjComm.Parameters("@BalRamt").Value

		lgstrData = Chr(11) & "" 
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgBalLamt,  ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgTotLamt,	 ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgThisLamt, ggAmtOfMoney.DecPoint, 0)
		If CDbl(lgTotLamt) <> CDbl(lgTotRamt) Then
			lgstrData = lgstrData & Chr(11) & ConvSPChars("대차착오")
		Else
			lgstrData = lgstrData & Chr(11) & ConvSPChars("합계")
		End If
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgThisRamt, ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgTotRamt,  ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgBalRamt,  ggAmtOfMoney.DecPoint, 0)
		lgstrData = lgstrData & Chr(11) & "" 
		lgstrData = lgstrData & Chr(11) & Chr(12)
       End If

   Else
      lgErrorStatus     = "YES"
      If lgObjComm.ActiveConnection.Errors.Count > 0 then
			strNativeErr = lgObjComm.ActiveConnection.Errors(0).NativeError
		End If

		Select Case Trim(strNativeErr)
			Case "8115"																'%1!을(를) 데이터 형식 %2!(으)로 변환하는 중 산술 오버플로 오류가 발생했습니다.
				Call DisplayMsgBox("121515", vbInformation, "", "", I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
		End Select
      Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
   End if

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
    On Error Resume Next
    Err.Clear
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

		Parent.strSp_Id		= "<%=lgSp_Id%>"

		Parent.ggoSpread.Source  = Parent.frm1.vspdData2
		Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data	
		Parent.DbQueryOk
    End If

</Script>