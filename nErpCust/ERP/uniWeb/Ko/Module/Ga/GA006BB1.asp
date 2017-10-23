<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3010bb1
'*  4. Program Name         : 구매품 재고금액 평가 
'*  5. Program Desc         : 원/부자재, 상품에 대한 재고금액 평가계산을 실행한다.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/01/09
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Ig Sung, Cho
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 
Call LoadBasisGlobalInf() 
	
On Error Resume Next
Err.Clear

Call HideStatusWnd 

Dim IntRetCD
Dim strMsg_cd   

   Call SubCreateCommandObject(lgObjComm)	
    
	With lgObjComm
		.CommandTimeOut =0
	    .CommandText = "USP_C_GOODS_DATA"
	    .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
		.Parameters.Append .CreateParameter("@yyyymm"     ,adVarXChar,adParamInput,LEN(Trim(Request("txtYyyymm"))), Trim(Request("txtYyyymm")))
	    .Parameters.Append .CreateParameter("@msg_cd"     ,adVarXChar,adParamOutput,6)

	    .Execute ,, adExecuteNoRecords

	End With

	If  Err.number = 0 Then
	    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

	    if  IntRetCD <> 1 then
	        strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
	        Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
	    end if
	Else
	    lgErrorStatus     = "YES"                                                         '☜: Set error status
	    Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if

	Call SubCloseCommandObject(lgObjComm)
	
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pConn,pRs,pErr)
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

<Script Language=vbscript>
'Dim strData
	Call Parent.ExeReflectOk
</Script>	
