<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c4102bb1
'*  4. Program Name         : 구매품 재고금액 평가 
'*  5. Program Desc         : 원/부자재, 상품에 대한 재고금액 평가계산을 실행한다.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2005/12/14
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     :HJO
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True		
Server.ScriptTimeOut = 10000
						'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 
Call LoadBasisGlobalInf() 

	
On Error Resume Next

Call HideStatusWnd 

Dim IntRetCD
Dim strMsg_cd   
Dim strYYYYMM

strYYYYMM=request("txtYYYYMM")


   Call SubCreateCommandObject(lgObjComm)	
    
	With lgObjComm

		.CommandTimeOut = 0
	    .CommandText = "usp_c_sales_cost_s"
	    .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@yyyymm"     ,adVarXChar,adParamInput,6,strYYYYMM)
	    .Parameters.Append .CreateParameter("@usr_id"     ,adVarXChar,adParamInput,13, gUsrID)
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
%>

<Script Language=vbscript>
'Dim strData
	Call Parent.ExeReflectOk
</Script>	
